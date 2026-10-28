using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;
using NLog;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.CodeExplorer.GitSync;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Api
{
    /// <summary>
    /// <see cref="IModuleStore"/> on top of the VBE. Every call runs on the UI thread, the only one the VBE can be talked to from.
    /// Code edits go through the code module, the way typing does, so attributes and a form's layout are left alone.
    /// Whole-module writes go through the VBE's own export and import, the only way the hidden header survives; the files for that live in a temp folder for the length of the call.
    /// Before a module is changed for the first time, or ever replaced or deleted, the version it had is exported to Rubberduck's backup folder.
    /// </summary>
    public class VbeModuleStore : IModuleStore
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly TimeSpan UiTimeout = TimeSpan.FromSeconds(60);
        private static readonly TimeSpan ParserWait = TimeSpan.FromSeconds(45);
        //How long the UI thread may wait for the parser to pause; only reached if a parse started after the wait above, and then it fails instead of hanging the editor.
        private const int PauseTimeoutMs = 10000;
        private const string Crlf = "\r\n";

        private readonly IProjectsProvider _projects;
        private readonly IVBE _vbe;
        private readonly IUiDispatcher _dispatcher;
        private readonly IParseManager _parseManager;
        private readonly IFileSystem _fileSystem;
        private readonly string _backupFolder;
        private readonly HashSet<string> _backedUp = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public VbeModuleStore(IProjectsProvider projects, IVBE vbe, IUiDispatcher dispatcher, IParseManager parseManager, IFileSystem fileSystem)
        {
            _projects = projects;
            _vbe = vbe;
            _dispatcher = dispatcher;
            _parseManager = parseManager;
            _fileSystem = fileSystem;
            _backupFolder = fileSystem.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Rubberduck", "ApiBackups", DateTime.Now.ToString("yyyyMMdd-HHmmss"));
        }

        #region Reading

        public IReadOnlyList<ProjectInfo> Projects()
            => OnUi(() => _projects.Projects().Select(p =>
            {
                var path = PathOf(p.Project);
                return new ProjectInfo
                {
                    Id = ProjectIds.Of(p.ProjectId, p.Project.Name, path),
                    Name = p.Project.Name,
                    Path = path,
                    Extension = string.IsNullOrEmpty(path) ? null : _fileSystem.Path.GetExtension(path)
                };
            }).ToList());

        //A document that was never saved has no file, and asking for it can throw.
        private static string PathOf(IVBProject project)
        {
            try
            {
                var name = project.FileName;
                return string.IsNullOrWhiteSpace(name) ? null : name;
            }
            catch (COMException)
            {
                return null;
            }
        }

        public IReadOnlyList<ModuleInfo> Modules(string project)
            => OnUi(() =>
            {
                var found = FindProject(project);
                var modules = new List<ModuleInfo>();
                using (var components = found.VBComponents)
                {
                    foreach (var component in components)
                    {
                        using (component)
                        {
                            modules.Add(new ModuleInfo
                            {
                                Project = found.Name,
                                Name = component.Name,
                                Kind = KindOf(component.Type),
                                LineCount = LineCount(component)
                            });
                        }
                    }
                }

                return (IReadOnlyList<ModuleInfo>)modules;
            });

        public ModuleContent Read(string project, string module, bool includeText)
            => OnUi(() =>
            {
                var found = FindProject(project);
                using (var components = found.VBComponents)
                using (var component = Find(components, module))
                {
                    var code = CodeOf(component);
                    var content = new ModuleContent
                    {
                        Project = found.Name,
                        ProjectPath = PathOf(found),
                        Name = component.Name,
                        Kind = KindOf(component.Type),
                        LineCount = LineCount(component),
                        Hash = ModuleApi.HashOf(code),
                        Code = code
                    };

                    if (includeText)
                    {
                        var temp = NewTempFolder();
                        try
                        {
                            var path = component.ExportAsSourceFile(_fileSystem.Path.Combine(temp, component.Name + component.Type.FileExtension()));
                            content.Text = _fileSystem.File.ReadAllText(path, GitSyncFolderProcessor.EncodingFor(path));
                        }
                        finally
                        {
                            DeleteFolder(temp);
                        }
                    }

                    return content;
                }
            });

        #endregion

        #region Writing

        public ModuleInfo ReplaceCode(string project, string module, string code, string expectedHash)
            => OnUiWrite(() =>
            {
                var found = FindProject(project);
                using (var components = found.VBComponents)
                using (var component = Find(components, module))
                {
                    CheckHash(component, expectedHash);
                    BackUpOnce(found, component);
                    using (var codeModule = component.CodeModule)
                    {
                        ReplaceAll(codeModule, code);
                    }

                    return Describe(found, component);
                }
            });

        public ModuleInfo EditLines(string project, string module, int startLine, int endLine, string replacement, string expectedHash)
            => OnUiWrite(() =>
            {
                var found = FindProject(project);
                using (var components = found.VBComponents)
                using (var component = Find(components, module))
                {
                    CheckHash(component, expectedHash);
                    using (var codeModule = component.CodeModule)
                    {
                        var count = codeModule.CountOfLines;
                        if (startLine > count + 1 || endLine > count)
                        {
                            throw ApiException.BadRequest($"The module has {count} line(s); lines {startLine}-{endLine} are out of range.");
                        }

                        BackUpOnce(found, component);
                        if (endLine >= startLine)
                        {
                            codeModule.DeleteLines(startLine, endLine - startLine + 1);
                        }

                        if (replacement.Length > 0)
                        {
                            codeModule.InsertLines(startLine, Normalize(replacement));
                        }
                    }

                    return Describe(found, component);
                }
            });

        public ModuleInfo ReplaceText(string project, string module, string text, string expectedHash)
            => OnUiWriteWithParserPause(() =>
            {
                var found = FindProject(project);
                using (var components = found.VBComponents)
                using (var existing = Find(components, module))
                {
                    CheckHash(existing, expectedHash);
                    var extension = existing.Type.FileExtension();
                    if (extension.Length == 0)
                    {
                        throw ApiException.BadRequest($"Modules of type {existing.Type} cannot be replaced as a whole; write their code instead.");
                    }

                    var declared = SyncText.ModuleAttributes(text);
                    if (declared.TryGetValue("VB_Name", out var name) && !name.Trim('"').Equals(existing.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new ApiException(422, "name_mismatch", $"The text declares the module '{name.Trim('"')}', not '{existing.Name}'. Renaming is not supported.");
                    }

                    var temp = NewTempFolder();
                    try
                    {
                        //Exporting first puts a form's binary next to the file, so the new text is imported with the layout the form already has.
                        var exported = existing.ExportAsSourceFile(_fileSystem.Path.Combine(temp, "old", existing.Name + extension));
                        var prepared = _fileSystem.Path.Combine(temp, "new", existing.Name + extension);
                        _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(prepared));
                        var binary = existing.Type.BinaryFileExtension();
                        if (binary.Length > 0 && _fileSystem.File.Exists(_fileSystem.Path.ChangeExtension(exported, binary)))
                        {
                            _fileSystem.File.Copy(_fileSystem.Path.ChangeExtension(exported, binary), _fileSystem.Path.ChangeExtension(prepared, binary));
                        }

                        _fileSystem.File.WriteAllText(prepared, Normalize(text), GitSyncFolderProcessor.EncodingFor(exported));
                        BackUp(found, existing, true);

                        Suspended(() =>
                        {
                            if (existing.Type == ComponentType.Document)
                            {
                                //Document modules cannot be removed; only their code is replaced.
                                using (components.ImportSourceFile(prepared)) { }
                                return;
                            }

                            components.Remove(existing);
                            try
                            {
                                using (components.Import(prepared)) { }
                            }
                            catch (Exception)
                            {
                                //Removed and not imported would leave the project without it.
                                using (components.Import(exported)) { }
                                throw;
                            }
                        });
                    }
                    finally
                    {
                        DeleteFolder(temp);
                    }

                    using (var replaced = Find(components, module))
                    {
                        GitSyncService.RemoveLeadingBlankLines(replaced);

                        //Attributes are not part of what the parser compares to see whether a module changed; an attribute-only write would otherwise go unnoticed until the code changes.
                        _parseManager.MarkAsModified(new QualifiedModuleName(replaced));
                        return Describe(found, replaced);
                    }
                }
            });

        public ModuleInfo Create(string project, string module, string kind, string code)
            => OnUiWrite(() =>
            {
                var found = FindProject(project);
                using (var components = found.VBComponents)
                {
                    using (var present = FindOrNull(components, module))
                    {
                        if (present != null)
                        {
                            throw ApiException.Conflict($"The project already has a module named '{module}'.");
                        }
                    }

                    var type = kind == ModuleKinds.Class ? ComponentType.ClassModule
                        : kind == ModuleKinds.Form ? ComponentType.UserForm
                        : ComponentType.StandardModule;

                    using (var created = components.Add(type))
                    {
                        created.Name = module;
                        if (code.Length > 0)
                        {
                            using (var codeModule = created.CodeModule)
                            {
                                ReplaceAll(codeModule, code);
                            }
                        }

                        return Describe(found, created);
                    }
                }
            });

        public void Delete(string project, string module, string expectedHash)
            => OnUiWriteWithParserPause(() =>
            {
                var found = FindProject(project);
                using (var components = found.VBComponents)
                using (var component = Find(components, module))
                {
                    CheckHash(component, expectedHash);
                    BackUp(found, component, true);
                    Suspended(() => components.RemoveSafely(component));
                    return true;
                }
            });

        #endregion

        #region Helpers

        //The VBE can only be talked to from its UI thread; the HTTP listener's threads hand the work over and wait.
        private T OnUi<T>(Func<T> work)
        {
            var task = _dispatcher.StartTask(work, CancellationToken.None);
            try
            {
                if (!task.Wait(UiTimeout))
                {
                    throw ApiException.Unavailable("The editor did not respond in time. It may be showing a dialog or running code.");
                }

                return task.Result;
            }
            catch (AggregateException exception)
            {
                var inner = exception.InnerException ?? exception;
                if (inner is COMException)
                {
                    throw ApiException.Unavailable($"The editor refused the request ({inner.Message}). It may be in break mode or running code.");
                }

                ExceptionDispatchInfo.Capture(inner).Throw();
                throw;
            }
        }

        //While a macro is running or stopped, the editor answers an edit with a "this action will reset your project" prompt that nobody can click for the caller
        //(and Excel soon adds its own "waiting for an OLE action" box on top), so a write is refused before it gets that far.
        private T OnUiWrite<T>(Func<T> work)
            => OnUi(() =>
            {
                if (!_vbe.IsInDesignMode)
                {
                    throw ApiException.NotInDesignMode();
                }

                return work();
            });

        //A write that pauses the parser must not do so from the UI thread while a parse is running: the pause waits for the parse to let go of its lock,
        //the parse needs the UI thread to read components, and neither ever finishes (Excel then shows "waiting for an OLE action"). So the wait for the parser
        //is done here, on the caller's thread, with the UI thread still free.
        private T OnUiWriteWithParserPause<T>(Func<T> work)
        {
            var timer = System.Diagnostics.Stopwatch.StartNew();
            while (!IsSettled(_parseManager.Status))
            {
                if (timer.Elapsed > ParserWait)
                {
                    throw ApiException.Unavailable($"Rubberduck is still parsing ({_parseManager.Status}), so the edit cannot be made yet. Try again in a moment.");
                }

                Thread.Sleep(100);
            }

            return OnUiWrite(work);
        }

        private static bool IsSettled(ParserState state)
            => state == ParserState.Ready || state == ParserState.Error || state == ParserState.ResolverError || state == ParserState.UnexpectedError;

        private void OnUi(Action work) => OnUi(() => { work(); return true; });

        private IVBProject FindProject(string key)
        {
            var matches = ProjectIds.Match(key, _projects.Projects().ToList(), p => (p.ProjectId, p.Project.Name, PathOf(p.Project)));
            switch (matches.Count)
            {
                case 0:
                    throw ApiException.NotFound($"No open project '{key}'.");
                case 1:
                    return matches[0].Project;
                default:
                    throw ApiException.Conflict($"More than one open project is named '{key}'; use its id (see /projects).");
            }
        }

        private static IVBComponent Find(IVBComponents components, string name)
            => FindOrNull(components, name) ?? throw ApiException.NotFound($"The project has no module '{name}'.");

        private static IVBComponent FindOrNull(IVBComponents components, string name) => GitSyncService.FindComponent(components, name);

        private static string KindOf(ComponentType type)
        {
            switch (type)
            {
                case ComponentType.StandardModule: return ModuleKinds.Standard;
                case ComponentType.ClassModule: return ModuleKinds.Class;
                case ComponentType.UserForm: return ModuleKinds.Form;
                case ComponentType.Document: return ModuleKinds.Document;
                default: return type.ToString().ToLowerInvariant();
            }
        }

        private static string CodeOf(IVBComponent component)
        {
            if (!component.HasCodeModule)
            {
                return string.Empty;
            }

            using (var codeModule = component.CodeModule)
            {
                return codeModule.Content() ?? string.Empty;
            }
        }

        private static int LineCount(IVBComponent component)
        {
            if (!component.HasCodeModule)
            {
                return 0;
            }

            using (var codeModule = component.CodeModule)
            {
                return codeModule.CountOfLines;
            }
        }

        private static ModuleInfo Describe(IVBProject project, IVBComponent component)
        {
            var code = CodeOf(component);
            return new ModuleInfo
            {
                Project = project.Name,
                ProjectPath = PathOf(project),
                Name = component.Name,
                Kind = KindOf(component.Type),
                LineCount = LineCount(component),
                Hash = ModuleApi.HashOf(code)
            };
        }

        private static void CheckHash(IVBComponent component, string expected)
        {
            if (!string.Equals(ModuleApi.HashOf(CodeOf(component)), expected, StringComparison.OrdinalIgnoreCase))
            {
                throw ApiException.PreconditionFailed($"'{component.Name}' changed since you read it. Read it again and redo the edit on the current code.");
            }
        }

        private static void ReplaceAll(ICodeModule module, string code)
        {
            if (module.CountOfLines > 0)
            {
                module.DeleteLines(1, module.CountOfLines);
            }

            if (code.Length > 0)
            {
                module.InsertLines(1, Normalize(code));
            }
        }

        private static string Normalize(string text) => text.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", Crlf);

        private void BackUpOnce(IVBProject project, IVBComponent component)
        {
            if (_backedUp.Add(project.Name + "/" + component.Name))
            {
                BackUp(project, component, false);
            }
        }

        //A change that cannot be undone from the editor is not made unless the old version is safe.
        private void BackUp(IVBProject project, IVBComponent component, bool always)
        {
            var extension = component.Type.FileExtension();
            if (extension.Length == 0)
            {
                return;
            }

            try
            {
                var folder = _fileSystem.Path.Combine(_backupFolder, Sanitize(project.Name));
                _fileSystem.Directory.CreateDirectory(folder);
                var name = always ? component.Name + "." + DateTime.Now.ToString("HHmmssfff") : component.Name;
                var path = component.ExportAsSourceFile(_fileSystem.Path.Combine(folder, name + extension));
                if (always && !_fileSystem.File.Exists(path))
                {
                    throw new System.IO.IOException("the export produced no file");
                }
            }
            catch (Exception exception) when (exception is System.IO.IOException || exception is UnauthorizedAccessException || exception is COMException)
            {
                throw ApiException.Unavailable($"The current version of '{component.Name}' could not be backed up, so it was not changed: {exception.Message}");
            }
        }

        private static string Sanitize(string name)
            => string.Concat(name.Select(c => System.IO.Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));

        private void Suspended(Action action)
        {
            Exception failure = null;
            Logger.Info("API asks the parser to pause for an edit.");
            //Any state where no parse is running will do. A project in Error must stay editable: fixing it is the reason to edit.
            var suspension = _parseManager.OnSuspendParser(this, new[] { ParserState.Ready, ParserState.Error, ParserState.ResolverError }, () =>
            {
                try
                {
                    action();
                }
                catch (Exception exception)
                {
                    failure = exception;
                }
            }, PauseTimeoutMs);

            Logger.Info($"API edit finished; the parser's pause ended with {suspension.Outcome}{(failure != null ? " (the edit failed: " + failure.Message + ")" : string.Empty)}.");
            if (failure != null)
            {
                ExceptionDispatchInfo.Capture(failure).Throw();
            }

            if (suspension.Outcome != SuspensionOutcome.Completed)
            {
                throw ApiException.Unavailable(suspension.Outcome == SuspensionOutcome.IncompatibleState
                    ? "Rubberduck is parsing right now, so the edit cannot be made. Try again in a moment."
                    : $"The parser could not be paused for the edit ({suspension.Outcome}). Try again in a moment.");
            }
        }

        private string NewTempFolder()
        {
            var folder = _fileSystem.Path.Combine(_fileSystem.Path.GetTempPath(), "RubberduckApi", Guid.NewGuid().ToString("N"));
            _fileSystem.Directory.CreateDirectory(folder);
            return folder;
        }

        private void DeleteFolder(string folder)
        {
            try
            {
                if (_fileSystem.Directory.Exists(folder))
                {
                    _fileSystem.Directory.Delete(folder, true);
                }
            }
            catch (System.IO.IOException)
            {
                //The temp folder is not worth failing over.
            }
        }

        #endregion
    }
}
