using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Text;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.PackageResources;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>A note of the project (see <see cref="NoteModule"/>), as the VBE exports it.</summary>
    public class NoteInput
    {
        public string ModuleName { get; set; }

        public string Code { get; set; }
    }

    /// <summary>
    /// What the sync does with notes: a note is a markdown file in the folder, not a module file. <c>CLAUDE.md</c> in the folder <c>Api.Core</c> is <c>Api\Core\CLAUDE.md</c>,
    /// and shows on GitHub as the markdown it is, with none of the comment marks the module needs to compile.
    /// Notes are listed with the modules, under the path of the file with forward slashes, which no module can be named.
    /// A markdown file the project has no note for is listed but never suggested for import: a repository holds other markdown, such as a README.
    /// </summary>
    public class GitSyncNotes
    {
        private readonly IFileSystem _fileSystem;
        private readonly GitSyncFolderProcessor _processor;

        public GitSyncNotes(IFileSystem fileSystem, GitSyncFolderProcessor processor)
        {
            _fileSystem = fileSystem;
            _processor = processor;
        }

        private class Side
        {
            public string ModuleName;
            public string Name;
            public string Relative;
            public string Markdown;
        }

        /// <summary>Compares the notes of the project with the markdown files of the folder. Adds what is out of sync to the plan, and what is in sync to its baselines.</summary>
        public void Plan(IEnumerable<NoteInput> notes, string root, GitSyncState state, GitIgnoreSet ignored, SyncPlan plan)
        {
            var inProject = ReadProject(notes);
            var inFolder = ReadFolder(root);

            foreach (var key in inProject.Keys.Union(inFolder.Keys, StringComparer.OrdinalIgnoreCase).OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
            {
                inProject.TryGetValue(key, out var module);
                inFolder.TryGetValue(key, out var file);
                var relative = (file ?? module).Relative;

                if (ignored.IsIgnored(relative, false))
                {
                    plan.Ignored.Add(key);
                    continue;
                }

                var item = new SyncItem
                {
                    Name = key,
                    NoteModule = module?.ModuleName,
                    RelativePath = relative,
                    ModuleExists = module != null,
                    FileExists = file != null,
                    ModuleSignature = module == null ? null : Signature(module.Markdown),
                    FileSignature = file == null ? null : Signature(file.Markdown)
                };

                state.Modules.TryGetValue(key, out var baseline);
                (item.Status, item.Suggested) = GitSyncPlanner.Classify(item, baseline);

                if (item.Status == SyncStatus.InSync)
                {
                    plan.InSyncCount++;
                    plan.InSyncBaselines[key] = BaselineOf(item);
                    continue;
                }

                if (!item.ModuleExists)
                {
                    item.Suggested = SyncAction.None;
                }

                if (module != null && file != null)
                {
                    item.ModuleText = module.Markdown;
                    item.FileText = file.Markdown;
                }

                plan.Items.Add(item);
            }
        }

        /// <summary>
        /// Carries out the chosen action of each note item, in the VBE. Stops nothing on an error: it is reported and the rest go on.
        /// Run while the parser is suspended, like the module items.
        /// </summary>
        public void Apply(IVBComponents components, GitSyncScan scan, IEnumerable<SyncItem> items, Action<string> backUp, GitSyncResult result)
        {
            foreach (var item in items)
            {
                try
                {
                    var path = System.IO.Path.GetFullPath(System.IO.Path.Combine(scan.Folder, item.RelativePath.Replace('/', '\\')));
                    string markdown;
                    switch (item.Chosen)
                    {
                        case SyncAction.Export:
                            markdown = ReadNote(components, item.NoteModule);
                            Write(path, markdown, backUp);
                            result.Exported++;
                            break;

                        case SyncAction.Import:
                            markdown = ReadFile(path);
                            Load(components, item, markdown);
                            result.Imported++;
                            break;

                        case SyncAction.Merge:
                            markdown = item.MergedText ?? string.Empty;
                            Load(components, item, markdown);
                            Write(path, markdown, backUp);
                            result.Merged++;
                            break;

                        default:
                            continue;
                    }

                    //What each side holds now, as the next scan will see it.
                    scan.State.Modules[item.Name] = BaselineOf(new SyncItem
                    {
                        RelativePath = item.RelativePath,
                        ModuleSignature = Signature(markdown),
                        FileSignature = Signature(ReadFile(path))
                    });
                }
                catch (Exception exception) when (exception is InvalidOperationException || exception is System.IO.IOException || exception is UnauthorizedAccessException)
                {
                    result.Errors.Add($"Could not sync '{item.Name}': {exception.Message}");
                }
            }
        }

        //Writes the markdown into the note's module, creating the module if the project has none: the folder comes from where the file is, the name from the file's name.
        private static void Load(IVBComponents components, SyncItem item, string markdown)
        {
            var displayName = item.RelativePath.Replace('/', '\\').Split('\\').Last();
            if (!NoteModule.IsValidName(displayName))
            {
                throw new InvalidOperationException($"'{displayName}' cannot be the name of a note.");
            }

            var tooLong = NoteModule.FirstLineTooLong(markdown);
            if (tooLong > 0)
            {
                throw new InvalidOperationException($"line {tooLong} is longer than the {NoteModule.MaxLineLength - 2} characters a note keeps.");
            }

            var directory = System.IO.Path.GetDirectoryName(item.RelativePath.Replace('/', '\\')) ?? string.Empty;
            var folder = GitSyncText.FolderPath(directory);

            IVBComponent component = null;
            try
            {
                string existing = null;
                if (item.NoteModule != null)
                {
                    component = GitSyncService.FindComponent(components, item.NoteModule);
                    existing = component == null ? null : Content(component);
                    if (existing != null && !NoteModule.IsNote(existing))
                    {
                        throw new InvalidOperationException($"module '{item.NoteModule}' is no longer a note.");
                    }
                }

                if (component == null)
                {
                    var moduleName = NoteModule.ModuleNameFor(displayName);
                    if (moduleName == null)
                    {
                        throw new InvalidOperationException($"'{displayName}' does not make a module name the VBE accepts.");
                    }

                    var taken = GitSyncService.FindComponent(components, moduleName);
                    if (taken != null)
                    {
                        taken.Dispose();
                        throw new InvalidOperationException($"the project already has a module named '{moduleName}'.");
                    }

                    component = components.Add(ComponentType.StandardModule);
                    component.Name = moduleName;
                }

                var annotations = existing == null ? new List<string>() : NoteModule.Annotations(existing).ToList();
                var code = NoteModule.ToModuleCode(markdown, NoteModule.WithFolder(NoteModule.WithMarker(annotations, displayName), folder));
                using (var codeModule = component.CodeModule)
                {
                    codeModule.Clear();
                    codeModule.InsertLines(1, code);
                }
            }
            finally
            {
                component?.Dispose();
            }
        }

        private static string ReadNote(IVBComponents components, string moduleName)
        {
            using (var component = GitSyncService.FindComponent(components, moduleName ?? string.Empty))
            {
                if (component == null)
                {
                    throw new InvalidOperationException("the project has no such note");
                }

                var content = Content(component);
                if (!NoteModule.IsNote(content))
                {
                    throw new InvalidOperationException($"module '{moduleName}' is no longer a note.");
                }

                return NoteModule.ToMarkdown(content);
            }
        }

        private static string Content(IVBComponent component)
        {
            using (var codeModule = component.CodeModule)
            {
                return codeModule.Content();
            }
        }

        private void Write(string path, string markdown, Action<string> backUp)
        {
            backUp(path);
            _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(path));
            //UTF-8 without a byte order mark: what GitHub and every tool expects of markdown.
            _fileSystem.File.WriteAllText(path, markdown.Replace("\r\n", "\n") + "\n", new UTF8Encoding(false));
        }

        private string ReadFile(string path)
        {
            return _fileSystem.File.ReadAllText(path, new UTF8Encoding(false, false)).TrimStart('﻿').Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd('\n').Replace("\n", "\r\n");
        }

        private static ModuleBaseline BaselineOf(SyncItem item) => new ModuleBaseline
        {
            Module = item.ModuleSignature,
            File = item.FileSignature,
            Path = item.RelativePath.Replace('\\', '/'),
            Folder = string.Empty
        };

        #region Reading both sides

        private static Dictionary<string, Side> ReadProject(IEnumerable<NoteInput> notes)
        {
            var sides = new Dictionary<string, Side>(StringComparer.OrdinalIgnoreCase);
            foreach (var note in notes)
            {
                var name = NoteModule.DisplayName(note.ModuleName, note.Code);
                var directory = GitSyncText.RelativeDirectory(NoteModule.FolderOf(NoteModule.Annotations(note.Code)));
                var relative = directory.Length == 0 ? name : directory + "\\" + name;
                var key = relative.Replace('\\', '/');
                if (!sides.ContainsKey(key))
                {
                    sides[key] = new Side { ModuleName = note.ModuleName, Name = name, Relative = relative, Markdown = NoteModule.ToMarkdown(note.Code) };
                }
            }

            return sides;
        }

        private Dictionary<string, Side> ReadFolder(string root)
        {
            var sides = new Dictionary<string, Side>(StringComparer.OrdinalIgnoreCase);
            foreach (var path in _processor.AllNoteFiles(root))
            {
                var relative = _processor.RelativePath(root, path);
                sides[relative.Replace('\\', '/')] = new Side { Name = _fileSystem.Path.GetFileName(path), Relative = relative, Markdown = ReadFile(path) };
            }

            return sides;
        }

        #endregion

        //Ignores the byte order mark, the kind of line ending and trailing blank lines, which tools add and remove freely.
        private static string Signature(string markdown)
            => PackageResource.HashOf(Encoding.UTF8.GetBytes(markdown.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd('\n')));
    }
}
