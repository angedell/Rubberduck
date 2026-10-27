using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Abstractions;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    public class GitSyncResult
    {
        public int Imported { get; set; }
        public int Exported { get; set; }
        public int Merged { get; set; }
        public List<string> Errors { get; } = new List<string>();
        public List<string> Notes { get; } = new List<string>();

        /// <summary>Where the versions that were overwritten were saved, or null if nothing was.</summary>
        public string BackupFolder { get; set; }
    }

    /// <summary>What a scan found, together with everything needed to apply it.</summary>
    public class GitSyncScan
    {
        public string Folder { get; set; }
        public SyncPlan Plan { get; set; }
        public ProjectConfig Config { get; set; }
        public GitSyncState State { get; set; }
        public string StatePath { get; set; }
        public IIndenterSettings Settings { get; set; }
        public GitIgnoreSet Ignored { get; set; }
        public SignatureCache Cache { get; set; }

        /// <summary>The folder had no rubberduck.json, so the scan wrote one.</summary>
        public bool ConfigCreated { get; set; }
    }

    /// <summary>
    /// Syncs a VBA project with a folder, once, when the user asks. <see cref="Scan"/> looks at both sides and changes nothing;
    /// <see cref="Apply"/> carries out the items the user picked. Modules always travel through the VBE's own export and import,
    /// because that is the only way their hidden attributes survive.
    /// Nothing is ever deleted from the project; files are only deleted when a module moves to another folder.
    /// </summary>
    /// <remarks>
    /// The VBE can only be talked to from the thread that called in, so everything that touches it runs there, one module after the other.
    /// The rest, which is where most of the time goes (reading files, indenting code, hashing), runs on several threads.
    /// </remarks>
    public class GitSyncService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IParseManager _parseManager;
        private readonly IFileSystem _fileSystem;
        private readonly IIndenter _indenter;
        private readonly IConfigurationService<Configuration> _configuration;

        private readonly GitSyncFolderProcessor _processor;
        private readonly ProjectConfigReader _configReader;
        private readonly ProjectConfigWriter _configWriter;
        private readonly GitSyncStateStore _stateStore;
        private readonly GitSyncLinks _links;

        public GitSyncService(
            IParseManager parseManager,
            IFileSystem fileSystem,
            IIndenter indenter,
            IConfigurationService<Configuration> configuration)
        {
            _parseManager = parseManager;
            _fileSystem = fileSystem;
            _indenter = indenter;
            _configuration = configuration;

            _processor = new GitSyncFolderProcessor(fileSystem);
            _configReader = new ProjectConfigReader(fileSystem);
            _configWriter = new ProjectConfigWriter(fileSystem);
            _stateStore = new GitSyncStateStore(fileSystem);
            _links = new GitSyncLinks(fileSystem);
        }

        /// <summary>The folder last used for the project, or null.</summary>
        public string RememberedFolder(IVBProject project) => _links.Folder(project.FileName);

        #region Scan

        /// <summary>Looks at the project and the folder and reports what is out of sync. Changes nothing.</summary>
        /// <exception cref="OperationCanceledException">The user cancelled.</exception>
        public GitSyncScan Scan(IVBProject project, string folder, ISyncProgress progress = null)
        {
            progress = progress ?? NoProgress.Instance;
            var timer = Stopwatch.StartNew();

            var root = _fileSystem.Path.GetFullPath(folder);
            var userSettings = _configuration.Read().UserSettings.IndenterSettings;
            var config = _configReader.Read(root, userSettings);
            var configCreated = false;
            if (!config.Found)
            {
                //A folder without a rubberduck.json gets one, with the user's indenter settings and the default sync settings, so there is something to commit and edit.
                configCreated = _configWriter.CreateDefault(root, userSettings).Created;
                config = _configReader.Read(root, userSettings);
            }

            var settings = config.Indenter ?? userSettings;
            var statePath = _stateStore.PathFor(config.Path, root);
            var state = _stateStore.Load(statePath);
            var ignored = _processor.LoadIgnoreRules(root, config.Sync.Ignore);
            var cache = SignatureCache.Load(_fileSystem, settings, null, config.Sync.IgnoreCaseOutsideStrings);
            var setup = timer.ElapsedMilliseconds;

            //The VBE part: one module after the other, on this thread.
            var temp = NewTempFolder();
            List<ModuleInput> modules;
            try
            {
                modules = ReadModules(project, temp, progress);
            }
            finally
            {
                DeleteFolder(temp);
            }

            var modulesDone = timer.ElapsedMilliseconds;

            //The rest needs no VBE and runs on several threads while this one keeps the progress window alive.
            var paths = _processor.AllSourceFiles(root).ToList();
            var files = Background(
                () => paths.AsParallel().WithCancellation(progress.Token).Select(path => ReadFile(root, path)).Where(file => file != null).ToList(),
                progress, SyncStage.ReadingFiles, () => 0, paths.Count);
            var filesDone = timer.ElapsedMilliseconds;

            var planner = new GitSyncPlanner(settings, _indenter, config.Sync, cache);
            var plan = Background(
                () => planner.Plan(modules, files, state, ignored, progress.Token),
                progress, SyncStage.Comparing, () => planner.Completed, modules.Count + files.Count);

            cache.Save(_fileSystem);

            Logger.Info($"Git sync scan of '{root}': {modules.Count} modules, {files.Count} files; setup {setup}ms, modules {modulesDone - setup}ms, "
                        + $"files {filesDone - modulesDone}ms, compare {timer.ElapsedMilliseconds - filesDone}ms, total {timer.ElapsedMilliseconds}ms; "
                        + $"signature cache {cache.Hits} hits / {cache.Misses} misses; {plan.Items.Count} out of sync.");

            return new GitSyncScan { Folder = root, Plan = plan, Config = config, State = state, StatePath = statePath, Settings = settings, Ignored = ignored, Cache = cache, ConfigCreated = configCreated };
        }

        private List<ModuleInput> ReadModules(IVBProject project, string temp, ISyncProgress progress)
        {
            var modules = new List<ModuleInput>();
            using (var components = project.VBComponents)
            {
                var total = components.Count;
                var done = 0;
                foreach (var component in components)
                {
                    using (component)
                    {
                        progress.Token.ThrowIfCancellationRequested();
                        progress.Report(SyncStage.ReadingModules, done++, total);

                        var extension = component.Type.FileExtension();
                        if (extension.Length == 0)
                        {
                            continue;
                        }

                        var path = component.ExportAsSourceFile(_fileSystem.Path.Combine(temp, component.Name + extension));
                        try
                        {
                            modules.Add(new ModuleInput
                            {
                                Name = component.Name,
                                Extension = extension,
                                ExportedText = _fileSystem.File.ReadAllText(path, GitSyncFolderProcessor.EncodingFor(path))
                            });
                        }
                        finally
                        {
                            DeleteWithBinary(path);
                        }
                    }
                }
            }

            return modules;
        }

        //Runs on a worker thread. The module's name is in the file; the file's own name is the fallback, as when the VBE imports it.
        private FileInput ReadFile(string root, string path)
        {
            var text = _fileSystem.File.ReadAllText(path, GitSyncFolderProcessor.EncodingFor(path));
            var name = SyncText.ModuleAttributes(text).TryGetValue("VB_Name", out var declared)
                ? declared.Trim('"')
                : _fileSystem.Path.GetFileNameWithoutExtension(path);

            return string.IsNullOrWhiteSpace(name)
                ? null
                : new FileInput { Name = name, RelativePath = _processor.RelativePath(root, path), Text = text };
        }

        /// <summary>
        /// Runs work on another thread and waits for it here, handling the progress window's messages meanwhile.
        /// </summary>
        private static T Background<T>(Func<T> work, ISyncProgress progress, SyncStage stage, Func<int> done, int total)
        {
            var task = Task.Run(work);
            while (!task.IsCompleted)
            {
                progress.Report(stage, done(), total);
                Thread.Sleep(30);
            }

            progress.Report(stage, total, total);
            try
            {
                return task.GetAwaiter().GetResult();
            }
            catch (AggregateException exception) when (exception.InnerExceptions.OfType<OperationCanceledException>().Any())
            {
                throw new OperationCanceledException(progress.Token);
            }
        }

        #endregion

        /// <summary>Adds the items to the ignore list of the config, which creates the config if the folder has none.</summary>
        public void Ignore(GitSyncScan scan, IEnumerable<SyncItem> items)
        {
            //Anchored, so that a file with the same name somewhere else is not caught.
            var patterns = items.Select(item => "/" + item.RelativePath.Replace('\\', '/')).ToList();
            _configWriter.AddIgnored(scan.Config.Path, scan.Folder, patterns);
        }

        #region Apply

        //What was done to an item, with the text to work out its new history from once the VBE is free again.
        private class Applied
        {
            public SyncItem Item;
            public string Relative;
            public string Folder;
            public string ModuleText;
            //The file's text as the VBE sees it.
            public string FileText;
        }

        //An import's file, made ready while the VBE is not needed yet.
        private class PreparedImport
        {
            public string Relative;
            public string Prepared;
            public string Copy;
            //Set when the rules were repaired: the repository's file gets the repaired text too, so both sides stay alike.
            public string RepairedFile;
            public string FilePath;
        }

        /// <summary>Carries out the chosen action of each item. The state is saved even if some items fail or the user cancels.</summary>
        public GitSyncResult Apply(IVBProject project, GitSyncScan scan, IEnumerable<SyncItem> items, ISyncProgress progress = null)
        {
            progress = progress ?? NoProgress.Instance;
            var timer = Stopwatch.StartNew();
            var result = new GitSyncResult();
            var chosen = items.Where(item => item.Chosen != SyncAction.None).ToList();

            foreach (var blocked in chosen.Where(item => item.IsBlocked).ToList())
            {
                result.Notes.Add($"'{blocked.Name}' was skipped: it breaks a rule that is enforced.");
                chosen.Remove(blocked);
            }

            var temp = NewTempFolder();
            var backup = _fileSystem.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Rubberduck", "GitSyncBackups", DateTime.Now.ToString("yyyyMMdd-HHmmss"));
            var applied = new List<Applied>();
            long prepared = 0, vbe = 0;
            try
            {
                //The files to import are read and made ready first, on several threads, so the parser is held still for as short a time as possible.
                var imports = Background(
                    () => chosen.Where(item => item.Chosen != SyncAction.Export).AsParallel().WithCancellation(progress.Token)
                        .Select(item => new KeyValuePair<SyncItem, PreparedImport>(item, PrepareImport(item, scan, temp, backup)))
                        .ToDictionary(pair => pair.Key, pair => pair.Value),
                    progress, SyncStage.Preparing, () => 0, chosen.Count);
                prepared = timer.ElapsedMilliseconds;

                Suspended(() =>
                {
                    using (var components = project.VBComponents)
                    {
                        var done = 0;
                        foreach (var item in chosen)
                        {
                            if (progress.Token.IsCancellationRequested)
                            {
                                result.Notes.Add($"Cancelled after {done} of {chosen.Count} item(s); the rest were not synced.");
                                break;
                            }

                            progress.Report(SyncStage.Syncing, done++, chosen.Count);
                            try
                            {
                                applied.Add(item.Chosen == SyncAction.Export
                                    ? Export(components, item, scan, temp, backup, result)
                                    : Import(components, item, imports[item], temp, backup, result));
                            }
                            catch (Exception exception)
                            {
                                Logger.Error(exception, $"Git sync: could not sync '{item.Name}'.");
                                result.Errors.Add($"Could not sync '{item.Name}': {exception.Message}");
                            }
                        }
                    }
                });
                vbe = timer.ElapsedMilliseconds - prepared;

                //The new history is worked out afterwards, on several threads, from the text already in hand.
                var baselines = applied.AsParallel()
                    .Select(a => new KeyValuePair<string, ModuleBaseline>(a.Item.Name, BaselineAfter(a, scan)))
                    .ToList();
                foreach (var baseline in baselines)
                {
                    scan.State.Modules[baseline.Key] = baseline.Value;
                }

                result.Imported = applied.Count(a => a.Item.Chosen == SyncAction.Import);
                result.Exported = applied.Count(a => a.Item.Chosen == SyncAction.Export);
                result.Merged = applied.Count(a => a.Item.Chosen == SyncAction.Merge);
                if (result.Merged > 0)
                {
                    result.Notes.Add($"{result.Merged} module(s) were merged from both sides; the project and the folder now hold the same text.");
                }
            }
            finally
            {
                DeleteFolder(temp);
            }

            foreach (var baseline in scan.Plan.InSyncBaselines)
            {
                scan.State.Modules[baseline.Key] = baseline.Value;
            }

            _stateStore.Save(scan.StatePath, scan.State);
            scan.Cache?.Save(_fileSystem);
            _links.Remember(project.FileName, scan.Folder);

            if (_fileSystem.Directory.Exists(backup))
            {
                result.BackupFolder = backup;
            }

            Logger.Info($"Git sync apply: {result.Imported} imported, {result.Exported} exported, {result.Errors.Count} errors; "
                        + $"prepare {prepared}ms, VBE {vbe}ms (parser suspended), history {timer.ElapsedMilliseconds - prepared - vbe}ms, total {timer.ElapsedMilliseconds}ms.");
            return result;
        }

        private void Suspended(Action action)
        {
            var suspension = _parseManager.OnSuspendParser(this, new[] { ParserState.Ready }, action);
            if (suspension.Outcome != SuspensionOutcome.Completed)
            {
                throw new InvalidOperationException(
                    $"The sync could not run because the parser could not be suspended ({suspension.Outcome}).",
                    suspension.EncounteredException);
            }
        }

        //Reads the file and writes the copy the VBE is going to import. Runs on a worker thread.
        private PreparedImport PrepareImport(SyncItem item, GitSyncScan scan, string temp, string backup)
        {
            var relative = item.RelativePath.Replace('/', '\\');
            var path = _fileSystem.Path.Combine(scan.Folder, relative);
            var directory = _fileSystem.Path.GetDirectoryName(relative) ?? string.Empty;

            //The VBE only reads Windows line endings, and the @Folder annotation has to match where the file is.
            //The file in the repository stays as it is; the VBE reads a corrected copy.
            //A merge is made from the text the user saw, not from what the file holds now; the file is kept in the backup before it is replaced.
            var merging = item.Chosen == SyncAction.Merge;
            if (merging)
            {
                BackUp(path, _fileSystem.Path.Combine(backup, "folder"));
            }

            var original = GitSyncText.NormalizeLineEndings(merging
                ? item.MergedText
                : _fileSystem.File.ReadAllText(path, GitSyncFolderProcessor.EncodingFor(path)));
            var repaired = SyncFixer.FixFolder(SyncFixer.Fix(item.Name, original, scan.Config.Sync), directory, scan.Config.Sync);
            var prepared = SyncText.PrepareForImport(repaired, directory);
            var copy = _fileSystem.Path.Combine(temp, "import", item.Name, _fileSystem.Path.GetFileName(path));
            _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(copy));
            _fileSystem.File.WriteAllText(copy, prepared, GitSyncFolderProcessor.EncodingFor(path));

            var binary = _fileSystem.Path.ChangeExtension(path, ComponentTypeExtensions.FormBinaryExtension);
            if (_fileSystem.File.Exists(binary))
            {
                _fileSystem.File.Copy(binary, _fileSystem.Path.ChangeExtension(copy, ComponentTypeExtensions.FormBinaryExtension), true);
            }

            return new PreparedImport
            {
                Relative = relative,
                Prepared = prepared,
                Copy = copy,
                FilePath = path,
                RepairedFile = merging || repaired != original ? repaired : null
            };
        }

        private Applied Export(IVBComponents components, SyncItem item, GitSyncScan scan, string temp, string backup, GitSyncResult result)
        {
            using (var component = FindComponent(components, item.Name))
            {
                if (component == null)
                {
                    throw new InvalidOperationException("the project has no such module");
                }

                var extension = component.Type.FileExtension();
                var exported = component.ExportAsSourceFile(_fileSystem.Path.Combine(temp, item.Name + extension));
                try
                {
                    var text = _fileSystem.File.ReadAllText(exported, GitSyncFolderProcessor.EncodingFor(exported));
                    var repaired = SyncFixer.Fix(item.Name, text, scan.Config.Sync);
                    var needsRepair = repaired != text;
                    if (needsRepair)
                    {
                        text = repaired;
                        _fileSystem.File.WriteAllText(exported, text, GitSyncFolderProcessor.EncodingFor(exported));
                    }

                    var folder = GitSyncText.GetFolderAnnotation(SyncText.CodeOf(text)) ?? string.Empty;
                    scan.State.Modules.TryGetValue(item.Name, out var baseline);

                    var relative = GitSyncPlanner.TargetPath(item.Name, extension, folder, baseline).Replace('/', '\\');
                    var target = _fileSystem.Path.Combine(scan.Folder, relative);

                    BackUp(target, backup);
                    _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(target));
                    CopyIfDifferent(exported, target);

                    var binaryExtension = component.Type.BinaryFileExtension();
                    if (binaryExtension.Length > 0)
                    {
                        var exportedBinary = _fileSystem.Path.ChangeExtension(exported, binaryExtension);
                        if (_fileSystem.File.Exists(exportedBinary))
                        {
                            var targetBinary = _fileSystem.Path.ChangeExtension(target, binaryExtension);
                            BackUp(targetBinary, backup);
                            CopyIfDifferent(exportedBinary, targetBinary);
                        }
                    }

                    //The module moved to another folder, so its old file goes.
                    if (item.FileExists && item.MovesFile)
                    {
                        var old = _fileSystem.Path.Combine(scan.Folder, item.RelativePath.Replace('/', '\\'));
                        if (!string.Equals(_fileSystem.Path.GetFullPath(old), _fileSystem.Path.GetFullPath(target), StringComparison.OrdinalIgnoreCase))
                        {
                            BackUp(old, backup);
                            DeleteWithBinary(old);
                        }
                    }

                    if (needsRepair)
                    {
                        //The project's module is brought in line with the file, through the VBE's own import so its attributes survive.
                        var copy = _fileSystem.Path.Combine(temp, "repaired", item.Name, item.Name + extension);
                        _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(copy));
                        _fileSystem.File.Copy(exported, copy, true);
                        var binary = component.Type.BinaryFileExtension();
                        if (binary.Length > 0 && _fileSystem.File.Exists(_fileSystem.Path.ChangeExtension(exported, binary)))
                        {
                            _fileSystem.File.Copy(_fileSystem.Path.ChangeExtension(exported, binary), _fileSystem.Path.ChangeExtension(copy, binary), true);
                        }

                        var imported = Import(components, item, new PreparedImport { Relative = relative, Prepared = text, Copy = copy }, temp, backup, result);
                        result.Notes.Add($"'{item.Name}' was repaired to follow the enforced rules.");
                        return imported;
                    }

                    //What was written is what was exported, so its text is at hand and the file need not be read back.
                    return new Applied { Item = item, Relative = relative, Folder = folder, ModuleText = text, FileText = text };
                }
                finally
                {
                    DeleteWithBinary(exported);
                }
            }
        }

        private Applied Import(IVBComponents components, SyncItem item, PreparedImport prepared, string temp, string backup, GitSyncResult result)
        {
            using (var existing = FindComponent(components, item.Name))
            {
                string saved = null;
                if (existing != null)
                {
                    _fileSystem.Directory.CreateDirectory(backup);
                    saved = existing.ExportAsSourceFile(_fileSystem.Path.Combine(backup, item.Name + existing.Type.FileExtension()));
                }

                try
                {
                    if (existing == null)
                    {
                        ImportNew(components, item.Name, prepared.Copy, result);
                    }
                    else if (existing.Type == ComponentType.Document)
                    {
                        //Document modules cannot be removed; only their code is replaced.
                        using (components.ImportSourceFile(prepared.Copy)) { }
                    }
                    else
                    {
                        components.Remove(existing);
                        using (components.Import(prepared.Copy)) { }
                    }
                }
                catch (Exception)
                {
                    RestoreAfterFailedImport(components, item.Name, saved, result);
                    throw;
                }
            }

            if (prepared.RepairedFile != null)
            {
                _fileSystem.File.WriteAllText(prepared.FilePath, prepared.RepairedFile, GitSyncFolderProcessor.EncodingFor(prepared.FilePath));
                result.Notes.Add($"'{item.Name}': its file was repaired to follow the enforced rules.");
            }

            //The VBE may have rewritten the header while importing, so the module is exported once more to see what it holds now.
            using (var imported = FindComponent(components, item.Name))
            {
                RemoveLeadingBlankLines(imported);

                var temporary = imported.ExportAsSourceFile(_fileSystem.Path.Combine(temp, item.Name + imported.Type.FileExtension()));
                try
                {
                    var text = _fileSystem.File.ReadAllText(temporary, GitSyncFolderProcessor.EncodingFor(temporary));
                    return new Applied
                    {
                        Item = item,
                        Relative = prepared.Relative,
                        Folder = GitSyncText.GetFolderAnnotation(SyncText.CodeOf(text)) ?? string.Empty,
                        ModuleText = text,
                        FileText = prepared.Prepared
                    };
                }
                finally
                {
                    DeleteWithBinary(temporary);
                }
            }
        }

        //Runs on a worker thread, after the VBE is free again. Both texts are already as they are written; no need to read anything back.
        private ModuleBaseline BaselineAfter(Applied applied, GitSyncScan scan)
        {
            var directory = _fileSystem.Path.GetDirectoryName(applied.Relative) ?? string.Empty;
            return new ModuleBaseline
            {
                Module = scan.Cache.Signature(applied.ModuleText, scan.Settings, _indenter),
                File = scan.Cache.Signature(SyncText.PrepareForImport(applied.FileText, directory), scan.Settings, _indenter),
                Path = applied.Relative.Replace('\\', '/'),
                Folder = applied.Folder
            };
        }

        //Importing a module that was removed first and then failed would leave the project without it; put the old one back.
        private void RestoreAfterFailedImport(IVBComponents components, string name, string saved, GitSyncResult result)
        {
            if (saved == null)
            {
                return;
            }

            try
            {
                using (var present = FindComponent(components, name))
                {
                    if (present == null)
                    {
                        using (components.Import(saved)) { }
                        result.Notes.Add($"'{name}' could not be imported; the project's version was put back.");
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception, $"Git sync: could not put '{name}' back.");
                result.Errors.Add($"'{name}' could not be put back. Its previous version is saved in {_fileSystem.Path.GetDirectoryName(saved)}");
            }
        }

        //The VBE often refuses to import a user form that does not exist yet; it works when an empty form of that name is there.
        private void ImportNew(IVBComponents components, string name, string path, GitSyncResult result)
        {
            try
            {
                using (components.Import(path)) { }
            }
            catch (Exception exception)
            {
                var isForm = string.Equals(_fileSystem.Path.GetExtension(path), ComponentTypeExtensions.FormExtension, StringComparison.OrdinalIgnoreCase);
                if (!isForm)
                {
                    throw;
                }

                Logger.Warn(exception, $"Could not import '{name}' as a new module; creating an empty user form instead.");
                using (var created = components.Add(ComponentType.UserForm))
                {
                    created.Name = name;
                }

                using (components.ImportSourceFile(path)) { }
                result.Notes.Add($"'{name}' could not be imported as a new form, so an empty form was created and filled from the file.");
            }
        }

        private static void RemoveLeadingBlankLines(IVBComponent component)
        {
            if (component == null)
            {
                return;
            }

            using (var module = component.CodeModule)
            {
                var lines = (module.Content() ?? string.Empty).Split(new[] { "\r\n" }, StringSplitOptions.None);
                var blank = lines.TakeWhile(line => line.Trim().Length == 0).Count();
                if (blank > 0 && blank < lines.Length)
                {
                    module.DeleteLines(1, blank);
                }
            }
        }

        private static IVBComponent FindComponent(IVBComponents components, string name)
        {
            IVBComponent found = null;
            foreach (var component in components)
            {
                if (found == null && component.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    found = component;
                }
                else
                {
                    component.Dispose();
                }
            }

            return found;
        }

        #endregion

        #region File helpers

        private string NewTempFolder()
        {
            var folder = _fileSystem.Path.Combine(_fileSystem.Path.GetTempPath(), "RubberduckGitSync", Guid.NewGuid().ToString("N"));
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

        private void BackUp(string path, string backup)
        {
            if (!_fileSystem.File.Exists(path))
            {
                return;
            }

            _fileSystem.Directory.CreateDirectory(backup);
            _fileSystem.File.Copy(path, _fileSystem.Path.Combine(backup, _fileSystem.Path.GetFileName(path)), true);
        }

        private void DeleteWithBinary(string path)
        {
            if (_fileSystem.File.Exists(path))
            {
                _fileSystem.File.Delete(path);
            }

            var binary = _fileSystem.Path.ChangeExtension(path, ComponentTypeExtensions.FormBinaryExtension);
            if (_fileSystem.File.Exists(binary))
            {
                _fileSystem.File.Delete(binary);
            }
        }

        private void CopyIfDifferent(string source, string target)
        {
            if (_fileSystem.File.Exists(target) && _fileSystem.File.ReadAllBytes(source).SequenceEqual(_fileSystem.File.ReadAllBytes(target)))
            {
                return;
            }

            _fileSystem.File.Copy(source, target, true);
        }

        #endregion
    }
}
