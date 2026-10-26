using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using FileSystemWatcher = System.IO.FileSystemWatcher;
using IOException = System.IO.IOException;
using NotifyFilters = System.IO.NotifyFilters;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    public class GitSyncResult
    {
        public int Imported { get; set; }
        public int Exported { get; set; }
        public int Conflicts { get; set; }
        /// <summary>The user declined to pick a winner for modules that exist on both sides, so nothing was synced.</summary>
        public bool Cancelled { get; set; }
        public int LineEndingsFixed { get; set; }
        public int FolderAnnotationsFixed { get; set; }
        public List<string> Errors { get; } = new List<string>();
        public List<string> Notes { get; } = new List<string>();
    }

    /// <summary>
    /// Keeps one VBA project in sync with one folder. For every module, a snapshot of both sides from the last sync
    /// tells which one changed since, and that decides the direction. Changes made while Rubberduck was not running
    /// are picked up as well, because the snapshots are persisted.
    /// Nothing is ever deleted from the project; files are only deleted when a module moves to another folder.
    /// </summary>
    public sealed class GitSyncSession : IDisposable
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private const int DebounceMilliseconds = 500;
        private const int PollMilliseconds = 2000;
        private static readonly TimeSpan QuietPeriod = TimeSpan.FromSeconds(2);

        private readonly string _manifestKey;
        private readonly GitSyncManifestStore _manifestStore;
        private readonly IProjectsProvider _projectsProvider;
        private readonly IParseManager _parseManager;
        private readonly IModuleNameFromFileExtractor _moduleNameExtractor;
        private readonly IFileSystem _fileSystem;
        private readonly GitSyncFolderProcessor _processor;
        private readonly string _tempFolder;

        private readonly object _lock = new object();
        private readonly HashSet<string> _pendingFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private ProjectManifest _manifest;
        //What the poll saw last for each module, to tell when someone is still typing.
        private readonly Dictionary<string, KeyValuePair<string, DateTime>> _lastSeen = new Dictionary<string, KeyValuePair<string, DateTime>>(StringComparer.OrdinalIgnoreCase);

        private FileSystemWatcher _watcher;
        private Timer _debounce;
        private Timer _poll;
        private int _polling;
        private bool _disposed;

        public GitSyncSession(
            string projectId,
            string root,
            string manifestKey,
            GitSyncManifestStore manifestStore,
            IProjectsProvider projectsProvider,
            IParseManager parseManager,
            IModuleNameFromFileExtractor moduleNameExtractor,
            IFileSystem fileSystem)
        {
            ProjectId = projectId;
            Root = fileSystem.Path.GetFullPath(root);
            _manifestKey = manifestKey;
            _manifestStore = manifestStore;
            _projectsProvider = projectsProvider;
            _parseManager = parseManager;
            _moduleNameExtractor = moduleNameExtractor;
            _fileSystem = fileSystem;
            _processor = new GitSyncFolderProcessor(fileSystem);
            _tempFolder = fileSystem.Path.Combine(fileSystem.Path.GetTempPath(), "RubberduckGitSync", Guid.NewGuid().ToString("N"));
        }

        public string ProjectId { get; }
        public string Root { get; }

        private Dictionary<string, ModuleSnapshot> Snapshots => _manifest.Modules;

        /// <summary>
        /// Synchronizes the folder with the project once and starts watching the folder.
        /// </summary>
        /// <param name="chooseWinner">Asked who wins for the modules that exist on both sides but were never synced and carry no stamp. Without it, those modules are skipped.</param>
        public GitSyncResult Start(Func<IReadOnlyList<InitialConflict>, InitialChoice> chooseWinner = null)
        {
            var result = new GitSyncResult();
            lock (_lock)
            {
                _fileSystem.Directory.CreateDirectory(Root);
                _manifest = _manifestStore.Load(_manifestKey, Root);

                var prepared = _processor.PrepareAll(Root).ToList();
                CountPrepared(prepared, result);

                Suspended(() => SyncModules(prepared.Select(file => file.Path).ToList(), true, result, null, chooseWinner));
                if (result.Cancelled)
                {
                    return result;
                }

                StartWatching();
            }

            return result;
        }

        /// <summary>
        /// The VBE raises no event when code is typed, and Rubberduck itself only notices edits when it reparses.
        /// So the code of every module is looked at every couple of seconds. This is also how the time of the last
        /// modification is known: it is the time the poll first saw the current code.
        /// </summary>
        private void Poll()
        {
            if (Interlocked.Exchange(ref _polling, 1) == 1)
            {
                return;
            }

            var result = new GitSyncResult();
            try
            {
                lock (_lock)
                {
                    if (_disposed || _manifest == null)
                    {
                        return;
                    }

                    var due = ModulesDueForSync();
                    if (due.Count == 0)
                    {
                        return;
                    }

                    //A conflict means importing, which needs the parser to hold still.
                    Suspended(() => SyncModules(due.Select(TrackedPath).Where(path => path != null).ToList(), false, result, due));
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Git sync failed to sync the changed modules.");
            }
            finally
            {
                Interlocked.Exchange(ref _polling, 0);
            }

            LogOutcome(result);
        }

        //Modules whose code differs from the snapshot and that nobody has typed in for a moment.
        private HashSet<string> ModulesDueForSync()
        {
            var due = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var project = _projectsProvider.Project(ProjectId);
            if (project == null)
            {
                return due;
            }

            var now = DateTime.UtcNow;
            var manifestChanged = false;
            using (var components = project.VBComponents)
            {
                foreach (var component in components)
                {
                    using (component)
                    {
                        if (component.Type.FileExtension().Length == 0)
                        {
                            continue;
                        }

                        var name = component.Name;
                        var hash = ModuleHash(component);

                        if (!_lastSeen.TryGetValue(name, out var seen) || seen.Key != hash)
                        {
                            _lastSeen[name] = new KeyValuePair<string, DateTime>(hash, now);
                            seen = _lastSeen[name];

                            if (Snapshots.TryGetValue(name, out var changedSnapshot) && changedSnapshot.ModuleHash != hash)
                            {
                                changedSnapshot.ModuleChangedUtc = now;
                                manifestChanged = true;
                            }
                        }

                        Snapshots.TryGetValue(name, out var snapshot);
                        if ((snapshot == null || snapshot.ModuleHash != hash) && now - seen.Value >= QuietPeriod)
                        {
                            due.Add(name);
                        }
                    }
                }
            }

            if (manifestChanged)
            {
                SaveManifest();
            }

            return due;
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

        private static void CountPrepared(IEnumerable<PreparedFile> prepared, GitSyncResult result)
        {
            foreach (var file in prepared)
            {
                result.LineEndingsFixed += file.LineEndingsFixed ? 1 : 0;
                result.FolderAnnotationsFixed += file.FolderAnnotationFixed ? 1 : 0;
            }
        }

        private static void LogOutcome(GitSyncResult result)
        {
            foreach (var error in result.Errors)
            {
                Logger.Warn(error);
            }

            foreach (var note in result.Notes)
            {
                Logger.Info(note);
            }
        }

        private void SaveManifest()
        {
            _manifestStore.Save(_manifestKey, _manifest);
        }

        private string TrackedPath(string name)
        {
            return Snapshots.TryGetValue(name, out var snapshot) && snapshot.RelativePath != null
                ? _fileSystem.Path.Combine(Root, snapshot.RelativePath)
                : null;
        }

        #region Watching the folder

        private void StartWatching()
        {
            _debounce = new Timer(_ => ProcessPendingFiles(), null, Timeout.Infinite, Timeout.Infinite);

            _watcher = new FileSystemWatcher(Root)
            {
                IncludeSubdirectories = true,
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size
            };
            _watcher.Changed += (sender, e) => OnFileEvent(e.FullPath);
            _watcher.Created += (sender, e) => OnFileEvent(e.FullPath);
            _watcher.Renamed += (sender, e) => OnFileEvent(e.FullPath);
            _watcher.EnableRaisingEvents = true;

            _poll = new Timer(_ => Poll(), null, PollMilliseconds, PollMilliseconds);
        }

        private void OnFileEvent(string path)
        {
            //A changed form binary means the form needs to be imported again.
            if (_fileSystem.Path.GetExtension(path).Equals(ComponentTypeExtensions.FormBinaryExtension, StringComparison.OrdinalIgnoreCase))
            {
                path = _fileSystem.Path.ChangeExtension(path, ComponentTypeExtensions.FormExtension);
            }

            if (!_processor.IsSourceFile(path))
            {
                return;
            }

            lock (_pendingFiles)
            {
                _pendingFiles.Add(path);
            }

            _debounce?.Change(DebounceMilliseconds, Timeout.Infinite);
        }

        private void ProcessPendingFiles()
        {
            List<string> files;
            lock (_pendingFiles)
            {
                files = _pendingFiles.ToList();
                _pendingFiles.Clear();
            }

            var result = new GitSyncResult();
            try
            {
                lock (_lock)
                {
                    if (_disposed)
                    {
                        return;
                    }

                    var prepared = new List<PreparedFile>();
                    var ignored = _processor.LoadIgnoreRules(Root);
                    foreach (var path in files.Where(file => _fileSystem.File.Exists(file) && !ignored.IsIgnored(RelativePathOf(file), false)))
                    {
                        try
                        {
                            prepared.Add(_processor.Prepare(Root, path));
                        }
                        catch (IOException)
                        {
                            //Most likely still being written by git or an editor; try again later.
                            lock (_pendingFiles)
                            {
                                _pendingFiles.Add(path);
                            }
                            _debounce?.Change(DebounceMilliseconds * 4, Timeout.Infinite);
                        }
                    }

                    CountPrepared(prepared, result);
                    if (prepared.Count > 0)
                    {
                        Suspended(() => SyncModules(prepared.Select(file => file.Path).ToList(), false, result));
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Git sync failed to process changed files.");
            }

            LogOutcome(result);
        }

        #endregion

        #region Syncing

        /// <param name="paths">The files to consider.</param>
        /// <param name="allModules">Whether modules without any of the given files are considered too.</param>
        /// <param name="onlyNames">If given, only these modules are considered.</param>
        /// <param name="chooseWinner">If given, asked who wins where both sides exist but nothing says which is newer; otherwise such modules are skipped.</param>
        private void SyncModules(
            ICollection<string> paths,
            bool allModules,
            GitSyncResult result,
            ISet<string> onlyNames = null,
            Func<IReadOnlyList<InitialConflict>, InitialChoice> chooseWinner = null)
        {
            var project = _projectsProvider.Project(ProjectId);
            if (project == null)
            {
                return;
            }

            var ignored = _processor.LoadIgnoreRules(Root);

            var fileByName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var path in paths)
            {
                var name = _moduleNameExtractor.ModuleName(path);
                if (name != null && !fileByName.ContainsKey(name))
                {
                    fileByName.Add(name, path);
                }
            }

            using (var components = project.VBComponents)
            {
                var names = new HashSet<string>(fileByName.Keys, StringComparer.OrdinalIgnoreCase);
                if (allModules || onlyNames != null)
                {
                    foreach (var component in components)
                    {
                        using (component)
                        {
                            names.Add(component.Name);
                        }
                    }
                }

                InitialChoice? winner = null;
                if (chooseWinner != null)
                {
                    var unresolved = fileByName
                        .Where(file => NeedsWinner(components, file.Key, file.Value, ignored))
                        .Select(file => new InitialConflict(file.Key, _fileSystem.File.GetLastWriteTime(file.Value)))
                        .ToList();
                    if (unresolved.Count > 0)
                    {
                        winner = chooseWinner(unresolved);
                        if (winner == InitialChoice.Cancel)
                        {
                            result.Cancelled = true;
                            return;
                        }
                    }
                }

                foreach (var name in names.Where(name => onlyNames == null || onlyNames.Contains(name)))
                {
                    try
                    {
                        var path = fileByName.TryGetValue(name, out var given) ? given : TrackedPath(name);
                        SyncModule(components, name, path, ignored, result, winner);
                    }
                    catch (Exception exception)
                    {
                        result.Errors.Add($"Could not sync '{name}': {exception.Message}");
                    }
                }
            }

            SaveManifest();
        }

        //Both sides exist, differ, and there is no snapshot and no complete pair of stamps to say which is newer.
        private bool NeedsWinner(IVBComponents components, string name, string filePath, GitIgnoreSet ignored)
        {
            if (Snapshots.ContainsKey(name) || ignored.IsIgnored(RelativePathOf(filePath), false))
            {
                return false;
            }

            using (var component = FindComponent(components, name))
            {
                if (component == null || component.Type.FileExtension().Length == 0)
                {
                    return false;
                }

                if (GitSyncStamp.StateOfCode(CodeOf(component)) != null && GitSyncStamp.StateOfFile(ReadSource(filePath)) != null)
                {
                    return false;
                }

                return !IsSameAsFile(component, filePath);
            }
        }

        private void SyncModule(IVBComponents components, string name, string filePath, GitIgnoreSet ignored, GitSyncResult result, InitialChoice? winner = null)
        {
            using (var component = FindComponent(components, name))
            {
                if (component != null && component.Type.FileExtension().Length == 0)
                {
                    return;
                }

                var fileExists = filePath != null
                                 && _fileSystem.File.Exists(filePath)
                                 && !ignored.IsIgnored(RelativePathOf(filePath), false);
                if (!fileExists)
                {
                    filePath = null;
                }

                Snapshots.TryGetValue(name, out var snapshot);

                var code = component == null ? null : CodeOf(component);
                var fileText = fileExists ? ReadSource(filePath) : null;
                var moduleStamp = GitSyncStamp.StateOfCode(code);
                var fileStamp = GitSyncStamp.StateOfFile(fileText);
                var fileWritten = fileExists ? _fileSystem.File.GetLastWriteTimeUtc(filePath) : DateTime.MinValue;
                var now = DateTime.UtcNow;

                //The stamps travel with the workbook and the repository, so they are the better guide; the local snapshot
                //only decides where one side has not been stamped yet.
                var decision = moduleStamp != null && fileStamp != null
                    ? SyncDirectionResolver.DecideByStamps(moduleStamp, fileStamp, fileWritten, snapshot?.ModuleChangedUtc ?? now)
                    : SyncDirectionResolver.Decide(
                        snapshot,
                        component == null ? null : ModuleHash(component),
                        fileExists ? HashOfFile(filePath) : null,
                        () => IsSameAsFile(component, filePath),
                        fileWritten,
                        now);

                var direction = decision.Direction;
                if (direction == SyncDirection.Ask)
                {
                    if (winner == null)
                    {
                        result.Notes.Add($"'{name}' exists in the project and in the folder, but nothing says which is newer. Run Git Sync... to choose.");
                        return;
                    }

                    direction = winner == InitialChoice.Folder ? SyncDirection.Import : SyncDirection.Export;
                }

                switch (direction)
                {
                    case SyncDirection.Record:
                        StampBoth(component, filePath, code, fileText, moduleStamp, fileStamp);
                        Record(name, filePath, component);
                        break;

                    case SyncDirection.Import:
                        if (decision.IsConflict)
                        {
                            BackUpModule(component, filePath, name, result);
                        }

                        StampFile(filePath, fileText, moduleStamp, fileStamp);
                        ImportFile(components, component, name, filePath);
                        result.Imported++;
                        break;

                    case SyncDirection.Export:
                        if (decision.IsConflict)
                        {
                            BackUpFile(filePath, name, result);
                        }

                        StampModule(component, code, moduleStamp, fileStamp);
                        if (ExportModule(component, name, ignored))
                        {
                            result.Exported++;
                        }
                        break;
                }
            }
        }

        #region Stamps

        private static string CodeOf(IVBComponent component)
        {
            using (var module = component.CodeModule)
            {
                return module.Content() ?? string.Empty;
            }
        }

        private string ReadSource(string path)
        {
            return _fileSystem.File.ReadAllText(path, GitSyncFolderProcessor.EncodingFor(path));
        }

        private static int NextRevision(StampState module, StampState file)
        {
            return Math.Max(module?.Revision ?? 0, file?.Revision ?? 0) + 1;
        }

        //Before a file is imported, its stamp has to describe its content; a hand edit or a merge may have invalidated it.
        private void StampFile(string path, string fileText, StampState moduleStamp, StampState fileStamp)
        {
            if (fileStamp != null && !fileStamp.Dirty)
            {
                return;
            }

            var hash = GitSyncStamp.HashOfCode(GitSyncStamp.CodeOfFile(fileText));
            WriteFileStamp(path, fileText, NextRevision(moduleStamp, fileStamp), hash, _fileSystem.File.GetLastWriteTimeUtc(path));
        }

        //Before a module is exported, its stamp has to describe its code.
        private void StampModule(IVBComponent component, string code, StampState moduleStamp, StampState fileStamp)
        {
            if (moduleStamp != null && !moduleStamp.Dirty)
            {
                return;
            }

            Snapshots.TryGetValue(component.Name, out var snapshot);
            SetModuleStamp(component, NextRevision(moduleStamp, fileStamp), GitSyncStamp.HashOfCode(code), snapshot?.ModuleChangedUtc ?? DateTime.UtcNow);
        }

        //Both sides hold the same code; give them the same stamp, keeping the revision if one side already has a valid stamp.
        private void StampBoth(IVBComponent component, string filePath, string code, string fileText, StampState moduleStamp, StampState fileStamp)
        {
            if (component == null || filePath == null)
            {
                return;
            }

            var hash = GitSyncStamp.HashOfCode(code);
            int revision;
            DateTime utc;
            if (moduleStamp != null && !moduleStamp.Dirty)
            {
                revision = moduleStamp.Revision;
                utc = moduleStamp.Utc ?? DateTime.UtcNow;
            }
            else if (fileStamp != null && !fileStamp.Dirty)
            {
                revision = fileStamp.Revision;
                utc = fileStamp.Utc ?? DateTime.UtcNow;
            }
            else
            {
                revision = NextRevision(moduleStamp, fileStamp);
                Snapshots.TryGetValue(component.Name, out var snapshot);
                utc = snapshot?.ModuleChangedUtc ?? DateTime.UtcNow;
            }

            if (moduleStamp == null || moduleStamp.Revision != revision || moduleStamp.Hash != hash || moduleStamp.Utc == null)
            {
                SetModuleStamp(component, revision, hash, utc);
            }

            if (fileStamp == null || fileStamp.Revision != revision || fileStamp.Hash != hash || fileStamp.Utc == null)
            {
                WriteFileStamp(filePath, fileText, revision, hash, utc);
            }
        }

        private void WriteFileStamp(string path, string fileText, int revision, string hash, DateTime utc)
        {
            _fileSystem.File.WriteAllText(path, GitSyncStamp.WithStamp(fileText, revision, hash, utc), GitSyncFolderProcessor.EncodingFor(path));
        }

        private static void SetModuleStamp(IVBComponent component, int revision, string hash, DateTime utc)
        {
            var stamp = GitSyncStamp.Format(revision, hash, utc);
            using (var module = component.CodeModule)
            {
                var lines = (module.Content() ?? string.Empty).Split(new[] { "\r\n" }, StringSplitOptions.None);
                var index = Array.FindIndex(lines, GitSyncStamp.IsStampLine);
                if (index >= 0)
                {
                    module.ReplaceLine(index + 1, stamp);
                }
                else
                {
                    module.InsertLines(1, stamp);
                }
            }
        }

        #endregion

        private void ImportFile(IVBComponents components, IVBComponent existing, string name, string path)
        {
            if (existing == null)
            {
                using (components.Import(path)) { }
            }
            else if (existing.Type == ComponentType.Document)
            {
                using (components.ImportSourceFile(path)) { }
            }
            else
            {
                components.Remove(existing);
                using (components.Import(path)) { }
            }

            using (var imported = FindComponent(components, name))
            {
                Record(name, path, imported);
            }
        }

        //Returns whether a file was written.
        private bool ExportModule(IVBComponent component, string name, GitIgnoreSet ignored)
        {
            var extension = component.Type.FileExtension();
            var temp = ExportToTemp(component);
            try
            {
                var text = _fileSystem.File.ReadAllText(temp, GitSyncFolderProcessor.EncodingFor(temp));
                var folder = GitSyncText.GetFolderAnnotation(text) ?? string.Empty;

                //The file stays where it is unless the module was moved to another virtual folder.
                Snapshots.TryGetValue(name, out var snapshot);
                var target = snapshot?.RelativePath != null && snapshot.Folder == folder
                    ? _fileSystem.Path.Combine(Root, snapshot.RelativePath)
                    : _fileSystem.Path.Combine(Root, GitSyncText.RelativeDirectory(folder), name + extension);

                //Modules that belong in an ignored folder stay in the project only.
                if (ignored.IsIgnored(RelativePathOf(target), false))
                {
                    Record(name, null, component);
                    return false;
                }

                var changed = CopyIfDifferent(temp, target);

                var binaryExtension = component.Type.BinaryFileExtension();
                if (binaryExtension.Length > 0)
                {
                    var tempBinary = _fileSystem.Path.ChangeExtension(temp, binaryExtension);
                    if (_fileSystem.File.Exists(tempBinary))
                    {
                        changed |= CopyIfDifferent(tempBinary, _fileSystem.Path.ChangeExtension(target, binaryExtension));
                    }
                }

                var oldPath = TrackedPath(name);
                if (oldPath != null && !PathsAreEqual(oldPath, target))
                {
                    DeleteWithBinary(oldPath, binaryExtension);
                }

                Record(name, target, component);
                return changed;
            }
            finally
            {
                DeleteTempFiles(temp);
            }
        }

        //Refreshes the snapshot: both sides are in agreement right now.
        private void Record(string name, string filePath, IVBComponent component)
        {
            var folder = string.Empty;
            if (component != null)
            {
                var temp = ExportToTemp(component);
                try
                {
                    folder = GitSyncText.GetFolderAnnotation(_fileSystem.File.ReadAllText(temp, GitSyncFolderProcessor.EncodingFor(temp))) ?? string.Empty;
                }
                finally
                {
                    DeleteTempFiles(temp);
                }
            }

            Snapshots[name] = new ModuleSnapshot
            {
                ModuleHash = component == null ? null : ModuleHash(component),
                FileHash = filePath == null ? null : HashOfFile(filePath),
                RelativePath = filePath == null ? null : RelativePathOf(filePath),
                Folder = folder,
                ModuleChangedUtc = null
            };
        }

        private void BackUpFile(string filePath, string name, GitSyncResult result)
        {
            var backup = filePath + ".conflict-" + Stamp();
            _fileSystem.File.Copy(filePath, backup, true);
            result.Conflicts++;
            result.Notes.Add($"'{name}' changed in both places; the project's version was kept. The file's version is saved as {backup}");
        }

        private void BackUpModule(IVBComponent component, string filePath, string name, GitSyncResult result)
        {
            var backup = filePath + ".module-conflict-" + Stamp();
            var temp = ExportToTemp(component);
            try
            {
                _fileSystem.File.Copy(temp, backup, true);
            }
            finally
            {
                DeleteTempFiles(temp);
            }

            result.Conflicts++;
            result.Notes.Add($"'{name}' changed in both places; the file's version was kept. The project's version is saved as {backup}");
        }

        private static string Stamp() => DateTime.UtcNow.ToString("yyyyMMddHHmmss");

        private bool IsSameAsFile(IVBComponent component, string path)
        {
            if (component == null || path == null)
            {
                return false;
            }

            var temp = ExportToTemp(component);
            try
            {
                return FilesAreEqual(temp, path);
            }
            finally
            {
                DeleteTempFiles(temp);
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

        //The code module's own ContentHash is based on string.GetHashCode, which is not stable between runs.
        private static string ModuleHash(IVBComponent component)
        {
            var content = new StringBuilder();
            using (var code = component.CodeModule)
            {
                content.Append(code.Content());
            }

            if (component.HasDesigner)
            {
                using (var controls = component.Controls)
                {
                    foreach (var control in controls)
                    {
                        using (control)
                        {
                            content.Append('\u0001').Append(control.Name);
                        }
                    }
                }
            }

            using (var sha = SHA1.Create())
            {
                return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(content.ToString())));
            }
        }

        #endregion

        #region File helpers

        private string RelativePathOf(string path)
        {
            var directory = _processor.RelativeDirectory(Root, path);
            var name = _fileSystem.Path.GetFileName(path);
            return directory.Length == 0 ? name : directory + "\\" + name;
        }

        private string ExportToTemp(IVBComponent component)
        {
            var path = _fileSystem.Path.Combine(_tempFolder, component.Name + component.Type.FileExtension());
            return component.ExportAsSourceFile(path);
        }

        private void DeleteTempFiles(string path)
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

        private void DeleteWithBinary(string path, string binaryExtension)
        {
            if (_fileSystem.File.Exists(path))
            {
                _fileSystem.File.Delete(path);
            }

            if (binaryExtension.Length > 0)
            {
                var binary = _fileSystem.Path.ChangeExtension(path, binaryExtension);
                if (_fileSystem.File.Exists(binary))
                {
                    _fileSystem.File.Delete(binary);
                }
            }
        }

        private bool PathsAreEqual(string first, string second)
        {
            return string.Equals(_fileSystem.Path.GetFullPath(first), _fileSystem.Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase);
        }

        private bool FilesAreEqual(string first, string second)
        {
            return _fileSystem.File.Exists(second)
                   && _fileSystem.File.ReadAllBytes(first).SequenceEqual(_fileSystem.File.ReadAllBytes(second));
        }

        private bool CopyIfDifferent(string source, string target)
        {
            if (FilesAreEqual(source, target))
            {
                return false;
            }

            _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(target));
            _fileSystem.File.Copy(source, target, true);
            return true;
        }

        private string HashOfFile(string path)
        {
            using (var sha = SHA1.Create())
            {
                return Convert.ToBase64String(sha.ComputeHash(_fileSystem.File.ReadAllBytes(path)));
            }
        }

        #endregion

        public void Dispose()
        {
            lock (_lock)
            {
                _disposed = true;
                _watcher?.Dispose();
                _debounce?.Dispose();
                _poll?.Dispose();
                _watcher = null;
                _debounce = null;
                _poll = null;
            }

            try
            {
                if (_fileSystem.Directory.Exists(_tempFolder))
                {
                    _fileSystem.Directory.Delete(_tempFolder, true);
                }
            }
            catch (IOException)
            {
                //The temp folder is not worth failing over.
            }
        }
    }
}
