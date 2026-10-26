using System;
using System.Collections.Generic;
using FileSystemWatcher = System.IO.FileSystemWatcher;
using IOException = System.IO.IOException;
using NotifyFilters = System.IO.NotifyFilters;
using System.IO.Abstractions;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
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
        public int LineEndingsFixed { get; set; }
        public int FolderAnnotationsFixed { get; set; }
        public List<string> Errors { get; } = new List<string>();
    }

    /// <summary>
    /// Keeps one VBA project in sync with one folder. Files win when the initial sync finds a difference.
    /// Afterwards, file changes are imported and module changes are exported as they happen.
    /// Nothing is ever deleted from the project; files are only deleted when a module moves to another folder.
    /// </summary>
    public sealed class GitSyncSession : IDisposable
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private const int DebounceMilliseconds = 500;

        private readonly IProjectsProvider _projectsProvider;
        private readonly IParseManager _parseManager;
        private readonly IModuleNameFromFileExtractor _moduleNameExtractor;
        private readonly IFileSystem _fileSystem;
        private readonly GitSyncFolderProcessor _processor;
        private readonly string _tempFolder;

        private readonly object _lock = new object();
        private readonly Dictionary<string, int> _componentHashes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _fileHashes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _modulePaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _moduleFolders = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _pendingFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private FileSystemWatcher _watcher;
        private Timer _debounce;
        private bool _disposed;

        public GitSyncSession(
            string projectId,
            string root,
            IProjectsProvider projectsProvider,
            IParseManager parseManager,
            IModuleNameFromFileExtractor moduleNameExtractor,
            IFileSystem fileSystem)
        {
            ProjectId = projectId;
            Root = fileSystem.Path.GetFullPath(root);
            _projectsProvider = projectsProvider;
            _parseManager = parseManager;
            _moduleNameExtractor = moduleNameExtractor;
            _fileSystem = fileSystem;
            _processor = new GitSyncFolderProcessor(fileSystem);
            _tempFolder = fileSystem.Path.Combine(fileSystem.Path.GetTempPath(), "RubberduckGitSync", Guid.NewGuid().ToString("N"));
        }

        public string ProjectId { get; }
        public string Root { get; }

        /// <summary>
        /// Synchronizes the folder with the project once and starts watching the folder.
        /// </summary>
        public GitSyncResult Start()
        {
            var result = new GitSyncResult();
            lock (_lock)
            {
                _fileSystem.Directory.CreateDirectory(Root);

                var prepared = _processor.PrepareAll(Root).ToList();
                CountPrepared(prepared, result);

                ImportFiles(prepared.Select(file => file.Path).ToList(), result);
                ExportChangedModulesCore(result);

                StartWatching();
            }

            return result;
        }

        /// <summary>
        /// Writes the modules changed since the last sync to the folder. Called when the parser is ready.
        /// </summary>
        public void ExportChangedModules()
        {
            var result = new GitSyncResult();
            lock (_lock)
            {
                if (_disposed)
                {
                    return;
                }

                ExportChangedModulesCore(result);
            }

            LogErrors(result);
        }

        private static void CountPrepared(IEnumerable<PreparedFile> prepared, GitSyncResult result)
        {
            foreach (var file in prepared)
            {
                result.LineEndingsFixed += file.LineEndingsFixed ? 1 : 0;
                result.FolderAnnotationsFixed += file.FolderAnnotationFixed ? 1 : 0;
            }
        }

        private static void LogErrors(GitSyncResult result)
        {
            foreach (var error in result.Errors)
            {
                Logger.Warn(error);
            }
        }

        #region Folder -> project

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
                    foreach (var path in files.Where(_fileSystem.File.Exists))
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
                    ImportFiles(prepared.Select(file => file.Path).ToList(), result);
                }
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "Git sync failed to import changed files.");
            }

            LogErrors(result);
        }

        private void ImportFiles(ICollection<string> paths, GitSyncResult result)
        {
            if (paths.Count == 0)
            {
                return;
            }

            var suspension = _parseManager.OnSuspendParser(this, new[] { ParserState.Ready }, () => ImportFilesCore(paths, result));
            if (suspension.Outcome != SuspensionOutcome.Completed)
            {
                throw new InvalidOperationException(
                    $"The files could not be imported because the parser could not be suspended ({suspension.Outcome}).",
                    suspension.EncounteredException);
            }
        }

        private void ImportFilesCore(IEnumerable<string> paths, GitSyncResult result)
        {
            var project = _projectsProvider.Project(ProjectId);
            if (project == null)
            {
                return;
            }

            using (var components = project.VBComponents)
            {
                foreach (var path in paths)
                {
                    try
                    {
                        ImportFile(components, path, result);
                    }
                    catch (Exception exception)
                    {
                        result.Errors.Add($"Could not import '{path}': {exception.Message}");
                    }
                }
            }
        }

        private void ImportFile(IVBComponents components, string path, GitSyncResult result)
        {
            var name = _moduleNameExtractor.ModuleName(path);
            if (name == null)
            {
                return;
            }

            var fileHash = HashOfFile(path);
            if (_fileHashes.TryGetValue(path, out var knownHash) && knownHash == fileHash)
            {
                return;
            }

            var text = _fileSystem.File.ReadAllText(path, GitSyncFolderProcessor.SourceEncoding);

            using (var component = FindComponent(components, name))
            {
                if (component != null && IsSameAsFile(component, path))
                {
                    Record(name, path, fileHash, text, component.ContentHash());
                    return;
                }

                if (component == null)
                {
                    using (components.Import(path)) { }
                }
                else if (component.Type == ComponentType.Document)
                {
                    using (components.ImportSourceFile(path)) { }
                }
                else
                {
                    components.Remove(component);
                    using (components.Import(path)) { }
                }
            }

            using (var imported = FindComponent(components, name))
            {
                Record(name, path, fileHash, text, imported?.ContentHash() ?? 0);
            }

            result.Imported++;
        }

        private void Record(string name, string path, string fileHash, string text, int componentHash)
        {
            _fileHashes[path] = fileHash;
            _modulePaths[name] = path;
            _moduleFolders[name] = GitSyncText.GetFolderAnnotation(text) ?? string.Empty;
            _componentHashes[name] = componentHash;
        }

        private bool IsSameAsFile(IVBComponent component, string path)
        {
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

        #endregion

        #region Project -> folder

        private void ExportChangedModulesCore(GitSyncResult result)
        {
            var project = _projectsProvider.Project(ProjectId);
            if (project == null)
            {
                return;
            }

            using (var components = project.VBComponents)
            {
                foreach (var component in components)
                {
                    using (component)
                    {
                        try
                        {
                            ExportIfChanged(component, result);
                        }
                        catch (Exception exception)
                        {
                            result.Errors.Add($"Could not export '{component.Name}': {exception.Message}");
                        }
                    }
                }
            }
        }

        private void ExportIfChanged(IVBComponent component, GitSyncResult result)
        {
            var extension = component.Type.FileExtension();
            if (extension.Length == 0)
            {
                return;
            }

            var name = component.Name;
            var hash = component.ContentHash();
            if (_componentHashes.TryGetValue(name, out var knownHash) && knownHash == hash)
            {
                return;
            }

            var temp = ExportToTemp(component);
            try
            {
                var text = _fileSystem.File.ReadAllText(temp, GitSyncFolderProcessor.SourceEncoding);
                var folder = GitSyncText.GetFolderAnnotation(text) ?? string.Empty;

                //The file stays where it is unless the module was moved to another virtual folder.
                var target = _modulePaths.TryGetValue(name, out var trackedPath)
                             && _moduleFolders.TryGetValue(name, out var trackedFolder)
                             && trackedFolder == folder
                    ? trackedPath
                    : _fileSystem.Path.Combine(Root, GitSyncText.RelativeDirectory(folder), name + extension);

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

                if (_modulePaths.TryGetValue(name, out var oldPath) && !PathsAreEqual(oldPath, target))
                {
                    DeleteWithBinary(oldPath, binaryExtension);
                }

                Record(name, target, HashOfFile(target), text, hash);
                result.Exported += changed ? 1 : 0;
            }
            finally
            {
                DeleteTempFiles(temp);
            }
        }

        #endregion

        #region File helpers

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
                _watcher = null;
                _debounce = null;
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
