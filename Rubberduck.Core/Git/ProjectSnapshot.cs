using System;
using System.Collections.Generic;
using System.Linq;
using System.IO.Abstractions;
using Rubberduck.UI.CodeExplorer.GitSync;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Git
{
    /// <summary>Makes a work tree hold exactly what a project holds.</summary>
    public interface IProjectSnapshot
    {
        /// <returns>What was done; its <see cref="GitSyncResult.Errors"/> are modules that could not be written, which leaves the tree behind the project.</returns>
        GitSyncResult Take(IVBProject project, string tree, ISyncProgress progress = null);

        /// <summary>The other way round for one module: its file in the tree replaces the module in the project. The module it replaces is backed up first.</summary>
        GitSyncResult Import(IVBProject project, string tree, string moduleName, ISyncProgress progress = null);

        /// <summary>Imports several files of the tree (relative paths, forward slashes) in one batch, so the parser is held still once. Files the project has no module for become modules.</summary>
        GitSyncResult ImportFiles(IVBProject project, string tree, IReadOnlyCollection<string> relativePaths, ISyncProgress progress = null);
    }

    /// <summary>
    /// A snapshot made with Git sync: the project is the truth. Every module that differs is exported over its file (the VBE's own export, so hidden attributes
    /// survive), and a file whose module is gone from the project is deleted. The tree belongs to Rubberduck, so deleting there loses nothing; a person's
    /// sync folder is never treated like this.
    /// </summary>
    public class GitProjectSnapshot : IProjectSnapshot
    {
        private readonly GitSyncService _sync;
        private readonly IFileSystem _fileSystem;

        public GitProjectSnapshot(GitSyncService sync, IFileSystem fileSystem)
        {
            _sync = sync;
            _fileSystem = fileSystem;
        }

        public GitSyncResult Take(IVBProject project, string tree, ISyncProgress progress = null)
        {
            var scan = _sync.Scan(project, tree, progress);

            var exports = scan.Plan.Items.Where(item => item.ModuleExists).ToList();
            foreach (var export in exports)
            {
                export.Chosen = SyncAction.Export;
            }

            //Resources of the document (ribbon, pictures, config) are kept even when the document has none: the config in the tree may be the one the scan just made.
            foreach (var gone in scan.Plan.Items.Where(item => !item.ModuleExists && !item.IsResource))
            {
                DeleteFile(_fileSystem.Path.Combine(tree, gone.RelativePath.Replace('/', '\\')));
                scan.State.Modules.Remove(gone.Name);
            }

            //Nothing the tree held before is worth keeping a copy of: git has it.
            var scratch = _fileSystem.Path.Combine(_fileSystem.Path.GetTempPath(), "rd-snapshot-" + Guid.NewGuid().ToString("N"));
            try
            {
                return _sync.Apply(project, scan, exports, progress, scratch);
            }
            finally
            {
                if (_fileSystem.Directory.Exists(scratch))
                {
                    _fileSystem.Directory.Delete(scratch, true);
                }
            }
        }

        public GitSyncResult Import(IVBProject project, string tree, string moduleName, ISyncProgress progress = null)
        {
            var scan = _sync.Scan(project, tree, progress);
            var item = scan.Plan.Items.FirstOrDefault(candidate => !candidate.IsResource && !candidate.IsNote && candidate.FileExists
                && string.Equals(candidate.Name, moduleName, StringComparison.OrdinalIgnoreCase));
            if (item == null)
            {
                var result = new GitSyncResult();
                result.Errors.Add($"The restored file of '{moduleName}' is the same as the module; there is nothing to import.");
                return result;
            }

            item.Chosen = SyncAction.Import;
            return _sync.Apply(project, scan, new[] { item }, progress);
        }

        public GitSyncResult ImportFiles(IVBProject project, string tree, IReadOnlyCollection<string> relativePaths, ISyncProgress progress = null)
        {
            var wanted = new HashSet<string>(relativePaths.Select(path => path.Replace('\\', '/')), StringComparer.OrdinalIgnoreCase);
            var scan = _sync.Scan(project, tree, progress);

            //Modules and notes; a ribbon, a picture or the config stored in the document is not something a pull changes in the project.
            var items = scan.Plan.Items.Where(item => !item.IsResource && item.FileExists && wanted.Contains(item.RelativePath.Replace('\\', '/'))).ToList();
            foreach (var item in items)
            {
                item.Chosen = SyncAction.Import;
            }

            var result = _sync.Apply(project, scan, items, progress);
            foreach (var skipped in wanted.Where(path => items.All(item => !string.Equals(item.RelativePath.Replace('\\', '/'), path, StringComparison.OrdinalIgnoreCase))))
            {
                result.Notes.Add($"'{skipped}' came with the pull but was not imported: it is not a module or note, or the project already holds the same text.");
            }

            return result;
        }

        private void DeleteFile(string path)
        {
            if (_fileSystem.File.Exists(path))
            {
                _fileSystem.File.Delete(path);
            }

            var binary = _fileSystem.Path.ChangeExtension(path, ".frx");
            if (_fileSystem.File.Exists(binary))
            {
                _fileSystem.File.Delete(binary);
            }
        }
    }
}
