using System;
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
        GitSyncResult Take(IVBProject project, string tree);
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

        public GitSyncResult Take(IVBProject project, string tree)
        {
            var scan = _sync.Scan(project, tree);

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
                return _sync.Apply(project, scan, exports, null, scratch);
            }
            finally
            {
                if (_fileSystem.Directory.Exists(scratch))
                {
                    _fileSystem.Directory.Delete(scratch, true);
                }
            }
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
