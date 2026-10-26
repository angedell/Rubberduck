using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using Newtonsoft.Json;
using NLog;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// What a module and its file looked like right after the last time they were synced.
    /// Comparing the current state of each side with this tells which side changed, i.e. the direction to sync in.
    /// </summary>
    public class ModuleSnapshot
    {
        /// <summary>Hash of the module's content in the project.</summary>
        public string ModuleHash { get; set; }

        /// <summary>Hash of the file's bytes, or null if the module has no file because it is ignored.</summary>
        public string FileHash { get; set; }

        /// <summary>The file's path relative to the sync root.</summary>
        public string RelativePath { get; set; }

        /// <summary>The @Folder annotation of the module at that time.</summary>
        public string Folder { get; set; }

        /// <summary>
        /// When a change of the module was first noticed. The VBE does not expose modification times,
        /// so this is the best available stand-in to settle conflicts. Null while the module is unchanged.
        /// </summary>
        public DateTime? ModuleChangedUtc { get; set; }
    }

    public class ProjectManifest
    {
        /// <summary>The snapshots only apply as long as the project is synced with this folder.</summary>
        public string Root { get; set; }
        public Dictionary<string, ModuleSnapshot> Modules { get; set; } = new Dictionary<string, ModuleSnapshot>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Persists the snapshots between sessions, per workbook, in the user's roaming profile (not in the synced folder, so git never sees it).
    /// </summary>
    public class GitSyncManifestStore
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IFileSystem _fileSystem;
        private readonly string _path;
        private readonly object _lock = new object();

        public GitSyncManifestStore(IFileSystem fileSystem, string path = null)
        {
            _fileSystem = fileSystem;
            _path = path ?? fileSystem.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "gitsync-state.json");
        }

        public ProjectManifest Load(string key, string root)
        {
            if (key == null)
            {
                return new ProjectManifest { Root = root };
            }

            lock (_lock)
            {
                var manifest = ReadAll().TryGetValue(key, out var found) ? found : null;
                if (manifest == null || !string.Equals(manifest.Root, root, StringComparison.OrdinalIgnoreCase))
                {
                    return new ProjectManifest { Root = root };
                }

                manifest.Modules = new Dictionary<string, ModuleSnapshot>(manifest.Modules ?? new Dictionary<string, ModuleSnapshot>(), StringComparer.OrdinalIgnoreCase);
                return manifest;
            }
        }

        public void Save(string key, ProjectManifest manifest)
        {
            if (key == null)
            {
                return;
            }

            lock (_lock)
            {
                try
                {
                    var all = ReadAll();
                    all[key] = manifest;
                    _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(_path));
                    _fileSystem.File.WriteAllText(_path, JsonConvert.SerializeObject(all, Formatting.Indented));
                }
                catch (Exception exception)
                {
                    Logger.Warn(exception, "Could not save the git sync state.");
                }
            }
        }

        private Dictionary<string, ProjectManifest> ReadAll()
        {
            try
            {
                if (_fileSystem.File.Exists(_path))
                {
                    var all = JsonConvert.DeserializeObject<Dictionary<string, ProjectManifest>>(_fileSystem.File.ReadAllText(_path));
                    if (all != null)
                    {
                        return new Dictionary<string, ProjectManifest>(all, StringComparer.OrdinalIgnoreCase);
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Warn(exception, "Could not read the git sync state; starting without snapshots.");
            }

            return new Dictionary<string, ProjectManifest>(StringComparer.OrdinalIgnoreCase);
        }
    }

    public enum SyncDirection
    {
        None,
        /// <summary>Write the module to the file.</summary>
        Export,
        /// <summary>Load the file into the module.</summary>
        Import,
        /// <summary>Both sides already agree; only the snapshot needs refreshing.</summary>
        Record,
        /// <summary>Both sides exist and differ, and nothing tells which one is newer: the user has to decide.</summary>
        Ask
    }

    public class SyncDecision
    {
        public SyncDecision(SyncDirection direction, bool isConflict = false)
        {
            Direction = direction;
            IsConflict = isConflict;
        }

        public SyncDirection Direction { get; }

        /// <summary>Both sides changed since the last sync; the winner was picked by modification time.</summary>
        public bool IsConflict { get; }
    }

    /// <summary>
    /// Decides in which direction a module and its file need to be synced. Pure, so it can be tested without a VBE.
    /// </summary>
    public static class SyncDirectionResolver
    {
        /// <summary>
        /// Decides by the stamps both sides carry, which works on any machine. A side is dirty if its code was edited since its stamp was written;
        /// of two clean sides, the one with the higher revision is the newer one.
        /// </summary>
        public static SyncDecision DecideByStamps(
            StampState module,
            StampState file,
            DateTime fileWrittenUtc,
            DateTime moduleChangedUtc)
        {
            SyncDecision ConflictByTime() => fileWrittenUtc > moduleChangedUtc
                ? new SyncDecision(SyncDirection.Import, true)
                : new SyncDecision(SyncDirection.Export, true);

            if (!module.Dirty && !file.Dirty)
            {
                if (module.Revision == file.Revision)
                {
                    if (module.Hash == file.Hash)
                    {
                        return new SyncDecision(SyncDirection.None);
                    }

                    //Same revision but different content: two machines went on from the same revision.
                    //The stamps say when each was last modified, which beats the local clocks.
                    if (module.Utc.HasValue && file.Utc.HasValue && module.Utc != file.Utc)
                    {
                        return file.Utc > module.Utc
                            ? new SyncDecision(SyncDirection.Import, true)
                            : new SyncDecision(SyncDirection.Export, true);
                    }

                    return ConflictByTime();
                }

                return module.Revision < file.Revision
                    ? new SyncDecision(SyncDirection.Import)
                    : new SyncDecision(SyncDirection.Export);
            }

            if (module.Dirty && !file.Dirty)
            {
                return file.Revision > module.Revision ? ConflictByTime() : new SyncDecision(SyncDirection.Export);
            }

            if (!module.Dirty)
            {
                return module.Revision > file.Revision ? ConflictByTime() : new SyncDecision(SyncDirection.Import);
            }

            return module.BodyHash == file.BodyHash ? new SyncDecision(SyncDirection.Record) : ConflictByTime();
        }

        /// <param name="snapshot">The state after the last sync, or null if the module was never synced.</param>
        /// <param name="moduleHash">The module's current hash, or null if the project has no such module.</param>
        /// <param name="fileHash">The file's current hash, or null if there is no file.</param>
        /// <param name="contentsEqual">Whether module and file hold the same content, if both exist.</param>
        /// <param name="fileWrittenUtc">The file's last write time, if it exists.</param>
        /// <param name="now">The current time, used as the module's change time if it was not recorded yet.</param>
        public static SyncDecision Decide(
            ModuleSnapshot snapshot,
            string moduleHash,
            string fileHash,
            Func<bool> contentsEqual,
            DateTime fileWrittenUtc,
            DateTime now)
        {
            if (moduleHash == null && fileHash == null)
            {
                return new SyncDecision(SyncDirection.None);
            }

            //Only one side exists: bring the other one into being.
            if (moduleHash == null)
            {
                return new SyncDecision(SyncDirection.Import);
            }

            if (fileHash == null)
            {
                return new SyncDecision(SyncDirection.Export);
            }

            //Never synced before, and both exist: no history to go by, unless they agree.
            if (snapshot == null)
            {
                return contentsEqual() ? new SyncDecision(SyncDirection.Record) : new SyncDecision(SyncDirection.Ask);
            }

            var moduleChanged = moduleHash != snapshot.ModuleHash;
            var fileChanged = fileHash != snapshot.FileHash;

            if (!moduleChanged && !fileChanged)
            {
                return new SyncDecision(SyncDirection.None);
            }

            if (moduleChanged && !fileChanged)
            {
                return new SyncDecision(SyncDirection.Export);
            }

            if (!moduleChanged)
            {
                return new SyncDecision(SyncDirection.Import);
            }

            //Both changed. If they changed to the same thing there is no conflict.
            if (contentsEqual())
            {
                return new SyncDecision(SyncDirection.Record);
            }

            var moduleChangedUtc = snapshot.ModuleChangedUtc ?? now;
            return fileWrittenUtc > moduleChangedUtc
                ? new SyncDecision(SyncDirection.Import, true)
                : new SyncDecision(SyncDirection.Export, true);
        }
    }
}
