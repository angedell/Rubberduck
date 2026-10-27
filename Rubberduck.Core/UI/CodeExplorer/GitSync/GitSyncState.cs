using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using Newtonsoft.Json;
using NLog;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// What a module and its file looked like when they were last synced. Comparing the current state of each side with this
    /// tells which side changed since, i.e. the direction to sync in.
    /// </summary>
    public class ModuleBaseline
    {
        /// <summary>Normalized signature of the module right after the sync.</summary>
        public string Module { get; set; }

        /// <summary>Normalized signature of the file right after the sync.</summary>
        public string File { get; set; }

        /// <summary>The file's path relative to the sync folder, with forward slashes.</summary>
        public string Path { get; set; }

        /// <summary>The @Folder annotation of the module at that time.</summary>
        public string Folder { get; set; }
    }

    /// <summary>
    /// The contents of <c>rubberduck.sync.json</c>. It is meant to be committed, so that everyone working on the repository starts from the same history.
    /// </summary>
    public class GitSyncState
    {
        public const string FileName = "rubberduck.sync.json";

        public int Version { get; set; } = 1;

        public SortedDictionary<string, ModuleBaseline> Modules { get; set; } = new SortedDictionary<string, ModuleBaseline>(StringComparer.OrdinalIgnoreCase);
    }

    public class GitSyncStateStore
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IFileSystem _fileSystem;

        public GitSyncStateStore(IFileSystem fileSystem)
        {
            _fileSystem = fileSystem;
        }

        /// <summary>The state file lives next to rubberduck.json, or in the sync folder if there is no config.</summary>
        public string PathFor(string configPath, string syncFolder)
        {
            var folder = configPath == null ? syncFolder : _fileSystem.Path.GetDirectoryName(configPath);
            return _fileSystem.Path.Combine(folder, GitSyncState.FileName);
        }

        public GitSyncState Load(string path)
        {
            try
            {
                if (_fileSystem.File.Exists(path))
                {
                    var state = JsonConvert.DeserializeObject<GitSyncState>(_fileSystem.File.ReadAllText(path));
                    if (state?.Modules != null)
                    {
                        state.Modules = new SortedDictionary<string, ModuleBaseline>(state.Modules, StringComparer.OrdinalIgnoreCase);
                        return state;
                    }
                }
            }
            catch (JsonException exception)
            {
                Logger.Warn(exception, "Could not read the git sync state; starting without history.");
            }

            return new GitSyncState();
        }

        public void Save(string path, GitSyncState state)
        {
            _fileSystem.File.WriteAllText(path, JsonConvert.SerializeObject(state, Formatting.Indented) + Environment.NewLine);
        }
    }
}
