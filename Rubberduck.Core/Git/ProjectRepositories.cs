using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using Newtonsoft.Json;

namespace Rubberduck.Git
{
    public enum ProjectRepositoryOutcome
    {
        /// <summary>The file has a repository already.</summary>
        Known,

        /// <summary>The file is where a known project used to be, and that is gone: it was moved or renamed.</summary>
        Moved,

        /// <summary>No repository yet; one was registered.</summary>
        New,

        /// <summary>A project of this name is known at another path that still exists: this file may be a copy of it. Nothing was registered.</summary>
        CopyOf
    }

    public class ProjectRepositoryResolution
    {
        public string Id { get; set; }
        public ProjectRepositoryOutcome Outcome { get; set; }

        /// <summary>For <see cref="ProjectRepositoryOutcome.CopyOf"/>: the project this may be a copy of, and where it is.</summary>
        public string CopyOfId { get; set; }

        public string CopyOfPath { get; set; }
    }

    /// <summary>
    /// Which repository belongs to which workbook. A repository is named by an id, not by the workbook's path, so that moving or renaming the file keeps its history.
    /// The map of id to last-known path and name lives in the user's AppData, next to the sync folders Git sync remembers.
    /// </summary>
    public class ProjectRepositories
    {
        private class Entry
        {
            public string Path { get; set; }
            public string Name { get; set; }
        }

        private readonly IFileSystem _fileSystem;
        private readonly string _mapFile;
        private readonly string _root;

        public ProjectRepositories(IFileSystem fileSystem)
            : this(fileSystem,
                System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "git-projects.json"),
                System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Rubberduck", "repos"))
        {
        }

        public ProjectRepositories(IFileSystem fileSystem, string mapFile, string root)
        {
            _fileSystem = fileSystem;
            _mapFile = mapFile;
            _root = root;
        }

        /// <summary>The folder that holds the repository (git directory and work tree) of the project with this id.</summary>
        public string RootOf(string id) => _fileSystem.Path.Combine(_root, id);

        /// <summary>Finds the repository of a saved workbook, moving it along if the file was moved, or registering a new one.</summary>
        public ProjectRepositoryResolution Resolve(string path, string name)
        {
            var map = Read();
            var full = _fileSystem.Path.GetFullPath(path);

            var known = map.FirstOrDefault(pair => SamePath(pair.Value.Path, full));
            if (known.Key != null)
            {
                return new ProjectRepositoryResolution { Id = known.Key, Outcome = ProjectRepositoryOutcome.Known };
            }

            //Same name, somewhere else. If the old file is gone the workbook was moved; if it is still there, this one is a copy.
            var namesake = map.Where(pair => string.Equals(pair.Value.Name, name, StringComparison.OrdinalIgnoreCase)).ToList();
            var moved = namesake.FirstOrDefault(pair => !_fileSystem.File.Exists(pair.Value.Path));
            if (moved.Key != null)
            {
                moved.Value.Path = full;
                Write(map);
                return new ProjectRepositoryResolution { Id = moved.Key, Outcome = ProjectRepositoryOutcome.Moved };
            }

            var original = namesake.FirstOrDefault();
            if (original.Key != null)
            {
                return new ProjectRepositoryResolution
                {
                    Outcome = ProjectRepositoryOutcome.CopyOf,
                    CopyOfId = original.Key,
                    CopyOfPath = original.Value.Path
                };
            }

            return new ProjectRepositoryResolution { Id = Register(map, full, name), Outcome = ProjectRepositoryOutcome.New };
        }

        /// <summary>
        /// Settles a <see cref="ProjectRepositoryOutcome.CopyOf"/>: the copy gets a repository of its own, either empty or starting from the original's history.
        /// </summary>
        public string Adopt(string path, string name, string copyOfId, bool continueHistory)
        {
            var map = Read();
            var id = Register(map, _fileSystem.Path.GetFullPath(path), name);
            if (continueHistory && map.ContainsKey(copyOfId) && _fileSystem.Directory.Exists(RootOf(copyOfId)))
            {
                CopyDirectory(RootOf(copyOfId), RootOf(id));
            }

            return id;
        }

        private string Register(Dictionary<string, Entry> map, string fullPath, string name)
        {
            var id = Guid.NewGuid().ToString("N");
            map[id] = new Entry { Path = fullPath, Name = name };
            Write(map);
            return id;
        }

        private bool SamePath(string a, string b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

        private void CopyDirectory(string source, string target)
        {
            _fileSystem.Directory.CreateDirectory(target);
            foreach (var file in _fileSystem.Directory.GetFiles(source))
            {
                _fileSystem.File.Copy(file, _fileSystem.Path.Combine(target, _fileSystem.Path.GetFileName(file)), true);
            }

            foreach (var directory in _fileSystem.Directory.GetDirectories(source))
            {
                CopyDirectory(directory, _fileSystem.Path.Combine(target, _fileSystem.Path.GetFileName(directory)));
            }
        }

        private Dictionary<string, Entry> Read()
        {
            try
            {
                if (_fileSystem.File.Exists(_mapFile))
                {
                    var map = JsonConvert.DeserializeObject<Dictionary<string, Entry>>(_fileSystem.File.ReadAllText(_mapFile));
                    if (map != null)
                    {
                        return map;
                    }
                }
            }
            catch (JsonException)
            {
                //An unreadable map would lose every history; keep the damaged file for a person to look at rather than overwrite it.
                _fileSystem.File.Move(_mapFile, _mapFile + ".damaged-" + DateTime.Now.ToString("yyyyMMdd-HHmmss"));
            }

            return new Dictionary<string, Entry>();
        }

        private void Write(Dictionary<string, Entry> map)
        {
            _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(_mapFile));
            _fileSystem.File.WriteAllText(_mapFile, JsonConvert.SerializeObject(map, Formatting.Indented));
        }
    }
}
