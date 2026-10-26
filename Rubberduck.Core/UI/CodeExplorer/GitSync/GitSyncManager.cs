using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// Singleton that owns the git-sync sessions, one per project, and remembers the folder chosen for each project.
    /// </summary>
    public class GitSyncManager : IDisposable
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IProjectsProvider _projectsProvider;
        private readonly IParseManager _parseManager;
        private readonly IModuleNameFromFileExtractor _moduleNameExtractor;
        private readonly IFileSystem _fileSystem;

        private readonly object _lock = new object();
        private readonly Dictionary<string, GitSyncSession> _sessions = new Dictionary<string, GitSyncSession>();

        public GitSyncManager(
            IProjectsProvider projectsProvider,
            IParseManager parseManager,
            IModuleNameFromFileExtractor moduleNameExtractor,
            IFileSystem fileSystem)
        {
            _projectsProvider = projectsProvider;
            _parseManager = parseManager;
            _moduleNameExtractor = moduleNameExtractor;
            _fileSystem = fileSystem;

            _parseManager.StateChanged += OnParserStateChanged;
        }

        private string SettingsFile => _fileSystem.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "gitsync.json");

        /// <summary>
        /// The folder the project is currently synced with, or null.
        /// </summary>
        public string ActiveFolder(IVBProject project)
        {
            lock (_lock)
            {
                return _sessions.TryGetValue(project.ProjectId, out var session) ? session.Root : null;
            }
        }

        /// <summary>
        /// The folder last used for the project, even in an earlier session, or null.
        /// </summary>
        public string RememberedFolder(IVBProject project)
        {
            if (string.IsNullOrWhiteSpace(project.FileName))
            {
                return null;
            }

            return ReadRememberedFolders().TryGetValue(project.FileName, out var folder) && _fileSystem.Directory.Exists(folder)
                ? folder
                : null;
        }

        public GitSyncResult Start(IVBProject project, string folder)
        {
            Stop(project);

            var session = new GitSyncSession(project.ProjectId, folder, _projectsProvider, _parseManager, _moduleNameExtractor, _fileSystem);
            GitSyncResult result;
            try
            {
                result = session.Start();
            }
            catch
            {
                session.Dispose();
                throw;
            }

            lock (_lock)
            {
                _sessions[project.ProjectId] = session;
            }

            Remember(project, session.Root);
            return result;
        }

        public void Stop(IVBProject project)
        {
            GitSyncSession session;
            lock (_lock)
            {
                if (!_sessions.TryGetValue(project.ProjectId, out session))
                {
                    return;
                }

                _sessions.Remove(project.ProjectId);
            }

            session.Dispose();
        }

        private void OnParserStateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State != ParserState.Ready)
            {
                return;
            }

            List<GitSyncSession> sessions;
            lock (_lock)
            {
                sessions = _sessions.Values.ToList();
            }

            //Never block the parser thread: a session may be waiting for the parser while holding its lock.
            foreach (var session in sessions)
            {
                Task.Run(() =>
                {
                    try
                    {
                        session.ExportChangedModules();
                    }
                    catch (Exception exception)
                    {
                        Logger.Error(exception, "Git sync failed to export changed modules.");
                    }
                });
            }
        }

        private void Remember(IVBProject project, string folder)
        {
            if (string.IsNullOrWhiteSpace(project.FileName))
            {
                return;
            }

            try
            {
                var folders = ReadRememberedFolders();
                folders[project.FileName] = folder;
                _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(SettingsFile));
                _fileSystem.File.WriteAllText(SettingsFile, JsonConvert.SerializeObject(folders, Formatting.Indented));
            }
            catch (Exception exception)
            {
                Logger.Warn(exception, "Could not remember the git sync folder.");
            }
        }

        private Dictionary<string, string> ReadRememberedFolders()
        {
            try
            {
                if (_fileSystem.File.Exists(SettingsFile))
                {
                    var folders = JsonConvert.DeserializeObject<Dictionary<string, string>>(_fileSystem.File.ReadAllText(SettingsFile));
                    if (folders != null)
                    {
                        return new Dictionary<string, string>(folders, StringComparer.OrdinalIgnoreCase);
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Warn(exception, "Could not read the remembered git sync folders.");
            }

            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        public void Dispose()
        {
            _parseManager.StateChanged -= OnParserStateChanged;

            List<GitSyncSession> sessions;
            lock (_lock)
            {
                sessions = _sessions.Values.ToList();
                _sessions.Clear();
            }

            foreach (var session in sessions)
            {
                session.Dispose();
            }
        }
    }
}
