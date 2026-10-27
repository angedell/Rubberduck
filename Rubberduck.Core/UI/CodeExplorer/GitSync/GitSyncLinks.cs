using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using Newtonsoft.Json;
using NLog;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// Remembers, per user, which folder each workbook was last synced with. This is the only thing the sync keeps outside the repository.
    /// </summary>
    public class GitSyncLinks
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IFileSystem _fileSystem;

        public GitSyncLinks(IFileSystem fileSystem)
        {
            _fileSystem = fileSystem;
        }

        private string SettingsFile => _fileSystem.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "gitsync.json");

        /// <summary>The folder last used for the workbook, or null.</summary>
        public string Folder(string projectFileName)
        {
            if (string.IsNullOrWhiteSpace(projectFileName))
            {
                return null;
            }

            return Read().TryGetValue(projectFileName, out var folder) && _fileSystem.Directory.Exists(folder)
                ? folder
                : null;
        }

        public void Remember(string projectFileName, string folder)
        {
            if (string.IsNullOrWhiteSpace(projectFileName))
            {
                return;
            }

            try
            {
                var folders = Read();
                folders[projectFileName] = folder;
                _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(SettingsFile));
                _fileSystem.File.WriteAllText(SettingsFile, JsonConvert.SerializeObject(folders, Formatting.Indented));
                GitSyncIndenterSettings.Forget();
            }
            catch (Exception exception)
            {
                Logger.Warn(exception, "Could not remember the git sync folder.");
            }
        }

        private Dictionary<string, string> Read()
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
    }
}
