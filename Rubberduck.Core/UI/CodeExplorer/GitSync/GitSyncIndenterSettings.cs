using System;
using System.Collections.Concurrent;
using System.IO.Abstractions;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// Lets the <c>rubberduck.json</c> of the folder a project is synced with decide how that project is indented,
    /// in the editor as much as in the sync, so that everyone working on the repository gets the same result.
    /// </summary>
    /// <remarks>
    /// Indenting a project asks for the settings once per module, so the answer is kept for a moment instead of
    /// looking for and parsing the file each time. A change to the file takes effect within <see cref="HoldFor"/>.
    /// </remarks>
    public static class GitSyncIndenterSettings
    {
        public static readonly TimeSpan HoldFor = TimeSpan.FromSeconds(2);

        private class Held
        {
            public DateTime Since;
            //Null when the project has no remembered folder or the folder has no config.
            public string ConfigPath;
            public IIndenterSettings Settings;
        }

        private static readonly ConcurrentDictionary<string, Held> Cache = new ConcurrentDictionary<string, Held>(StringComparer.OrdinalIgnoreCase);

        /// <param name="projectFileName">The file name of the workbook, which is what the sync folder is remembered by. May be null.</param>
        /// <param name="userSettings">The user's own settings, used wherever the file has no say.</param>
        /// <param name="now">For tests.</param>
        public static IIndenterSettings For(IFileSystem fileSystem, string projectFileName, IIndenterSettings userSettings, DateTime? now = null)
        {
            if (string.IsNullOrWhiteSpace(projectFileName))
            {
                return userSettings;
            }

            var time = now ?? DateTime.UtcNow;
            if (Cache.TryGetValue(projectFileName, out var held) && time - held.Since < HoldFor)
            {
                //No config means the user's own settings, which may have changed since, so they are not held.
                return held.ConfigPath == null ? userSettings : held.Settings;
            }

            var folder = new GitSyncLinks(fileSystem).Folder(projectFileName);
            var config = folder == null ? null : new ProjectConfigReader(fileSystem).Read(folder, userSettings);
            Cache[projectFileName] = new Held { Since = time, ConfigPath = config?.Path, Settings = config?.Indenter };

            return config?.Indenter ?? userSettings;
        }

        /// <summary>Forgets what is held, e.g. after the folder of a project changed.</summary>
        public static void Forget() => Cache.Clear();
    }
}
