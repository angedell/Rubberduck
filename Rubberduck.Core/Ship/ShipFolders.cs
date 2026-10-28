using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace Rubberduck.Ship
{
    /// <summary>What a folder setting needs to know about the machine, apart from files: variables and the places Windows keeps for a user.</summary>
    public interface IShipEnvironment
    {
        string GetEnvironmentVariable(string name);

        /// <summary>The path of a Windows folder by the name used in the settings (Documents, Desktop, Downloads, Music, Pictures, Videos), or null if there is none.</summary>
        string KnownFolder(string name);
    }

    public class ShipEnvironment : IShipEnvironment
    {
        private static readonly Guid DownloadsFolder = new Guid("374DE290-123F-4565-9164-39C4925E467B");

        public string GetEnvironmentVariable(string name) => Environment.GetEnvironmentVariable(name);

        public string KnownFolder(string name)
        {
            switch ((name ?? string.Empty).ToLowerInvariant())
            {
                case "documents": return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                case "desktop": return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                case "music": return Environment.GetFolderPath(Environment.SpecialFolder.MyMusic);
                case "pictures": return Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
                case "videos": return Environment.GetFolderPath(Environment.SpecialFolder.MyVideos);
                case "downloads": return SHGetKnownFolderPath(DownloadsFolder, 0, IntPtr.Zero, out var path) == 0 ? Take(path) : null;
                default: return null;
            }
        }

        private static string Take(IntPtr path)
        {
            try
            {
                return Marshal.PtrToStringUni(path);
            }
            finally
            {
                Marshal.FreeCoTaskMem(path);
            }
        }

        [DllImport("shell32.dll")]
        private static extern int SHGetKnownFolderPath([MarshalAs(UnmanagedType.LPStruct)] Guid rfid, uint flags, IntPtr token, out IntPtr path);
    }

    /// <summary>The folder a setting names on this machine, or why it names none.</summary>
    public class ShipFolderResult
    {
        /// <summary>The full path; null when nothing is set or the setting cannot be resolved here.</summary>
        public string Path { get; set; }

        /// <summary>Why <see cref="Path"/> is null when something was set; null when nothing was.</summary>
        public string Problem { get; set; }
    }

    /// <summary>
    /// Turns the output folder as written in the settings into a folder of this machine: relative to the document's folder, with <c>../</c>, with <c>%ENV%</c> variables,
    /// by the name of a Windows folder (Documents, Desktop, Downloads...), OneDrive, or the main Dropbox folder.
    /// </summary>
    public class ShipFolders
    {
        public static readonly string[] KnownFolders = { "Documents", "Desktop", "Downloads", "Music", "Pictures", "Videos" };
        private static readonly string[] OneDriveVariables = { "OneDrive", "OneDriveCommercial", "OneDriveConsumer" };
        private static readonly Regex Variable = new Regex("%([^%\\r\\n]+)%");

        /// <summary>The environment variables an output folder can use: the ones that name a folder of the user's.</summary>
        public static readonly HashSet<string> AllowedVariables = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "USERPROFILE", "HOMEDRIVE", "HOMEPATH", "APPDATA", "LOCALAPPDATA", "PROGRAMDATA", "PUBLIC", "TEMP", "TMP", "SYSTEMDRIVE",
            "OneDrive", "OneDriveCommercial", "OneDriveConsumer"
        };

        private readonly IFileSystem _fileSystem;
        private readonly IShipEnvironment _environment;

        public ShipFolders(IFileSystem fileSystem, IShipEnvironment environment)
        {
            _fileSystem = fileSystem;
            _environment = environment;
        }

        /// <param name="written">The folder as it is in the settings.</param>
        /// <param name="documentFolder">The folder of the saved document, for a relative folder.</param>
        public ShipFolderResult Resolve(string written, string documentFolder)
        {
            if (string.IsNullOrWhiteSpace(written))
            {
                return new ShipFolderResult();
            }

            var missing = new List<string>();
            var refused = new List<string>();
            var text = Variable.Replace(written.Trim(), match =>
            {
                //The settings can be written by an agent, and the folder a setting resolves to is shown back: only variables that hold a folder are expanded, never a secret.
                if (!AllowedVariables.Contains(match.Groups[1].Value))
                {
                    refused.Add(match.Groups[1].Value);
                    return match.Value;
                }

                var value = _environment.GetEnvironmentVariable(match.Groups[1].Value);
                if (string.IsNullOrEmpty(value))
                {
                    missing.Add(match.Groups[1].Value);
                    return match.Value;
                }

                return value;
            });

            if (refused.Count > 0)
            {
                return Failed($"{string.Join(", ", refused.Select(r => "%" + r + "%"))} cannot be used in an output folder. The variables that can: {string.Join(", ", AllowedVariables.Select(v => "%" + v + "%"))}.");
            }

            if (missing.Count > 0)
            {
                return Failed($"The environment variable {string.Join(", ", missing.Select(m => "%" + m + "%"))} is not set on this machine.");
            }

            text = text.Replace('/', '\\');
            var segments = text.Split(new[] { '\\' }, 2);
            var first = segments[0];
            var rest = segments.Length > 1 ? segments[1] : string.Empty;

            var named = Named(first, out var problem);
            if (problem != null)
            {
                return Failed(problem);
            }

            if (named != null)
            {
                return Done(_fileSystem.Path.Combine(named, rest));
            }

            if (_fileSystem.Path.IsPathRooted(text))
            {
                return Done(text);
            }

            if (string.IsNullOrWhiteSpace(documentFolder))
            {
                return Failed($"'{written}' is relative to the folder of the document, which has not been saved yet.");
            }

            return Done(_fileSystem.Path.Combine(documentFolder, text));
        }

        /// <summary>The folder a name stands for: null if it is not one of the names, with <paramref name="problem"/> set if it is one that this machine lacks.</summary>
        private string Named(string name, out string problem)
        {
            problem = null;
            var known = KnownFolders.FirstOrDefault(k => k.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (known != null)
            {
                var path = _environment.KnownFolder(known);
                problem = string.IsNullOrEmpty(path) ? $"Windows has no {known} folder for this user." : null;
                return string.IsNullOrEmpty(path) ? null : path;
            }

            if (name.Equals("OneDrive", StringComparison.OrdinalIgnoreCase))
            {
                var path = OneDriveVariables.Select(_environment.GetEnvironmentVariable).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
                problem = path == null ? "OneDrive is not set up for this user on this machine." : null;
                return path;
            }

            if (name.Equals("Dropbox", StringComparison.OrdinalIgnoreCase))
            {
                var path = DropboxFolder();
                problem = path == null ? "The main Dropbox folder could not be found: Dropbox's info.json is missing or unreadable." : null;
                return path;
            }

            return null;
        }

        private string DropboxFolder()
        {
            foreach (var variable in new[] { "LOCALAPPDATA", "APPDATA" })
            {
                var root = _environment.GetEnvironmentVariable(variable);
                if (string.IsNullOrEmpty(root))
                {
                    continue;
                }

                try
                {
                    var info = _fileSystem.Path.Combine(root, "Dropbox", "info.json");
                    if (!_fileSystem.File.Exists(info))
                    {
                        continue;
                    }

                    var json = JObject.Parse(_fileSystem.File.ReadAllText(info));
                    var account = json["personal"] ?? json.Properties().Select(p => p.Value).FirstOrDefault(v => v is JObject);
                    var path = (string)account?["path"];
                    if (!string.IsNullOrWhiteSpace(path))
                    {
                        return path;
                    }
                }
                catch (Exception exception) when (exception is Newtonsoft.Json.JsonException || exception is System.IO.IOException || exception is UnauthorizedAccessException)
                {
                    //Try the next place.
                }
            }

            return null;
        }

        private ShipFolderResult Done(string path)
        {
            try
            {
                return new ShipFolderResult { Path = _fileSystem.Path.GetFullPath(path) };
            }
            catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException || exception is System.IO.PathTooLongException)
            {
                return Failed($"'{path}' is not a usable folder: {exception.Message}");
            }
        }

        private static ShipFolderResult Failed(string problem) => new ShipFolderResult { Problem = problem };

        /// <summary>Checks that the folder is there and that a file can be made in it. Throws <see cref="ShipRefusedException"/> with reason folderMissing or folderNotWritable.</summary>
        /// <param name="probeWrite">False to check only that the folder is there: a preview must not write into a folder it was merely pointed at.</param>
        public void Validate(string folder, bool probeWrite = true)
        {
            if (string.IsNullOrWhiteSpace(folder) || !_fileSystem.Directory.Exists(folder))
            {
                throw new ShipRefusedException("folderMissing", $"The folder '{folder}' does not exist. Create it, or choose another one.");
            }

            if (!probeWrite)
            {
                return;
            }

            var probe = _fileSystem.Path.Combine(folder, ".rubberduck-ship-" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                using (_fileSystem.File.Create(probe))
                {
                }

                _fileSystem.File.Delete(probe);
            }
            catch (Exception exception) when (exception is UnauthorizedAccessException || exception is System.IO.IOException || exception is System.Security.SecurityException)
            {
                throw new ShipRefusedException("folderNotWritable", $"Rubberduck cannot write to the folder '{folder}': {exception.Message}");
            }
        }
    }
}
