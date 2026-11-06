using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Abstractions;
using System.Linq;

namespace Rubberduck.Git
{
    public class GitInstallation
    {
        public GitInstallation(string path, Version version)
        {
            Path = path;
            Version = version;
        }

        /// <summary>The full path of git.exe.</summary>
        public string Path { get; }

        public Version Version { get; }
    }

    /// <summary>
    /// Finds a git that is good enough. The order is the path in the settings, git on PATH, the copy Rubberduck installed itself,
    /// and Program Files. A git that is older than <see cref="GitVersion.Minimum"/>, or does not answer, counts as missing.
    /// </summary>
    public class GitLocator
    {
        private readonly IFileSystem _fileSystem;
        private readonly Func<string, Version> _probe;
        private readonly string _configuredPath;
        private readonly string _pathVariable;
        private readonly string _ownFolder;
        private readonly string _programFiles;

        public GitLocator(IFileSystem fileSystem, Func<string, Version> probe, string configuredPath, string pathVariable, string ownFolder, string programFiles)
        {
            _fileSystem = fileSystem;
            _probe = probe;
            _configuredPath = configuredPath;
            _pathVariable = pathVariable;
            _ownFolder = ownFolder;
            _programFiles = programFiles;
        }

        /// <summary>Where Rubberduck puts the gits it installs: one folder per version, each with <c>cmd\git.exe</c> in it.</summary>
        public static string DefaultOwnFolder
            => System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Rubberduck", "git");

        public static string ExecutableIn(string versionFolder) => System.IO.Path.Combine(versionFolder, "cmd", "git.exe");

        /// <summary>The first usable git, or null if there is none.</summary>
        public GitInstallation Find()
        {
            foreach (var candidate in Candidates())
            {
                if (!_fileSystem.File.Exists(candidate))
                {
                    continue;
                }

                var version = _probe(candidate);
                if (GitVersion.IsSupported(version))
                {
                    return new GitInstallation(candidate, version);
                }
            }

            return null;
        }

        private IEnumerable<string> Candidates()
        {
            if (!string.IsNullOrWhiteSpace(_configuredPath))
            {
                yield return _configuredPath;
            }

            foreach (var folder in (_pathVariable ?? string.Empty).Split(';').Select(entry => entry.Trim().Trim('"')).Where(entry => entry.Length > 0))
            {
                yield return _fileSystem.Path.Combine(folder, "git.exe");
            }

            foreach (var own in OwnInstallations())
            {
                yield return own;
            }

            if (!string.IsNullOrEmpty(_programFiles))
            {
                yield return _fileSystem.Path.Combine(_programFiles, "Git", "cmd", "git.exe");
            }
        }

        //Newest first, so that an installer that left an older version behind does not win.
        private IEnumerable<string> OwnInstallations()
        {
            if (string.IsNullOrEmpty(_ownFolder) || !_fileSystem.Directory.Exists(_ownFolder))
            {
                return Enumerable.Empty<string>();
            }

            return _fileSystem.Directory.GetDirectories(_ownFolder)
                .Select(folder => new { Folder = folder, Version = Version.TryParse(_fileSystem.Path.GetFileName(folder), out var parsed) ? parsed : null })
                .Where(entry => entry.Version != null)
                .OrderByDescending(entry => entry.Version)
                .Select(entry => ExecutableIn(entry.Folder));
        }

        /// <summary>Runs <c>git --version</c>; null if it does not start, does not finish in time or says something else.</summary>
        public static Version RunVersion(string gitPath)
        {
            try
            {
                var info = new ProcessStartInfo(gitPath, "--version")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using (var process = Process.Start(info))
                {
                    var output = process.StandardOutput.ReadToEndAsync();
                    if (!process.WaitForExit(10000))
                    {
                        try { process.Kill(); } catch (InvalidOperationException) { }
                        return null;
                    }

                    return process.ExitCode == 0 ? GitVersion.Parse(output.Result) : null;
                }
            }
            catch (System.ComponentModel.Win32Exception)
            {
                return null;
            }
        }
    }
}
