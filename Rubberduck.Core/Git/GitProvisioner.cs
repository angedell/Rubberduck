using System;
using System.IO.Abstractions;

namespace Rubberduck.Git
{
    /// <summary>Makes sure git is there for the git client, and fetches it when the user agrees.</summary>
    public interface IGitProvisioner
    {
        /// <summary>The first usable git: the configured path first, then PATH, Rubberduck's own copy, Program Files. Null if there is none.</summary>
        GitInstallation Find(string configuredPath);

        /// <summary>The folder <see cref="Install"/> puts git in.</summary>
        string TargetFolder { get; }

        string DownloadUrl { get; }

        /// <summary>The size of the download, in bytes.</summary>
        long DownloadSize { get; }

        /// <summary>Downloads MinGit, checks it against the hash Rubberduck knows and unpacks it into <see cref="TargetFolder"/>.</summary>
        /// <returns>The path of git.exe.</returns>
        /// <exception cref="System.IO.IOException">The download, the check or the unpacking failed; the message says why.</exception>
        string Install();
    }

    public class GitProvisioner : IGitProvisioner
    {
        private readonly IFileSystem _fileSystem;
        private readonly MinGitRelease _release = MinGitRelease.ForThisMachine;

        public GitProvisioner(IFileSystem fileSystem)
        {
            _fileSystem = fileSystem;
        }

        public string TargetFolder => GitLocator.DefaultOwnFolder;

        public string DownloadUrl => _release.Url;

        public long DownloadSize => _release.Size;

        public GitInstallation Find(string configuredPath)
            => new GitLocator(_fileSystem, GitLocator.RunVersion, configuredPath, Environment.GetEnvironmentVariable("PATH"), TargetFolder,
                Environment.GetEnvironmentVariable("ProgramFiles")).Find();

        public string Install()
            => new MinGitInstaller(_fileSystem, TargetFolder, MinGitInstaller.Download, GitLocator.RunVersion).Install(_release);
    }
}
