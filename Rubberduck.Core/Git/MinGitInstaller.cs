using System;
using System.IO.Abstractions;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Security.Cryptography;

namespace Rubberduck.Git
{
    /// <summary>A MinGit build Rubberduck knows the hash of. Raising the version means raising these three values together.</summary>
    public class MinGitRelease
    {
        public MinGitRelease(string version, string url, string sha256, long size)
        {
            Version = version;
            Url = url;
            Sha256 = sha256;
            Size = size;
        }

        public string Version { get; }
        public string Url { get; }

        /// <summary>Hex, as published with the release.</summary>
        public string Sha256 { get; }

        public long Size { get; }

        public static readonly MinGitRelease Win64 = new MinGitRelease("2.56.0.4",
            "https://github.com/git-for-windows/git/releases/download/v2.56.0.windows.4/MinGit-2.56.0.4-64-bit.zip",
            "e3bf78351fe777fba47ddee5950082f8a98bb0f02604fdc7ba3ee2cb68ffa877", 39846926);

        public static readonly MinGitRelease Win32 = new MinGitRelease("2.56.0.4",
            "https://github.com/git-for-windows/git/releases/download/v2.56.0.windows.4/MinGit-2.56.0.4-32-bit.zip",
            "6b3d71c7536d3c85baebaeb26d461e74c755895ccc1f4c817d0e289470419ccc", 39575289);

        public static MinGitRelease ForThisMachine => Environment.Is64BitOperatingSystem ? Win64 : Win32;
    }

    /// <summary>
    /// Puts MinGit, the portable build of Git for Windows, in Rubberduck's own folder. It needs no administrator rights and leaves
    /// the system and PATH alone. Nothing is kept unless the zip has the hash Rubberduck knows and the git inside answers.
    /// </summary>
    public class MinGitInstaller
    {
        private readonly IFileSystem _fileSystem;
        private readonly string _targetFolder;
        private readonly Func<string, byte[]> _download;
        private readonly Func<string, Version> _probe;

        public MinGitInstaller(IFileSystem fileSystem, string targetFolder, Func<string, byte[]> download, Func<string, Version> probe)
        {
            _fileSystem = fileSystem;
            _targetFolder = targetFolder;
            _download = download;
            _probe = probe;
        }

        /// <summary>Downloads the release and installs it.</summary>
        /// <returns>The path of git.exe.</returns>
        /// <exception cref="System.IO.IOException">The download, the hash check or the unpacking failed; the message says why. Nothing is left behind.</exception>
        public string Install(MinGitRelease release)
        {
            byte[] zip;
            try
            {
                zip = _download(release.Url);
            }
            catch (WebException exception)
            {
                throw new System.IO.IOException($"Git could not be downloaded from {release.Url}: {exception.Message}", exception);
            }

            return Unpack(zip, release);
        }

        /// <summary>Installs from a MinGit zip the user already has, which must be the release Rubberduck knows.</summary>
        public string InstallFromFile(string zipPath, MinGitRelease release)
            => Unpack(_fileSystem.File.ReadAllBytes(zipPath), release);

        private string Unpack(byte[] zip, MinGitRelease release)
        {
            if (!string.Equals(Hash(zip), release.Sha256, StringComparison.OrdinalIgnoreCase))
            {
                throw new System.IO.IOException($"The file is not MinGit {release.Version}: its SHA-256 hash is not the one Rubberduck expects. It was not installed.");
            }

            var target = _fileSystem.Path.Combine(_targetFolder, release.Version);
            var staging = target + ".installing-" + Guid.NewGuid().ToString("N");
            try
            {
                Extract(zip, staging);

                var git = GitLocator.ExecutableIn(staging);
                if (!_fileSystem.File.Exists(git))
                {
                    throw new System.IO.IOException("The download holds no cmd\\git.exe.");
                }

                if (!GitVersion.IsSupported(_probe(git)))
                {
                    throw new System.IO.IOException("The git that was unpacked does not run on this machine.");
                }

                if (_fileSystem.Directory.Exists(target))
                {
                    _fileSystem.Directory.Delete(target, true);
                }

                _fileSystem.Directory.Move(staging, target);
                return GitLocator.ExecutableIn(target);
            }
            catch (System.IO.InvalidDataException exception)
            {
                throw new System.IO.IOException("The download is not a zip file: " + exception.Message, exception);
            }
            finally
            {
                if (_fileSystem.Directory.Exists(staging))
                {
                    _fileSystem.Directory.Delete(staging, true);
                }
            }
        }

        private void Extract(byte[] zip, string folder)
        {
            var root = _fileSystem.Path.GetFullPath(folder).TrimEnd('\\') + "\\";
            _fileSystem.Directory.CreateDirectory(folder);
            using (var archive = new ZipArchive(new System.IO.MemoryStream(zip), ZipArchiveMode.Read))
            {
                foreach (var entry in archive.Entries)
                {
                    var path = _fileSystem.Path.GetFullPath(_fileSystem.Path.Combine(folder, entry.FullName.Replace('/', '\\')));

                    //An entry that climbs out of the folder would write somewhere it has no business.
                    if (!path.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new System.IO.IOException($"The zip holds an entry outside its folder ('{entry.FullName}'); it was not installed.");
                    }

                    if (entry.FullName.EndsWith("/", StringComparison.Ordinal))
                    {
                        _fileSystem.Directory.CreateDirectory(path);
                        continue;
                    }

                    _fileSystem.Directory.CreateDirectory(_fileSystem.Path.GetDirectoryName(path));
                    using (var source = entry.Open())
                    using (var destination = _fileSystem.File.Create(path))
                    {
                        source.CopyTo(destination);
                    }
                }
            }
        }

        private static string Hash(byte[] bytes)
        {
            using (var sha = SHA256.Create())
            {
                return string.Concat(sha.ComputeHash(bytes).Select(b => b.ToString("x2")));
            }
        }

        /// <summary>Downloads through the system's proxy, with the user's own credentials for it.</summary>
        public static byte[] Download(string url)
        {
            //GitHub insists on TLS 1.2, which .NET Framework does not offer by default.
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            using (var client = new WebClient())
            {
                client.Proxy = WebRequest.GetSystemWebProxy();
                client.Proxy.Credentials = CredentialCache.DefaultCredentials;
                return client.DownloadData(url);
            }
        }
    }
}
