using System;
using System.IO.Abstractions;
using System.IO.Compression;
using System.Net;

namespace Rubberduck.Api
{
    /// <summary>Makes sure Sysinternals ProcDump is there for the dump routes, and fetches it when the user agrees.</summary>
    public interface IProcDumpInstaller
    {
        /// <summary>Where ProcDump was found, or null.</summary>
        string Find();

        /// <summary>The folder <see cref="Install"/> puts ProcDump in.</summary>
        string TargetFolder { get; }

        /// <summary>Downloads ProcDump from Microsoft and unpacks procdump.exe into <see cref="TargetFolder"/>.</summary>
        /// <returns>The path of procdump.exe.</returns>
        /// <exception cref="System.IO.IOException">The download or unpacking failed; the message says why.</exception>
        string Install();
    }

    public class ProcDumpInstaller : IProcDumpInstaller
    {
        public const string DownloadUrl = "https://download.sysinternals.com/files/Procdump.zip";

        private readonly Func<string> _find;
        private readonly IFileSystem _fileSystem;
        private readonly Func<string, byte[]> _download;

        public ProcDumpInstaller(DumpApi dump, IFileSystem fileSystem)
            : this(dump.FindTool, fileSystem, System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Rubberduck", "Tools"), Download)
        {
        }

        public ProcDumpInstaller(Func<string> find, IFileSystem fileSystem, string targetFolder, Func<string, byte[]> download)
        {
            _find = find;
            _fileSystem = fileSystem;
            TargetFolder = targetFolder;
            _download = download;
        }

        public string Find() => _find();

        public string TargetFolder { get; }

        public string Install()
        {
            byte[] zip;
            try
            {
                zip = _download(DownloadUrl);
            }
            catch (WebException exception)
            {
                throw new System.IO.IOException($"ProcDump could not be downloaded from {DownloadUrl}: {exception.Message}", exception);
            }

            try
            {
                using (var archive = new ZipArchive(new System.IO.MemoryStream(zip), ZipArchiveMode.Read))
                {
                    var entry = archive.GetEntry("procdump.exe") ?? throw new System.IO.IOException("The download holds no procdump.exe.");
                    _fileSystem.Directory.CreateDirectory(TargetFolder);
                    var target = _fileSystem.Path.Combine(TargetFolder, "procdump.exe");
                    using (var source = entry.Open())
                    using (var destination = _fileSystem.File.Create(target))
                    {
                        source.CopyTo(destination);
                    }

                    return target;
                }
            }
            catch (System.IO.InvalidDataException exception)
            {
                throw new System.IO.IOException("The download is not a zip file: " + exception.Message, exception);
            }
        }

        private static byte[] Download(string url)
        {
            //The server insists on TLS 1.2, which .NET Framework does not offer by default.
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            using (var client = new WebClient())
            {
                return client.DownloadData(url);
            }
        }
    }
}
