using System;
using System.IO;
using System.IO.Abstractions;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using NUnit.Framework;
using Rubberduck.Git;

namespace RubberduckTests.Git
{
    [TestFixture]
    public class GitInstallTests
    {
        private string _folder;

        [SetUp]
        public void SetUp() => _folder = Path.Combine(Path.GetTempPath(), "rd-git-" + Guid.NewGuid());

        [TearDown]
        public void TearDown()
        {
            if (Directory.Exists(_folder))
            {
                Directory.Delete(_folder, true);
            }
        }

        private static byte[] Zip(params string[] entryNames)
        {
            using (var stream = new MemoryStream())
            {
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
                {
                    foreach (var name in entryNames)
                    {
                        using (var writer = new StreamWriter(archive.CreateEntry(name).Open(), Encoding.ASCII))
                        {
                            writer.Write("MZ");
                        }
                    }
                }

                return stream.ToArray();
            }
        }

        private static MinGitRelease ReleaseOf(byte[] zip)
        {
            using (var sha = SHA256.Create())
            {
                return new MinGitRelease("2.56.0.4", "https://example.test/MinGit.zip", string.Concat(sha.ComputeHash(zip).Select(b => b.ToString("x2"))), zip.Length);
            }
        }

        private MinGitInstaller Installer(Func<string, byte[]> download, Func<string, Version> probe = null)
            => new MinGitInstaller(new FileSystem(), _folder, download, probe ?? (path => new Version(2, 56, 0)));

        private string Create(string relative)
        {
            var path = Path.Combine(_folder, relative);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, "MZ");
            return path;
        }

        [TestCase("git version 2.43.0.windows.1", "2.43.0")]
        [TestCase("git version 2.31.1\r\n", "2.31.1")]
        [TestCase("git version 2.56", "2.56.0")]
        public void Parse_ReadsTheVersionLine(string output, string expected)
            => Assert.AreEqual(Version.Parse(expected), GitVersion.Parse(output));

        [TestCase("")]
        [TestCase(null)]
        [TestCase("usage: git [--version]")]
        public void Parse_ReturnsNullForAnythingElse(string output) => Assert.IsNull(GitVersion.Parse(output));

        [Test]
        public void IsSupported_RejectsOldAndMissingVersions()
        {
            Assert.IsFalse(GitVersion.IsSupported(null));
            Assert.IsFalse(GitVersion.IsSupported(new Version(2, 30, 9)));
            Assert.IsTrue(GitVersion.IsSupported(new Version(2, 31, 0)));
        }

        [Test]
        public void Find_PrefersTheConfiguredPathOverPath()
        {
            var configured = Create(@"configured\git.exe");
            var onPath = Create(@"bin\git.exe");

            var found = new GitLocator(new FileSystem(), p => new Version(2, 40, 0), configured, Path.GetDirectoryName(onPath), null, null).Find();

            Assert.AreEqual(configured, found.Path);
        }

        [Test]
        public void Find_SkipsAGitThatIsTooOld()
        {
            var old = Create(@"old\git.exe");
            var current = Create(@"new\git.exe");

            var found = new GitLocator(new FileSystem(), p => p == old ? new Version(2, 20, 0) : new Version(2, 40, 0), null,
                Path.GetDirectoryName(old) + ";" + Path.GetDirectoryName(current), null, null).Find();

            Assert.AreEqual(current, found.Path);
        }

        [Test]
        public void Find_UsesTheNewestOfRubberducksOwnInstallations()
        {
            Create(@"own\2.9.0.1\cmd\git.exe");
            var newest = Create(@"own\2.56.0.4\cmd\git.exe");

            var found = new GitLocator(new FileSystem(), p => new Version(2, 56, 0), null, null, Path.Combine(_folder, "own"), null).Find();

            Assert.AreEqual(newest, found.Path);
        }

        [Test]
        public void Find_ReturnsNullWhenThereIsNone()
            => Assert.IsNull(new GitLocator(new FileSystem(), p => new Version(2, 40, 0), null, @"C:\nowhere", Path.Combine(_folder, "own"), Path.Combine(_folder, "pf")).Find());

        [Test]
        public void Install_UnpacksIntoAFolderNamedAfterTheVersion()
        {
            var zip = Zip("cmd/git.exe", "etc/gitconfig");
            string asked = null;

            var git = Installer(url => { asked = url; return zip; }).Install(ReleaseOf(zip));

            Assert.AreEqual("https://example.test/MinGit.zip", asked);
            Assert.AreEqual(Path.Combine(_folder, "2.56.0.4", "cmd", "git.exe"), git);
            Assert.IsTrue(File.Exists(git));
            Assert.IsTrue(File.Exists(Path.Combine(_folder, "2.56.0.4", "etc", "gitconfig")));
        }

        [Test]
        public void Install_RefusesAZipWithAnotherHashAndLeavesNothing()
        {
            var zip = Zip("cmd/git.exe");
            var release = new MinGitRelease("2.56.0.4", "https://example.test/MinGit.zip", new string('0', 64), zip.Length);

            var exception = Assert.Throws<IOException>(() => Installer(url => zip).Install(release));

            StringAssert.Contains("hash", exception.Message);
            Assert.IsFalse(Directory.Exists(_folder) && Directory.GetFileSystemEntries(_folder).Any());
        }

        [Test]
        public void Install_RefusesAnEntryThatClimbsOutOfTheFolder()
        {
            var zip = Zip("cmd/git.exe", "../evil.txt");

            Assert.Throws<IOException>(() => Installer(url => zip).Install(ReleaseOf(zip)));

            Assert.IsFalse(File.Exists(Path.Combine(Path.GetDirectoryName(_folder), "evil.txt")));
            Assert.IsFalse(Directory.GetDirectories(_folder).Any());
        }

        [Test]
        public void Install_RefusesAGitThatDoesNotRunAndLeavesNothing()
        {
            var zip = Zip("cmd/git.exe");

            Assert.Throws<IOException>(() => Installer(url => zip, path => null).Install(ReleaseOf(zip)));

            Assert.IsFalse(Directory.GetDirectories(_folder).Any());
        }

        [Test]
        public void Install_ReportsAFailedDownload()
        {
            var exception = Assert.Throws<IOException>(() => Installer(url => throw new WebException("offline")).Install(MinGitRelease.Win64));

            StringAssert.Contains("offline", exception.Message);
        }

        [Test]
        public void InstallFromFile_AcceptsTheKnownZip()
        {
            var zip = Zip("cmd/git.exe");
            Directory.CreateDirectory(_folder);
            var file = Path.Combine(_folder, "MinGit.zip");
            File.WriteAllBytes(file, zip);

            var git = Installer(url => throw new InvalidOperationException("no download expected")).InstallFromFile(file, ReleaseOf(zip));

            Assert.IsTrue(File.Exists(git));
        }

        [Test]
        public void TheKnownReleasesPointAtGitHubAndHaveHashes()
        {
            foreach (var release in new[] { MinGitRelease.Win64, MinGitRelease.Win32 })
            {
                StringAssert.StartsWith("https://github.com/git-for-windows/git/releases/download/", release.Url);
                Assert.AreEqual(64, release.Sha256.Length);
            }
        }
    }
}
