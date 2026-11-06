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

        private static (Rubberduck.UI.Settings.GeneralSettingsViewModel viewModel, Moq.Mock<IGitProvisioner> git, Moq.Mock<Rubberduck.Interaction.IMessageBox> box) ViewModel(bool installed, bool agrees)
        {
            var git = new Moq.Mock<IGitProvisioner>();
            git.Setup(g => g.Find(Moq.It.IsAny<string>())).Returns(installed ? new GitInstallation(@"C:\Git\cmd\git.exe", new Version(2, 56, 0)) : null);
            git.SetupGet(g => g.TargetFolder).Returns(@"C:\Rubberduck\git");
            git.SetupGet(g => g.DownloadUrl).Returns("https://example.test/MinGit.zip");
            git.SetupGet(g => g.DownloadSize).Returns(40000000);
            var box = new Moq.Mock<Rubberduck.Interaction.IMessageBox>();
            box.Setup(b => b.ConfirmYesNo(Moq.It.IsAny<string>(), Moq.It.IsAny<string>(), Moq.It.IsAny<bool>())).Returns(agrees);
            var types = new Moq.Mock<Rubberduck.Settings.IExperimentalTypesProvider>();
            types.SetupGet(t => t.ExperimentalTypes).Returns(new Type[0]);
            var config = new Rubberduck.Settings.Configuration(new Rubberduck.Settings.UserSettings(new Rubberduck.Settings.GeneralSettings(), new Rubberduck.Settings.HotkeySettings(), null, null, null, null, null, null));
            var viewModel = new Rubberduck.UI.Settings.GeneralSettingsViewModel(config, new Moq.Mock<Rubberduck.Common.IOperatingSystem>().Object, box.Object,
                new Moq.Mock<Rubberduck.VBEditor.VbeRuntime.Settings.IVbeSettings>().Object, types.Object, null, null, null, null, null, git.Object);
            viewModel.GitClientEnabled = true;
            return (viewModel, git, box);
        }

        [Test]
        public void Ticking_WhenGitIsThere_AsksNothing()
        {
            var (viewModel, git, box) = ViewModel(installed: true, agrees: true);

            viewModel.EnsureGitCommand.Execute(null);

            Assert.IsTrue(viewModel.GitClientEnabled);
            box.Verify(b => b.ConfirmYesNo(Moq.It.IsAny<string>(), Moq.It.IsAny<string>(), Moq.It.IsAny<bool>()), Moq.Times.Never);
            git.Verify(g => g.Install(), Moq.Times.Never);
        }

        [Test]
        public void Ticking_WhenGitIsMissing_NamesTheDownloadAndInstallsOnYes()
        {
            var (viewModel, git, box) = ViewModel(installed: false, agrees: true);

            viewModel.EnsureGitCommand.Execute(null);

            Assert.IsTrue(viewModel.GitClientEnabled);
            git.Verify(g => g.Install(), Moq.Times.Once);
            box.Verify(b => b.ConfirmYesNo(Moq.It.Is<string>(text => text.Contains("https://example.test/MinGit.zip") && text.Contains(@"C:\Rubberduck\git") && text.Contains("About 38 MB")),
                Moq.It.IsAny<string>(), Moq.It.IsAny<bool>()), Moq.Times.Once);
        }

        [Test]
        public void Ticking_WhenGitIsMissing_AndTheUserSaysNo_DownloadsNothingAndTakesTheTickBack()
        {
            var (viewModel, git, _) = ViewModel(installed: false, agrees: false);

            viewModel.EnsureGitCommand.Execute(null);

            Assert.IsFalse(viewModel.GitClientEnabled);
            git.Verify(g => g.Install(), Moq.Times.Never);
        }

        [Test]
        public void Ticking_WhenTheInstallFails_TellsTheUserAndTakesTheTickBack()
        {
            var (viewModel, git, box) = ViewModel(installed: false, agrees: true);
            git.Setup(g => g.Install()).Throws(new IOException("hash mismatch"));

            viewModel.EnsureGitCommand.Execute(null);

            Assert.IsFalse(viewModel.GitClientEnabled);
            box.Verify(b => b.NotifyWarn(Moq.It.Is<string>(text => text.StartsWith("hash mismatch")), Moq.It.IsAny<string>()), Moq.Times.Once);
        }
    }
}
