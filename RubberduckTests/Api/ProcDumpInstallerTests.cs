using System;
using System.IO;
using System.IO.Abstractions;
using System.IO.Compression;
using System.Net;
using System.Text;
using Moq;
using NUnit.Framework;
using Rubberduck.Api;
using Rubberduck.Interaction;
using Rubberduck.Common;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;
using GeneralSettings = Rubberduck.Settings.GeneralSettings;

namespace RubberduckTests.Api
{
    [TestFixture]
    public class ProcDumpInstallerTests
    {
        private static byte[] Zip(string entryName)
        {
            using (var stream = new MemoryStream())
            {
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
                using (var writer = new StreamWriter(archive.CreateEntry(entryName).Open(), Encoding.ASCII))
                {
                    writer.Write("MZ");
                }

                return stream.ToArray();
            }
        }

        private string _folder;

        [SetUp]
        public void SetUp() => _folder = Path.Combine(Path.GetTempPath(), "rd-procdump-" + Guid.NewGuid());

        [TearDown]
        public void TearDown()
        {
            if (Directory.Exists(_folder))
            {
                Directory.Delete(_folder, true);
            }
        }

        private ProcDumpInstaller Installer(Func<string, byte[]> download)
            => new ProcDumpInstaller(() => null, new FileSystem(), _folder, download);

        [Test]
        public void Install_UnpacksProcDumpIntoTheToolsFolder()
        {
            string asked = null;
            var installer = Installer(url => { asked = url; return Zip("procdump.exe"); });

            var path = installer.Install();

            Assert.AreEqual(ProcDumpInstaller.DownloadUrl, asked);
            Assert.AreEqual(Path.Combine(_folder, "procdump.exe"), path);
            Assert.AreEqual("MZ", File.ReadAllText(path));
        }

        [Test]
        public void Install_WhenTheDownloadFails_SaysSo()
        {
            var installer = Installer(_ => throw new WebException("offline"));

            var failure = Assert.Throws<IOException>(() => installer.Install());

            StringAssert.Contains("offline", failure.Message);
        }

        [Test]
        public void Install_WhenTheZipHoldsNoProcDump_SaysSo()
        {
            var installer = Installer(_ => Zip("readme.txt"));

            StringAssert.Contains("procdump.exe", Assert.Throws<IOException>(() => installer.Install()).Message);
        }

        [Test]
        public void Install_WhenWhatCameIsNotAZip_SaysSo()
        {
            var installer = Installer(_ => Encoding.ASCII.GetBytes("<html>proxy login</html>"));

            StringAssert.Contains("not a zip", Assert.Throws<IOException>(() => installer.Install()).Message);
        }

        private static (GeneralSettingsViewModel viewModel, Mock<IProcDumpInstaller> installer, Mock<IMessageBox> box) ViewModel(bool installed, bool agrees)
        {
            var installer = new Mock<IProcDumpInstaller>();
            installer.Setup(i => i.Find()).Returns(installed ? @"C:\Tools\procdump.exe" : null);
            installer.SetupGet(i => i.TargetFolder).Returns(@"C:\Tools");
            var box = new Mock<IMessageBox>();
            box.Setup(b => b.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>())).Returns(agrees);
            var types = new Mock<IExperimentalTypesProvider>();
            types.SetupGet(t => t.ExperimentalTypes).Returns(new Type[0]);
            var config = new Configuration(new UserSettings(new GeneralSettings(), new HotkeySettings(), null, null, null, null, null, null));
            var viewModel = new GeneralSettingsViewModel(config, new Mock<IOperatingSystem>().Object, box.Object, new Mock<Rubberduck.VBEditor.VbeRuntime.Settings.IVbeSettings>().Object,
                types.Object, null, null, null, null, installer.Object);
            viewModel.ApiAllowProcDump = true;
            return (viewModel, installer, box);
        }

        [Test]
        public void Ticking_WhenProcDumpIsThere_AsksNothing()
        {
            var (viewModel, installer, box) = ViewModel(installed: true, agrees: true);

            viewModel.EnsureProcDumpCommand.Execute(null);

            Assert.IsTrue(viewModel.ApiAllowProcDump);
            box.Verify(b => b.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>()), Times.Never);
            installer.Verify(i => i.Install(), Times.Never);
        }

        [Test]
        public void Ticking_WhenProcDumpIsMissing_OffersTheDownload_AndInstallsOnYes()
        {
            var (viewModel, installer, _) = ViewModel(installed: false, agrees: true);

            viewModel.EnsureProcDumpCommand.Execute(null);

            Assert.IsTrue(viewModel.ApiAllowProcDump);
            installer.Verify(i => i.Install(), Times.Once);
        }

        [Test]
        public void Ticking_WhenProcDumpIsMissing_AndTheUserSaysNo_TakesTheTickBack()
        {
            var (viewModel, installer, _) = ViewModel(installed: false, agrees: false);

            viewModel.EnsureProcDumpCommand.Execute(null);

            Assert.IsFalse(viewModel.ApiAllowProcDump);
            installer.Verify(i => i.Install(), Times.Never);
        }

        [Test]
        public void Ticking_WhenTheDownloadFails_TellsTheUser_AndTakesTheTickBack()
        {
            var (viewModel, installer, box) = ViewModel(installed: false, agrees: true);
            installer.Setup(i => i.Install()).Throws(new IOException("offline"));

            viewModel.EnsureProcDumpCommand.Execute(null);

            Assert.IsFalse(viewModel.ApiAllowProcDump);
            box.Verify(b => b.NotifyWarn("offline", It.IsAny<string>()), Times.Once);
        }
    }
}
