using System.IO;
using System.IO.Abstractions;
using System.Linq;
using System.Text;
using Moq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    public class LogApiTests
    {
        private const string LogPath = @"C:\Users\x\AppData\Roaming\Rubberduck\Logs\RubberduckLog.txt";

        private class Control : ILogControl
        {
            public string ConfiguredLevel { get; set; } = "Trace";
            public string EffectiveLevel { get; set; } = "Trace";
            public string LogFile => LogPath;
            public string Set;
            public void SetLevel(string level) { Set = level; ConfiguredLevel = EffectiveLevel = level; }
        }

        private static readonly string Log =
            "2026-10-27 06:28:22.9468;INFO-2.5.92.13;Rubberduck._Extension;Rubberduck is shutting down.;\r\n" +
            "2026-10-27 06:28:22.9623;TRACE-2.5.92.13;Rubberduck.VBEditor.WindowsApi.SubclassManager;Releasing subclass for hWnd 0x00090D3C.;\r\n" +
            "2026-10-27 06:28:23.0000;WARN-2.5.92.13;Rubberduck.Navigation.CodeExplorer.CodeExplorerProjectViewModel;No resources listed for 'C:\\a.xlsm'.;\r\n" +
            "2026-10-27 06:28:23.1000;ERROR-2.5.92.13;Rubberduck.PackageResources.PackageResourceService;Boom;System.IO.IOException: locked\r\n" +
            "   at Rubberduck.Something()\r\n" +
            "2026-10-27 06:28:23.2000;DEBUG-2.5.92.13;Rubberduck.Api.ApiServer;Listening on 127.0.0.1:7437.;\r\n";

        private Control _control;
        private string _content;
        private bool _exists;
        private int _asked;
        private bool _allow;
        private LogApi _api;

        [SetUp]
        public void SetUp()
        {
            _control = new Control();
            _content = Log;
            _exists = true;
            _asked = 0;
            _allow = true;

            var file = new Mock<IFile>();
            file.Setup(f => f.Exists(LogPath)).Returns(() => _exists);
            file.Setup(f => f.Open(LogPath, FileMode.Open, FileAccess.Read, It.IsAny<FileShare>()))
                .Returns(() => new MemoryStream(Encoding.UTF8.GetBytes(_content)));
            var fileSystem = new Mock<IFileSystem>();
            fileSystem.Setup(f => f.File).Returns(file.Object);

            _api = new LogApi(_control, fileSystem.Object, new ApiWriteApproval(_ => { _asked++; return _allow; }));
        }

        [Test]
        public void Parse_SplitsTheEntries_AndKeepsAStackTraceWithItsEntry()
        {
            var entries = LogApi.Parse(Log);

            Assert.AreEqual(5, entries.Count);
            Assert.AreEqual("Warn", entries[2].Level);
            Assert.AreEqual("Rubberduck.Navigation.CodeExplorer.CodeExplorerProjectViewModel", entries[2].Logger);
            Assert.AreEqual("No resources listed for 'C:\\a.xlsm'.", entries[2].Message);
            StringAssert.StartsWith("Boom;System.IO.IOException: locked\n   at Rubberduck.Something()", entries[3].Message);
        }

        [Test]
        public void Get_GivesTheLatestEntries_OldestFirst()
        {
            var report = _api.Get(2, null, null, null);

            CollectionAssert.AreEqual(new[] { "Error", "Debug" }, report.Entries.Select(e => e.Level));
            Assert.AreEqual(5, report.Total);
            Assert.IsTrue(report.Truncated);
        }

        [Test]
        public void MinLevel_LeavesOutTheMoreVerbose()
        {
            var report = _api.Get(null, "warn", null, null);

            CollectionAssert.AreEqual(new[] { "Warn", "Error" }, report.Entries.Select(e => e.Level));
        }

        [Test]
        public void Contains_LooksAtTheMessageAndTheClass_AnyCase()
        {
            Assert.AreEqual(1, _api.Get(null, null, "NO RESOURCES", null).Entries.Count);
            Assert.AreEqual(1, _api.Get(null, null, "packageresourceservice", null).Entries.Count);
        }

        [Test]
        public void Logger_FiltersByClassName()
            => Assert.AreEqual(1, _api.Get(null, null, null, "ApiServer").Entries.Count);

        [TestCase(-1)]
        [TestCase(2001)]
        public void ATailOutOfRange_Is400(int tail)
            => Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Get(tail, null, null, null)).Status);

        [TestCase("loud")]
        [TestCase("off")]
        public void AMinLevelThatIsNotALevel_Is400(string level)
            => Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Get(null, level, null, null)).Status);

        [Test]
        public void WhenTheSettingDidNotTakeEffect_TheReportSaysSo()
        {
            _control.ConfiguredLevel = "Trace";
            _control.EffectiveLevel = "Off";

            var report = _api.Get(null, null, null, null);

            Assert.AreEqual("Trace", report.ConfiguredLevel);
            Assert.AreEqual("Off", report.EffectiveLevel);
            StringAssert.Contains("has not taken effect", report.Note);
        }

        [Test]
        public void WhenLoggingIsOff_TheReportSaysNothingNewIsWritten()
        {
            _control.ConfiguredLevel = _control.EffectiveLevel = "Off";

            StringAssert.Contains("Logging is off", _api.Get(null, null, null, null).Note);
        }

        [Test]
        public void WithoutALogFile_TheReportSaysSo_InsteadOfFailing()
        {
            _exists = false;

            var report = _api.Get(null, null, null, null);

            StringAssert.Contains("no log file", report.Note);
            Assert.IsEmpty(report.Entries);
            Assert.AreEqual(LogPath, report.Path);
        }

        [Test]
        public void SetLevel_AsksForApproval_ThenSetsTheCanonicalName()
        {
            var report = _api.SetLevel("agent", "debug");

            Assert.AreEqual(1, _asked);
            Assert.AreEqual("Debug", _control.Set);
            Assert.AreEqual("Debug", report.ConfiguredLevel);
            StringAssert.Contains("Verbose", report.Note);
        }

        [Test]
        public void SetLevel_WithoutApproval_ChangesNothing()
        {
            _allow = false;

            Assert.Throws<ApiException>(() => _api.SetLevel("agent", "Trace"));
            Assert.IsNull(_control.Set);
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase("chatty")]
        public void SetLevel_WithAnythingElse_Is400_AndAsksNobody(string level)
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.SetLevel("agent", level)).Status);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Router_ServesTheLogAndTheLevel()
        {
            var router = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api, log: _api);

            var read = router.Handle(new ApiRequest { Method = "GET", Path = ApiRouter.Base + "/log", Client = "t" });
            read.Body.GetType();
            Assert.AreEqual(200, read.Status);
            Assert.AreEqual(5, ((LogReport)read.Body).Entries.Count);

            var put = router.Handle(new ApiRequest { Method = "PUT", Path = ApiRouter.Base + "/log", Body = new JObject { ["level"] = "Info" }.ToString(), Client = "t" });
            Assert.AreEqual(200, put.Status);
            Assert.AreEqual("Info", _control.Set);

            Assert.AreEqual(405, router.Handle(new ApiRequest { Method = "DELETE", Path = ApiRouter.Base + "/log", Client = "t" }).Status);
        }

        [Test]
        public void WithoutALogApi_TheRouteIsNotThere()
        {
            var router = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api);

            Assert.AreEqual(404, router.Handle(new ApiRequest { Method = "GET", Path = ApiRouter.Base + "/log", Client = "t" }).Status);
        }
    }
}
