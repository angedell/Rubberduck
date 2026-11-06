using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Threading;
using Moq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    public class DumpApiTests
    {
        private const string Local = @"C:\Users\x\AppData\Local";
        private const string Tool = Local + @"\Rubberduck\Tools\procdump.exe";
        private const string Folder = Local + @"\Rubberduck\Dumps";

        private class Runner : IProcessRunner
        {
            public string File;
            public string Arguments;
            public int? ExitCode = 0;
            public Action Effect = () => { };

            public (int? ExitCode, string Output) Run(string file, string arguments, TimeSpan timeout)
            {
                File = file;
                Arguments = arguments;
                Effect();
                return (ExitCode, "ProcDump output");
            }
        }

        private class Analyzer : IDumpAnalyzer
        {
            public string File;
            public int Calls;

            public DumpAnalysis Analyze(string file, CancellationToken token)
            {
                File = file;
                Calls++;
                return new DumpAnalysis { Summary = "stuck" };
            }
        }

        private class Launched : IBackgroundProcess
        {
            public bool HasExited { get; set; }
            public int? ExitCode => HasExited ? 0 : (int?)null;
            public string Output => "monitoring";
            public bool Killed;

            public void Kill()
            {
                Killed = true;
                HasExited = true;
            }
        }

        private class Launcher : IProcessLauncher
        {
            public string File;
            public string Arguments;
            public Launched Process = new Launched();

            public IBackgroundProcess Start(string file, string arguments)
            {
                File = file;
                Arguments = arguments;
                return Process;
            }
        }

        private readonly Dictionary<string, DateTime> _written = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _directories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _environment = new Dictionary<string, string>();
        private Runner _runner;
        private Analyzer _analyzer;
        private Launcher _launcher;
        private int _asked;
        private bool _allow;
        private DumpApi _api;

        [SetUp]
        public void SetUp()
        {
            _files.Clear();
            _written.Clear();
            _directories.Clear();
            _directories.Add(Folder);
            _environment.Clear();
            _environment["LOCALAPPDATA"] = Local;
            _runner = new Runner();
            _analyzer = new Analyzer();
            _launcher = new Launcher();
            _asked = 0;
            _allow = true;
            _activated = true;

            var system = new Mock<IFileSystem>();
            system.SetupGet(s => s.Path).Returns(new FileSystem().Path);
            var file = new Mock<IFile>();
            file.Setup(f => f.Exists(It.IsAny<string>())).Returns<string>(p => _files.Contains(p));
            file.Setup(f => f.GetLastWriteTimeUtc(It.IsAny<string>())).Returns<string>(p => _written.TryGetValue(p, out var when) ? when : DateTime.UtcNow);
            file.Setup(f => f.Delete(It.IsAny<string>())).Callback<string>(p => _files.Remove(p));
            system.SetupGet(s => s.File).Returns(file.Object);
            var directory = new Mock<IDirectory>();
            directory.Setup(d => d.Exists(It.IsAny<string>())).Returns<string>(d => _directories.Contains(d));
            directory.Setup(d => d.GetFiles(It.IsAny<string>(), "*.dmp")).Returns<string, string>((d, _) => _files.Where(f => f.StartsWith(d + @"\") && f.EndsWith(".dmp")).ToArray());
            system.SetupGet(s => s.Directory).Returns(directory.Object);
            var info = new Mock<IFileInfoFactory>();
            var length = new Mock<IFileInfo>();
            length.SetupGet(i => i.Length).Returns(1234);
            info.Setup(i => i.FromFileName(It.IsAny<string>())).Returns(length.Object);
            system.SetupGet(s => s.FileInfo).Returns(info.Object);

            _api = new DumpApi(new ApiWriteApproval(_ => { _asked++; return _allow; }), system.Object, _runner, () => 4242, key => _environment.TryGetValue(key, out var value) ? value : null, _analyzer, _launcher, activated: () => _activated);
        }

        private bool _activated = true;

        [Test]
        public void Capture_WhenProcDumpIsSwitchedOff_IsForbidden_AndAsksNothing()
        {
            _files.Add(Tool);
            _activated = false;

            var failure = Assert.Throws<ApiException>(() => _api.Capture("agent", null, null));

            Assert.AreEqual(403, failure.Status);
            Assert.AreEqual("procdump_off", failure.Code);
            Assert.AreEqual(0, _asked);
            Assert.IsNull(_runner.File);
        }

        [Test]
        public void StartMonitor_WhenProcDumpIsSwitchedOff_IsForbidden()
        {
            _files.Add(Tool);
            _activated = false;

            var failure = Assert.Throws<ApiException>(() => _api.StartMonitor("agent", null, null, null, null, null));

            Assert.AreEqual("procdump_off", failure.Code);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Status_WhenProcDumpIsSwitchedOff_StillAnswers()
        {
            _activated = false;

            Assert.DoesNotThrow(() => _api.Status());
        }

        [Test]
        public void Capture_WithoutProcDump_SaysWhereToPutIt_AndAsksNothing()
        {
            var failure = Assert.Throws<ApiException>(() => _api.Capture("agent", null, null));

            Assert.AreEqual(DumpApi.ToolNotFound, failure.Code);
            StringAssert.Contains(@"Rubberduck\Tools", failure.Message);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Status_ReportsTheToolOnlyWhenItIsThere()
        {
            Assert.IsNull(_api.Status().Tool);
            Assert.IsNotNull(_api.Status().Note);

            _files.Add(Tool);

            Assert.AreEqual(Tool, _api.Status().Tool);
            Assert.IsNull(_api.Status().Note);
        }

        [Test]
        public void Capture_RunsProcDumpOnThisProcess_AndReturnsTheNewFile()
        {
            _files.Add(Tool);
            _files.Add(Folder + @"\EXCEL_old.dmp");
            _runner.Effect = () => _files.Add(Folder + @"\EXCEL_new.dmp");

            var report = _api.Capture("agent", null, null);

            Assert.AreEqual(Tool, _runner.File);
            Assert.AreEqual($"-accepteula -ma 4242 \"{Folder}\"", _runner.Arguments);
            Assert.AreEqual(Folder + @"\EXCEL_new.dmp", report.File);
            Assert.AreEqual(1234, report.SizeBytes);
            Assert.AreEqual(1, _asked);
        }

        [Test]
        public void Capture_WritesToTheFolderTheEnvironmentNames()
        {
            _files.Add(Tool);
            _environment["RUBBERDUCK_DUMP_FOLDER"] = @"D:\Shared\DebugRD";
            _directories.Add(@"D:\Shared\DebugRD");
            _runner.Effect = () => _files.Add(@"D:\Shared\DebugRD\EXCEL_new.dmp");

            var report = _api.Capture("agent", null, null);

            Assert.AreEqual(@"D:\Shared\DebugRD\EXCEL_new.dmp", report.File);
            StringAssert.Contains(@"""D:\Shared\DebugRD""", _runner.Arguments);
        }

        [Test]
        public void Capture_Mini_UsesTheSmallDump()
        {
            _files.Add(Tool);
            _runner.Effect = () => _files.Add(Folder + @"\EXCEL_new.dmp");

            _api.Capture("agent", "mini", 30);

            StringAssert.Contains("-mm", _runner.Arguments);
        }

        [Test]
        public void Capture_WhenTheUserSaysNo_RunsNothing()
        {
            _files.Add(Tool);
            _allow = false;

            Assert.Throws<ApiException>(() => _api.Capture("agent", null, null));

            Assert.IsNull(_runner.File);
        }

        [Test]
        public void Capture_WhenNoFileAppears_SaysSo()
        {
            _files.Add(Tool);
            _runner.ExitCode = null;

            var failure = Assert.Throws<ApiException>(() => _api.Capture("agent", null, null));

            Assert.AreEqual(503, failure.Status);
            StringAssert.Contains("did not finish", failure.Message);
        }

        [TestCase("huge")]
        public void Capture_RejectsAnUnknownKind(string kind)
        {
            _files.Add(Tool);

            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Capture("agent", kind, null)).Status);
        }

        private string Dump(string name, int minutesAgo = 0)
        {
            var path = Folder + @"\" + name;
            _files.Add(path);
            _written[path] = DateTime.UtcNow.AddMinutes(-minutesAgo);
            return path;
        }

        [Test]
        public void Capture_WithAnalyze_ReadsTheNewDump()
        {
            _files.Add(Tool);
            _runner.Effect = () => Dump("EXCEL_new.dmp");

            var report = _api.Capture("agent", null, null, analyze: true);

            Assert.AreEqual("stuck", report.Analysis.Summary);
            Assert.AreEqual(Folder + @"\EXCEL_new.dmp", _analyzer.File);
        }

        [Test]
        public void Capture_WithoutAnalyze_ReadsNothing()
        {
            _files.Add(Tool);
            _runner.Effect = () => Dump("EXCEL_new.dmp");

            Assert.IsNull(_api.Capture("agent", null, null).Analysis);
            Assert.AreEqual(0, _analyzer.Calls);
        }

        [TestCase(@"C:\Users\x\SARO Dropbox\Debug", "dropbox")]
        [TestCase(@"C:\Users\x\OneDrive - Contoso\Debug", "onedrive")]
        public void Status_UsesTheLocalFolderInsteadOfASyncedOne_AndSaysSo(string folder, string service)
        {
            _files.Add(Tool);
            _environment["RUBBERDUCK_DUMP_FOLDER"] = folder;

            var report = _api.Status();

            Assert.AreEqual(Folder, report.Folder);
            Assert.IsNull(report.Warning);
            StringAssert.Contains(folder, report.Note);
        }

        [Test]
        public void Capture_WithASyncedFolderNamed_WritesToTheLocalOne()
        {
            _files.Add(Tool);
            _environment["RUBBERDUCK_DUMP_FOLDER"] = @"C:\Users\x\SARO Dropbox\Debug";
            _runner.Effect = () => Dump("EXCEL_new.dmp");

            var report = _api.Capture("agent", null, null);

            Assert.AreEqual(Folder + @"\EXCEL_new.dmp", report.File);
            StringAssert.Contains($"\"{Folder}\"", _runner.Arguments);
        }

        [Test]
        public void Status_WarnsWhenNoLocalFolderIsAvailable()
        {
            _environment["LOCALAPPDATA"] = @"C:\Users\x\Dropbox\AppData";
            _environment["RUBBERDUCK_DUMP_FOLDER"] = @"C:\Users\x\SARO Dropbox\Debug";

            var report = _api.Status();

            Assert.AreEqual(@"C:\Users\x\SARO Dropbox\Debug", report.Folder);
            StringAssert.Contains("dropbox", report.Warning);
        }

        [Test]
        public void Status_SaysNothingOfTheFolderWhenItIsLocal()
        {
            Assert.IsNull(_api.Status().Warning);
        }

        [Test]
        public void Analyze_Latest_ReadsTheNewestDump_AndAsksNothing()
        {
            Dump("old.dmp", 30);
            var newest = Dump("new.dmp", 1);

            var analysis = _api.Analyze("agent", "latest");

            Assert.AreEqual(newest, _analyzer.File);
            Assert.AreEqual("stuck", analysis.Summary);
            Assert.AreEqual(1234, analysis.SizeBytes);
            Assert.AreEqual(0, _asked);
            Assert.IsTrue(_files.Contains(newest));
        }

        [Test]
        public void Analyze_ByName_ReadsThatDump()
        {
            var old = Dump("old.dmp", 30);
            Dump("new.dmp", 1);

            _api.Analyze("agent", "OLD.dmp");

            Assert.AreEqual(old, _analyzer.File);
        }

        [Test]
        public void Analyze_WithDeleteAfter_AsksFirst_ThenDeletes()
        {
            var file = Dump("new.dmp");

            var analysis = _api.Analyze("agent", "new.dmp", deleteAfter: true);

            Assert.IsTrue(analysis.Deleted);
            Assert.IsFalse(_files.Contains(file));
            Assert.AreEqual(1, _asked);
        }

        [Test]
        public void Analyze_WithDeleteAfter_WhenTheUserSaysNo_ReadsAndDeletesNothing()
        {
            var file = Dump("new.dmp");
            _allow = false;

            Assert.Throws<ApiException>(() => _api.Analyze("agent", "new.dmp", deleteAfter: true));

            Assert.AreEqual(0, _analyzer.Calls);
            Assert.IsTrue(_files.Contains(file));
        }

        [Test]
        public void Analyze_WithNoDump_IsNotFound()
        {
            Assert.AreEqual(404, Assert.Throws<ApiException>(() => _api.Analyze("agent", null)).Status);
        }

        [TestCase(@"..\..\secret.dmp")]
        [TestCase(@"C:\Windows\win.ini")]
        [TestCase("sub/other.dmp")]
        [TestCase("notes.txt")]
        public void Delete_RefusesAnythingButAFileNameOfTheFolder(string name)
        {
            _files.Add(@"C:\Windows\win.ini");

            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Delete("agent", name)).Status);

            Assert.IsTrue(_files.Contains(@"C:\Windows\win.ini"));
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Delete_ANameThatIsNotThere_IsNotFound()
        {
            Dump("a.dmp");

            Assert.AreEqual(404, Assert.Throws<ApiException>(() => _api.Delete("agent", "b.dmp")).Status);
        }

        [Test]
        public void Delete_RemovesTheFile_ReportsWhatItFreed()
        {
            var a = Dump("a.dmp");
            var b = Dump("b.dmp");

            var report = _api.Delete("agent", "a.dmp");

            CollectionAssert.AreEqual(new[] { a }, report.Deleted);
            Assert.AreEqual(1234, report.FreedBytes);
            CollectionAssert.AreEqual(new[] { b }, report.Existing);
            Assert.AreEqual(1, _asked);
        }

        [Test]
        public void Delete_WhenTheUserSaysNo_KeepsTheFile()
        {
            var file = Dump("a.dmp");
            _allow = false;

            Assert.Throws<ApiException>(() => _api.Delete("agent", "a.dmp"));

            Assert.IsTrue(_files.Contains(file));
        }

        [Test]
        public void DeleteAll_RemovesEveryDump_KeepingTheNewestWhenAsked()
        {
            Dump("a.dmp", 30);
            Dump("b.dmp", 20);
            var newest = Dump("c.dmp", 1);

            var report = _api.DeleteAll("agent", 1);

            Assert.AreEqual(2, report.Deleted.Count);
            CollectionAssert.AreEqual(new[] { newest }, report.Existing);
        }

        [Test]
        public void DeleteAll_WithNothingToDelete_AsksNothing()
        {
            Dump("a.dmp");

            var report = _api.DeleteAll("agent", 1);

            Assert.IsEmpty(report.Deleted);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void DeleteAll_AFileThatIsInUse_IsReportedNotThrown()
        {
            var busy = Dump("busy.dmp", 1);
            Dump("free.dmp", 2);
            var file = new Mock<IFile>();
            file.Setup(f => f.GetLastWriteTimeUtc(It.IsAny<string>())).Returns<string>(p => _written[p]);
            file.Setup(f => f.Delete(busy)).Throws(new System.IO.IOException("in use"));
            file.Setup(f => f.Delete(It.Is<string>(p => p != busy))).Callback<string>(p => _files.Remove(p));
            var system = new Mock<IFileSystem>();
            system.SetupGet(s => s.Path).Returns(new FileSystem().Path);
            system.SetupGet(s => s.File).Returns(file.Object);
            var directory = new Mock<IDirectory>();
            directory.Setup(d => d.Exists(It.IsAny<string>())).Returns(true);
            directory.Setup(d => d.GetFiles(It.IsAny<string>(), "*.dmp")).Returns<string, string>((d, _) => _files.Where(f => f.EndsWith(".dmp")).ToArray());
            system.SetupGet(s => s.Directory).Returns(directory.Object);
            var length = new Mock<IFileInfo>();
            length.SetupGet(i => i.Length).Returns(10);
            var info = new Mock<IFileInfoFactory>();
            info.Setup(i => i.FromFileName(It.IsAny<string>())).Returns(length.Object);
            system.SetupGet(s => s.FileInfo).Returns(info.Object);
            var api = new DumpApi(new ApiWriteApproval(_ => true), system.Object, _runner, () => 4242, key => key == "LOCALAPPDATA" ? Local : null, _analyzer, _launcher);

            var report = api.DeleteAll("agent", null);

            Assert.AreEqual(1, report.Deleted.Count);
            Assert.AreEqual(1, report.Failed.Count);
            StringAssert.Contains("in use", report.Failed[0]);
        }

        [Test]
        public void StartMonitor_WatchesForAHang_ByDefault_AndAsksFirst()
        {
            _files.Add(Tool);

            var report = _api.StartMonitor("agent", null, null, null, null, null);

            Assert.IsTrue(report.Running);
            Assert.AreEqual($"-accepteula -ma -h -n 2 4242 \"{Folder}\"", _launcher.Arguments);
            Assert.AreEqual(Tool, _launcher.File);
            Assert.AreEqual("a hung window", report.Triggers);
            Assert.AreEqual(1, _asked);
        }

        [Test]
        public void StartMonitor_CombinesTheTriggersAsked()
        {
            _files.Add(Tool);

            _api.StartMonitor("agent", "mini", true, true, 90, 5);

            Assert.AreEqual($"-accepteula -mm -h -e -c 90 -s 10 -n 5 4242 \"{Folder}\"", _launcher.Arguments);
        }

        [Test]
        public void StartMonitor_WithOnlyExceptions_DoesNotAlsoWatchForAHang()
        {
            _files.Add(Tool);

            _api.StartMonitor("agent", null, null, true, null, null);

            StringAssert.DoesNotContain(" -h", _launcher.Arguments);
            StringAssert.Contains(" -e", _launcher.Arguments);
        }

        [Test]
        public void StartMonitor_WithNothingToWatchFor_IsRefused()
        {
            _files.Add(Tool);

            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.StartMonitor("agent", null, false, false, null, null)).Status);
            Assert.IsNull(_launcher.File);
        }

        [TestCase(0)]
        [TestCase(11)]
        public void StartMonitor_RejectsAnAbsurdNumberOfDumps(int count)
        {
            _files.Add(Tool);

            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.StartMonitor("agent", null, null, null, null, count)).Status);
        }

        [Test]
        public void StartMonitor_WhenTheUserSaysNo_StartsNothing()
        {
            _files.Add(Tool);
            _allow = false;

            Assert.Throws<ApiException>(() => _api.StartMonitor("agent", null, null, null, null, null));

            Assert.IsNull(_launcher.File);
        }

        [Test]
        public void StartMonitor_WithoutProcDump_SaysWhereToPutIt()
        {
            Assert.AreEqual(DumpApi.ToolNotFound, Assert.Throws<ApiException>(() => _api.StartMonitor("agent", null, null, null, null, null)).Code);
        }

        [Test]
        public void StartMonitor_WhileOneRuns_IsAConflict()
        {
            _files.Add(Tool);
            _api.StartMonitor("agent", null, null, null, null, null);

            Assert.AreEqual(409, Assert.Throws<ApiException>(() => _api.StartMonitor("agent", null, null, null, null, null)).Status);
        }

        [Test]
        public void StartMonitor_AfterItEnded_StartsAgain()
        {
            _files.Add(Tool);
            _api.StartMonitor("agent", null, null, null, null, null);
            _launcher.Process.HasExited = true;

            Assert.DoesNotThrow(() => _api.StartMonitor("agent", null, null, null, null, null));
        }

        [Test]
        public void MonitorStatus_BeforeAnyStart_SaysSo()
        {
            var report = _api.MonitorStatus(true);

            Assert.IsFalse(report.Running);
            StringAssert.Contains("No monitor", report.Note);
        }

        [Test]
        public void MonitorStatus_ListsOnlyTheDumpsItCaught_AndReadsTheNewest()
        {
            _files.Add(Tool);
            Dump("before.dmp", 60);
            _api.StartMonitor("agent", null, null, null, null, null);
            Dump("caught1.dmp", 5);
            var latest = Dump("caught2.dmp", 1);

            var report = _api.MonitorStatus(true);

            Assert.AreEqual(2, report.NewDumps.Count);
            Assert.AreEqual(latest, report.NewDumps[0]);
            Assert.AreEqual(latest, _analyzer.File);
            Assert.AreEqual("stuck", report.Analysis.Summary);
        }

        [Test]
        public void MonitorStatus_WithoutAnalyze_ReadsNothing()
        {
            _files.Add(Tool);
            _api.StartMonitor("agent", null, null, null, null, null);
            Dump("caught.dmp");

            _api.MonitorStatus(false);

            Assert.AreEqual(0, _analyzer.Calls);
        }

        [Test]
        public void StopMonitor_EndsProcDump()
        {
            _files.Add(Tool);
            _api.StartMonitor("agent", null, null, null, null, null);

            var report = _api.StopMonitor();

            Assert.IsTrue(_launcher.Process.Killed);
            Assert.IsFalse(report.Running);
        }
    }
}
