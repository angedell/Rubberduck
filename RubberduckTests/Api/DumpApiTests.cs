using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
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

        private readonly HashSet<string> _files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _directories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _environment = new Dictionary<string, string>();
        private Runner _runner;
        private int _asked;
        private bool _allow;
        private DumpApi _api;

        [SetUp]
        public void SetUp()
        {
            _files.Clear();
            _directories.Clear();
            _directories.Add(Folder);
            _environment.Clear();
            _environment["LOCALAPPDATA"] = Local;
            _runner = new Runner();
            _asked = 0;
            _allow = true;

            var system = new Mock<IFileSystem>();
            system.SetupGet(s => s.Path).Returns(new FileSystem().Path);
            var file = new Mock<IFile>();
            file.Setup(f => f.Exists(It.IsAny<string>())).Returns<string>(p => _files.Contains(p));
            file.Setup(f => f.GetLastWriteTimeUtc(It.IsAny<string>())).Returns(DateTime.UtcNow);
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

            _api = new DumpApi(new ApiWriteApproval(_ => { _asked++; return _allow; }), system.Object, _runner, () => 4242, key => _environment.TryGetValue(key, out var value) ? value : null);
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
    }
}
