using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    internal class FakeDiagnostics : IDiagnosticsSource
    {
        public List<Diagnostic> Found = new List<Diagnostic>
        {
            new Diagnostic { Module = "Module1", Inspection = "Hint one", Severity = "hint", StartLine = 9 },
            new Diagnostic { Module = "Module1", Inspection = "Error late", Severity = "error", StartLine = 7 },
            new Diagnostic { Module = "Module1", Inspection = "Error early", Severity = "error", StartLine = 2 },
            new Diagnostic { Module = "Class1", Inspection = "Warning", Severity = "warning", StartLine = 1 },
            new Diagnostic { Module = "Module1", Inspection = "Suggestion", Severity = "suggestion", StartLine = 4 }
        };

        public string AskedProject, AskedModule;

        public DiagnosticsReport Run(string project, string module, TimeSpan? wait = null)
        {
            AskedProject = project;
            AskedModule = module;
            return new DiagnosticsReport
            {
                Project = project,
                Module = module,
                ParserState = "Ready",
                Diagnostics = Found.Where(d => module == null || d.Module == module).ToList()
            };
        }
    }

    /// <summary>A window that behaves like the editor's: sending a line echoes it and appends what the line "prints".</summary>
    internal class FakeWindow : IImmediateWindow
    {
        public string Text = string.Empty;
        public List<string> Sent = new List<string>();
        public Func<string, string> Output = line => line.StartsWith("?") ? " " + line.Substring(1) + " \r\n" : string.Empty;
        public bool Throws;

        public string Read()
        {
            if (Throws) throw ApiException.Unavailable("no window");
            return Text;
        }

        public string Peek() => Throws ? null : Text;

        public void Send(string line)
        {
            Sent.Add(line);
            Text += line + "\r\n" + Output(line);
        }
    }

    internal class ImmediateFixture
    {
        public FakeWindow Window = new FakeWindow();
        public int Asked;
        public bool Allow = true;
        public ImmediateApi Api;
    }

    internal static class Fakes
    {
        public static DiagnosticsApi Diagnostics() => new DiagnosticsApi(new FakeDiagnostics());

        public static ImmediateFixture Immediate(Func<bool> enabled, Func<bool> canRead = null)
        {
            var fixture = new ImmediateFixture();
            var clock = TimeSpan.Zero;
            fixture.Api = new ImmediateApi(
                fixture.Window,
                new ApiExecuteApproval(_ => { fixture.Asked++; return fixture.Allow; }),
                enabled,
                step => clock += step,
                () => clock,
                canRead: canRead);
            return fixture;
        }
    }

    [TestFixture]
    [Category("Api")]
    public class DiagnosticsApiTests
    {
        private FakeDiagnostics _source;
        private DiagnosticsApi _api;

        [SetUp]
        public void SetUp()
        {
            _source = new FakeDiagnostics();
            _api = new DiagnosticsApi(_source);
        }

        [Test]
        public void Results_AreWorstFirst_ThenInReadingOrder()
        {
            var names = _api.Get("p", null, null, null).Diagnostics.Select(d => d.Inspection).ToList();

            CollectionAssert.AreEqual(new[] { "Error early", "Error late", "Warning", "Suggestion", "Hint one" }, names);
        }

        [Test]
        public void CountsBySeverity_CoverEverythingFound_EvenWhatMinSeverityDrops()
        {
            var report = _api.Get("p", null, "error", null);

            Assert.AreEqual(2, report.Total);
            Assert.AreEqual(2, report.CountsBySeverity["error"]);
            Assert.AreEqual(1, report.CountsBySeverity["hint"]);
            Assert.AreEqual(5, report.CountsBySeverity.Values.Sum());
        }

        [Test]
        public void MinSeverity_DropsTheLessSevere()
        {
            var report = _api.Get("p", null, "warning", null);

            Assert.AreEqual(3, report.Total);
            CollectionAssert.AreEquivalent(new[] { "error", "error", "warning" }, report.Diagnostics.Select(d => d.Severity));
        }

        [Test]
        public void Limit_CutsTheList_AndSaysSo()
        {
            var report = _api.Get("p", null, null, 2);

            Assert.AreEqual(2, report.Diagnostics.Count);
            Assert.AreEqual(5, report.Total);
            Assert.IsTrue(report.Truncated);
            Assert.IsTrue(report.Diagnostics.All(d => d.Severity == "error"));
        }

        [Test]
        public void Module_IsPassedOn()
        {
            _api.Get("MyProject", "Class1", null, null);

            Assert.AreEqual("MyProject", _source.AskedProject);
            Assert.AreEqual("Class1", _source.AskedModule);
        }

        [Test]
        public void BlankModule_MeansTheWholeProject()
        {
            _api.Get("p", " ", null, null);

            Assert.IsNull(_source.AskedModule);
        }

        [TestCase("loud")]
        public void UnknownSeverity_Is400(string severity)
            => Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Get("p", null, severity, null)).Status);

        [TestCase(0)]
        [TestCase(1001)]
        public void LimitOutOfRange_Is400(int limit)
            => Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Get("p", null, null, limit)).Status);
    }

    [TestFixture]
    [Category("Api")]
    public class DiagnosticsRouteTests
    {
        private ApiRouter _router;

        [SetUp]
        public void SetUp()
            => _router = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api);

        private ApiResponse Get(string path, params string[] query)
        {
            var request = new ApiRequest { Method = "GET", Path = path };
            for (var i = 0; i < query.Length; i += 2) request.Query[query[i]] = query[i + 1];
            return _router.Handle(request);
        }

        [Test]
        public void Project_ListsEverything()
            => Assert.AreEqual(5, ((DiagnosticsReport)Get("/api/v1/projects/p1/diagnostics").Body).Total);

        [Test]
        public void ModuleInThePath_IsOneModule()
            => Assert.AreEqual(4, ((DiagnosticsReport)Get("/api/v1/projects/p1/modules/Module1/diagnostics").Body).Total);

        [Test]
        public void ModuleInTheQuery_IsOneModule()
            => Assert.AreEqual(1, ((DiagnosticsReport)Get("/api/v1/projects/p1/diagnostics", "module", "Class1").Body).Total);

        [Test]
        public void QueryFilters_Apply()
            => Assert.AreEqual(2, ((DiagnosticsReport)Get("/api/v1/projects/p1/diagnostics", "minSeverity", "error").Body).Total);

        [Test]
        public void BadLimit_Is400()
            => Assert.AreEqual(400, Get("/api/v1/projects/p1/diagnostics", "limit", "many").Status);

        [Test]
        public void NeedsNoApproval()
        {
            var asked = 0;
            var router = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => { asked++; return false; })), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api);

            router.Handle(new ApiRequest { Method = "GET", Path = "/api/v1/projects/p1/diagnostics" });

            Assert.AreEqual(0, asked);
        }
    }

    [TestFixture]
    [Category("Api")]
    public class ImmediateApiTests
    {
        private ImmediateFixture _fixture;
        private bool _enabled;

        [SetUp]
        public void SetUp()
        {
            _enabled = true;
            _fixture = Fakes.Immediate(() => _enabled);
        }

        private ImmediateRun Run(string command, int? wait = null) => _fixture.Api.Run("agent", command, wait);

        [Test]
        public void WhenNotAllowed_NothingWorks_AndNobodyIsAsked()
        {
            _enabled = false;

            Assert.AreEqual("immediate_disabled", Assert.Throws<ApiException>(() => _fixture.Api.Read()).Code);
            Assert.AreEqual("immediate_disabled", Assert.Throws<ApiException>(() => Run("?1")).Code);
            Assert.AreEqual(0, _fixture.Asked);
            Assert.IsEmpty(_fixture.Window.Sent);
        }

        [Test]
        public void Run_ReturnsWhatTheCommandAdded()
        {
            _fixture.Window.Text = "old stuff\r\n";

            var result = Run("?6*7");

            Assert.AreEqual("?6*7\r\n 6*7 \r\n", result.Output);
            Assert.IsTrue(result.Completed);
        }

        [Test]
        public void Run_SendsEachLine_InOrder_SkippingBlankOnes()
        {
            Run("x = 1\r\n\r\n?x\n");

            CollectionAssert.AreEqual(new[] { "x = 1", "?x" }, _fixture.Window.Sent);
        }

        [Test]
        public void FirstCommandAsks_LaterOnesDoNot()
        {
            Run("?1");
            Run("?2");

            Assert.AreEqual(1, _fixture.Asked);
        }

        [Test]
        public void Denied_RunsNothing()
        {
            _fixture.Allow = false;

            Assert.AreEqual("write_denied", Assert.Throws<ApiException>(() => Run("Kill \"x\"")).Code);
            Assert.IsEmpty(_fixture.Window.Sent);
        }

        [Test]
        public void Read_ReturnsTheText_AndTheEndOfALongOne()
        {
            _fixture.Window.Text = new string('a', 30000) + "END";

            var content = _fixture.Api.Read();

            Assert.IsTrue(content.Truncated);
            Assert.AreEqual(20000, content.Text.Length);
            StringAssert.EndsWith("END", content.Text);
        }

        [Test]
        public void ARunningCommand_IsReportedAsNotCompleted()
        {
            //The window changes at every look, as with code that keeps printing; the clock moves 100 ms per look.
            var busy = new BusyWindow();
            var api = new ImmediateApi(busy, new ApiExecuteApproval(_ => true), () => true, step => busy.Tick(), () => TimeSpan.FromMilliseconds(100 * busy.Reads));

            var result = api.Run("agent", "Loop", 2000);

            Assert.IsFalse(result.Completed);
        }

        [Test]
        public void ACommandThatNeverAppearsInTheWindow_IsNotReportedAsRun()
        {
            //The characters went elsewhere (Excel's cell editor): the window did not change at all.
            var clock = TimeSpan.Zero;
            var api = new ImmediateApi(new SilentWindow(), new ApiExecuteApproval(_ => true), () => true, step => clock += step, () => clock);

            var exception = Assert.Throws<ApiException>(() => api.Run("agent", "Logger.Level = 3", 1000));

            Assert.AreEqual("not_echoed", exception.Code);
        }

        [TestCase("")]
        [TestCase("   ")]
        public void EmptyCommand_Is400(string command)
            => Assert.AreEqual(400, Assert.Throws<ApiException>(() => Run(command)).Status);

        [Test]
        public void TooManyLines_Is400()
            => Assert.AreEqual(400, Assert.Throws<ApiException>(() => Run(string.Join("\n", Enumerable.Repeat("x = 1", ImmediateApi.MaxLines + 1)))).Status);

        [Test]
        public void CharactersTheWindowCannotTake_AreRefused_BeforeAnythingIsAsked()
        {
            var exception = Assert.Throws<ApiException>(() => Run("?\"你\""));

            Assert.AreEqual("unsupported_character", exception.Code);
            Assert.AreEqual(0, _fixture.Asked);
        }

        [Test]
        public void WaitOutOfRange_Is400()
            => Assert.AreEqual(400, Assert.Throws<ApiException>(() => Run("?1", ImmediateApi.MaxWaitMs + 1)).Status);

        [Test]
        public void UnreadableWindow_Is503()
        {
            _fixture.Window.Throws = true;

            Assert.AreEqual(503, Assert.Throws<ApiException>(() => Run("?1")).Status);
        }

        [TestCase("a\r\nb\r\n", "a\r\nb\r\nc\r\n", "c\r\n")]
        [TestCase("", "all new", "all new\r\n")]
        [TestCase("same", "same", "")]
        public void Delta_IsWhatWasAppended(string before, string after, string expected)
            => Assert.AreEqual(expected, ImmediateApi.Delta(before, after));

        [Test]
        public void Delta_IgnoresADifferentLastLineEndingOfTheOldText()
        {
            //The text pattern gave back other line breaks the second time.
            var before = "?1\r\n 1 \r\n";
            var after = "?1\n 1 \n?2\n 2 \n";

            Assert.AreEqual("?2\r\n 2 \r\n", ImmediateApi.Delta(before, after));
        }

        [Test]
        public void Delta_EndsEveryAnswerWithALineBreak()
            => Assert.AreEqual("Book1\r\n", ImmediateApi.Delta("?x\r\n", "?x\r\nBook1"));

        [Test]
        public void Delta_SurvivesTheWindowDroppingItsOldestLines()
        {
            var before = "one\r\ntwo\r\nthree\r\n";
            var after = "two\r\nthree\r\nfour\r\n";

            Assert.AreEqual("four\r\n", ImmediateApi.Delta(before, after));
        }

        private class SilentWindow : IImmediateWindow
        {
            public string Read() => "old text\r\n";
            public string Peek() => Read();
            public void Send(string line) { }
        }

        private class BusyWindow : IImmediateWindow
        {
            public int Reads;
            private int _text;
            public void Tick() => _text++;
            public string Read() { Reads++; return "line " + _text; }
            public string Peek() => Read();
            public void Send(string line) { }
        }
    }

    [TestFixture]
    [Category("Api")]
    public class ImmediateRouteTests
    {
        private bool _enabled;
        private ImmediateFixture _fixture;
        private ApiRouter _router;

        [SetUp]
        public void SetUp()
        {
            _enabled = true;
            _fixture = Fakes.Immediate(() => _enabled);
            _router = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), _fixture.Api);
        }

        private ApiResponse Send(string method, object body = null)
            => _router.Handle(new ApiRequest { Method = method, Path = "/api/v1/immediate", Body = body == null ? null : JObject.FromObject(body).ToString(), Client = "t" });

        [Test]
        public void Get_ReadsTheWindow()
        {
            _fixture.Window.Text = "hello";

            Assert.AreEqual("hello", ((ImmediateContent)Send("GET").Body).Text);
        }

        [Test]
        public void Post_RunsTheCommand()
        {
            var run = (ImmediateRun)Send("POST", new { command = "?1+1" }).Body;

            StringAssert.Contains("1+1", run.Output);
        }

        [Test]
        public void Post_WhenDisabled_Is403()
        {
            _enabled = false;

            Assert.AreEqual(403, Send("POST", new { command = "?1" }).Status);
        }

        [Test]
        public void Post_WithoutACommand_Is400()
            => Assert.AreEqual(400, Send("POST", new { waitMs = 100 }).Status);

        [Test]
        public void OtherMethods_Are405()
            => Assert.AreEqual(405, Send("DELETE").Status);
    }

    [TestFixture]
    [Category("Api")]
    public class ApiSetupCommandTests
    {
        [Test]
        public void UrlAcl_NamesThePortAndTheUser()
            => Assert.AreEqual("netsh http add urlacl url=http://+:7437/ user=\"PC\\Ann\"", ApiSetupCommands.UrlAcl(7437, "PC\\Ann"));

        [Test]
        public void Firewall_IsLimitedToThePortAndPrivateNetworks()
        {
            var command = ApiSetupCommands.Firewall(9000);

            StringAssert.Contains("localport=9000", command);
            StringAssert.Contains("profile=private,domain", command);
            StringAssert.DoesNotContain("public", command);
        }
    }
}
