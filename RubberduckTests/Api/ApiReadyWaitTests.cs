using System;
using System.Collections.Generic;
using System.Threading;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Api
{
    [TestFixture]
    [Category("Api")]
    public class ApiReadyWaitTests
    {
        private ManualResetEventSlim _release;
        private int _catchUps;
        private Exception _failure;
        private ReadyWait _ready;
        private ApiRouter _router;
        private McpEndpoint _mcp;


        [SetUp]
        public void SetUp()
        {
            _release = new ManualResetEventSlim(true);
            _catchUps = 0;
            _failure = null;
            _ready = new ReadyWait(
                key => key == "missing" ? throw ApiException.NotFound("No open project 'missing'.") : ("id1", "Book1"),
                () =>
                {
                    Interlocked.Increment(ref _catchUps);
                    _release.Wait();
                    if (_failure != null)
                    {
                        throw _failure;
                    }

                    return ParserState.Ready;
                });

            var insight = new InsightApi(new FakeInsight());
            _router = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), new DiagnosticsApi(new Diag()), Fakes.Immediate(() => false).Api,
                insight: insight, tests: new TestsApi(new FakeTests(), new ApiExecuteApproval(_ => true), () => true), ready: _ready);
            _mcp = new McpEndpoint(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), new DiagnosticsApi(new Diag()), Fakes.Immediate(() => false).Api,
                insight: insight, ready: _ready);
        }

        [TearDown]
        public void TearDown() => _release.Set();

        private class Diag : IDiagnosticsSource
        {
            public DiagnosticsReport Run(string project, string module, TimeSpan? wait = null) => new DiagnosticsReport { Project = project };
        }

        private ApiResponse Get(string path, string wait = "ready", string waitSeconds = null)
        {
            var request = new ApiRequest { Method = "GET", Path = path };
            if (wait != null)
            {
                request.Query["wait"] = wait;
            }

            if (waitSeconds != null)
            {
                request.Query["waitSeconds"] = waitSeconds;
            }

            return _router.Handle(request);
        }

        [TestCase("/api/v1/projects/p1/parser-errors")]
        [TestCase("/api/v1/projects/p1/symbols")]
        [TestCase("/api/v1/projects/p1/symbols/references")]
        [TestCase("/api/v1/projects/p1/symbols/definition")]
        [TestCase("/api/v1/projects/p1/structure")]
        [TestCase("/api/v1/projects/p1/metrics")]
        [TestCase("/api/v1/projects/p1/tests")]
        public void WhenTheParserIsNotReadyInTime_TheAnswerIs202Running(string path)
        {
            _release.Reset();

            var response = Get(path, "ready", "0");

            Assert.AreEqual(202, response.Status);
            var pending = (ReadyPending)response.Body;
            Assert.AreEqual("running", pending.Status);
            Assert.AreEqual("Book1", pending.Project);
            Assert.AreEqual(ReadyWait.RetryAfterSeconds, pending.RetryAfterSeconds);
        }

        [Test]
        public void WhenTheParserIsReady_TheReadGoesAhead()
        {
            var response = Get("/api/v1/projects/p1/structure");

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual(1, _catchUps);
        }

        [Test]
        public void WithoutWait_NothingWaitsAndNothingIsStarted()
        {
            _release.Reset();

            var response = Get("/api/v1/projects/p1/structure", wait: null, waitSeconds: "0");

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual(0, _catchUps);
        }

        [Test]
        public void AskingAgainJoinsTheCatchUpInProgress_AndReadsOnceItIsDone()
        {
            _release.Reset();
            Assert.AreEqual(202, Get("/api/v1/projects/p1/structure", "ready", "0").Status);
            Assert.AreEqual(202, Get("/api/v1/projects/p1/structure", "ready", "0").Status);

            _release.Set();
            var response = Get("/api/v1/projects/p1/structure", "ready", "5");

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual(1, _catchUps, "one catch-up for the three calls");
        }

        [Test]
        public void TheWaitIsNotOnTheCallersThreadAlone_ItEndsWithinTheCap()
        {
            _release.Reset();
            var timer = System.Diagnostics.Stopwatch.StartNew();

            var response = Get("/api/v1/projects/p1/metrics", "ready", "1");

            Assert.AreEqual(202, response.Status);
            Assert.Less(timer.Elapsed, TimeSpan.FromSeconds(10));
        }

        [Test]
        public void ACatchUpThatFails_IsReportedAsTheReadWould()
        {
            _failure = ApiException.NotInDesignMode();

            var response = Get("/api/v1/projects/p1/structure");

            Assert.AreEqual(409, response.Status);
        }

        [Test]
        public void AfterAFailure_TheNextCallStartsOver()
        {
            _failure = ApiException.NotInDesignMode();
            Get("/api/v1/projects/p1/structure");
            _failure = null;

            Assert.AreEqual(200, Get("/api/v1/projects/p1/structure").Status);
            Assert.AreEqual(2, _catchUps);
        }

        [Test]
        public void AMissingProject_Is404_WithoutStartingAnything()
        {
            var response = Get("/api/v1/projects/missing/structure");

            Assert.AreEqual(404, response.Status);
            Assert.AreEqual(0, _catchUps);
        }

        [TestCase("now")]
        [TestCase("")]
        public void AnyOtherWait_IsABadRequest(string wait)
        {
            Assert.AreEqual(400, Get("/api/v1/projects/p1/structure", wait).Status);
        }

        [TestCase("-1")]
        [TestCase("301")]
        [TestCase("soon")]
        public void AnInvalidWaitSeconds_IsABadRequest(string seconds)
        {
            Assert.AreEqual(400, Get("/api/v1/projects/p1/structure", "ready", seconds).Status);
        }

        [Test]
        public void Libraries_DoNotUseTheParser_SoWaitIsNotApplied()
        {
            _release.Reset();

            var response = Get("/api/v1/projects/p1/libraries", "ready", "0");

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual(0, _catchUps);
        }

        [Test]
        public void OnlyReadsAreGated()
        {
            Assert.IsFalse(ReadyWait.Applies("POST", new[] { "projects", "p1", "tests", "run" }));
            Assert.IsFalse(ReadyWait.Applies("GET", new[] { "projects", "p1", "diagnostics" }));
            Assert.IsFalse(ReadyWait.Applies("GET", new[] { "projects", "p1", "modules" }));
            Assert.IsTrue(ReadyWait.Applies("GET", new[] { "projects", "p1", "symbols", "references" }));
        }

        private JObject CallTool(string tool, JObject arguments)
        {
            var body = new JObject
            {
                ["jsonrpc"] = "2.0",
                ["id"] = 1,
                ["method"] = "tools/call",
                ["params"] = new JObject { ["name"] = tool, ["arguments"] = arguments }
            };
            var response = _mcp.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Body = body.ToString() });
            return (JObject)((JObject)response.Body)["result"];
        }

        [Test]
        public void TheMcpTools_TakeWaitAndAnswerRunning()
        {
            _release.Reset();

            var result = CallTool("get_structure", new JObject { ["project"] = "p1", ["wait"] = "ready", ["waitSeconds"] = 0 });

            Assert.IsFalse(result.Value<bool>("isError"));
            StringAssert.Contains("\"running\"", (string)result["content"][0]["text"]);
        }

        [Test]
        public void TheMcpTools_ReadAsBeforeWithoutWait()
        {
            _release.Reset();

            var result = CallTool("get_structure", new JObject { ["project"] = "p1" });

            StringAssert.Contains("App.Core", (string)result["content"][0]["text"]);
            Assert.AreEqual(0, _catchUps);
        }

        [Test]
        public void TheMcpToolsThatReadTheParser_OfferWait()
        {
            var body = new JObject { ["jsonrpc"] = "2.0", ["id"] = 1, ["method"] = "tools/list" };
            var response = _mcp.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Body = body.ToString() });
            var tools = (JArray)((JObject)((JObject)response.Body)["result"])["tools"];

            var offering = new List<string>();
            foreach (JObject tool in tools)
            {
                if (((JObject)tool["inputSchema"]["properties"]).ContainsKey("wait"))
                {
                    offering.Add((string)tool["name"]);
                }
            }

            CollectionAssert.AreEquivalent(
                new[] { "get_parser_errors", "list_symbols", "find_references", "find_definition", "get_structure", "get_metrics" },
                offering);
        }
    }
}
