using System;
using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    [Category("Api")]
    public class ImmediateCaptureTests
    {
        private FakeWindow _window;
        private ImmediateCapture _capture;
        private TimeSpan _clock;
        private Action _whileWaiting;

        [SetUp]
        public void SetUp()
        {
            _window = new FakeWindow { Text = "old line\r\n" };
            _clock = TimeSpan.Zero;
            _whileWaiting = null;
            _capture = new ImmediateCapture(_window, () => true, step => { _clock += step; _whileWaiting?.Invoke(); }, () => _clock);
            _capture.Capture(false);
        }

        private void Print(string line) => _window.Text += line + "\r\n";

        [Test]
        public void WhatWasThereBeforeTheFirstLook_IsHistory()
            => Assert.AreEqual(0, _capture.Get(null, null, null, 0).LineCount);

        [Test]
        public void PrintedLines_AreKept_EvenWhenNobodyAskedAtTheTime()
        {
            Print("a");
            _capture.Capture(false);
            Print("b");
            _capture.Capture(false);

            var output = _capture.Get(null, null, null, 0);

            Assert.AreEqual("a\nb", output.Text);
            Assert.AreEqual(2, output.Next);
        }

        [Test]
        public void Since_ReturnsOnlyWhatIsNew()
        {
            Print("a");
            var first = _capture.Get(null, null, null, 0);
            Print("b");

            var second = _capture.Get(first.Next, null, null, 0);

            Assert.AreEqual("b", second.Text);
            Assert.AreEqual(0, _capture.Get(second.Next, null, null, 0).LineCount);
        }

        [Test]
        public void Limit_WithSince_GivesTheOldestFirst_AndSaysThereIsMore()
        {
            Print("1");
            Print("2");
            Print("3");

            var page = _capture.Get(0, 2, null, 0);

            Assert.AreEqual("1\n2", page.Text);
            Assert.IsTrue(page.More);
            Assert.AreEqual("3", _capture.Get(page.Next, 2, null, 0).Text);
        }

        [Test]
        public void Limit_WithoutSince_GivesTheNewest()
        {
            Print("1");
            Print("2");
            Print("3");

            Assert.AreEqual("2\n3", _capture.Get(null, 2, null, 0).Text);
        }

        [Test]
        public void Contains_FiltersLines_IgnoringCase()
        {
            Print("[Orders.Total] qty=2");
            Print("something else");
            Print("[orders.total] qty=3");

            var output = _capture.Get(0, null, "[ORDERS.TOTAL]", 0);

            Assert.AreEqual("[Orders.Total] qty=2\n[orders.total] qty=3", output.Text);
        }

        [Test]
        public void ALineBeingTyped_IsNotOutput_UntilItHasStoodStill()
        {
            _window.Text += "?1+";
            _capture.Capture(false);

            _window.Text += "1\r\n 2 \r\n";

            Assert.AreEqual("?1+1\n 2 ", _capture.Get(null, null, null, 0).Text);
        }

        [Test]
        public void ALastLineWithoutLineBreak_CountsOnceItStaysTheSame()
        {
            _window.Text += "no break";
            _capture.Capture(false);
            _capture.Capture(false);

            Assert.AreEqual("no break", _capture.Get(null, null, null, 0).Text);
        }

        [Test]
        public void TheEmptyPromptLineAfterOutput_IsNotALine_AndTheNextCommandReplacesIt()
        {
            _window.Text += "a\r\n\r\n";
            Assert.AreEqual(1, _capture.Get(null, null, null, 0).LineCount);

            _window.Text = "old line\r\na\r\nFoo\r\n";

            var output = _capture.Get(0, null, null, 0);
            Assert.AreEqual("a\nFoo", output.Text);
            Assert.AreEqual(2, output.Next);
        }

        [Test]
        public void BlankLines_AreKept()
        {
            Print("a");
            Print("");
            Print("b");

            Assert.AreEqual("a\n\nb", _capture.Get(null, null, null, 0).Text);
        }

        [Test]
        public void WhenTheWindowDropsItsOldestLines_OnlyTheNewOnesAreAdded()
        {
            _window.Text = "one\r\ntwo\r\nthree\r\n";
            _capture.Capture(false);
            _window.Text = "two\r\nthree\r\nfour\r\n";

            _capture.Capture(false);

            Assert.AreEqual("four", _capture.Get(null, 1, null, 0).Text);
        }

        [Test]
        public void ClearingTheWindow_DoesNotLoseWhatIsPrintedAfterwards()
        {
            Print("before");
            _capture.Capture(false);
            _window.Text = string.Empty;
            _capture.Capture(false);
            Print("after");

            Assert.AreEqual("after", _capture.Get(null, 1, null, 0).Text);
        }

        [Test]
        public void TheBuffer_KeepsTheNewestLines_AndSaysWhenSinceIsOlderThanThat()
        {
            for (var i = 0; i < ImmediateCapture.Capacity + 10; i++)
            {
                Print("line " + i);
            }

            var output = _capture.Get(0, ImmediateCapture.MaxLimit, null, 0);

            Assert.IsTrue(output.Missed);
            Assert.AreEqual("line 10", output.Text.Split('\n').First());
            Assert.IsTrue(output.More);
        }

        [Test]
        public void Wait_ReturnsAsSoonAsSomethingIsPrinted()
        {
            _whileWaiting = () => { if (_clock >= TimeSpan.FromMilliseconds(300)) { Print("late"); _whileWaiting = null; } };

            var output = _capture.Get(0, null, null, 5000);

            Assert.AreEqual("late", output.Text);
            Assert.Less(_clock, TimeSpan.FromMilliseconds(1000));
        }

        [Test]
        public void Wait_GivesUp_WhenNothingComes()
        {
            var output = _capture.Get(0, null, null, 1000);

            Assert.AreEqual(0, output.LineCount);
            Assert.GreaterOrEqual(_clock, TimeSpan.FromMilliseconds(1000));
        }

        [Test]
        public void AWindowThatCannotBeRead_IsReported_NotAnsweredWithNothing()
        {
            _window.Throws = true;

            Assert.Throws<ApiException>(() => _capture.Get(null, null, null, 0));
        }

        [Test]
        public void ClosedWindow_IsNotAnError_ForTheBackgroundLook()
        {
            _window.Throws = true;

            Assert.IsFalse(_capture.Capture(false));
        }

        [TestCase(0, null)]
        [TestCase(1001, null)]
        [TestCase(null, -1)]
        [TestCase(null, 30001)]
        public void BadNumbers_AreRefused(int? limit, int? waitMs)
            => Assert.Throws<ApiException>(() => _capture.Get(null, limit, null, waitMs ?? 0));
    }

    /// <summary>The user may allow agents to read the Immediate window without letting them run anything.</summary>
    [TestFixture]
    [Category("Api")]
    public class ImmediateReadOnlyTests
    {
        private bool _run, _read;
        private ImmediateFixture _fixture;
        private ApiRouter _router;

        [SetUp]
        public void SetUp()
        {
            _run = false;
            _read = true;
            _fixture = Fakes.Immediate(() => _run, () => _read);
            _router = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), _fixture.Api);
        }

        private ApiResponse Rest(string method, string path, string body = null)
            => _router.Handle(new ApiRequest { Method = method, Path = path, Body = body, Query = new System.Collections.Generic.Dictionary<string, string>(), Client = "t" });

        private JObject Rpc(string method)
            => (JObject)_router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Body = new JObject { ["jsonrpc"] = "2.0", ["id"] = 1, ["method"] = method }.ToString(), Client = "t" }).Body;

        private System.Collections.Generic.List<string> ToolNames()
            => Rpc("tools/list")["result"]["tools"].Select(t => (string)t["name"]).ToList();

        [Test]
        public void Reading_Works_WithoutTheRightToRun()
        {
            _fixture.Window.Text = "hello";

            Assert.AreEqual(200, Rest("GET", "/api/v1/immediate").Status);
            Assert.AreEqual(200, Rest("GET", "/api/v1/immediate/output").Status);
        }

        [Test]
        public void Running_IsRefused_AndNothingIsAsked_AndNothingTyped()
        {
            var response = Rest("POST", "/api/v1/immediate", "{\"command\":\"Kill \\\"x\\\"\"}");

            Assert.AreEqual(403, response.Status);
            StringAssert.Contains("allowed reading", (string)JObject.FromObject(response.Body)["error"]["message"]);
            Assert.AreEqual(0, _fixture.Asked);
            CollectionAssert.IsEmpty(_fixture.Window.Sent);
        }

        [Test]
        public void RunningImpliesReading()
        {
            _read = false;
            _run = true;

            Assert.AreEqual(200, Rest("GET", "/api/v1/immediate/output").Status);
        }

        [Test]
        public void WhenNeitherIsAllowed_BothAreRefused()
        {
            _read = false;

            Assert.AreEqual(403, Rest("GET", "/api/v1/immediate").Status);
            Assert.AreEqual(403, Rest("GET", "/api/v1/immediate/output").Status);
            Assert.AreEqual(403, Rest("POST", "/api/v1/immediate", "{\"command\":\"?1\"}").Status);
        }

        [Test]
        public void Mcp_OffersTheReadingTools_ButNotRun()
        {
            var names = ToolNames();

            CollectionAssert.Contains(names, "read_immediate");
            CollectionAssert.Contains(names, "read_debug_output");
            CollectionAssert.DoesNotContain(names, "run_immediate");
        }

        [Test]
        public void Mcp_Instructions_TellTheAgentToAskTheUserToRunTheCode()
        {
            var instructions = (string)Rpc("initialize")["result"]["instructions"];

            StringAssert.Contains("DEBUGGING WITH Debug.Print", instructions);
            StringAssert.Contains("ask the user to run the code", instructions);
            StringAssert.DoesNotContain("run the code with run_immediate", instructions);
        }

        [Test]
        public void Mcp_WithRunAllowed_OffersEverything()
        {
            _run = true;

            CollectionAssert.IsSupersetOf(ToolNames(), new[] { "read_immediate", "read_debug_output", "run_immediate" });
            StringAssert.Contains("run the code with run_immediate", (string)Rpc("initialize")["result"]["instructions"]);
        }

        [Test]
        public void Mcp_WithNeither_OffersNothingOfIt_AndGivesNoAdvice()
        {
            _read = false;

            CollectionAssert.DoesNotContain(ToolNames(), "read_debug_output");
            CollectionAssert.DoesNotContain(ToolNames(), "read_immediate");
            StringAssert.DoesNotContain("DEBUGGING WITH Debug.Print", (string)Rpc("initialize")["result"]["instructions"]);
        }
    }

    [TestFixture]
    [Category("Api")]
    public class ImmediateOutputRouteTests
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

        private ApiResponse Get(string query = "")
            => _router.Handle(new ApiRequest
            {
                Method = "GET",
                Path = "/api/v1/immediate/output",
                Query = query.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries).Select(p => p.Split('=')).ToDictionary(p => p[0], p => p[1], StringComparer.OrdinalIgnoreCase),
                Client = "t"
            });

        [Test]
        public void Get_AnswersWithTheCursor()
        {
            var answer = (ImmediateOutput)Get("limit=5").Body;

            Assert.AreEqual(0, answer.Next);
            Assert.AreEqual(0, answer.LineCount);
        }

        [Test]
        public void Get_IsForbidden_WhileTheImmediateWindowIsOffLimits()
        {
            _enabled = false;

            Assert.AreEqual(403, Get().Status);
        }

        [Test]
        public void Get_RefusesANumberThatIsNone()
            => Assert.AreEqual(400, Get("since=abc").Status);
    }
}
