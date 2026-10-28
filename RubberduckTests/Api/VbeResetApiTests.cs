using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    [Category("Api")]
    public class VbeResetApiTests
    {
        private string _mode;
        private string _modeAfterReset;
        private bool _resetResult;
        private int _resets;
        private int _asked;
        private bool _allow;
        private BlockingDialog _dialog;
        private ApiRouter _router;

        private class Dialogs : IDialogWatcher
        {
            private readonly VbeResetApiTests _owner;
            public Dialogs(VbeResetApiTests owner) => _owner = owner;
            public BlockingDialog Find() => _owner._dialog;
            public bool Acknowledge(BlockingDialog dialog) => false;
        }

        [SetUp]
        public void SetUp()
        {
            _mode = ParserSync.ModeRunningOrBreak;
            _modeAfterReset = ParserSync.ModeDesign;
            _resetResult = true;
            _resets = 0;
            _asked = 0;
            _allow = true;
            _dialog = null;

            var reset = new VbeResetApi(
                new ApiWriteApproval(_ => { _asked++; return _allow; }),
                () => _mode,
                () => { _resets++; _mode = _modeAfterReset; return _resetResult; },
                new Dialogs(this));
            _router = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api, vbeReset: reset);
        }

        private static JObject Json(ApiResponse response) => JObject.FromObject(response.Body);

        private ApiResponse Send(string method)
            => _router.Handle(new ApiRequest { Method = method, Path = "/api/v1/vbe/reset", Client = "agent" });

        [Test]
        public void Reset_WhenStopped_PressesReset_AndReportsDesign()
        {
            var response = Send("POST");

            Assert.AreEqual(200, response.Status);
            Assert.IsTrue((bool)Json(response)["Reset"]);
            Assert.AreEqual("design", (string)Json(response)["VbeMode"]);
            Assert.AreEqual(1, _resets);
            Assert.AreEqual(1, _asked);
        }

        [Test]
        public void Reset_WhenInDesignMode_DoesNothing_AndAsksNothing()
        {
            _mode = ParserSync.ModeDesign;

            var body = Json(Send("POST"));

            Assert.IsFalse((bool)body["Reset"]);
            Assert.AreEqual("design", (string)body["VbeMode"]);
            Assert.AreEqual(0, _resets);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Reset_WhenUserSaysNo_Is403_AndPressesNothing()
        {
            _allow = false;

            Assert.AreEqual(403, Send("POST").Status);
            Assert.AreEqual(0, _resets);
        }

        [Test]
        public void Reset_WhenModeStaysStopped_Is409_NotSuccess()
        {
            _modeAfterReset = ParserSync.ModeRunningOrBreak;
            _resetResult = false;

            var response = Send("POST");

            Assert.AreEqual(409, response.Status);
            StringAssert.Contains(VbeResetApi.ResetFailed, Json(response).ToString());
            StringAssert.Contains("running_or_break", Json(response).ToString());
        }

        [Test]
        public void Reset_WhenADialogBlocks_Is409_AndQuotesTheDialog()
        {
            _modeAfterReset = ParserSync.ModeUnresponsive;
            _resetResult = false;
            _dialog = new BlockingDialog { Title = "Microsoft Visual Basic", Message = "Run-time error '11'" };

            var response = Send("POST");

            Assert.AreEqual(409, response.Status);
            StringAssert.Contains("Run-time error '11'", Json(response).ToString());
        }

        [Test]
        public void Reset_ReportedDesignButResultFalse_Is409()
        {
            _resetResult = false;

            Assert.AreEqual(409, Send("POST").Status);
        }

        [Test]
        public void Reset_RejectsOtherMethods()
        {
            Assert.AreEqual(405, Send("GET").Status);
        }

        [Test]
        public void WithoutResetSupport_TheRouteDoesNotExist()
        {
            var router = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api);

            Assert.AreEqual(404, router.Handle(new ApiRequest { Method = "POST", Path = "/api/v1/vbe/reset", Client = "agent" }).Status);
        }

        [Test]
        public void Mcp_OffersResetVbe_AndRunsIt()
        {
            JObject Call(object message)
                => Json(_router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Body = JObject.FromObject(message).ToString(), Client = "agent" }));

            var tools = Call(new { jsonrpc = "2.0", id = 1, method = "tools/list" })["result"]["tools"].Select(t => (string)t["name"]).ToList();
            var result = Call(new { jsonrpc = "2.0", id = 2, method = "tools/call", @params = new { name = "reset_vbe", arguments = new { } } });

            CollectionAssert.Contains(tools, "reset_vbe");
            Assert.IsFalse((bool)result["result"]["isError"]);
            StringAssert.Contains("design", (string)result["result"]["content"][0]["text"]);
        }
    }
}
