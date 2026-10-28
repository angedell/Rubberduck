using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    internal class FakeEncapsulateField : IEncapsulateFieldSource
    {
        public EncapsulateFieldPreview Shown = new EncapsulateFieldPreview
        {
            Module = "Module1",
            Field = "count",
            Strategy = EncapsulateFieldRequest.BackingField,
            Valid = true,
            Code = "Public Property Get Count() As Long",
            Properties = { new EncapsulatedProperty { Field = "count", PropertyName = "Count", BackingField = "mCount", CanBeReadWrite = true } }
        };

        public List<string> Calls = new List<string>();
        public EncapsulateFieldRequest LastRequest;

        public EncapsulateFieldPreview Preview(string project, EncapsulateFieldRequest request)
        {
            Calls.Add("preview");
            LastRequest = request;
            return Shown;
        }

        public EncapsulateFieldResult Apply(string project, EncapsulateFieldRequest request)
        {
            Calls.Add("apply");
            LastRequest = request;
            return new EncapsulateFieldResult { Project = project, Module = "Module1", Field = "count", Properties = Shown.Properties };
        }
    }

    [TestFixture]
    [Category("Api")]
    public class RefactorEncapsulateFieldApiTests
    {
        private FakeEncapsulateField _source;
        private int _asked;
        private bool _allow;
        private RefactorEncapsulateFieldApi _api;

        [SetUp]
        public void SetUp()
        {
            _source = new FakeEncapsulateField();
            _asked = 0;
            _allow = true;
            _api = new RefactorEncapsulateFieldApi(_source, new ApiWriteApproval(_ => { _asked++; return _allow; }));
        }

        private static EncapsulateFieldRequest Request() => new EncapsulateFieldRequest { Module = "Module1", Name = "count" };

        [Test]
        public void Preview_ShowsTheCode_ChangesNothing_AsksNobody()
        {
            var preview = _api.Preview("p", Request());

            StringAssert.Contains("Property Get Count", preview.Code);
            CollectionAssert.AreEqual(new[] { "preview" }, _source.Calls);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Encapsulate_LooksFirst_ThenAsksTheUser_ThenActs()
        {
            var result = _api.Encapsulate("agent", "p", Request());

            Assert.AreEqual("count", result.Field);
            CollectionAssert.AreEqual(new[] { "preview", "apply" }, _source.Calls);
            Assert.AreEqual(1, _asked);
        }

        [Test]
        public void Encapsulate_WithoutApproval_IsForbidden_AndChangesNothing()
        {
            _allow = false;

            Assert.AreEqual(403, Assert.Throws<ApiException>(() => _api.Encapsulate("agent", "p", Request())).Status);
            CollectionAssert.DoesNotContain(_source.Calls, "apply");
        }

        [Test]
        public void AModuleIsRequired()
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Encapsulate("agent", "p", new EncapsulateFieldRequest { Name = "count" })).Status);
            Assert.IsEmpty(_source.Calls);
        }

        [Test]
        public void AnUnknownStrategy_IsABadRequest()
        {
            var request = Request();
            request.Strategy = "magic";

            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Preview("p", request)).Status);
        }

        [TestCase("backingField")]
        [TestCase("UDTMEMBER")]
        public void TheTwoStrategies_AreAccepted(string strategy)
        {
            var request = Request();
            request.Strategy = strategy;

            Assert.DoesNotThrow(() => _api.Preview("p", request));
        }

        [Test]
        public void AnInvalidEncapsulation_IsRefused_BeforeTheUserIsAsked()
        {
            _source.Shown.Valid = false;
            _source.Shown.Problems.Add("'Count' clashes with a procedure.");

            var failure = Assert.Throws<ApiException>(() => _api.Encapsulate("agent", "p", Request()));

            Assert.AreEqual(422, failure.Status);
            StringAssert.Contains("clashes", failure.Message);
            Assert.AreEqual(0, _asked);
            CollectionAssert.DoesNotContain(_source.Calls, "apply");
        }
    }

    [TestFixture]
    [Category("Api")]
    public class RefactorEncapsulateFieldRouteTests
    {
        private FakeEncapsulateField _source;
        private int _asked;
        private ApiRouter _router;

        [SetUp]
        public void SetUp()
        {
            _source = new FakeEncapsulateField();
            _asked = 0;
            var approval = new ApiWriteApproval(_ => { _asked++; return true; });
            _router = new ApiRouter(new ModuleApi(new FakeStore(), approval), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api,
                renameRefactor: new RefactorRenameApi(new FakeRename(), approval),
                encapsulateRefactor: new RefactorEncapsulateFieldApi(_source, approval));
        }

        private ApiResponse Post(string path, string body, string method = "POST")
            => _router.Handle(new ApiRequest { Method = method, Path = ApiRouter.Base + path, Body = body, Client = "agent" });

        private const string Body = "{\"module\":\"Module1\",\"name\":\"count\",\"propertyName\":\"Total\",\"readOnly\":true,\"strategy\":\"backingField\"}";

        [Test]
        public void Preview_Route_PassesTheOptions()
        {
            var response = Post("/projects/p/refactor/encapsulate-field/preview", Body);

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual("Total", _source.LastRequest.PropertyName);
            Assert.AreEqual(true, _source.LastRequest.ReadOnly);
            Assert.AreEqual("backingField", _source.LastRequest.Strategy);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Encapsulate_Route_Acts()
        {
            var response = Post("/projects/p/refactor/encapsulate-field", Body);

            Assert.AreEqual(200, response.Status);
            Assert.IsInstanceOf<EncapsulateFieldResult>(response.Body);
            Assert.AreEqual(1, _asked);
        }

        [Test]
        public void ReadOnly_IsOptional()
        {
            Post("/projects/p/refactor/encapsulate-field/preview", "{\"module\":\"Module1\",\"name\":\"count\"}");

            Assert.IsNull(_source.LastRequest.ReadOnly);
            Assert.IsNull(_source.LastRequest.PropertyName);
        }

        [Test]
        public void BadRequests_AreRefused()
        {
            Assert.AreEqual(405, Post("/projects/p/refactor/encapsulate-field", Body, "GET").Status);
            Assert.AreEqual(400, Post("/projects/p/refactor/encapsulate-field", null).Status);
            Assert.AreEqual(400, Post("/projects/p/refactor/encapsulate-field", "{\"module\":\"M\",\"readOnly\":\"yes\"}").Status);
            Assert.AreEqual(400, Post("/projects/p/refactor/encapsulate-field", "{\"module\":\"M\",\"propertyName\":3}").Status);
            Assert.AreEqual(404, Post("/projects/p/refactor/encapsulate-field/other", Body).Status);
        }

        [Test]
        public void WithoutTheFeature_TheRouteIsNotFound()
        {
            var bare = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api);

            var response = bare.Handle(new ApiRequest { Method = "POST", Path = ApiRouter.Base + "/projects/p/refactor/encapsulate-field", Body = Body, Client = "agent" });

            Assert.AreEqual(404, response.Status);
        }

        private JObject Rpc(string method, JObject parameters = null)
        {
            var body = new JObject { ["jsonrpc"] = "2.0", ["id"] = 1, ["method"] = method };
            if (parameters != null) body["params"] = parameters;
            return JObject.FromObject(_router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Body = body.ToString(), Client = "agent" }).Body);
        }

        [Test]
        public void TheTools_AreListed_NextToTheOthers()
        {
            var names = Rpc("tools/list")["result"]["tools"].Select(t => t.Value<string>("name")).ToList();

            CollectionAssert.IsSubsetOf(new[] { "preview_encapsulate_field", "encapsulate_field", "preview_rename", "rename_symbol", "read_module" }, names);
        }

        [Test]
        public void TheTools_Work()
        {
            var arguments = new JObject { ["project"] = "p", ["module"] = "Module1", ["name"] = "count", ["readOnly"] = false };

            var previewed = Rpc("tools/call", new JObject { ["name"] = "preview_encapsulate_field", ["arguments"] = arguments })["result"];
            var done = Rpc("tools/call", new JObject { ["name"] = "encapsulate_field", ["arguments"] = arguments })["result"];

            Assert.IsFalse(previewed.Value<bool>("isError"));
            StringAssert.Contains("Property Get Count", previewed["content"][0].Value<string>("text"));
            Assert.IsFalse(done.Value<bool>("isError"));
            Assert.AreEqual(false, _source.LastRequest.ReadOnly);
        }

        [Test]
        public void ATool_NeedsAProject()
        {
            var result = Rpc("tools/call", new JObject { ["name"] = "encapsulate_field", ["arguments"] = new JObject { ["module"] = "Module1" } })["result"];

            Assert.IsTrue(result.Value<bool>("isError"));
            CollectionAssert.IsEmpty(_source.Calls);
        }
    }
}
