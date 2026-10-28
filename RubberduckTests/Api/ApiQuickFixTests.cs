using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    internal class FakeQuickFixes : IQuickFixSource
    {
        public List<string> Calls = new List<string>();
        public string LastScope;
        public bool Stale;

        public QuickFixOptions List(string project, string result)
        {
            Calls.Add("list");
            if (Stale) throw ApiException.Conflict("The diagnostics are out of date.");
            if (result == "nope") throw ApiException.NotFound("No diagnostic has that id in this project.");
            return new QuickFixOptions
            {
                Project = project,
                Result = result,
                Module = "Module1",
                Inspection = "Unused variable",
                Fixes = new List<QuickFixOption>
                {
                    new QuickFixOption { Name = "RemoveUnusedDeclarationQuickFix", Description = "Remove", CanFixInModule = true },
                    new QuickFixOption { Name = "IgnoreOnceQuickFix", Description = "Ignore once", CanFixInModule = false }
                }
            };
        }

        public QuickFixOutcome Apply(string project, string result, string fix, string scope)
        {
            Calls.Add("apply");
            LastScope = scope;
            return new QuickFixOutcome { Project = project, Result = result, Fix = fix, Scope = scope, Fixed = scope == "module" ? 3 : 1 };
        }
    }

    [TestFixture]
    [Category("Api")]
    public class QuickFixApiTests
    {
        private FakeQuickFixes _source;
        private int _asked;
        private bool _allow;
        private QuickFixApi _api;

        [SetUp]
        public void SetUp()
        {
            _source = new FakeQuickFixes();
            _asked = 0;
            _allow = true;
            _api = new QuickFixApi(_source, new ApiWriteApproval(_ => { _asked++; return _allow; }));
        }

        [Test]
        public void List_NeedsAResultId_AndAsksNobody()
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.List("p", " ")).Status);

            var options = _api.List("p", "abc");

            Assert.AreEqual(2, options.Fixes.Count);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Apply_AsksTheUser_ThenFixesOneResult()
        {
            var outcome = _api.Apply("agent", "p", "abc", "removeunuseddeclarationquickfix", null);

            Assert.AreEqual(1, _asked);
            Assert.AreEqual("result", _source.LastScope);
            Assert.AreEqual("RemoveUnusedDeclarationQuickFix", outcome.Fix);
            Assert.AreEqual(1, outcome.Fixed);
        }

        [Test]
        public void Apply_ToTheModule_FixesAllOfThem()
        {
            var outcome = _api.Apply("agent", "p", "abc", "RemoveUnusedDeclarationQuickFix", "Module");

            Assert.AreEqual("module", _source.LastScope);
            Assert.AreEqual(3, outcome.Fixed);
        }

        [Test]
        public void Apply_ToTheModule_IsRefusedForAFixThatCannot_BeforeAskingTheUser()
        {
            var failure = Assert.Throws<ApiException>(() => _api.Apply("agent", "p", "abc", "IgnoreOnceQuickFix", "module"));

            Assert.AreEqual(422, failure.Status);
            Assert.AreEqual(0, _asked);
            CollectionAssert.DoesNotContain(_source.Calls, "apply");
        }

        [Test]
        public void Apply_AnUnknownFix_IsRefused_ListingTheRealOnes()
        {
            var failure = Assert.Throws<ApiException>(() => _api.Apply("agent", "p", "abc", "Whatever", null));

            Assert.AreEqual(422, failure.Status);
            StringAssert.Contains("RemoveUnusedDeclarationQuickFix", failure.Message);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Apply_BadScope_OrMissingFields_AreBadRequests()
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Apply("agent", "p", "abc", "X", "project")).Status);
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Apply("agent", "p", "abc", null, null)).Status);
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Apply("agent", "p", null, "X", null)).Status);
        }

        [Test]
        public void Apply_WhenTheUserSaysNo_ChangesNothing()
        {
            _allow = false;

            var failure = Assert.Throws<ApiException>(() => _api.Apply("agent", "p", "abc", "RemoveUnusedDeclarationQuickFix", null));

            Assert.AreEqual(403, failure.Status);
            CollectionAssert.DoesNotContain(_source.Calls, "apply");
        }

        [Test]
        public void StaleDiagnostics_AreAConflict_NotAWrongFix()
        {
            _source.Stale = true;

            Assert.AreEqual(409, Assert.Throws<ApiException>(() => _api.Apply("agent", "p", "abc", "RemoveUnusedDeclarationQuickFix", null)).Status);
            CollectionAssert.DoesNotContain(_source.Calls, "apply");
        }
    }

    [TestFixture]
    [Category("Api")]
    public class QuickFixRouteTests
    {
        private FakeQuickFixes _source;
        private ApiRouter _router;

        [SetUp]
        public void SetUp()
        {
            _source = new FakeQuickFixes();
            var approval = new ApiWriteApproval(_ => true);
            _router = new ApiRouter(new ModuleApi(new FakeStore(), approval), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api,
                quickFix: new QuickFixApi(_source, approval));
        }

        private ApiResponse Send(string method, string path, string body = null, string result = null)
        {
            var request = new ApiRequest { Method = method, Path = ApiRouter.Base + path, Body = body, Client = "agent" };
            if (result != null) request.Query["result"] = result;
            return _router.Handle(request);
        }

        [Test]
        public void Get_ListsTheFixes()
        {
            var response = Send("GET", "/projects/p/quickfix", result: "abc");

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual(2, ((QuickFixOptions)response.Body).Fixes.Count);
        }

        [Test]
        public void Get_WithoutAResult_IsABadRequest_AndAnUnknownOneIsNotFound()
        {
            Assert.AreEqual(400, Send("GET", "/projects/p/quickfix").Status);
            Assert.AreEqual(404, Send("GET", "/projects/p/quickfix", result: "nope").Status);
        }

        [Test]
        public void Post_AppliesTheFix()
        {
            var response = Send("POST", "/projects/p/quickfix", "{\"result\":\"abc\",\"fix\":\"RemoveUnusedDeclarationQuickFix\",\"scope\":\"module\"}");

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual(3, ((QuickFixOutcome)response.Body).Fixed);
        }

        [Test]
        public void Post_NeedsABody_AndOtherMethodsAreRefused()
        {
            Assert.AreEqual(400, Send("POST", "/projects/p/quickfix").Status);
            Assert.AreEqual(405, Send("DELETE", "/projects/p/quickfix").Status);
        }

        [Test]
        public void WithoutTheFeature_TheRouteIsNotFound()
        {
            var bare = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api);

            Assert.AreEqual(404, bare.Handle(new ApiRequest { Method = "GET", Path = ApiRouter.Base + "/projects/p/quickfix", Client = "agent" }).Status);
        }

        private JObject Rpc(string method, JObject parameters = null)
        {
            var body = new JObject { ["jsonrpc"] = "2.0", ["id"] = 1, ["method"] = method };
            if (parameters != null) body["params"] = parameters;
            return JObject.FromObject(_router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Body = body.ToString(), Client = "agent" }).Body);
        }

        [Test]
        public void TheTools_AreListed_AndWork()
        {
            var tools = Rpc("tools/list")["result"]["tools"].Select(t => t.Value<string>("name")).ToList();
            CollectionAssert.IsSubsetOf(new[] { "list_quick_fixes", "apply_quick_fix" }, tools);

            var listed = Rpc("tools/call", new JObject { ["name"] = "list_quick_fixes", ["arguments"] = new JObject { ["project"] = "p", ["result"] = "abc" } })["result"];
            var applied = Rpc("tools/call", new JObject
            {
                ["name"] = "apply_quick_fix",
                ["arguments"] = new JObject { ["project"] = "p", ["result"] = "abc", ["fix"] = "RemoveUnusedDeclarationQuickFix" }
            })["result"];

            Assert.IsFalse(listed.Value<bool>("isError"));
            Assert.IsFalse(applied.Value<bool>("isError"));
            CollectionAssert.AreEqual(new[] { "list", "list", "apply" }, _source.Calls);
        }
    }
}
