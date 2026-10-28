using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    internal class FakeMoveToFolder : IMoveToFolderSource
    {
        public List<string> Calls = new List<string>();
        public MoveToFolderRequest LastRequest;
        public bool Changes = true;
        public bool Valid = true;

        public MoveToFolderPreview Preview(string project, MoveToFolderRequest request)
        {
            Calls.Add("preview");
            LastRequest = request;
            return new MoveToFolderPreview
            {
                Project = project,
                Folder = request.Folder,
                Valid = Valid,
                Problems = Valid ? new List<string>() : new List<string> { "A folder name cannot be empty." },
                Moves = request.Modules.Select(m => new ModuleMove { Module = m, FromFolder = "VBAProject", Changes = Changes }).ToList()
            };
        }

        public MoveToFolderResult Apply(string project, MoveToFolderRequest request)
        {
            Calls.Add("apply");
            LastRequest = request;
            return new MoveToFolderResult { Project = project, Folder = request.Folder, Moved = request.Modules };
        }
    }

    [TestFixture]
    [Category("Api")]
    public class RefactorMoveToFolderApiTests
    {
        private FakeMoveToFolder _source;
        private int _asked;
        private bool _allow;
        private RefactorMoveToFolderApi _api;

        [SetUp]
        public void SetUp()
        {
            _source = new FakeMoveToFolder();
            _asked = 0;
            _allow = true;
            _api = new RefactorMoveToFolderApi(_source, new ApiWriteApproval(_ => { _asked++; return _allow; }));
        }

        private static MoveToFolderRequest Request(params string[] modules)
            => new MoveToFolderRequest { Modules = modules.Length == 0 ? new List<string> { "Module1" } : modules.ToList(), Folder = "App.Core" };

        [Test]
        public void Preview_ChangesNothing_AsksNobody()
        {
            var preview = _api.Preview("p", Request());

            Assert.IsTrue(preview.Moves.Single().Changes);
            CollectionAssert.AreEqual(new[] { "preview" }, _source.Calls);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Move_LooksFirst_ThenAsksTheUser_ThenMoves()
        {
            var result = _api.Move("agent", "p", Request("Module1", "Class1"));

            CollectionAssert.AreEqual(new[] { "Module1", "Class1" }, result.Moved);
            CollectionAssert.AreEqual(new[] { "preview", "apply" }, _source.Calls);
            Assert.AreEqual(1, _asked);
        }

        [Test]
        public void Move_WithoutApproval_IsForbidden_AndMovesNothing()
        {
            _allow = false;

            Assert.AreEqual(403, Assert.Throws<ApiException>(() => _api.Move("agent", "p", Request())).Status);
            CollectionAssert.DoesNotContain(_source.Calls, "apply");
        }

        [Test]
        public void WhenNothingWouldChange_NothingIsWritten_AndNobodyIsAsked()
        {
            _source.Changes = false;

            var result = _api.Move("agent", "p", Request());

            CollectionAssert.AreEqual(new[] { "Module1" }, result.Unchanged);
            CollectionAssert.IsEmpty(result.Moved);
            CollectionAssert.AreEqual(new[] { "preview" }, _source.Calls);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void AnInvalidFolder_IsRefused_BeforeTheUserIsAsked()
        {
            _source.Valid = false;

            var failure = Assert.Throws<ApiException>(() => _api.Move("agent", "p", Request()));

            Assert.AreEqual(422, failure.Status);
            Assert.AreEqual(0, _asked);
            CollectionAssert.DoesNotContain(_source.Calls, "apply");
        }

        [Test]
        public void ARequestNeedsModulesAndAFolder()
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Move("agent", "p", new MoveToFolderRequest { Folder = "A" })).Status);
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Move("agent", "p", new MoveToFolderRequest { Modules = { "M" } })).Status);
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Move("agent", "p", new MoveToFolderRequest { Modules = { "M" }, Folder = "" })).Status);
            Assert.IsEmpty(_source.Calls);
        }

        [Test]
        public void TooManyModules_AreRefused()
        {
            var many = Enumerable.Range(1, MoveToFolderRequest.MaxModules + 1).Select(i => "M" + i).ToArray();

            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Preview("p", Request(many))).Status);
        }

        [TestCase("App.Core", 0)]
        [TestCase("Solo", 0)]
        [TestCase("", 1)]
        [TestCase("App..Core", 1)]
        [TestCase(".App", 1)]
        [TestCase("App.", 1)]
        [TestCase("App\nCore", 1)]
        public void FolderNames_AreCheckedLikeTheDialogDoes(string folder, int problems)
        {
            Assert.AreEqual(problems, RefactorMoveToFolderApi.FolderProblems(folder).Count);
        }
    }

    [TestFixture]
    [Category("Api")]
    public class RefactorMoveToFolderRouteTests
    {
        private FakeMoveToFolder _source;
        private int _asked;
        private ApiRouter _router;

        [SetUp]
        public void SetUp()
        {
            _source = new FakeMoveToFolder();
            _asked = 0;
            var approval = new ApiWriteApproval(_ => { _asked++; return true; });
            _router = new ApiRouter(new ModuleApi(new FakeStore(), approval), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api,
                renameRefactor: new RefactorRenameApi(new FakeRename(), approval),
                encapsulateRefactor: new RefactorEncapsulateFieldApi(new FakeEncapsulateField(), approval),
                moveRefactor: new RefactorMoveToFolderApi(_source, approval));
        }

        private ApiResponse Post(string path, string body, string method = "POST")
            => _router.Handle(new ApiRequest { Method = method, Path = ApiRouter.Base + path, Body = body, Client = "agent" });

        [Test]
        public void Preview_Route_TakesOneModuleOrAList_WithoutRepeats()
        {
            var response = Post("/projects/p/refactor/move-to-folder/preview", "{\"module\":\"Module1\",\"modules\":[\"Class1\",\"module1\"],\"folder\":\"App.Core\"}");

            Assert.AreEqual(200, response.Status);
            CollectionAssert.AreEqual(new[] { "Module1", "Class1" }, _source.LastRequest.Modules);
            Assert.AreEqual("App.Core", _source.LastRequest.Folder);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Move_Route_Moves()
        {
            var response = Post("/projects/p/refactor/move-to-folder", "{\"module\":\"Module1\",\"folder\":\"App.Core\"}");

            Assert.AreEqual(200, response.Status);
            Assert.IsInstanceOf<MoveToFolderResult>(response.Body);
            Assert.AreEqual(1, _asked);
        }

        [Test]
        public void BadRequests_AreRefused()
        {
            Assert.AreEqual(405, Post("/projects/p/refactor/move-to-folder", "{}", "GET").Status);
            Assert.AreEqual(400, Post("/projects/p/refactor/move-to-folder", null).Status);
            Assert.AreEqual(400, Post("/projects/p/refactor/move-to-folder", "{\"modules\":\"Module1\",\"folder\":\"A\"}").Status);
            Assert.AreEqual(400, Post("/projects/p/refactor/move-to-folder", "{\"modules\":[1],\"folder\":\"A\"}").Status);
            Assert.AreEqual(400, Post("/projects/p/refactor/move-to-folder", "{\"folder\":\"A\"}").Status);
            Assert.AreEqual(404, Post("/projects/p/refactor/move-to-folder/other", "{}").Status);
        }

        [Test]
        public void WithoutTheFeature_TheRouteIsNotFound()
        {
            var bare = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api);

            var response = bare.Handle(new ApiRequest { Method = "POST", Path = ApiRouter.Base + "/projects/p/refactor/move-to-folder", Body = "{}", Client = "agent" });

            Assert.AreEqual(404, response.Status);
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
            var names = Rpc("tools/list")["result"]["tools"].Select(t => t.Value<string>("name")).ToList();
            CollectionAssert.IsSubsetOf(new[] { "preview_move_to_folder", "move_to_folder", "rename_symbol", "encapsulate_field" }, names);

            var arguments = new JObject { ["project"] = "p", ["modules"] = new JArray("Module1", "Class1"), ["folder"] = "App.Core" };
            var moved = Rpc("tools/call", new JObject { ["name"] = "move_to_folder", ["arguments"] = arguments })["result"];

            Assert.IsFalse(moved.Value<bool>("isError"));
            CollectionAssert.AreEqual(new[] { "Module1", "Class1" }, _source.LastRequest.Modules);
            Assert.AreEqual(1, _asked);
        }

        [Test]
        public void ATool_ReportsRefusalsInsideTheResult()
        {
            _source.Valid = false;
            var arguments = new JObject { ["project"] = "p", ["module"] = "Module1", ["folder"] = "App..Core" };

            var result = Rpc("tools/call", new JObject { ["name"] = "move_to_folder", ["arguments"] = arguments })["result"];

            Assert.IsTrue(result.Value<bool>("isError"));
            StringAssert.Contains("refactoring_refused", result["content"][0].Value<string>("text"));
        }
    }
}
