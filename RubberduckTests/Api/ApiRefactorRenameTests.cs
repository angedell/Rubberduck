using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    internal class FakeRename : IRenameSource
    {
        public RenamePreview Preview_ = new RenamePreview
        {
            Valid = true,
            Target = new RenameTarget { Module = "Module1", Name = "Alpha", Kind = "Function", Line = 2 },
            ReferenceCount = 3,
            ModulesAffected = new List<string> { "Module1", "Module2" }
        };

        public List<string> Calls = new List<string>();
        public RenameRequest LastRequest;

        public RenamePreview Preview(string project, RenameRequest request)
        {
            Calls.Add("preview");
            LastRequest = request;
            Preview_.NewName = request.NewName;
            return Preview_;
        }

        public RenameResult Apply(string project, RenameRequest request)
        {
            Calls.Add("apply");
            LastRequest = request;
            return new RenameResult { Project = project, OldName = "Alpha", NewName = request.NewName, ReferenceCount = Preview_.ReferenceCount };
        }
    }

    [TestFixture]
    [Category("Api")]
    public class RefactorRenameApiTests
    {
        private FakeRename _source;
        private int _asked;
        private bool _allow;
        private RefactorRenameApi _api;

        [SetUp]
        public void SetUp()
        {
            _source = new FakeRename();
            _asked = 0;
            _allow = true;
            _api = new RefactorRenameApi(_source, new ApiWriteApproval(_ => { _asked++; return _allow; }));
        }

        private static RenameRequest Request(string newName = "Beta") => new RenameRequest { Module = "Module1", Name = "Alpha", NewName = newName };

        [Test]
        public void Preview_ChangesNothing_AndAsksNobody()
        {
            var preview = _api.Preview("p", Request());

            Assert.AreEqual(3, preview.ReferenceCount);
            CollectionAssert.AreEqual(new[] { "preview" }, _source.Calls);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Rename_LooksFirst_ThenAsksTheUser_ThenRenames()
        {
            var result = _api.Rename("agent", "p", Request());

            Assert.AreEqual("Beta", result.NewName);
            CollectionAssert.AreEqual(new[] { "preview", "apply" }, _source.Calls);
            Assert.AreEqual(1, _asked);
        }

        [Test]
        public void Rename_WithoutApproval_IsForbidden_AndRenamesNothing()
        {
            _allow = false;

            var failure = Assert.Throws<ApiException>(() => _api.Rename("agent", "p", Request()));

            Assert.AreEqual(403, failure.Status);
            CollectionAssert.DoesNotContain(_source.Calls, "apply");
        }

        [TestCase(null, "Beta")]
        [TestCase(" ", "Beta")]
        [TestCase("Module1", null)]
        [TestCase("Module1", " ")]
        public void ARequestNeedsAModuleAndANewName(string module, string newName)
        {
            var failure = Assert.Throws<ApiException>(() => _api.Rename("agent", "p", new RenameRequest { Module = module, Name = "Alpha", NewName = newName }));

            Assert.AreEqual(400, failure.Status);
            Assert.AreEqual(0, _asked);
            Assert.IsEmpty(_source.Calls);
        }

        [Test]
        public void AnInvalidName_IsRefused_BeforeTheUserIsAsked()
        {
            _source.Preview_.Valid = false;
            _source.Preview_.Problems.Add("Names cannot start with a digit.");

            var failure = Assert.Throws<ApiException>(() => _api.Rename("agent", "p", Request("1Beta")));

            Assert.AreEqual(422, failure.Status);
            Assert.AreEqual(0, _asked);
            CollectionAssert.DoesNotContain(_source.Calls, "apply");
        }

        [Test]
        public void ConflictingNames_NeedAllowConflicts()
        {
            _source.Preview_.Conflicts.Add("VBAProject.Module2.Beta");

            Assert.AreEqual(409, Assert.Throws<ApiException>(() => _api.Rename("agent", "p", Request())).Status);
            Assert.AreEqual(0, _asked);

            var request = Request();
            request.AllowConflicts = true;
            _api.Rename("agent", "p", request);

            CollectionAssert.Contains(_source.Calls, "apply");
        }

        [Test]
        public void ADifferentTarget_NeedsToBeAccepted()
        {
            _source.Preview_.ResolvedTarget = new RenameTarget { Module = "IShape", Name = "Area", Kind = "Function", Line = 1 };
            _source.Preview_.ResolvedBecause = "interface member";

            var failure = Assert.Throws<ApiException>(() => _api.Rename("agent", "p", Request()));

            Assert.AreEqual(409, failure.Status);
            StringAssert.Contains("acceptResolvedTarget", failure.Message);

            var request = Request();
            request.AcceptResolvedTarget = true;
            _api.Rename("agent", "p", request);

            CollectionAssert.Contains(_source.Calls, "apply");
        }
    }

    [TestFixture]
    [Category("Api")]
    public class RefactorRenameRouteTests
    {
        private FakeRename _source;
        private int _asked;
        private ApiRouter _router;

        [SetUp]
        public void SetUp()
        {
            _source = new FakeRename();
            _asked = 0;
            var approval = new ApiWriteApproval(_ => { _asked++; return true; });
            _router = new ApiRouter(new ModuleApi(new FakeStore(), approval), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api,
                renameRefactor: new RefactorRenameApi(_source, approval));
        }

        private ApiResponse Post(string path, string body, string method = "POST")
            => _router.Handle(new ApiRequest { Method = method, Path = ApiRouter.Base + path, Body = body, Client = "agent" });

        private const string Body = "{\"module\":\"Module1\",\"name\":\"Alpha\",\"line\":2,\"newName\":\"Beta\",\"allowConflicts\":true}";

        [Test]
        public void Preview_Route_AnswersWhatWouldHappen()
        {
            var response = Post("/projects/p/refactor/rename/preview", Body);

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual(3, ((RenamePreview)response.Body).ReferenceCount);
            Assert.AreEqual(0, _asked);
            Assert.AreEqual(2, _source.LastRequest.Line);
            Assert.IsTrue(_source.LastRequest.AllowConflicts);
        }

        [Test]
        public void Rename_Route_Renames()
        {
            var response = Post("/projects/p/refactor/rename", Body);

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual("Beta", ((RenameResult)response.Body).NewName);
            Assert.AreEqual(1, _asked);
        }

        [Test]
        public void Rename_Route_WantsPost_AndABody_AndWellTypedFields()
        {
            Assert.AreEqual(405, Post("/projects/p/refactor/rename", Body, "GET").Status);
            Assert.AreEqual(405, Post("/projects/p/refactor/rename/preview", Body, "PUT").Status);
            Assert.AreEqual(400, Post("/projects/p/refactor/rename", null).Status);
            Assert.AreEqual(400, Post("/projects/p/refactor/rename", "{\"module\":\"M\",\"newName\":\"B\",\"line\":\"two\"}").Status);
            Assert.AreEqual(400, Post("/projects/p/refactor/rename", "{\"module\":\"M\",\"newName\":\"B\",\"allowConflicts\":\"yes\"}").Status);
        }

        [Test]
        public void AnotherRefactoringPath_IsNotFound()
        {
            Assert.AreEqual(404, Post("/projects/p/refactor/nonsense", Body).Status);
            Assert.AreEqual(404, Post("/projects/p/refactor/rename/other", Body).Status);
        }

        [Test]
        public void WithoutTheFeature_TheRouteIsNotFound()
        {
            var bare = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api);

            var response = bare.Handle(new ApiRequest { Method = "POST", Path = ApiRouter.Base + "/projects/p/refactor/rename", Body = Body, Client = "agent" });

            Assert.AreEqual(404, response.Status);
        }

        private JObject Rpc(string method, JObject parameters = null)
        {
            var body = new JObject { ["jsonrpc"] = "2.0", ["id"] = 1, ["method"] = method };
            if (parameters != null) body["params"] = parameters;
            return JObject.FromObject(_router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Body = body.ToString(), Client = "agent" }).Body);
        }

        [Test]
        public void TheTools_AreListed_WithTheirRequiredFields()
        {
            var tools = Rpc("tools/list")["result"]["tools"].ToList();

            CollectionAssert.IsSubsetOf(new[] { "preview_rename", "rename_symbol" }, tools.Select(t => t.Value<string>("name")));
            var rename = tools.Single(t => t.Value<string>("name") == "rename_symbol");
            CollectionAssert.AreEquivalent(new[] { "project", "module", "newName" }, rename["inputSchema"]["required"].Select(r => r.Value<string>()));
        }

        [Test]
        public void TheTools_Work()
        {
            var arguments = new JObject { ["project"] = "p", ["module"] = "Module1", ["name"] = "Alpha", ["newName"] = "Beta" };

            var previewed = Rpc("tools/call", new JObject { ["name"] = "preview_rename", ["arguments"] = arguments })["result"];
            var renamed = Rpc("tools/call", new JObject { ["name"] = "rename_symbol", ["arguments"] = arguments })["result"];

            Assert.IsFalse(previewed.Value<bool>("isError"));
            Assert.IsFalse(renamed.Value<bool>("isError"));
            StringAssert.Contains("Beta", renamed["content"][0].Value<string>("text"));
            CollectionAssert.AreEqual(new[] { "preview", "preview", "apply" }, _source.Calls);
        }

        [Test]
        public void ATool_ReportsRefusalsInsideTheResult()
        {
            _source.Preview_.Valid = false;
            _source.Preview_.Problems.Add("Reserved word.");
            var arguments = new JObject { ["project"] = "p", ["module"] = "Module1", ["name"] = "Alpha", ["newName"] = "Sub" };

            var result = Rpc("tools/call", new JObject { ["name"] = "rename_symbol", ["arguments"] = arguments })["result"];

            Assert.IsTrue(result.Value<bool>("isError"));
            StringAssert.Contains("refactoring_refused", result["content"][0].Value<string>("text"));
        }
    }

    [TestFixture]
    [Category("Api")]
    public class RefactorTargetsTests
    {
        private static TargetCandidate Candidate(string name, string kind, int line, int column = 1, string parent = "Module1")
            => new TargetCandidate { Name = name, Kind = kind, Line = line, Column = column, EndColumn = column + name.Length, Parent = parent, Tag = $"{kind}:{name}@{line}" };

        private static readonly List<TargetCandidate> Declared = new List<TargetCandidate>
        {
            Candidate("Alpha", "Function", 2),
            Candidate("Count", "Variable", 1),
            Candidate("Index", "Variable", 5, 20, "Alpha"),
            Candidate("Index", "Variable", 9, 20, "Beta"),
            Candidate("Value", "PropertyGet", 12),
            Candidate("Value", "PropertyLet", 16),
            Candidate("Value", "PropertySet", 20)
        };

        [Test]
        public void ByName_IsCaseInsensitive()
        {
            Assert.AreEqual("Function:Alpha@2", RefactorTargets.Choose(Declared, "Module1", "alpha", null, null, null));
        }

        [Test]
        public void ByNameWithKind_NarrowsTheMatches()
        {
            Assert.AreEqual("Variable:Count@1", RefactorTargets.Choose(Declared, "Module1", "Count", "variable", null, null));
        }

        [Test]
        public void NotUnique_ListsTheMatches_AndALineSettlesIt()
        {
            var failure = Assert.Throws<ApiException>(() => RefactorTargets.Choose(Declared, "Module1", "Index", null, null, null));

            Assert.AreEqual(409, failure.Status);
            StringAssert.Contains("line 5", failure.Message);
            StringAssert.Contains("line 9", failure.Message);
            Assert.AreEqual("Variable:Index@9", RefactorTargets.Choose(Declared, "Module1", "Index", null, 9, null));
        }

        [Test]
        public void TheAccessorsOfOneProperty_AreOneTarget()
        {
            Assert.AreEqual("PropertyGet:Value@12", RefactorTargets.Choose(Declared, "Module1", "Value", null, null, null));
        }

        [Test]
        public void ByPositionAlone_NeedsALineThatHasADeclaration()
        {
            Assert.AreEqual("Variable:Count@1", RefactorTargets.Choose(Declared, "Module1", null, null, 1, 3));
            Assert.AreEqual(404, Assert.Throws<ApiException>(() => RefactorTargets.Choose(Declared, "Module1", null, null, 99, null)).Status);
        }

        [Test]
        public void AColumnOutsideTheName_MatchesNothing()
        {
            Assert.AreEqual(404, Assert.Throws<ApiException>(() => RefactorTargets.Choose(Declared, "Module1", "Count", null, 1, 40)).Status);
        }

        [Test]
        public void UnknownName_IsNotFound()
        {
            Assert.AreEqual(404, Assert.Throws<ApiException>(() => RefactorTargets.Choose(Declared, "Module1", "Nope", null, null, null)).Status);
        }

        [TestCase(null, null, null)]
        [TestCase("Alpha", 0, null)]
        [TestCase("Alpha", 1, 0)]
        [TestCase("Alpha", null, 3)]
        public void BadArguments_AreABadRequest(string name, int? line, int? column)
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => RefactorTargets.Choose(Declared, "Module1", name, null, line, column)).Status);
        }
    }
}
