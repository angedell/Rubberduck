using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    internal class FakeInsight : IInsightSource
    {
        public List<ParserError> Errors = new List<ParserError>();
        public List<Symbol> Declared = new List<Symbol>
        {
            new Symbol { Module = "Module1", Name = "Beta", Kind = "Procedure", StartLine = 9, EndLine = 12 },
            new Symbol { Module = "Module1", Name = "Alpha", Kind = "Function", StartLine = 2, EndLine = 6 },
            new Symbol { Module = "Module1", Name = "Count", Kind = "Variable", StartLine = 1, EndLine = 1 }
        };

        public string AskedKind;
        public bool AskedLocals;
        public int AskedLimit;

        public ParserErrorsReport ParserErrors(string project) => new ParserErrorsReport { Project = project, ParserState = Errors.Count > 0 ? "Error" : "Ready", Errors = Errors };

        public SymbolsReport Symbols(string project, string module, bool includeLocals, string kind)
        {
            AskedKind = kind;
            AskedLocals = includeLocals;
            return new SymbolsReport { Project = project, Module = module, Symbols = Declared.ToList() };
        }

        public ReferencesReport References(string project, string module, string name, int limit)
        {
            AskedLimit = limit;
            return new ReferencesReport { Project = project };
        }

        public DefinitionReport Definition(string project, string module, int line, int column) => new DefinitionReport { Found = false };
        public IReadOnlyList<LibraryReference> Libraries(string project) => new[] { new LibraryReference { Name = "Excel" } };
        public IReadOnlyList<StructureFolder> Structure(string project) => new[] { new StructureFolder { Folder = "App.Core", Modules = { "Module1" } } };
        public IReadOnlyList<MetricValue> Metrics(string project, string module) => new List<MetricValue>();
    }

    internal class FakeTests : ITestSource
    {
        public List<string> Runs = new List<string>();
        public int Failed;

        public IReadOnlyList<TestInfo> List(string project, string module)
            => new[] { new TestInfo { Module = "Tests1", Name = "Works", Outcome = "Unknown" } };

        public TestRunReport Run(string project, string module, string test, TimeSpan wait)
        {
            Runs.Add($"{module}/{test}/{wait.TotalSeconds}");
            return TestsApi.Summarize(project, true, new[]
            {
                new TestInfo { Module = "Tests1", Name = "Works", Outcome = Failed > 0 ? "Failed" : "Succeeded" }
            });
        }
    }

    internal class FakeDialogs : IDialogWatcher
    {
        public BlockingDialog Showing;
        public bool Clicked;
        public bool EditorError = true;

        public BlockingDialog Find() => Showing;

        public bool Acknowledge(BlockingDialog dialog)
        {
            Clicked = true;
            Showing = null;
            return true;
        }
    }

    [TestFixture]
    [Category("Api")]
    public class InsightApiTests
    {
        private FakeInsight _source;
        private InsightApi _api;

        [SetUp]
        public void SetUp()
        {
            _source = new FakeInsight();
            _api = new InsightApi(_source);
        }

        [Test]
        public void Symbols_NeedAModule()
        {
            var failure = Assert.Throws<ApiException>(() => _api.Symbols("p", " ", false, null));

            Assert.AreEqual(400, failure.Status);
        }

        [Test]
        public void Symbols_AreCounted_AndEmptyKindMeansAll()
        {
            var report = _api.Symbols("p", "Module1", true, "");

            Assert.AreEqual(3, report.Total);
            Assert.IsNull(_source.AskedKind);
            Assert.IsTrue(_source.AskedLocals);
        }

        [TestCase(0)]
        [TestCase(InsightApi.MaxReferenceLimit + 1)]
        public void References_RejectABadLimit(int limit)
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.References("p", null, "Alpha", limit)).Status);
        }

        [Test]
        public void References_DefaultLimit_IsUsedWhenNoneGiven()
        {
            _api.References("p", null, "Alpha", null);

            Assert.AreEqual(InsightApi.DefaultReferenceLimit, _source.AskedLimit);
        }

        [Test]
        public void References_NeedAName()
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.References("p", null, "", null)).Status);
        }

        [TestCase(0, 1)]
        [TestCase(1, 0)]
        public void Definition_PositionsAreOneBased(int line, int column)
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Definition("p", "Module1", line, column)).Status);
        }
    }

    [TestFixture]
    [Category("Api")]
    public class TestsApiTests
    {
        private FakeTests _source;
        private bool _enabled;
        private int _asked;
        private TestsApi _api;

        [SetUp]
        public void SetUp()
        {
            _source = new FakeTests();
            _enabled = true;
            _asked = 0;
            _api = new TestsApi(_source, new ApiExecuteApproval(_ => { _asked++; return true; }), () => _enabled);
        }

        [Test]
        public void Listing_NeedsNeitherTheOptInNorAnApproval()
        {
            _enabled = false;

            Assert.AreEqual(1, _api.List("p", null).Count);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Running_IsRefused_WhenRunningCodeIsOff()
        {
            _enabled = false;

            var failure = Assert.Throws<ApiException>(() => _api.Run("agent", "p", null, null, null));

            Assert.AreEqual(403, failure.Status);
            Assert.AreEqual("execute_disabled", failure.Code);
            Assert.IsEmpty(_source.Runs);
        }

        [Test]
        public void Running_AsksOnce_AndPassesTheFilters()
        {
            _api.Run("agent", "p", "Tests1", "Works", 5);
            _api.Run("agent", "p", "Tests1", null, null);

            Assert.AreEqual(1, _asked);
            CollectionAssert.AreEqual(new[] { "Tests1/Works/5", "Tests1//60" }, _source.Runs);
        }

        [Test]
        public void Running_WithoutApproval_RunsNothing()
        {
            var api = new TestsApi(_source, new ApiExecuteApproval(_ => false), () => true);

            Assert.AreEqual(403, Assert.Throws<ApiException>(() => api.Run("agent", "p", null, null, null)).Status);
            Assert.IsEmpty(_source.Runs);
        }

        [Test]
        public void ATest_NeedsItsModule()
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Run("agent", "p", null, "Works", null)).Status);
        }

        [TestCase(0)]
        [TestCase(301)]
        public void Waiting_IsBounded(int seconds)
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Run("agent", "p", null, null, seconds)).Status);
        }

        [Test]
        public void Summary_CountsOutcomes()
        {
            var report = TestsApi.Summarize("p", true, new[]
            {
                new TestInfo { Outcome = "Succeeded" }, new TestInfo { Outcome = "Succeeded" }, new TestInfo { Outcome = "Failed" }, new TestInfo { Outcome = "Ignored" }
            });

            Assert.AreEqual(2, report.Succeeded);
            Assert.AreEqual(1, report.Failed);
            Assert.AreEqual(1, report.Ignored);
        }
    }

    [TestFixture]
    [Category("Api")]
    public class CheckApiTests
    {
        private FakeInsight _insight;
        private FakeDiagnostics _diagnostics;
        private FakeTests _tests;
        private bool _runCode;

        [SetUp]
        public void SetUp()
        {
            _insight = new FakeInsight();
            _diagnostics = new FakeDiagnostics { Found = new List<Diagnostic>() };
            _tests = new FakeTests();
            _runCode = true;
        }

        private CheckApi Check()
            => new CheckApi(new InsightApi(_insight), new DiagnosticsApi(_diagnostics),
                new TestsApi(_tests, new ApiExecuteApproval(_ => true), () => _runCode));

        [Test]
        public void ParseErrors_SkipTheInspections_AndTheTests()
        {
            _insight.Errors.Add(new ParserError { Module = "Module1", Line = 3, Message = "bad" });

            var report = Check().Check("agent", "p", null, true);

            Assert.IsFalse(report.Clean);
            Assert.IsNull(report.Diagnostics);
            Assert.IsNull(report.Tests);
            Assert.IsNull(_diagnostics.AskedProject);
        }

        [Test]
        public void Hints_DoNotSpoilACleanResult()
        {
            _diagnostics.Found.Add(new Diagnostic { Module = "Module1", Severity = "hint" });
            _diagnostics.Found.Add(new Diagnostic { Module = "Module1", Severity = "suggestion" });

            Assert.IsTrue(Check().Check("agent", "p", null, false).Clean);
        }

        [Test]
        public void Hints_AreCountedAsSuppressed_SoCleanIsNotMistakenForNothingFound()
        {
            _diagnostics.Found.Add(new Diagnostic { Module = "Module1", Severity = "hint" });
            _diagnostics.Found.Add(new Diagnostic { Module = "Module1", Severity = "suggestion" });
            _diagnostics.Found.Add(new Diagnostic { Module = "Module1", Severity = "suggestion" });

            var report = Check().Check("agent", "p", null, false);

            Assert.AreEqual(1, report.Suppressed["hint"]);
            Assert.AreEqual(2, report.Suppressed["suggestion"]);
            Assert.IsFalse(report.Suppressed.ContainsKey("warning"));
        }

        [Test]
        public void AWarning_IsNotClean()
        {
            _diagnostics.Found.Add(new Diagnostic { Module = "Module1", Severity = "warning" });

            Assert.IsFalse(Check().Check("agent", "p", null, false).Clean);
        }

        [Test]
        public void SkippedTests_AreReported_SoCleanIsNotMistakenForTested()
        {
            var report = Check().Check("agent", "p", null, false);

            Assert.IsTrue(report.Clean);
            Assert.IsNull(report.Tests);
            StringAssert.Contains("runTests", report.TestsNotRun);
        }

        [Test]
        public void TestsThatRan_NeedNoSkippedNote()
        {
            Assert.IsNull(Check().Check("agent", "p", null, true).TestsNotRun);
        }

        [Test]
        public void Tests_RunOnlyWhenAsked()
        {
            Check().Check("agent", "p", null, false);
            Assert.IsEmpty(_tests.Runs);

            Check().Check("agent", "p", "Module1", true);
            Assert.AreEqual(1, _tests.Runs.Count);
        }

        [Test]
        public void AFailingTest_IsNotClean()
        {
            _tests.Failed = 1;

            Assert.IsFalse(Check().Check("agent", "p", null, true).Clean);
        }

        [Test]
        public void Tests_AreRefused_WhenRunningCodeIsOff()
        {
            _runCode = false;

            Assert.AreEqual(403, Assert.Throws<ApiException>(() => Check().Check("agent", "p", null, true)).Status);
        }
    }

    [TestFixture]
    [Category("Api")]
    public class BatchEditTests
    {
        private FakeStore _store;
        private int _asked;
        private ModuleApi _api;

        [SetUp]
        public void SetUp()
        {
            _store = new FakeStore();
            _store.Code["Module2"] = "Sub B()\r\nEnd Sub";
            _asked = 0;
            _api = new ModuleApi(_store, new ApiWriteApproval(_ => { _asked++; return true; }));
        }

        private string HashOf(string module) => ModuleApi.HashOf(_store.Code[module]);

        [Test]
        public void Edits_AreAppliedInOrder_WithOneApproval()
        {
            var result = _api.Batch("agent", "p", new[]
            {
                new ModuleEdit { Op = "replaceCode", Module = "Module1", ExpectedHash = HashOf("Module1"), Code = "Sub X()\r\nEnd Sub" },
                new ModuleEdit { Op = "editLines", Module = "Module2", ExpectedHash = HashOf("Module2"), StartLine = 1, EndLine = 1, Replacement = "Sub C()" },
                new ModuleEdit { Op = "create", Module = "Module3", Kind = "class", Code = "Option Explicit" }
            });

            CollectionAssert.AreEqual(new[] { "ReplaceCode", "EditLines", "Create" }, _store.Calls);
            CollectionAssert.AreEqual(new[] { "Module1", "Module2", "Module3" }, result.Modules.Select(m => m.Name));
            Assert.AreEqual(1, _asked);
        }

        [Test]
        public void AStaleHash_AnywhereInTheBatch_ChangesNothing()
        {
            var failure = Assert.Throws<ApiException>(() => _api.Batch("agent", "p", new[]
            {
                new ModuleEdit { Op = "replaceCode", Module = "Module1", ExpectedHash = HashOf("Module1"), Code = "Sub X()\r\nEnd Sub" },
                new ModuleEdit { Op = "replaceCode", Module = "Module2", ExpectedHash = "stale", Code = "Sub Y()\r\nEnd Sub" }
            }));

            Assert.AreEqual(412, failure.Status);
            StringAssert.Contains("Module2", failure.Message);
            Assert.IsEmpty(_store.Calls);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void TheSameModuleTwice_IsRefused()
        {
            var failure = Assert.Throws<ApiException>(() => _api.Batch("agent", "p", new[]
            {
                new ModuleEdit { Op = "editLines", Module = "Module1", ExpectedHash = HashOf("Module1"), StartLine = 1, EndLine = 1, Replacement = "a" },
                new ModuleEdit { Op = "editLines", Module = "module1", ExpectedHash = HashOf("Module1"), StartLine = 2, EndLine = 2, Replacement = "b" }
            }));

            Assert.AreEqual(400, failure.Status);
            Assert.IsEmpty(_store.Calls);
        }

        [Test]
        public void CreatingAnExistingModule_IsAConflict_BeforeAnythingChanges()
        {
            var failure = Assert.Throws<ApiException>(() => _api.Batch("agent", "p", new[]
            {
                new ModuleEdit { Op = "replaceCode", Module = "Module1", ExpectedHash = HashOf("Module1"), Code = "Sub X()\r\nEnd Sub" },
                new ModuleEdit { Op = "create", Module = "Module2", Kind = "standard", Code = "" }
            }));

            Assert.AreEqual(409, failure.Status);
            Assert.IsEmpty(_store.Calls);
        }

        [Test]
        public void AFailureWhileApplying_SaysWhatWasAlreadyChanged()
        {
            //Module1 is changed behind the batch's back, after the checks: the second edit's store call fails.
            var racing = new RacingStore(_store);
            var api = new ModuleApi(racing, new ApiWriteApproval(_ => true));

            var failure = Assert.Throws<ApiException>(() => api.Batch("agent", "p", new[]
            {
                new ModuleEdit { Op = "replaceCode", Module = "Module1", ExpectedHash = HashOf("Module1"), Code = "Sub X()\r\nEnd Sub" },
                new ModuleEdit { Op = "replaceCode", Module = "Module2", ExpectedHash = HashOf("Module2"), Code = "Sub Y()\r\nEnd Sub" }
            }));

            Assert.AreEqual(412, failure.Status);
            StringAssert.Contains("Already changed: Module1", failure.Message);
            StringAssert.DoesNotContain("Sub Y", _store.Code["Module2"]);
        }

        [TestCase(null)]
        [TestCase("explode")]
        public void AnUnknownOperation_IsRefused(string op)
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Batch("agent", "p", new[] { new ModuleEdit { Op = op, Module = "Module1" } })).Status);
        }

        [Test]
        public void AnEmptyBatch_IsRefused()
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Batch("agent", "p", new ModuleEdit[0])).Status);
        }

        [Test]
        public void HeaderLinesInCode_AreRefused_EvenInTheLastEdit()
        {
            var failure = Assert.Throws<ApiException>(() => _api.Batch("agent", "p", new[]
            {
                new ModuleEdit { Op = "replaceCode", Module = "Module1", ExpectedHash = HashOf("Module1"), Code = "Sub X()\r\nEnd Sub" },
                new ModuleEdit { Op = "replaceCode", Module = "Module2", ExpectedHash = HashOf("Module2"), Code = "Attribute VB_Name = \"x\"" }
            }));

            Assert.AreEqual(422, failure.Status);
            Assert.IsEmpty(_store.Calls);
        }

        /// <summary>Lets the first write through, then changes Module2 so the next one finds a stale hash.</summary>
        private class RacingStore : IModuleStore
        {
            private readonly FakeStore _inner;

            public RacingStore(FakeStore inner) { _inner = inner; }

            public IReadOnlyList<ProjectInfo> Projects() => _inner.Projects();
            public IReadOnlyList<ModuleInfo> Modules(string project) => _inner.Modules(project);
            public ModuleContent Read(string project, string module, bool includeText) => _inner.Read(project, module, includeText);

            public ModuleInfo ReplaceCode(string project, string module, string code, string expectedHash)
            {
                var result = _inner.ReplaceCode(project, module, code, expectedHash);
                _inner.Code["Module2"] = "Sub B()\r\nEnd Sub '";
                return result;
            }

            public ModuleInfo EditLines(string project, string module, int startLine, int endLine, string replacement, string expectedHash)
                => _inner.EditLines(project, module, startLine, endLine, replacement, expectedHash);

            public ModuleInfo ReplaceText(string project, string module, string text, string expectedHash) => _inner.ReplaceText(project, module, text, expectedHash);
            public ModuleInfo Create(string project, string module, string kind, string code) => _inner.Create(project, module, kind, code);
            public void Delete(string project, string module, string expectedHash) => _inner.Delete(project, module, expectedHash);
        }
    }

    [TestFixture]
    [Category("Api")]
    public class ImmediateDialogTests
    {
        private FakeWindow _window;
        private FakeDialogs _dialogs;
        private ImmediateApi _api;

        [SetUp]
        public void SetUp()
        {
            _window = new FakeWindow();
            _dialogs = new FakeDialogs();
            var clock = TimeSpan.Zero;
            _api = new ImmediateApi(_window, new ApiExecuteApproval(_ => true), () => true, step => clock += step, () => clock, _dialogs);
        }

        [Test]
        public void AnEditorErrorDialog_IsAcknowledged_AndReported()
        {
            _window.Output = line =>
            {
                _dialogs.Showing = new BlockingDialog { Title = "Microsoft Visual Basic", Message = "Run-time error '11':\nDivision by zero", IsEditorError = true };
                return string.Empty;
            };

            var run = _api.Run("agent", "Call Boom", null);

            Assert.IsTrue(_dialogs.Clicked);
            CollectionAssert.AreEqual(new[] { "Run-time error '11':\nDivision by zero" }, run.Errors);
            Assert.IsNull(run.BlockedByDialog);
            Assert.IsTrue(run.Completed);
        }

        [Test]
        public void AnyOtherDialog_IsLeftAlone_AndReportedAsBlocking()
        {
            _window.Output = line =>
            {
                _dialogs.Showing = new BlockingDialog { Title = "Excel", Message = "Save changes?", Buttons = { "Yes", "No" }, IsEditorError = false };
                return string.Empty;
            };

            var run = _api.Run("agent", "Call Ask", null);

            Assert.IsFalse(_dialogs.Clicked);
            Assert.IsFalse(run.Completed);
            Assert.AreEqual("Save changes?", run.BlockedByDialog.Message);
            CollectionAssert.AreEqual(new[] { "Yes", "No" }, run.BlockedByDialog.Buttons);
        }

        [Test]
        public void WithoutADialog_NothingIsReported()
        {
            var run = _api.Run("agent", "?1+1", null);

            Assert.IsEmpty(run.Errors);
            Assert.IsNull(run.BlockedByDialog);
        }
    }

    [TestFixture]
    [Category("Api")]
    public class InsightRouteTests
    {
        private FakeInsight _insight;
        private FakeTests _tests;
        private FakeDiagnostics _diagnostics;
        private FakeStore _store;
        private bool _runCode;
        private ApiRouter _router;

        [SetUp]
        public void SetUp()
        {
            _insight = new FakeInsight();
            _tests = new FakeTests();
            _diagnostics = new FakeDiagnostics { Found = new List<Diagnostic>() };
            _store = new FakeStore();
            _runCode = true;
            var insight = new InsightApi(_insight);
            var diagnostics = new DiagnosticsApi(_diagnostics);
            var testsApi = new TestsApi(_tests, new ApiExecuteApproval(_ => true), () => _runCode);
            _router = new ApiRouter(new ModuleApi(_store, new ApiWriteApproval(_ => true)), diagnostics, Fakes.Immediate(() => false).Api,
                null, insight, testsApi, new CheckApi(insight, diagnostics, testsApi));
        }

        private ApiResponse Get(string path, params string[] query)
        {
            var request = new ApiRequest { Method = "GET", Path = ApiRouter.Base + path, Client = "agent" };
            for (var i = 0; i < query.Length; i += 2) request.Query[query[i]] = query[i + 1];
            return _router.Handle(request);
        }

        private ApiResponse Post(string path, string body)
            => _router.Handle(new ApiRequest { Method = "POST", Path = ApiRouter.Base + path, Body = body, Client = "agent" });

        private static JObject Json(ApiResponse response) => JObject.FromObject(response.Body);

        [Test]
        public void ParserErrors_Route()
        {
            _insight.Errors.Add(new ParserError { Module = "Module1", Line = 4, Message = "bad" });

            var response = Get("/projects/p/parser-errors");

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual("Error", Json(response).Value<string>("ParserState"));
        }

        [Test]
        public void Symbols_Route_PassesTheQuery()
        {
            var response = Get("/projects/p/symbols", "module", "Module1", "includeLocals", "true", "kind", "Function");

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual("Function", _insight.AskedKind);
            Assert.IsTrue(_insight.AskedLocals);
        }

        [Test]
        public void Symbols_WithoutAModule_IsABadRequest()
        {
            Assert.AreEqual(400, Get("/projects/p/symbols").Status);
        }

        [Test]
        public void References_AndDefinition_Routes()
        {
            Assert.AreEqual(200, Get("/projects/p/symbols/references", "name", "Alpha", "limit", "5").Status);
            Assert.AreEqual(5, _insight.AskedLimit);
            Assert.AreEqual(200, Get("/projects/p/symbols/definition", "module", "Module1", "line", "3", "column", "4").Status);
            Assert.AreEqual(400, Get("/projects/p/symbols/references", "name", "Alpha", "limit", "many").Status);
        }

        [Test]
        public void Libraries_Structure_Metrics_Routes()
        {
            Assert.AreEqual("Excel", Json(Get("/projects/p/libraries"))["libraries"][0].Value<string>("Name"));
            Assert.AreEqual("App.Core", Json(Get("/projects/p/structure"))["folders"][0].Value<string>("Folder"));
            Assert.AreEqual(200, Get("/projects/p/metrics").Status);
        }

        [Test]
        public void Tests_List_NeedsNoOptIn()
        {
            _runCode = false;

            Assert.AreEqual(200, Get("/projects/p/tests").Status);
        }

        [Test]
        public void Tests_Run_IsForbidden_WhenOff()
        {
            _runCode = false;

            var response = Post("/projects/p/tests/run", "{}");

            Assert.AreEqual(403, response.Status);
            Assert.IsEmpty(_tests.Runs);
        }

        [Test]
        public void Tests_Run_PassesTheFilters()
        {
            var response = Post("/projects/p/tests/run", "{\"module\":\"Tests1\",\"test\":\"Works\",\"waitSeconds\":9}");

            Assert.AreEqual(200, response.Status);
            CollectionAssert.AreEqual(new[] { "Tests1/Works/9" }, _tests.Runs);
        }

        [Test]
        public void Check_Route_AcceptsAnEmptyBody()
        {
            var response = Post("/projects/p/check", "");

            Assert.AreEqual(200, response.Status);
            Assert.IsTrue(Json(response).Value<bool>("Clean"));
        }

        [Test]
        public void Edits_Route_AppliesTheBatch()
        {
            var hash = ModuleApi.HashOf(_store.Code["Module1"]);

            var response = Post("/projects/p/edits", "{\"edits\":[{\"op\":\"editLines\",\"module\":\"Module1\",\"expectedHash\":\"" + hash + "\",\"startLine\":1,\"endLine\":1,\"replacement\":\"Option Explicit '\"}]}");

            Assert.AreEqual(200, response.Status);
            CollectionAssert.AreEqual(new[] { "EditLines" }, _store.Calls);
        }

        [Test]
        public void Edits_Route_RejectsABodyWithoutEdits()
        {
            Assert.AreEqual(400, Post("/projects/p/edits", "{}").Status);
        }

        [Test]
        public void Routes_AreAbsent_WhenTheirApiIsNotThere()
        {
            var bare = new ApiRouter(new ModuleApi(_store, new ApiWriteApproval(_ => true)), new DiagnosticsApi(_diagnostics), Fakes.Immediate(() => false).Api);

            var response = bare.Handle(new ApiRequest { Method = "GET", Path = ApiRouter.Base + "/projects/p/symbols", Client = "agent" });

            Assert.AreEqual(404, response.Status);
        }

        [Test]
        public void ARunBody_ThatIsNotAnObject_IsABadRequest()
        {
            Assert.AreEqual(400, Post("/projects/p/tests/run", "[]").Status);
        }
    }

    [TestFixture]
    [Category("Api")]
    public class InsightMcpTests
    {
        private bool _runCode;
        private ApiRouter _router;
        private FakeStore _store;

        [SetUp]
        public void SetUp()
        {
            _runCode = false;
            _store = new FakeStore();
            var diagnostics = new DiagnosticsApi(new FakeDiagnostics { Found = new List<Diagnostic>() });
            var insight = new InsightApi(new FakeInsight());
            var tests = new TestsApi(new FakeTests(), new ApiExecuteApproval(_ => true), () => _runCode);
            _router = new ApiRouter(new ModuleApi(_store, new ApiWriteApproval(_ => true)), diagnostics, Fakes.Immediate(() => false).Api,
                null, insight, tests, new CheckApi(insight, diagnostics, tests));
        }

        private JObject Rpc(string method, JObject parameters = null)
        {
            var body = new JObject { ["jsonrpc"] = "2.0", ["id"] = 1, ["method"] = method };
            if (parameters != null) body["params"] = parameters;
            var response = _router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Body = body.ToString(), Client = "agent" });
            return JObject.FromObject(response.Body);
        }

        private List<string> ToolNames() => Rpc("tools/list")["result"]["tools"].Select(t => t.Value<string>("name")).ToList();

        [Test]
        public void TheNewTools_AreListed()
        {
            CollectionAssert.IsSubsetOf(
                new[] { "get_parser_errors", "list_symbols", "find_references", "find_definition", "list_libraries", "get_structure", "get_metrics", "list_tests", "check_work", "batch_edit" },
                ToolNames());
        }

        [Test]
        public void RunTests_IsOffered_OnlyWhenRunningCodeIsAllowed()
        {
            CollectionAssert.DoesNotContain(ToolNames(), "run_tests");

            _runCode = true;

            CollectionAssert.Contains(ToolNames(), "run_tests");
        }

        [Test]
        public void Initialize_AnnouncesResources()
        {
            Assert.IsNotNull(Rpc("initialize")["result"]["capabilities"]["resources"]);
        }

        [Test]
        public void Resources_ListModulesAndReports()
        {
            var uris = Rpc("resources/list")["result"]["resources"].Select(r => r.Value<string>("uri")).ToList();

            CollectionAssert.Contains(uris, "rubberduck://VBAProject/modules/Module1");
            CollectionAssert.Contains(uris, "rubberduck://VBAProject/diagnostics");
            CollectionAssert.Contains(uris, "rubberduck://VBAProject/parser-errors");
        }

        [Test]
        public void Resources_ReadAModule()
        {
            var contents = Rpc("resources/read", new JObject { ["uri"] = "rubberduck://VBAProject/modules/Module1" })["result"]["contents"][0];

            Assert.AreEqual(_store.Code["Module1"], contents.Value<string>("text"));
            Assert.AreEqual("text/plain", contents.Value<string>("mimeType"));
        }

        [Test]
        public void Resources_ReadReports()
        {
            var contents = Rpc("resources/read", new JObject { ["uri"] = "rubberduck://VBAProject/parser-errors" })["result"]["contents"][0];

            Assert.AreEqual("application/json", contents.Value<string>("mimeType"));
            StringAssert.Contains("Ready", contents.Value<string>("text"));
        }

        [Test]
        public void Resources_UnknownUri_IsAnError()
        {
            Assert.IsNotNull(Rpc("resources/read", new JObject { ["uri"] = "file:///etc/passwd" })["error"]);
        }

        [Test]
        public void BatchEdit_Tool_Works()
        {
            var hash = ModuleApi.HashOf(_store.Code["Module1"]);
            var edits = new JArray(new JObject { ["op"] = "replaceCode", ["module"] = "Module1", ["expectedHash"] = hash, ["code"] = "Sub Z()\r\nEnd Sub" });

            var result = Rpc("tools/call", new JObject { ["name"] = "batch_edit", ["arguments"] = new JObject { ["project"] = "p", ["edits"] = edits } })["result"];

            Assert.IsFalse(result.Value<bool>("isError"));
            CollectionAssert.AreEqual(new[] { "ReplaceCode" }, _store.Calls);
        }
    }
}
