using System;
using System.Collections.Generic;
using System.Linq;
using Moq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;
using Rubberduck.Ship;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Api
{
    [TestFixture]
    [Category("Api")]
    public class ApiShipTests
    {
        private class FakeShip : IShipService
        {
            public ShipSettings Settings = new ShipSettings { OutputFolder = "..\\out" };
            public bool Unattended;
            public ShipRefusedException Refuse;
            public List<string> Calls = new List<string>();
            public ShipOnExisting? LastOnExisting;
            public ShipSettings Written;
            public Func<bool> InsideWith = () => false;
            public bool PreviewedInsideWith;
            public int RootCount = 3;

            public bool CanShip(IVBProject project) => true;
            public ShipSettings ReadSettings(IVBProject project) => Settings;
            public void WriteSettings(IVBProject project, ShipSettings settings) { Calls.Add("WriteSettings"); Written = settings; Settings = settings; }
            public string ResolveOutputFolder(IVBProject project, ShipSettings settings) => settings.OutputFolder == null ? null : @"C:\work\out";

            public ShipReport Preview(IVBProject project, ShipSettings settings)
            {
                Calls.Add("Preview");
                PreviewedInsideWith |= InsideWith();
                if (Refuse != null) throw Refuse;
                return Report(true);
            }

            public ShipReport Ship(IVBProject project, ShipSettings settings, ShipOnExisting onExisting, IProgress<string> progress)
            {
                Calls.Add("Ship");
                LastOnExisting = onExisting;
                if (Refuse != null) throw Refuse;
                return Report(false);
            }

            public bool IsUnattended(IVBProject project) => Unattended;
            public void SetUnattended(IVBProject project, bool value) { Calls.Add("SetUnattended"); Unattended = value; }

            private ShipReport Report(bool dryRun)
            {
                var report = new ShipReport { DryRun = dryRun, OutputPath = @"C:\work\out\Book.xlsm" };
                for (var i = 0; i < RootCount; i++) report.Roots.Add("Module1.Root" + i + ": public procedure of a standard module");
                var notes = new ShipStepReport { Step = ShipSteps.Notes, Removed = 2 };
                notes.Details.Add("Note_1");
                report.Steps.Add(notes);
                report.Steps.Add(new ShipStepReport { Step = ShipSteps.Comments, Removed = 40 });
                var shaking = new ShipStepReport { Step = ShipSteps.TreeShaking, Removed = 2 };
                shaking.Details.Add("Module1.OldHelper");
                shaking.Details.Add("Module2.mLegacy");
                report.Steps.Add(shaking);
                report.DynamicCalls.Add("Module2.RunByName line 14: Application.Run with a computed name");
                report.Unparsed.Add("Broken: syntax error");
                return report;
            }
        }

        private class FakeProjects : IShipProjects
        {
            public int Lookups;
            public bool Inside;

            public IVBProject Resolve(string project)
            {
                Lookups++;
                if (project == "missing") throw ApiException.NotFound("No open project 'missing'.");
                return new Mock<IVBProject>().Object;
            }

            public T With<T>(string project, Func<IVBProject, T> work, TimeSpan? timeout = null)
            {
                Lookups++;
                if (project == "missing") throw ApiException.NotFound("No open project 'missing'.");
                Inside = true;
                try
                {
                    return work(new Mock<IVBProject>().Object);
                }
                finally
                {
                    Inside = false;
                }
            }
        }

        private FakeShip _service;
        private FakeProjects _projects;
        private int _asked;
        private bool _allow;
        private ApiRouter _router;

        [SetUp]
        public void SetUp()
        {
            _service = new FakeShip();
            _projects = new FakeProjects();
            _service.InsideWith = () => _projects.Inside;
            _asked = 0;
            _allow = true;
            _router = Router(_service);
        }

        private ApiRouter Router(IShipService service)
        {
            var approval = new ApiWriteApproval(_ => { _asked++; return _allow; });
            var ship = new ShipApi(service, _projects, approval);
            return new ApiRouter(new ModuleApi(new FakeStore(), approval), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api, ship: ship);
        }

        private ApiResponse Send(string method, string path, object body = null)
            => _router.Handle(new ApiRequest { Method = method, Path = "/api/v1" + path, Body = body == null ? null : JObject.FromObject(body).ToString(), Client = "agent" });

        private static JObject Json(ApiResponse response) => JObject.FromObject(response.Body);

        [Test]
        public void Settings_Get_ShowsThemResolvedFolderAndUnattendedReadOnly()
        {
            _service.Unattended = true;

            var response = Send("GET", "/projects/p/ship/settings");

            Assert.AreEqual(200, response.Status);
            var json = Json(response);
            Assert.AreEqual("xlsm", (string)json["format"]);
            Assert.AreEqual("..\\out", (string)json["outputFolder"]);
            Assert.AreEqual(@"C:\work\out", (string)json["resolvedOutputFolder"]);
            Assert.AreEqual("ask", (string)json["onExisting"]);
            Assert.IsTrue((bool)json["steps"]["comments"]);
            Assert.IsTrue((bool)json["unattended"]);
            Assert.IsTrue((bool)json["unattendedReadOnly"]);
            Assert.AreEqual(0, _asked, "a read needs no approval");
        }

        [Test]
        public void Settings_Put_ChangesWhatIsSent_AndNeedsApproval()
        {
            var response = Send("PUT", "/projects/p/ship/settings", new { format = "xlam", onExisting = "keepBoth", steps = new { comments = false } });

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual(1, _asked);
            Assert.AreEqual(ShipFormat.Xlam, _service.Written.Format);
            Assert.AreEqual(ShipOnExisting.KeepBoth, _service.Written.OnExisting);
            Assert.IsFalse(_service.Written.IsOn(ShipSteps.Comments));
            Assert.AreEqual("..\\out", _service.Written.OutputFolder, "what was not sent stays");
            Assert.AreEqual("xlam", (string)Json(response)["format"]);
        }

        [Test]
        public void Settings_Get_ShowsTreeShakingOffByDefault_WithItsModeAndKeepList()
        {
            var json = Json(Send("GET", "/projects/p/ship/settings"));

            Assert.IsFalse((bool)json["steps"]["treeShaking"]);
            Assert.IsTrue((bool)json["steps"]["comments"]);
            Assert.AreEqual("conservative", (string)json["treeShaking"]["mode"]);
            Assert.AreEqual(0, ((JArray)json["treeShaking"]["keep"]).Count);
        }

        [Test]
        public void Settings_Put_SetsTreeShaking_ModeKeepAndStep()
        {
            var response = Send("PUT", "/projects/p/ship/settings", new { steps = new { treeShaking = true }, treeShaking = new { mode = "entryPointsOnly", keep = new[] { "Module1.Go", "Helper" } } });

            Assert.AreEqual(200, response.Status);
            Assert.IsTrue(_service.Written.IsOn(ShipSteps.TreeShaking));
            Assert.AreEqual(ShipTreeShakingMode.EntryPointsOnly, _service.Written.TreeShaking.Mode);
            CollectionAssert.AreEqual(new[] { "Module1.Go", "Helper" }, _service.Written.TreeShaking.Keep);
            var json = Json(response);
            Assert.IsTrue((bool)json["steps"]["treeShaking"]);
            Assert.AreEqual("entryPointsOnly", (string)json["treeShaking"]["mode"]);
            CollectionAssert.AreEqual(new[] { "Module1.Go", "Helper" }, json["treeShaking"]["keep"].Select(k => (string)k));
        }

        [Test]
        public void Settings_Put_KeepReplacesTheListAndModeAloneLeavesTheKeepList()
        {
            Send("PUT", "/projects/p/ship/settings", new { treeShaking = new { keep = new[] { "A", "B" } } });
            Send("PUT", "/projects/p/ship/settings", new { treeShaking = new { mode = "entryPointsOnly" } });
            CollectionAssert.AreEqual(new[] { "A", "B" }, _service.Written.TreeShaking.Keep);

            Send("PUT", "/projects/p/ship/settings", new { treeShaking = new { keep = new[] { "C" } } });
            CollectionAssert.AreEqual(new[] { "C" }, _service.Written.TreeShaking.Keep);
            Assert.AreEqual(ShipTreeShakingMode.EntryPointsOnly, _service.Written.TreeShaking.Mode);
        }

        [Test]
        public void Settings_Put_RefusesWhatIsNotValidForTreeShaking_AndWritesNothing()
        {
            var bad = new Dictionary<string, object>
            {
                ["unknown mode"] = new { treeShaking = new { mode = "aggressive" } },
                ["mode not a string"] = new { treeShaking = new { mode = 3 } },
                ["keep not an array"] = new { treeShaking = new { keep = "Module1.Go" } },
                ["keep entry not a string"] = new { treeShaking = new { keep = new object[] { "A", 3 } } },
                ["keep too many"] = new { treeShaking = new { keep = Enumerable.Range(0, 501).Select(i => "N" + i).ToArray() } },
                ["keep entry too long"] = new { treeShaking = new { keep = new[] { new string('a', 201) } } },
                ["keep control character"] = new { treeShaking = new { keep = new[] { "Evil" + (char)7 + "Name" } } },
                ["keep newline"] = new { treeShaking = new { keep = new[] { "A\nB" } } },
                ["treeShaking not an object"] = new { treeShaking = "on" },
                ["unknown treeShaking property"] = new { treeShaking = new { everything = true } },
                ["valid mode with an invalid keep"] = new { treeShaking = new { mode = "entryPointsOnly", keep = new[] { new string('a', 201) } } }
            };

            foreach (var item in bad)
            {
                var response = Send("PUT", "/projects/p/ship/settings", item.Value);
                Assert.AreEqual(400, response.Status, item.Key);
            }

            Assert.IsNull(_service.Written, "nothing is written when the request is invalid");
            Assert.AreEqual(ShipTreeShakingMode.Conservative, _service.Settings.TreeShaking.Mode);
        }

        [TestCase("{ \"keep\": [ \" A \", \"Mod.B\" ], \"mode\": \"  EntryPointsOnly \" }")]
        [TestCase("{ \"keep\": [], \"mode\": \"CONSERVATIVE\" }")]
        [TestCase("{ \"keep\": [ \"A\", 3, \"\", \"B\" ] }")]
        [TestCase("{ \"keep\": [ \"A\", \"Bad\\u0007\" ] }")]
        [TestCase("{ \"keep\": \"A\" }")]
        [TestCase("{ \"mode\": \"aggressive\" }")]
        [TestCase("{ \"mode\": 3 }")]
        public void TheSettingsFileAndThePutRouteApplyTheSameRulesWithTheSameWords(string section)
        {
            var config = new Rubberduck.Settings.ProjectConfigReader(null).Parse("{ \"ship\": { \"treeShaking\": " + section + " } }", null);
            var response = Send("PUT", "/projects/p/ship/settings", new { treeShaking = JObject.Parse(section) });

            if (config.Errors.Count == 0)
            {
                Assert.AreEqual(200, response.Status);
                CollectionAssert.AreEqual(config.Ship.TreeShaking.Keep, _service.Written.TreeShaking.Keep);
                Assert.AreEqual(config.Ship.TreeShaking.Mode, _service.Written.TreeShaking.Mode);
            }
            else
            {
                Assert.AreEqual(400, response.Status);
                Assert.IsNull(_service.Written);
                Assert.AreEqual(config.Errors[0], (string)Json(response)["error"]["message"]);
            }
        }

        [Test]
        public void Settings_Put_AcceptsExactlyTheLimits()
        {
            var response = Send("PUT", "/projects/p/ship/settings", new { treeShaking = new { keep = Enumerable.Range(0, 500).Select(i => i == 0 ? new string('a', 200) : "N" + i).ToArray() } });

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual(500, _service.Written.TreeShaking.Keep.Count);
        }

        [Test]
        public void Preview_CarriesTheRemovalsByNameAndTheDynamicCalls()
        {
            var json = Json(Send("GET", "/projects/p/ship/preview"));

            var step = json["steps"].Single(x => (string)x["step"] == "treeShaking");
            Assert.AreEqual(2, (int)step["removed"]);
            CollectionAssert.AreEqual(new[] { "Module1.OldHelper", "Module2.mLegacy" }, step["details"].Select(d => (string)d));
            Assert.AreEqual(2, (int)json["counts"]["treeShaking"]);
            var calls = (JArray)json["dynamicCalls"];
            Assert.AreEqual(1, calls.Count);
            StringAssert.Contains("Application.Run", (string)calls[0]);
        }

        [Test]
        public void Ship_AnswersCarryTheDynamicCallsToo()
        {
            var notConfirmed = Json(Send("POST", "/projects/p/ship", new { }));
            var shipped = Json(Send("POST", "/projects/p/ship", new { confirm = true, onExisting = "keepBoth" }));

            Assert.AreEqual(1, ((JArray)notConfirmed["dynamicCalls"]).Count);
            Assert.AreEqual(1, ((JArray)shipped["dynamicCalls"]).Count);
            Assert.AreEqual("Module1.OldHelper", (string)shipped["steps"].Single(x => (string)x["step"] == "treeShaking")["details"][0]);
        }

        [Test]
        public void Settings_Put_RefusedWithoutApproval_WritesNothing()
        {
            _allow = false;

            var response = Send("PUT", "/projects/p/ship/settings", new { format = "xlam" });

            Assert.AreEqual(403, response.Status);
            CollectionAssert.DoesNotContain(_service.Calls, "WriteSettings");
        }

        [TestCase(true)]
        [TestCase(false)]
        public void Settings_Put_CannotSetUnattended(bool value)
        {
            var response = Send("PUT", "/projects/p/ship/settings", new { unattended = value });

            Assert.AreEqual(400, response.Status);
            Assert.AreEqual("unattended_read_only", (string)Json(response)["error"]["code"]);
            Assert.IsEmpty(_service.Calls);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Settings_Put_RefusesBadValues()
        {
            Assert.AreEqual(400, Send("PUT", "/projects/p/ship/settings", new { format = "exe" }).Status);
            Assert.AreEqual(400, Send("PUT", "/projects/p/ship/settings", new { onExisting = "overwrite" }).Status);
            Assert.AreEqual(400, Send("PUT", "/projects/p/ship/settings", new { steps = new { nothing = true } }).Status);
            Assert.AreEqual(400, Send("PUT", "/projects/p/ship/settings", new { bogus = 1 }).Status);
            Assert.AreEqual(400, Send("PUT", "/projects/p/ship/settings").Status);
            Assert.IsEmpty(_service.Calls);
        }

        [Test]
        public void Preview_RunsOffTheEditorThread_BecauseTreeShakingWaitsForTheParserThatWaitsForIt()
        {
            Send("GET", "/projects/p/ship/preview");
            Send("POST", "/projects/p/ship");
            Send("POST", "/projects/p/ship", new { confirm = true });

            CollectionAssert.AreEqual(new[] { "Preview", "Preview", "Ship" }, _service.Calls);
            Assert.IsFalse(_service.PreviewedInsideWith, "IShipService.Preview must not run inside IShipProjects.With");
        }

        [Test]
        public void Preview_And_Ship_ListWhatKeptTheCode_WithTheTotal()
        {
            var preview = Json(Send("GET", "/projects/p/ship/preview"));
            var shipped = Json(Send("POST", "/projects/p/ship", new { confirm = true, onExisting = "keepBoth" }));

            Assert.AreEqual(3, ((JArray)preview["roots"]).Count);
            Assert.AreEqual(3, (int)preview["rootsTotal"]);
            StringAssert.Contains("public procedure", (string)preview["roots"][0]);
            Assert.AreEqual(3, ((JArray)shipped["roots"]).Count);
        }

        [Test]
        public void Roots_AreCappedAt500_ButTheTotalIsTold()
        {
            _service.RootCount = 800;

            var json = Json(Send("GET", "/projects/p/ship/preview"));

            Assert.AreEqual(500, ((JArray)json["roots"]).Count);
            Assert.AreEqual(800, (int)json["rootsTotal"]);
        }

        [Test]
        public void Preview_IsReadOnly_AndNeedsNoApproval()
        {
            var response = Send("GET", "/projects/p/ship/preview");

            Assert.AreEqual(200, response.Status);
            var json = Json(response);
            Assert.IsTrue((bool)json["dryRun"]);
            Assert.AreEqual(2, (int)json["counts"]["notes"]);
            Assert.AreEqual(40, (int)json["counts"]["comments"]);
            Assert.AreEqual("Broken: syntax error", (string)json["unparsedModules"][0]);
            Assert.AreEqual(0, _asked);
            CollectionAssert.AreEqual(new[] { "Preview" }, _service.Calls);
        }

        [Test]
        public void Ship_WithoutConfirm_OnlyPreviews_AndSaysWhatToSend()
        {
            var response = Send("POST", "/projects/p/ship");

            Assert.AreEqual(200, response.Status);
            var json = Json(response);
            Assert.IsFalse((bool)json["shipped"]);
            Assert.IsTrue((bool)json["confirmRequired"]);
            StringAssert.Contains("\"confirm\": true", (string)json["note"]);
            CollectionAssert.AreEqual(new[] { "Preview" }, _service.Calls);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Ship_WithConfirm_NeedsApproval_AndReturnsTheReport()
        {
            var response = Send("POST", "/projects/p/ship", new { confirm = true });

            Assert.AreEqual(200, response.Status);
            var json = Json(response);
            Assert.IsTrue((bool)json["shipped"]);
            Assert.IsFalse((bool)json["dryRun"]);
            Assert.AreEqual(@"C:\work\out\Book.xlsm", (string)json["outputPath"]);
            Assert.AreEqual(2, (int)json["counts"]["notes"]);
            Assert.AreEqual(1, _asked);
            CollectionAssert.AreEqual(new[] { "Ship" }, _service.Calls);
            Assert.AreEqual(ShipOnExisting.Ask, _service.LastOnExisting);
        }

        [Test]
        public void Ship_Denied_ShipsNothing()
        {
            _allow = false;

            var response = Send("POST", "/projects/p/ship", new { confirm = true });

            Assert.AreEqual(403, response.Status);
            CollectionAssert.DoesNotContain(_service.Calls, "Ship");
        }

        [Test]
        public void Ship_ConfirmMustBeABoolean()
        {
            Assert.AreEqual(400, Send("POST", "/projects/p/ship", new { confirm = "yes" }).Status);
        }

        [Test]
        public void Ship_Unattended_WaivesConfirmAndApproval()
        {
            _service.Unattended = true;

            var response = Send("POST", "/projects/p/ship");

            Assert.AreEqual(200, response.Status);
            Assert.IsTrue((bool)Json(response)["shipped"]);
            Assert.IsTrue((bool)Json(response)["unattended"]);
            Assert.AreEqual(0, _asked);
            CollectionAssert.AreEqual(new[] { "Ship" }, _service.Calls);
        }

        [TestCase("keepBoth", ShipOnExisting.KeepBoth)]
        [TestCase("KEEPBOTH", ShipOnExisting.KeepBoth)]
        [TestCase("ask", ShipOnExisting.Ask)]
        public void Ship_OnExistingInTheBody_IsPassedOn(string text, ShipOnExisting expected)
        {
            Send("POST", "/projects/p/ship", new { confirm = true, onExisting = text });

            Assert.AreEqual(expected, _service.LastOnExisting);
        }

        [Test]
        public void Ship_Replace_IsLeftToTheUser_UnlessTheyAreUnattended()
        {
            var refused = Send("POST", "/projects/p/ship", new { confirm = true, onExisting = "replace" });

            Assert.AreEqual(403, refused.Status);
            StringAssert.Contains("replace_needs_user", Json(refused).ToString());
            Assert.AreEqual(0, _asked);
            CollectionAssert.DoesNotContain(_service.Calls, "Ship");

            _service.Unattended = true;
            Assert.AreEqual(200, Send("POST", "/projects/p/ship", new { confirm = true, onExisting = "replace" }).Status);
            Assert.AreEqual(ShipOnExisting.Replace, _service.LastOnExisting);
        }

        [Test]
        public void Ship_BadOnExisting_IsABadRequest_AndAsksNobody()
        {
            Assert.AreEqual(400, Send("POST", "/projects/p/ship", new { confirm = true, onExisting = "clobber" }).Status);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Ship_WhenTheFileExists_Is409_WithPathAndChoices()
        {
            _service.Refuse = new ShipRefusedException("exists", "Book.xlsm exists.") { Path = @"C:\work\out\Book.xlsm" };

            var response = Send("POST", "/projects/p/ship", new { confirm = true });

            Assert.AreEqual(409, response.Status);
            var error = Json(response)["error"];
            Assert.AreEqual("ship_exists", (string)error["code"]);
            Assert.AreEqual("exists", (string)error["reason"]);
            Assert.AreEqual(@"C:\work\out\Book.xlsm", (string)error["path"]);
            CollectionAssert.AreEqual(new[] { "ask", "replace", "keepBoth" }, error["choices"].Select(c => (string)c).ToArray());
            StringAssert.Contains("onExisting", (string)error["message"]);
        }

        [TestCase("unsaved", 409)]
        [TestCase("unsupportedHost", 422)]
        [TestCase("notExcel", 422)]
        [TestCase("folderMissing", 422)]
        [TestCase("folderNotWritable", 422)]
        [TestCase("vbomNotTrusted", 422)]
        [TestCase("excelUnavailable", 503)]
        [TestCase("somethingNew", 500)]
        public void Ship_OtherRefusals_MapToStatuses_WithTheReasonInTheBody(string reason, int status)
        {
            _service.Refuse = new ShipRefusedException(reason, "because " + reason);

            var response = Send("POST", "/projects/p/ship", new { confirm = true });

            Assert.AreEqual(status, response.Status);
            Assert.AreEqual(reason, (string)Json(response)["error"]["reason"]);
            Assert.AreEqual("ship_" + reason, (string)Json(response)["error"]["code"]);
            StringAssert.Contains("because " + reason, (string)Json(response)["error"]["message"]);
        }

        [Test]
        public void Preview_Refusal_IsMapped()
        {
            _service.Refuse = new ShipRefusedException("unsaved", "Save the workbook first.");

            Assert.AreEqual(409, Send("GET", "/projects/p/ship/preview").Status);
        }

        [Test]
        public void UnknownProject_Is404()
        {
            Assert.AreEqual(404, Send("GET", "/projects/missing/ship/settings").Status);
        }

        [Test]
        public void WrongMethods_Are405()
        {
            Assert.AreEqual(405, Send("GET", "/projects/p/ship").Status);
            Assert.AreEqual(405, Send("POST", "/projects/p/ship/preview").Status);
            Assert.AreEqual(405, Send("DELETE", "/projects/p/ship/settings").Status);
            Assert.AreEqual(404, Send("GET", "/projects/p/ship/other").Status);
        }

        [Test]
        public void Routes_ListsEachShipRoute_AndOmitsThemWithoutTheService()
        {
            var listed = ((IEnumerable<string>)Json(Send("GET", ""))["routes"].Select(r => (string)r)).ToList();
            CollectionAssert.IsSupersetOf(listed, new[]
            {
                "GET,PUT /api/v1/projects/{p}/ship/settings",
                "GET /api/v1/projects/{p}/ship/preview",
                "POST /api/v1/projects/{p}/ship"
            });

            var bare = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api);
            var without = Json(bare.Handle(new ApiRequest { Method = "GET", Path = "/api/v1", Client = "t" }))["routes"].Select(r => (string)r).ToList();
            Assert.IsFalse(without.Any(r => r.Contains("/ship")));
            Assert.AreEqual(404, bare.Handle(new ApiRequest { Method = "GET", Path = "/api/v1/projects/p/ship/settings", Client = "t" }).Status);
        }

        private JObject Call(string tool, object arguments)
            => Json(_router.Handle(new ApiRequest
            {
                Method = "POST",
                Path = "/mcp",
                Client = "agent",
                Body = JObject.FromObject(new { jsonrpc = "2.0", id = 1, method = "tools/call", @params = new { name = tool, arguments } }).ToString()
            }))["result"] as JObject;

        [Test]
        public void Mcp_OffersTheFourTools_WithTheCaveats()
        {
            var response = Json(_router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Client = "t", Body = "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"tools/list\"}" }));
            var tools = response["result"]["tools"].ToDictionary(t => (string)t["name"], t => (string)t["description"]);

            CollectionAssert.IsSubsetOf(new[] { "get_ship_settings", "set_ship_settings", "preview_ship", "ship_project" }, tools.Keys);
            foreach (var name in new[] { "get_ship_settings", "set_ship_settings", "preview_ship", "ship_project" })
            {
                StringAssert.Contains("never from the open project", tools[name], name);
                StringAssert.Contains("saved", tools[name], name);
                StringAssert.Contains("Preview first", tools[name], name);
                StringAssert.Contains(".xlam", tools[name], name);
            }

            StringAssert.Contains("unattended", tools["set_ship_settings"]);
            StringAssert.Contains("confirm", tools["ship_project"]);
        }

        [Test]
        public void Mcp_ToolDescriptions_WarnThatTreeShakingIsOffByDefaultAndCanRemoveNeededCode()
        {
            var response = Json(_router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Client = "t", Body = "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"tools/list\"}" }));
            var tools = response["result"]["tools"].ToDictionary(t => (string)t["name"], t => (string)t["description"]);

            foreach (var name in new[] { "get_ship_settings", "set_ship_settings", "preview_ship", "ship_project" })
            {
                StringAssert.Contains("Tree shaking", tools[name], name);
                StringAssert.Contains("OFF by default", tools[name], name);
                StringAssert.Contains("preview", tools[name].ToLowerInvariant(), name);
            }

            StringAssert.Contains("CAN REMOVE CODE THE USER NEEDS", tools["set_ship_settings"]);
            StringAssert.Contains("dynamicCalls", tools["preview_ship"]);
            var setProperties = response["result"]["tools"].Single(t => (string)t["name"] == "set_ship_settings")["inputSchema"]["properties"];
            Assert.IsNotNull(setProperties["treeShaking"]);
        }

        [Test]
        public void Mcp_Instructions_TellAgentsHowToWriteCallsTreeShakingCanFollow()
        {
            var response = Json(_router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Client = "t", Body = "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"initialize\",\"params\":{}}" }));
            var instructions = (string)response["result"]["instructions"];

            StringAssert.Contains("Application.Run", instructions);
            StringAssert.Contains("CallByName", instructions);
            StringAssert.Contains("Application.OnTime", instructions);
            StringAssert.Contains("'@EntryPoint", instructions);
            StringAssert.Contains("keep list", instructions);
            StringAssert.Contains("off by default", instructions);
            StringAssert.Contains("preview_ship", instructions);
        }

        [Test]
        public void Mcp_Tools_DoTheSame()
        {
            var settings = Call("get_ship_settings", new { project = "p" });
            var preview = Call("preview_ship", new { project = "p" });
            var notConfirmed = Call("ship_project", new { project = "p" });
            var shipped = Call("ship_project", new { project = "p", confirm = true, onExisting = "keepBoth" });
            var unattended = Call("set_ship_settings", new { project = "p", unattended = true });

            Assert.IsFalse((bool)settings["isError"]);
            Assert.IsFalse((bool)preview["isError"]);
            StringAssert.Contains("\"confirmRequired\": true", (string)notConfirmed["content"][0]["text"]);
            Assert.IsFalse((bool)shipped["isError"]);
            StringAssert.Contains("\"shipped\": true", (string)shipped["content"][0]["text"]);
            Assert.AreEqual(ShipOnExisting.KeepBoth, _service.LastOnExisting);
            Assert.IsTrue((bool)unattended["isError"]);
            StringAssert.Contains("unattended_read_only", (string)unattended["content"][0]["text"]);
            Assert.IsFalse(_service.Unattended);
            CollectionAssert.DoesNotContain(_service.Calls, "SetUnattended");
        }

        [Test]
        public void Mcp_ExistsError_NamesPathAndChoices()
        {
            _service.Refuse = new ShipRefusedException("exists", "in the way") { Path = @"C:\work\out\Book.xlsm" };

            var result = Call("ship_project", new { project = "p", confirm = true });

            Assert.IsTrue((bool)result["isError"]);
            var text = (string)result["content"][0]["text"];
            StringAssert.Contains("ship_exists", text);
            StringAssert.Contains(@"C:\work\out\Book.xlsm", text);
            StringAssert.Contains("keepBoth", text);
        }
    }
}
