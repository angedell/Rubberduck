using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using Moq;
using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    /// <summary>A store that keeps modules in memory and enforces the hash like the real one.</summary>
    internal class FakeStore : IModuleStore
    {
        public Dictionary<string, string> Code = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["Module1"] = "Option Explicit\r\nSub A()\r\nEnd Sub" };
        public List<string> Calls = new List<string>();

        private ModuleInfo Info(string name) => new ModuleInfo { Project = "VBAProject", Name = name, Kind = "standard", LineCount = 1, Hash = ModuleApi.HashOf(Code[name]) };

        private void Check(string name, string hash)
        {
            if (!Code.ContainsKey(name)) throw ApiException.NotFound("no module");
            if (ModuleApi.HashOf(Code[name]) != hash) throw ApiException.PreconditionFailed("stale");
        }

        public IReadOnlyList<ProjectInfo> Projects() => new[] { new ProjectInfo { Id = "p1", Name = "VBAProject" } };
        public IReadOnlyList<ModuleInfo> Modules(string project) => Code.Keys.Select(Info).ToList();

        public ModuleContent Read(string project, string module, bool includeText)
        {
            if (!Code.ContainsKey(module)) throw ApiException.NotFound("no module");
            var info = Info(module);
            return new ModuleContent { Project = info.Project, Name = info.Name, Kind = info.Kind, Hash = info.Hash, Code = Code[module], Text = includeText ? "Attribute VB_Name = \"" + module + "\"\r\n" + Code[module] : null };
        }

        public ModuleInfo ReplaceCode(string project, string module, string code, string expectedHash)
        {
            Check(module, expectedHash); Calls.Add("ReplaceCode"); Code[module] = code; return Info(module);
        }

        public ModuleInfo EditLines(string project, string module, int startLine, int endLine, string replacement, string expectedHash)
        {
            Check(module, expectedHash); Calls.Add("EditLines");
            var lines = Code[module].Split(new[] { "\r\n" }, StringSplitOptions.None).ToList();
            if (endLine >= startLine) lines.RemoveRange(startLine - 1, endLine - startLine + 1);
            if (replacement.Length > 0) lines.InsertRange(startLine - 1, replacement.Split(new[] { "\r\n" }, StringSplitOptions.None));
            Code[module] = string.Join("\r\n", lines);
            return Info(module);
        }

        public ModuleInfo ReplaceText(string project, string module, string text, string expectedHash)
        {
            Check(module, expectedHash); Calls.Add("ReplaceText"); Code[module] = text; return Info(module);
        }

        public ModuleInfo Create(string project, string module, string kind, string code)
        {
            if (Code.ContainsKey(module)) throw ApiException.Conflict("exists");
            Calls.Add("Create"); Code[module] = code; return Info(module);
        }

        public void Delete(string project, string module, string expectedHash)
        {
            Check(module, expectedHash); Calls.Add("Delete"); Code.Remove(module);
        }
    }

    internal static class InMemory
    {
        /// <summary>Just enough of a file system for <see cref="ApiToken"/>: text files by path.</summary>
        public static IFileSystem FileSystem()
        {
            var files = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var file = new Mock<IFile>();
            file.Setup(f => f.Exists(It.IsAny<string>())).Returns<string>(files.ContainsKey);
            file.Setup(f => f.ReadAllText(It.IsAny<string>())).Returns<string>(path => files[path]);
            file.Setup(f => f.WriteAllText(It.IsAny<string>(), It.IsAny<string>())).Callback<string, string>((path, text) => files[path] = text);
            var directory = new Mock<IDirectory>();
            var path_ = new Mock<IPath>();
            path_.Setup(p => p.GetDirectoryName(It.IsAny<string>())).Returns<string>(System.IO.Path.GetDirectoryName);

            var fileSystem = new Mock<IFileSystem>();
            fileSystem.Setup(f => f.File).Returns(file.Object);
            fileSystem.Setup(f => f.Directory).Returns(directory.Object);
            fileSystem.Setup(f => f.Path).Returns(path_.Object);
            return fileSystem.Object;
        }
    }

    [TestFixture]
    [Category("Api")]
    public class ApiTokenTests
    {
        [Test]
        public void Token_IsKeptAcrossInstances()
        {
            var fileSystem = InMemory.FileSystem();
            var first = new ApiToken(fileSystem, @"C:\cfg\api-token.txt").Value;

            Assert.AreEqual(first, new ApiToken(fileSystem, @"C:\cfg\api-token.txt").Value);
        }

        [Test]
        public void Regenerate_LocksOutTheOldToken()
        {
            var token = new ApiToken(InMemory.FileSystem(), @"C:\cfg\api-token.txt");
            var old = token.Value;

            token.Regenerate();

            Assert.IsFalse(token.Matches(old));
            Assert.IsTrue(token.Matches(token.Value));
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase("rd_wrong")]
        public void Matches_RejectsWrongTokens(string presented)
        {
            Assert.IsFalse(new ApiToken(InMemory.FileSystem(), @"C:\cfg\t.txt").Matches(presented));
        }
    }

    [TestFixture]
    [Category("Api")]
    public class ApiRequestGuardTests
    {
        private ApiToken _token;
        private bool _network;
        private DateTime _now = new DateTime(2026, 1, 1, 12, 0, 0, DateTimeKind.Utc);

        [SetUp]
        public void SetUp()
        {
            _token = new ApiToken(InMemory.FileSystem(), @"C:\cfg\t.txt");
            _network = false;
        }

        private ApiRequestGuard Guard() => new ApiRequestGuard(_token, () => _network, () => _now);
        private string Bearer => "Bearer " + _token.Value;

        [Test]
        public void ValidRequest_Passes() => Assert.DoesNotThrow(() => Guard().Check("127.0.0.1:7437", null, Bearer, "1.2.3.4"));

        [Test]
        public void MissingToken_Is401()
            => Assert.AreEqual(401, Assert.Throws<ApiException>(() => Guard().Check("localhost:7437", null, null, "r")).Status);

        [Test]
        public void AnyOriginHeader_IsRefused_EvenWithTheRightToken()
            => Assert.AreEqual("origin_refused", Assert.Throws<ApiException>(() => Guard().Check("localhost:7437", "http://evil.example", Bearer, "r")).Code);

        [TestCase("evil.example:7437")]
        [TestCase("192.168.1.5:7437")]
        public void NonLoopbackHost_IsRefused_WhenNetworkIsOff(string host)
            => Assert.AreEqual("host_refused", Assert.Throws<ApiException>(() => Guard().Check(host, null, Bearer, "r")).Code);

        [Test]
        public void NonLoopbackHost_Passes_WhenNetworkIsOn()
        {
            _network = true;
            Assert.DoesNotThrow(() => Guard().Check("192.168.1.5:7437", null, Bearer, "r"));
        }

        [TestCase("[::1]:7437")]
        [TestCase("LOCALHOST")]
        public void LoopbackHostForms_Pass(string host) => Assert.DoesNotThrow(() => Guard().Check(host, null, Bearer, "r"));

        [Test]
        public void TenWrongTokens_LockTheAddressOut_ForAMinute()
        {
            var guard = Guard();
            for (var i = 0; i < 10; i++)
            {
                Assert.Throws<ApiException>(() => guard.Check("localhost", null, "Bearer nope", "9.9.9.9"));
            }

            Assert.AreEqual(429, Assert.Throws<ApiException>(() => guard.Check("localhost", null, Bearer, "9.9.9.9")).Status);
            Assert.DoesNotThrow(() => guard.Check("localhost", null, Bearer, "8.8.8.8"));

            _now = _now.AddSeconds(61);
            Assert.DoesNotThrow(() => guard.Check("localhost", null, Bearer, "9.9.9.9"));
        }
    }

    [TestFixture]
    [Category("Api")]
    public class ApiWriteApprovalTests
    {
        [Test]
        public void Yes_IsAskedOnce()
        {
            var asked = 0;
            var approval = new ApiWriteApproval(_ => { asked++; return true; });

            approval.Require("agent");
            approval.Require("agent");

            Assert.AreEqual(1, asked);
            Assert.IsTrue(approval.IsApproved);
        }

        [Test]
        public void No_IsRemembered_ForAMinute_ThenAsksAgain()
        {
            var now = DateTime.UtcNow;
            var asked = 0;
            var approval = new ApiWriteApproval(_ => { asked++; return false; }, () => now);

            Assert.AreEqual("write_denied", Assert.Throws<ApiException>(() => approval.Require("agent")).Code);
            Assert.Throws<ApiException>(() => approval.Require("agent"));
            Assert.AreEqual(1, asked);

            now = now.AddSeconds(61);
            Assert.Throws<ApiException>(() => approval.Require("agent"));
            Assert.AreEqual(2, asked);
        }

        [Test]
        public void Reset_MakesTheNextWriteAskAgain()
        {
            var asked = 0;
            var approval = new ApiWriteApproval(_ => { asked++; return true; });
            approval.Require("agent");

            approval.Reset();
            approval.Require("agent");

            Assert.AreEqual(2, asked);
        }
    }

    [TestFixture]
    [Category("Api")]
    public class ApiRouterTests
    {
        private FakeStore _store;
        private int _asked;
        private bool _allow;
        private ApiRouter _router;

        [SetUp]
        public void SetUp()
        {
            _store = new FakeStore();
            _asked = 0;
            _allow = true;
            _router = new ApiRouter(new ModuleApi(_store, new ApiWriteApproval(_ => { _asked++; return _allow; })), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api);
        }

        private ApiResponse Send(string method, string path, object body = null, string ifMatch = null)
            => _router.Handle(new ApiRequest { Method = method, Path = path, Body = body == null ? null : JObject.FromObject(body).ToString(), IfMatch = ifMatch, Client = "test" });

        private static string ErrorCode(ApiResponse response) => (string)JObject.FromObject(response.Body)["error"]["code"];
        private string Hash => ModuleApi.HashOf(_store.Code["Module1"]);

        [Test]
        public void ListProjects_NeedsNoApproval()
        {
            var response = Send("GET", "/api/v1/projects");

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void ReadModule_ReturnsCodeAndHash_AsETag()
        {
            var response = Send("GET", "/api/v1/projects/p1/modules/Module1");

            var module = (ModuleContent)response.Body;
            Assert.AreEqual(Hash, module.Hash);
            Assert.AreEqual(Hash, response.ETag);
            Assert.IsNull(module.Text);
        }

        [Test]
        public void ReadModule_WithIncludeText_ReturnsTheHeader()
        {
            var request = new ApiRequest { Method = "GET", Path = "/api/v1/projects/p1/modules/Module1" };
            request.Query["include"] = "text";

            Assert.IsNotNull(((ModuleContent)_router.Handle(request).Body).Text);
        }

        [Test]
        public void Write_WithoutHash_Is428_AndNothingIsAsked()
        {
            var response = Send("PUT", "/api/v1/projects/p1/modules/Module1", new { code = "x" });

            Assert.AreEqual(428, response.Status);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Write_WithStaleHash_Is412_AndLeavesTheCode()
        {
            var response = Send("PUT", "/api/v1/projects/p1/modules/Module1", new { code = "x", expectedHash = "stale" });

            Assert.AreEqual(412, response.Status);
            StringAssert.StartsWith("Option Explicit", _store.Code["Module1"]);
        }

        [Test]
        public void Write_WithHashInBody_ReplacesCode_AndReturnsTheNewHash()
        {
            var response = Send("PUT", "/api/v1/projects/p1/modules/Module1", new { code = "Sub B()\r\nEnd Sub", expectedHash = Hash });

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual("Sub B()\r\nEnd Sub", _store.Code["Module1"]);
            Assert.AreEqual(ModuleApi.HashOf("Sub B()\r\nEnd Sub"), response.ETag);
        }

        [Test]
        public void Write_WithIfMatchHeader_UsesItAsTheHash()
        {
            var response = Send("PUT", "/api/v1/projects/p1/modules/Module1", new { code = "' new" }, "\"" + Hash + "\"");

            Assert.AreEqual(200, response.Status);
        }

        [Test]
        public void Write_WithText_ReplacesTheWholeModule()
        {
            Send("PUT", "/api/v1/projects/p1/modules/Module1", new { text = "Attribute VB_Name = \"Module1\"\r\n' x", expectedHash = Hash });

            CollectionAssert.AreEqual(new[] { "ReplaceText" }, _store.Calls);
        }

        [Test]
        public void Write_WithTextLackingTheHeader_IsRefused_AndChangesNothing()
        {
            var response = Send("PUT", "/api/v1/projects/p1/modules/Module1", new { text = "Assert.AreEqual 0, sut.Count", expectedHash = Hash });

            Assert.AreEqual(422, response.Status);
            Assert.AreEqual("text_not_a_module", ErrorCode(response));
            Assert.IsEmpty(_store.Calls);
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void Write_WithTextOfAnotherModule_IsRefused()
        {
            var response = Send("PUT", "/api/v1/projects/p1/modules/Module1", new { text = "Attribute VB_Name = \"Other\"\r\n' x", expectedHash = Hash });

            Assert.AreEqual(422, response.Status);
            Assert.IsEmpty(_store.Calls);
        }

        [Test]
        public void Batch_ReplaceTextLackingTheHeader_FailsTheWholeBatch_BeforeAnythingIsChanged()
        {
            var edits = new object[]
            {
                new { op = "replaceCode", module = "Module1", expectedHash = Hash, code = "' fine" },
            };
            Assert.AreEqual(200, Send("POST", "/api/v1/projects/p1/edits", new { edits }).Status);
            _store.Calls.Clear();

            var bad = new object[] { new { op = "replaceText", module = "Module1", expectedHash = Hash, text = "just a line" } };
            var response = Send("POST", "/api/v1/projects/p1/edits", new { edits = bad });

            Assert.AreEqual(422, response.Status);
            Assert.IsEmpty(_store.Calls);
        }

        [Test]
        public void Write_WithAttributes_SetsThemWithoutTheCallerRebuildingTheHeader()
        {
            var response = Send("PUT", "/api/v1/projects/p1/modules/Module1", new { attributes = new { module = new { VB_PredeclaredId = true } }, expectedHash = Hash });

            Assert.AreEqual(200, response.Status);
            CollectionAssert.AreEqual(new[] { "ReplaceText" }, _store.Calls);
            StringAssert.Contains("Attribute VB_PredeclaredId = True", _store.Code["Module1"]);
        }

        [Test]
        public void Read_WithIncludeAttributes_GivesThemAsData_WithoutTheText()
        {
            _store.Code["Module1"] = "Attribute VB_PredeclaredId = True\r\nOption Explicit";
            var request = new ApiRequest { Method = "GET", Path = "/api/v1/projects/p1/modules/Module1", Client = "test" };
            request.Query["include"] = "attributes";

            var content = (ModuleContent)_router.Handle(request).Body;

            Assert.AreEqual(true, (bool)content.Attributes["module"]["VB_PredeclaredId"]);
            Assert.IsNull(content.Text);
        }

        [Test]
        public void ApprovalIsAskedOnce_AcrossWrites()
        {
            Send("PUT", "/api/v1/projects/p1/modules/Module1", new { code = "' 1", expectedHash = Hash });
            Send("PUT", "/api/v1/projects/p1/modules/Module1", new { code = "' 2", expectedHash = Hash });

            Assert.AreEqual(1, _asked);
            Assert.AreEqual(2, _store.Calls.Count);
        }

        [Test]
        public void DeniedWrite_Is403_AndChangesNothing()
        {
            _allow = false;

            var response = Send("PUT", "/api/v1/projects/p1/modules/Module1", new { code = "' no", expectedHash = Hash });

            Assert.AreEqual(403, response.Status);
            Assert.AreEqual("write_denied", ErrorCode(response));
            Assert.IsEmpty(_store.Calls);
        }

        [Test]
        public void AttributeLinesInCode_AreRefused()
        {
            var response = Send("PUT", "/api/v1/projects/p1/modules/Module1", new { code = "Attribute VB_Name = \"X\"\r\nSub A()\r\nEnd Sub", expectedHash = Hash });

            Assert.AreEqual(422, response.Status);
            Assert.AreEqual("header_in_code", ErrorCode(response));
        }

        [Test]
        public void EditLines_ReplacesTheRange()
        {
            Send("PATCH", "/api/v1/projects/p1/modules/Module1", new { startLine = 2, endLine = 3, replacement = "Sub Z()\r\nEnd Sub", expectedHash = Hash });

            Assert.AreEqual("Option Explicit\r\nSub Z()\r\nEnd Sub", _store.Code["Module1"]);
        }

        [Test]
        public void EditLines_EndOneBeforeStart_Inserts()
        {
            Send("PATCH", "/api/v1/projects/p1/modules/Module1", new { startLine = 1, endLine = 0, replacement = "' top", expectedHash = Hash });

            StringAssert.StartsWith("' top\r\nOption Explicit", _store.Code["Module1"]);
        }

        [Test]
        public void EditLines_WithBadRange_Is400()
            => Assert.AreEqual(400, Send("PATCH", "/api/v1/projects/p1/modules/Module1", new { startLine = 0, endLine = 1, replacement = "x", expectedHash = Hash }).Status);

        [Test]
        public void Create_Returns201()
        {
            var response = Send("POST", "/api/v1/projects/p1/modules", new { name = "modNew", kind = "standard", code = "Option Explicit" });

            Assert.AreEqual(201, response.Status);
            Assert.IsTrue(_store.Code.ContainsKey("modNew"));
        }

        [TestCase("1bad")]
        [TestCase("has space")]
        [TestCase("")]
        public void Create_WithInvalidName_Is400(string name)
            => Assert.AreEqual(400, Send("POST", "/api/v1/projects/p1/modules", new { name, kind = "standard" }).Status);

        [Test]
        public void Create_WithUnknownKind_Is400()
            => Assert.AreEqual(400, Send("POST", "/api/v1/projects/p1/modules", new { name = "modX", kind = "document" }).Status);

        [Test]
        public void Create_WhenItExists_Is409()
            => Assert.AreEqual(409, Send("POST", "/api/v1/projects/p1/modules", new { name = "Module1", kind = "class" }).Status);

        [Test]
        public void Delete_NeedsTheHash_AndReturns204()
        {
            Assert.AreEqual(428, Send("DELETE", "/api/v1/projects/p1/modules/Module1").Status);
            Assert.AreEqual(204, Send("DELETE", "/api/v1/projects/p1/modules/Module1", null, Hash).Status);
            Assert.IsFalse(_store.Code.ContainsKey("Module1"));
        }

        [Test]
        public void MalformedJson_Is400()
        {
            var response = _router.Handle(new ApiRequest { Method = "PUT", Path = "/api/v1/projects/p1/modules/Module1", Body = "{nope" });

            Assert.AreEqual(400, response.Status);
        }

        [Test]
        public void UnknownRoute_Is404()
            => Assert.AreEqual(404, Send("GET", "/api/v1/cells").Status);

        [Test]
        public void WrongMethod_Is405()
            => Assert.AreEqual(405, Send("POST", "/api/v1/projects").Status);

        [Test]
        public void ModuleNamesInThePath_AreUnescaped()
        {
            _store.Code["Mod One"] = "x";

            Assert.AreEqual(200, Send("GET", "/api/v1/projects/p1/modules/Mod%20One").Status);
        }
    }

    [TestFixture]
    [Category("Api")]
    public class McpEndpointTests
    {
        private FakeStore _store;
        private ApiRouter _router;
        private bool _immediateOn;

        [SetUp]
        public void SetUp()
        {
            _immediateOn = false;
            _store = new FakeStore();
            _router = new ApiRouter(new ModuleApi(_store, new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => _immediateOn).Api);
        }

        private JObject Rpc(string method, object parameters = null, int? id = 1)
        {
            var message = new JObject { ["jsonrpc"] = "2.0", ["method"] = method };
            if (id.HasValue) message["id"] = id.Value;
            if (parameters != null) message["params"] = JObject.FromObject(parameters);
            var response = _router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Body = message.ToString(), Client = "test" });
            return response.Body as JObject;
        }

        private JObject Call(string tool, object arguments) => Rpc("tools/call", new { name = tool, arguments });
        private static string TextOf(JObject result) => (string)result["result"]["content"][0]["text"];

        [Test]
        public void Initialize_AnnouncesTools()
        {
            var result = Rpc("initialize", new { protocolVersion = "2025-03-26" });

            Assert.IsNotNull(result["result"]["capabilities"]["tools"]);
            Assert.AreEqual("rubberduck", (string)result["result"]["serverInfo"]["name"]);
        }

        [TestCase(true)]
        [TestCase(false)]
        public void TheDebugPrintAdvice_IsGivenOnlyWhileTheToolForItIsOffered(bool immediateOn)
        {
            _immediateOn = immediateOn;

            var instructions = (string)Rpc("initialize", new { protocolVersion = "2025-03-26" })["result"]["instructions"];
            var offered = Rpc("tools/list")["result"]["tools"].Any(t => (string)t["name"] == "read_debug_output");

            Assert.AreEqual(immediateOn, offered);
            Assert.AreEqual(immediateOn, instructions.Contains("DEBUGGING WITH Debug.Print"));
            Assert.AreEqual(immediateOn, instructions.Contains("#Const DebugLog"));
        }

        [Test]
        public void ToolResults_UseCamelCase_LikeTheRestApi()
        {
            _immediateOn = true;

            var answer = JObject.Parse(TextOf(Call("read_debug_output", new { })));

            CollectionAssert.AreEquivalent(new[] { "text", "lineCount", "next", "more", "missed" }, answer.Properties().Select(p => p.Name));
        }

        [Test]
        public void Notification_Gets202_AndNoBody()
        {
            var response = _router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Body = "{\"jsonrpc\":\"2.0\",\"method\":\"notifications/initialized\"}" });

            Assert.AreEqual(202, response.Status);
            Assert.IsNull(response.Body);
        }

        [Test]
        public void ToolsList_HasTheModuleTools()
        {
            var names = Rpc("tools/list")["result"]["tools"].Select(t => (string)t["name"]).ToList();

            CollectionAssert.IsSupersetOf(names, new[] { "list_projects", "list_modules", "read_module", "write_module", "edit_lines", "create_module", "delete_module", "get_diagnostics" });
            CollectionAssert.DoesNotContain(names, "run_immediate");
        }

        [Test]
        public void ImmediateTools_AreOffered_OnlyWhenTheUserAllowsThem()
        {
            _immediateOn = true;

            var names = Rpc("tools/list")["result"]["tools"].Select(t => (string)t["name"]).ToList();

            CollectionAssert.IsSupersetOf(names, new[] { "read_immediate", "run_immediate" });
        }

        [Test]
        public void ReadModule_ReturnsTheHash()
        {
            var result = Call("read_module", new { project = "p1", module = "Module1" });

            Assert.IsFalse((bool)result["result"]["isError"]);
            StringAssert.Contains(ModuleApi.HashOf(_store.Code["Module1"]), TextOf(result));
        }

        [Test]
        public void WriteWithStaleHash_IsAToolError_NotAProtocolError()
        {
            var result = Call("write_module", new { project = "p1", module = "Module1", expectedHash = "stale", code = "x" });

            Assert.IsTrue((bool)result["result"]["isError"]);
            StringAssert.StartsWith("stale_hash", TextOf(result));
        }

        [Test]
        public void Write_AppliesTheCode()
        {
            Call("write_module", new { project = "p1", module = "Module1", expectedHash = ModuleApi.HashOf(_store.Code["Module1"]), code = "' via mcp" });

            Assert.AreEqual("' via mcp", _store.Code["Module1"]);
        }

        [Test]
        public void UnknownMethod_IsMinus32601()
            => Assert.AreEqual(-32601, (int)Rpc("nope")["error"]["code"]);

        [Test]
        public void GarbageBody_IsAParseError()
        {
            var response = _router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Body = "not json" });

            Assert.AreEqual(-32700, (int)((JObject)response.Body)["error"]["code"]);
        }

        [Test]
        public void Get_IsNotAllowed()
            => Assert.AreEqual(405, _router.Handle(new ApiRequest { Method = "GET", Path = "/mcp" }).Status);
    }

    [TestFixture]
    [Category("Api")]
    public class ModuleHashTests
    {
        [Test]
        public void LineEndingsAndTrailingNewlines_DoNotChangeTheHash()
            => Assert.AreEqual(ModuleApi.HashOf("a\r\nb"), ModuleApi.HashOf("a\nb\n"));

        [Test]
        public void DifferentCode_HasADifferentHash()
            => Assert.AreNotEqual(ModuleApi.HashOf("a"), ModuleApi.HashOf("b"));
    }
}
