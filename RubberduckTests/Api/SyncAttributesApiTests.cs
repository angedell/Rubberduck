using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    public class SyncAttributesApiTests
    {
        private const string Header = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Todo\"\r\nAttribute VB_GlobalNameSpace = False\r\n";

        //Annotations that ask for attributes the header does not have.
        private const string OutOfSync = Header
            + "'@PredeclaredId\r\n'@Exposed\r\n'@ModuleDescription \"A list\"\r\nOption Explicit\r\n\r\n"
            + "'@DefaultMember\r\n'@Description(\"Gets one\")\r\nPublic Function Item(ByVal i As Long) As String\r\nEnd Function\r\n\r\n"
            + "'@Enumerator\r\nPublic Function NewEnum() As IUnknown\r\nEnd Function\r\n\r\n"
            + "'@MemberAttribute VB_MemberFlags, \"40\"\r\nPublic Sub Clear()\r\nEnd Sub\r\n";

        private class Store : IModuleStore
        {
            public string Raw;
            public int Writes;

            private string Hash => ModuleApi.HashOf(Raw);

            public IReadOnlyList<ProjectInfo> Projects() => new[] { new ProjectInfo { Id = "p1", Name = "P" } };
            public IReadOnlyList<ModuleInfo> Modules(string project) => new[] { new ModuleInfo { Name = "Todo", Hash = Hash } };

            public ModuleContent Read(string project, string module, bool includeText)
            {
                if (!string.Equals(module, "Todo", StringComparison.OrdinalIgnoreCase)) throw ApiException.NotFound("no module");
                return new ModuleContent { Project = "P", Name = "Todo", Kind = "class", Hash = Hash, Text = includeText ? Raw : null };
            }

            public ModuleInfo ReplaceText(string project, string module, string text, string expectedHash)
            {
                if (expectedHash != Hash) throw ApiException.PreconditionFailed("stale");
                Writes++;
                Raw = text;
                return new ModuleInfo { Name = "Todo", Hash = Hash };
            }

            public ModuleInfo ReplaceCode(string project, string module, string code, string expectedHash) => throw new NotSupportedException();
            public ModuleInfo EditLines(string project, string module, int startLine, int endLine, string replacement, string expectedHash) => throw new NotSupportedException();
            public ModuleInfo Create(string project, string module, string kind, string code) => throw new NotSupportedException();
            public void Delete(string project, string module, string expectedHash) => throw new NotSupportedException();
        }

        private Store _store;
        private int _asked;
        private bool _approve;
        private bool _design;

        [SetUp]
        public void SetUp()
        {
            _store = new Store { Raw = OutOfSync };
            _asked = 0;
            _approve = true;
            _design = true;
        }

        private SyncAttributesApi Api() => new SyncAttributesApi(_store, new ApiWriteApproval(_ => { _asked++; return _approve; }), () => _design);

        private static string Describe(SyncAttributesReport report)
            => string.Join(";", report.Changes.Select(c => $"{c.Scope}.{c.Attribute}={c.To}"));

        [Test]
        public void Preview_lists_what_the_annotations_ask_for_and_changes_nothing()
        {
            var report = Api().Preview("p1", "Todo");

            Assert.That(report.InSync, Is.False);
            Assert.That(report.Applied, Is.False);
            Assert.That(Describe(report), Is.EqualTo(
                "module.VB_PredeclaredId=True;module.VB_Exposed=True;module.VB_Description=A list;"
                + "Item.VB_UserMemId=0;Item.VB_Description=Gets one;NewEnum.VB_UserMemId=-4;Clear.VB_MemberFlags=40"));
            Assert.That(report.Changes.First(c => c.Attribute == "VB_Exposed").Because, Is.EqualTo("@Exposed"));
            Assert.That(report.Changes.First(c => c.Scope == "NewEnum").Because, Is.EqualTo("@Enumerator"));
            Assert.That(_store.Raw, Is.EqualTo(OutOfSync));
            Assert.That(_asked, Is.EqualTo(0));
        }

        [Test]
        public void Apply_writes_the_attributes_and_afterwards_all_is_in_sync()
        {
            var api = Api();
            var hash = api.Preview("p1", "Todo").Hash;

            var applied = api.Apply("agent", "p1", "Todo", hash);

            Assert.That(applied.Applied, Is.True);
            Assert.That(applied.NewHash, Is.Not.Null);
            Assert.That(_store.Writes, Is.EqualTo(1));
            var attributes = ModuleAttributes.Parse(_store.Raw);
            Assert.That(attributes["module"]["VB_PredeclaredId"].Value<bool>(), Is.True);
            Assert.That(attributes["module"]["VB_Description"].Value<string>(), Is.EqualTo("A list"));
            Assert.That(attributes["members"]["Item"]["VB_UserMemId"].Value<int>(), Is.EqualTo(0));
            Assert.That(attributes["members"]["Item"]["VB_Description"].Value<string>(), Is.EqualTo("Gets one"));
            Assert.That(attributes["members"]["NewEnum"]["VB_UserMemId"].Value<int>(), Is.EqualTo(-4));
            Assert.That(_store.Raw, Does.Contain("Public Function Item(ByVal i As Long) As String\r\nAttribute Item."));
            Assert.That(_store.Raw, Does.Contain("'@DefaultMember"), "annotations are not touched");
            Assert.That(api.Preview("p1", "Todo").InSync, Is.True);
        }

        [Test]
        public void Apply_without_changes_writes_nothing_and_asks_nobody()
        {
            var api = Api();
            api.Apply("agent", "p1", "Todo", api.Preview("p1", "Todo").Hash);
            _asked = 0;
            _store.Writes = 0;

            var again = api.Apply("agent", "p1", "Todo", ModuleApi.HashOf(_store.Raw));

            Assert.That(again.InSync, Is.True);
            Assert.That(again.Applied, Is.False);
            Assert.That(_store.Writes, Is.EqualTo(0));
            Assert.That(_asked, Is.EqualTo(0));
        }

        [Test]
        public void A_changed_attribute_is_reported_with_its_old_value()
        {
            _store.Raw = Header.Replace("Attribute VB_GlobalNameSpace = False\r\n", "Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Exposed = False\r\n")
                         + "'@Exposed\r\nOption Explicit\r\n";

            var change = Api().Preview("p1", "Todo").Changes.Single();

            Assert.That(change.Attribute, Is.EqualTo("VB_Exposed"));
            Assert.That(change.From.Value<bool>(), Is.False);
            Assert.That(change.To.Value<bool>(), Is.True);
        }

        [Test]
        public void The_name_attribute_follows_the_module()
        {
            _store.Raw = Header.Replace("\"Todo\"", "\"Module1\"") + "Option Explicit\r\n";

            var change = Api().Preview("p1", "Todo").Changes.Single();

            Assert.That(change.Attribute, Is.EqualTo("VB_Name"));
            Assert.That(change.To.Value<string>(), Is.EqualTo("Todo"));
        }

        [Test]
        public void Apply_needs_a_hash_and_refuses_a_stale_one()
        {
            var api = Api();

            Assert.That(Assert.Throws<ApiException>(() => api.Apply("agent", "p1", "Todo", null)).Status, Is.EqualTo(428));
            Assert.That(Assert.Throws<ApiException>(() => api.Apply("agent", "p1", "Todo", "0000")).Status, Is.EqualTo(412));
            Assert.That(_store.Writes, Is.EqualTo(0));
        }

        [Test]
        public void Apply_is_refused_outside_design_mode_but_preview_is_not()
        {
            _design = false;
            var api = Api();

            var failure = Assert.Throws<ApiException>(() => api.Apply("agent", "p1", "Todo", ModuleApi.HashOf(OutOfSync)));

            Assert.That(failure.Status, Is.EqualTo(409));
            Assert.That(failure.Code, Is.EqualTo("vbe_not_in_design_mode"));
            Assert.That(api.Preview("p1", "Todo").InSync, Is.False);
        }

        [Test]
        public void Apply_needs_the_users_approval()
        {
            _approve = false;
            var api = Api();

            var failure = Assert.Throws<ApiException>(() => api.Apply("agent", "p1", "Todo", ModuleApi.HashOf(OutOfSync)));

            Assert.That(failure.Code, Is.EqualTo("write_denied"));
            Assert.That(_store.Writes, Is.EqualTo(0));
        }

        private ApiRouter Router()
            => new ApiRouter(new ModuleApi(_store, new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api, syncAttributes: Api());

        [Test]
        public void The_route_previews_on_get_and_on_a_dry_run_and_applies_on_post()
        {
            var router = Router();
            var path = "/api/v1/projects/p1/modules/Todo/sync-attributes";

            var get = router.Handle(new ApiRequest { Method = "GET", Path = path, Client = "c" });
            var dry = router.Handle(new ApiRequest { Method = "POST", Path = path, Body = "{\"dryRun\":true}", Client = "c" });
            Assert.That(get.Status, Is.EqualTo(200));
            Assert.That(((SyncAttributesReport)dry.Body).Applied, Is.False);
            Assert.That(_store.Writes, Is.EqualTo(0));

            var noHash = router.Handle(new ApiRequest { Method = "POST", Path = path, Body = "{}", Client = "c" });
            Assert.That(noHash.Status, Is.EqualTo(428));

            var hash = ((SyncAttributesReport)get.Body).Hash;
            var post = router.Handle(new ApiRequest { Method = "POST", Path = path, IfMatch = "\"" + hash + "\"", Client = "c" });
            Assert.That(post.Status, Is.EqualTo(200));
            Assert.That(((SyncAttributesReport)post.Body).Applied, Is.True);

            _design = false;
            var stopped = router.Handle(new ApiRequest { Method = "POST", Path = path, IfMatch = hash, Client = "c" });
            Assert.That(stopped.Status, Is.EqualTo(409));
        }

        [Test]
        public void The_mcp_tool_previews_by_default_and_applies_on_request()
        {
            var router = Router();

            JObject Call(object arguments)
            {
                var body = JObject.FromObject(new { jsonrpc = "2.0", id = 1, method = "tools/call", @params = new { name = "sync_attributes", arguments } }).ToString();
                var response = router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Body = body, Client = "c" });
                return JObject.FromObject(response.Body)["result"] as JObject;
            }

            var preview = Call(new { project = "p1", module = "Todo" });
            var report = JObject.Parse((string)preview["content"][0]["text"]);
            Assert.That((bool)report.GetValue("inSync", StringComparison.OrdinalIgnoreCase), Is.False);
            Assert.That(_store.Writes, Is.EqualTo(0));

            var applied = Call(new { project = "p1", module = "Todo", apply = true, expectedHash = (string)report.GetValue("hash", StringComparison.OrdinalIgnoreCase) });
            Assert.That((bool?)applied["isError"] ?? false, Is.False);
            Assert.That(_store.Writes, Is.EqualTo(1));
        }
    }
}
