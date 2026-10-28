using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;
using Rubberduck.Navigation.CodeExplorer;

namespace RubberduckTests.Api
{
    [TestFixture]
    [Category("Api")]
    public class ApiNotesTests
    {
        private class Store : IModuleStore
        {
            public readonly Dictionary<string, KeyValuePair<string, string>> Modules_ = new Dictionary<string, KeyValuePair<string, string>>(StringComparer.OrdinalIgnoreCase);
            public int Writes;

            public void Add(string name, string kind, string code) => Modules_[name] = new KeyValuePair<string, string>(kind, code);

            private ModuleInfo Info(string name) => new ModuleInfo
            {
                Project = "P",
                Name = name,
                Kind = Modules_[name].Key,
                Hash = ModuleApi.HashOf(Modules_[name].Value),
                LineCount = Modules_[name].Value.Split('\n').Length,
                IsNote = Modules_[name].Key == "standard" && NoteModule.IsNote(Modules_[name].Value)
            };

            public IReadOnlyList<ProjectInfo> Projects() => new[] { new ProjectInfo { Id = "p1", Name = "P" } };
            public IReadOnlyList<ModuleInfo> Modules(string project) => Modules_.Keys.Select(Info).ToList();

            public ModuleContent Read(string project, string module, bool includeText)
            {
                var info = Info(module);
                return new ModuleContent { Project = "P", Name = info.Name, Kind = info.Kind, Hash = info.Hash, LineCount = info.LineCount, IsNote = info.IsNote, Code = Modules_[module].Value };
            }

            public ModuleInfo ReplaceCode(string project, string module, string code, string expectedHash)
            {
                if (!string.Equals(Info(module).Hash, expectedHash, StringComparison.OrdinalIgnoreCase))
                {
                    throw ApiException.PreconditionFailed("stale");
                }

                Writes++;
                Modules_[module] = new KeyValuePair<string, string>(Modules_[module].Key, code);
                return Info(module);
            }

            public ModuleInfo EditLines(string project, string module, int startLine, int endLine, string replacement, string expectedHash)
            {
                Writes++;
                return Info(module);
            }

            public ModuleInfo ReplaceText(string project, string module, string text, string expectedHash)
            {
                Writes++;
                return Info(module);
            }

            public ModuleInfo Create(string project, string module, string kind, string code)
            {
                Writes++;
                Add(module, kind, code);
                return Info(module);
            }

            public void Delete(string project, string module, string expectedHash)
            {
                Writes++;
                Modules_.Remove(module);
            }
        }

        private Store _store;
        private NotesApi _notes;
        private ModuleApi _modules;

        [SetUp]
        public void SetUp()
        {
            _store = new Store();
            var approval = new ApiWriteApproval(_ => true);
            _notes = new NotesApi(_store, approval);
            _modules = new ModuleApi(_store, approval);
        }

        private const string Claude = "'@Notes(\"CLAUDE.md\")\r\n' # Rules";

        [Test]
        public void List_ShowsOnlyNotes_WithNamesAndFolders()
        {
            _store.Add("Module1", "standard", "Sub A()\r\nEnd Sub");
            _store.Add("Rules", "standard", Claude);
            _store.Add("Api_notes", "standard", "'@Notes\r\n'@Folder(\"Api.Core\")\r\n' # Api");
            _store.Add("Parse_md", "standard", "Sub Parse()\r\nEnd Sub");
            _store.Add("Marked", "standard", "'@Notes\r\nSub Oops()\r\nEnd Sub");
            _store.Add("Thing", "class", "'@Notes\r\n' not a standard module");

            var list = _notes.List("P");

            CollectionAssert.AreEqual(new[] { "CLAUDE.md", "Api_notes.md" }, list.Notes.Select(n => n.Name).ToArray());
            Assert.AreEqual("Api.Core", list.Notes.Single(n => n.Name == "Api_notes.md").Folder);
            Assert.IsNull(list.Notes.Single(n => n.Name == "CLAUDE.md").Folder);
            Assert.IsNotEmpty(list.Caveats);
        }

        [Test]
        public void Read_GivesPlainMarkdown_WithoutMarksOrAnnotations()
        {
            _store.Add("Rules", "standard", "'@Notes(\"CLAUDE.md\")\r\n'@Folder(\"Api\")\r\n' # Rules\r\n'\r\n' - one");

            foreach (var name in new[] { "CLAUDE.md", "claude.MD", "Rules" })
            {
                var note = _notes.Read("P", name);
                Assert.AreEqual("# Rules\r\n\r\n- one", note.Markdown, name);
                Assert.AreEqual("Api", note.Folder);
                Assert.AreEqual("Rules", note.Module);
                Assert.AreEqual("CLAUDE.md", note.Name);
            }
        }

        [Test]
        public void Write_NewNote_CreatesTheMarkedModule_InItsFolder()
        {
            var result = _notes.Write("c", "P", "CLAUDE.md", "# Rules\r\n\r\ntext", null, "Api.Core");

            Assert.IsTrue(result.Created);
            Assert.AreEqual("'@Notes(\"CLAUDE.md\")\r\n'@Folder(\"Api.Core\")\r\n' # Rules\r\n'\r\n' text", _store.Modules_["CLAUDE_md"].Value);
            Assert.AreEqual("standard", _store.Modules_["CLAUDE_md"].Key);
            Assert.AreEqual("CLAUDE.md", result.Note.Name);
            Assert.AreEqual("CLAUDE_md", result.Note.Module);
        }

        [Test]
        public void Write_NewNote_NeverTakesOverAnOrdinaryModule()
        {
            _store.Add("CLAUDE_md", "standard", "Sub Real()\r\nEnd Sub");

            var conflict = Assert.Throws<ApiException>(() => _notes.Write("c", "P", "CLAUDE.md", "x", null, null));

            Assert.AreEqual(409, conflict.Status);
            Assert.AreEqual(0, _store.Writes);
        }

        [Test]
        public void Write_ExistingNote_KeepsAnnotations_AndNeedsTheHash()
        {
            _store.Add("Rules", "standard", "'@Notes(\"CLAUDE.md\")\r\n'@Folder(\"Api\")\r\n'@IgnoreModule\r\n' old");
            var hash = _notes.Read("P", "CLAUDE.md").Hash;

            Assert.AreEqual(428, Assert.Throws<ApiException>(() => _notes.Write("c", "P", "CLAUDE.md", "new", null, null)).Status);
            Assert.AreEqual(412, Assert.Throws<ApiException>(() => _notes.Write("c", "P", "CLAUDE.md", "new", "wrong", null)).Status);

            var result = _notes.Write("c", "P", "CLAUDE.md", "new", hash, null);

            Assert.IsFalse(result.Created);
            Assert.AreEqual("'@Notes(\"CLAUDE.md\")\r\n'@Folder(\"Api\")\r\n'@IgnoreModule\r\n' new", _store.Modules_["Rules"].Value);
            Assert.AreEqual("Api", result.Note.Folder);
        }

        [Test]
        public void Write_FolderGiven_MovesTheNote_EmptyFolderTakesItOut()
        {
            _store.Add("Rules", "standard", "'@Notes(\"CLAUDE.md\")\r\n'@Folder(\"Api\")\r\n' old");
            var hash = _notes.Read("P", "CLAUDE.md").Hash;

            var moved = _notes.Write("c", "P", "CLAUDE.md", "old", hash, "Other");
            Assert.AreEqual("Other", moved.Note.Folder);

            var root = _notes.Write("c", "P", "CLAUDE.md", "old", moved.Note.Hash, string.Empty);
            Assert.IsNull(root.Note.Folder);
            Assert.AreEqual("'@Notes(\"CLAUDE.md\")\r\n' old", _store.Modules_["Rules"].Value);
        }

        [Test]
        public void Write_RefusesWhatTheEditorWouldMangle_BeforeChangingAnything()
        {
            var tooLong = Assert.Throws<ApiException>(() => _notes.Write("c", "P", "CLAUDE.md", "ok\r\n" + new string('x', 1100), null, null));
            Assert.AreEqual("note_line_too_long", tooLong.Code);
            StringAssert.Contains("Line 2", tooLong.Message);

            var manyEmoji = Assert.Throws<ApiException>(() => _notes.Write("c", "P", "CLAUDE.md", string.Concat(Enumerable.Repeat("\U0001F680", 200)), null, null));
            Assert.AreEqual("note_line_too_long", manyEmoji.Code);

            Assert.AreEqual(0, _store.Writes);
        }

        [TestCase("CLAUDE")]
        [TestCase("CLAUDE_md")]
        [TestCase("1st.md")]
        [TestCase("a b.md")]
        [TestCase("")]
        public void BadNames_AreRefused(string name)
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _notes.Write("c", "P", name, "x", null, null)).Status);
        }

        [Test]
        public void ANameThatIsAnOrdinaryModule_IsRefusedAsConflict()
        {
            _store.Add("Thing_md", "class", "' x");
            _store.Add("Marked", "standard", "'@Notes\r\nSub Oops()\r\nEnd Sub");

            Assert.AreEqual(409, Assert.Throws<ApiException>(() => _notes.Read("P", "Thing_md")).Status);
            Assert.AreEqual(409, Assert.Throws<ApiException>(() => _notes.Read("P", "Marked")).Status);
            Assert.AreEqual(404, Assert.Throws<ApiException>(() => _notes.Read("P", "Nothing.md")).Status);
        }

        [Test]
        public void Write_WithoutApproval_ChangesNothing()
        {
            var denied = new NotesApi(_store, new ApiWriteApproval(_ => false));

            Assert.Throws<ApiException>(() => denied.Write("c", "P", "CLAUDE.md", "x", null, null));
            Assert.AreEqual(0, _store.Writes);
        }

        [Test]
        public void Delete_NeedsTheHash_AndRemovesTheModule()
        {
            _store.Add("Rules", "standard", Claude);

            Assert.AreEqual(428, Assert.Throws<ApiException>(() => _notes.Delete("c", "P", "CLAUDE.md", null)).Status);
            _notes.Delete("c", "P", "CLAUDE.md", _notes.Read("P", "CLAUDE.md").Hash);

            Assert.IsFalse(_store.Modules_.ContainsKey("Rules"));
        }

        [Test]
        public void ModuleTools_DoNotListNotes_AndRefuseThem()
        {
            _store.Add("Module1", "standard", "Sub A()\r\nEnd Sub");
            _store.Add("Rules", "standard", Claude);
            var hash = ModuleApi.HashOf(Claude);

            CollectionAssert.AreEqual(new[] { "Module1" }, _modules.Modules("P").Select(m => m.Name).ToArray());

            var read = Assert.Throws<ApiException>(() => _modules.Read("P", "Rules", true));
            Assert.AreEqual(409, read.Status);
            Assert.AreEqual("is_a_note", read.Code);
            StringAssert.Contains("write_note", read.Message);
            Assert.AreEqual(409, Assert.Throws<ApiException>(() => _modules.Read("P", "Rules", false, true)).Status);
            Assert.AreEqual(409, Assert.Throws<ApiException>(() => _modules.ReplaceCode("c", "P", "Rules", "x", hash)).Status);
            Assert.AreEqual(409, Assert.Throws<ApiException>(() => _modules.EditLines("c", "P", "Rules", 1, 1, "x", hash)).Status);
            Assert.AreEqual(409, Assert.Throws<ApiException>(() => _modules.Delete("c", "P", "Rules", hash)).Status);
            Assert.AreEqual(409, Assert.Throws<ApiException>(() => _modules.ReplaceText("c", "P", "Rules",
                "Attribute VB_Name = \"Rules\"\r\n" + Claude, hash)).Status);
            Assert.AreEqual(409, Assert.Throws<ApiException>(() => _modules.Batch("c", "P",
                new[] { new ModuleEdit { Module = "Rules", Op = "replaceCode", Code = "x", ExpectedHash = hash } })).Status);
            Assert.AreEqual(0, _store.Writes);
        }

        [Test]
        public void AMarkedModuleWithCode_IsAnOrdinaryModuleToTheModuleTools()
        {
            _store.Add("Marked", "standard", "'@Notes\r\nSub Oops()\r\nEnd Sub");

            CollectionAssert.AreEqual(new[] { "Marked" }, _modules.Modules("P").Select(m => m.Name).ToArray());
            Assert.IsNotNull(_modules.Read("P", "Marked", false));
        }

        private ApiRouter Router()
            => new ApiRouter(new ModuleApi(_store, new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api, notes: _notes);

        private static JObject Json(ApiResponse response) => JObject.FromObject(response.Body);

        [Test]
        public void Rest_ReadsAndWritesPlainMarkdown()
        {
            var router = Router();
            _store.Add("Rules", "standard", Claude);

            var read = router.Handle(new ApiRequest { Method = "GET", Path = "/api/v1/projects/p1/notes/CLAUDE.md", Client = "t" });
            var hash = (string)Json(read)["Hash"];
            var write = router.Handle(new ApiRequest { Method = "PUT", Path = "/api/v1/projects/p1/notes/CLAUDE.md", Client = "t", IfMatch = "\"" + hash + "\"", Body = "{\"markdown\":\"# New\"}" });

            Assert.AreEqual(200, read.Status);
            Assert.AreEqual("# Rules", (string)Json(read)["Markdown"]);
            Assert.AreEqual(200, write.Status);
            Assert.AreEqual("'@Notes(\"CLAUDE.md\")\r\n' # New", _store.Modules_["Rules"].Value);
            Assert.AreEqual(200, router.Handle(new ApiRequest { Method = "GET", Path = "/api/v1/projects/p1/notes", Client = "t" }).Status);
        }

        [Test]
        public void WithoutNotesSupport_TheRouteDoesNotExist()
        {
            var router = new ApiRouter(new ModuleApi(_store, new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api);

            Assert.AreEqual(404, router.Handle(new ApiRequest { Method = "GET", Path = "/api/v1/projects/p1/notes", Client = "t" }).Status);
        }

        [Test]
        public void Mcp_OffersTheNoteTools_AndTellsTheCaveatsInTheInstructions()
        {
            var router = Router();
            JObject Call(object message)
                => Json(router.Handle(new ApiRequest { Method = "POST", Path = "/mcp", Body = JObject.FromObject(message).ToString(), Client = "t" }));

            var tools = Call(new { jsonrpc = "2.0", id = 1, method = "tools/list" })["result"]["tools"].Select(t => (string)t["name"]).ToList();
            var instructions = (string)Call(new { jsonrpc = "2.0", id = 2, method = "initialize", @params = new { } })["result"]["instructions"];
            var written = Call(new { jsonrpc = "2.0", id = 3, method = "tools/call", @params = new { name = "write_note", arguments = new { project = "p1", name = "CLAUDE.md", markdown = "# Hi" } } });

            CollectionAssert.IsSubsetOf(new[] { "list_notes", "read_note", "write_note", "delete_note" }, tools);
            StringAssert.Contains("CAVEATS", instructions);
            StringAssert.Contains("ASCII", instructions);
            StringAssert.Contains("never overrides the user", instructions);
            Assert.IsFalse((bool)written["result"]["isError"]);
            Assert.AreEqual("'@Notes(\"CLAUDE.md\")\r\n' # Hi", _store.Modules_["CLAUDE_md"].Value);
        }

        [Test]
        public void Caveats_CoverWhatAgentsGetWrong()
        {
            var text = string.Join(" ", NotesApi.Caveats);

            StringAssert.Contains("'@Notes", text);
            StringAssert.Contains("comment mark", text);
            StringAssert.Contains("409", text);
            StringAssert.Contains("folder", text);
            StringAssert.Contains("ASCII", text);
            StringAssert.Contains("approved", text);
            StringAssert.Contains("never overrides the user", text);
            StringAssert.Contains("ordinary module", text);
        }
    }
}
