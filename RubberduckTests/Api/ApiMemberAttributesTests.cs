using System;
using System.Collections.Generic;
using System.Linq;
using Moq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    public class ApiMemberAttributesTests
    {
        private const string Header = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Todo\"\r\nAttribute VB_PredeclaredId = True\r\n";

        /// <summary>The editor as far as attributes go: replacing the code drops the attributes of the members, the module's own stay.</summary>
        private class Editor : IModuleStore
        {
            public string Raw = Header + "Option Explicit\r\nPublic Function Hello() As String\r\nAttribute Hello.VB_UserMemId = 0\r\n    Hello = \"hi\"\r\nEnd Function\r\n";
            public List<string> Calls = new List<string>();

            private IEnumerable<string> Lines => Raw.Split(new[] { "\r\n" }, StringSplitOptions.None);

            private string Code => string.Join("\r\n", Lines
                .SkipWhile(l => !l.StartsWith("Option") && !l.StartsWith("Public") && !l.StartsWith("Private"))
                .Where(l => !l.StartsWith("Attribute ")));

            private ModuleInfo Info() => new ModuleInfo { Project = "P", Name = "Todo", Kind = "class", Hash = ModuleApi.HashOf(Code) };

            public IReadOnlyList<ProjectInfo> Projects() => new[] { new ProjectInfo { Id = "p1", Name = "P" } };
            public IReadOnlyList<ModuleInfo> Modules(string project) => new[] { Info() };

            public ModuleContent Read(string project, string module, bool includeText)
                => new ModuleContent { Project = "P", Name = "Todo", Kind = "class", Hash = Info().Hash, Code = Code, Text = includeText ? Raw : null };

            public ModuleInfo ReplaceCode(string project, string module, string code, string expectedHash)
            {
                Calls.Add("ReplaceCode");
                Raw = Header + code;
                return Info();
            }

            public ModuleInfo EditLines(string project, string module, int startLine, int endLine, string replacement, string expectedHash)
            {
                Calls.Add("EditLines");
                var lines = Code.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList();
                lines.RemoveRange(startLine - 1, endLine - startLine + 1);
                lines.InsertRange(startLine - 1, replacement.Split(new[] { "\r\n" }, StringSplitOptions.None));
                Raw = Header + string.Join("\r\n", lines);
                return Info();
            }

            public ModuleInfo ReplaceText(string project, string module, string text, string expectedHash)
            {
                Calls.Add("ReplaceText");
                Raw = text;
                return Info();
            }

            public ModuleInfo Create(string project, string module, string kind, string code) => throw new NotSupportedException();
            public void Delete(string project, string module, string expectedHash) => throw new NotSupportedException();
        }

        private Editor _editor;
        private ModuleApi _api;

        [SetUp]
        public void SetUp()
        {
            _editor = new Editor();
            _api = new ModuleApi(_editor, new ApiWriteApproval(_ => true));
        }

        private string Hash => _editor.Read("P", "Todo", false).Hash;

        [Test]
        public void ReplaceCode_PutsBackTheAttributesOfMembersTheNewCodeStillHas()
        {
            var info = _api.ReplaceCode("agent", "P", "Todo", "Option Explicit\r\nPublic Function Hello() As String\r\n    Hello = \"changed\"\r\nEnd Function", Hash);

            StringAssert.Contains("Attribute Hello.VB_UserMemId = 0", _editor.Raw);
            StringAssert.Contains("Hello.VB_UserMemId", info.Notes.Single());
        }

        [Test]
        public void ReplaceCode_DoesNotResurrectTheAttributesOfAMemberThatIsGone()
        {
            var info = _api.ReplaceCode("agent", "P", "Todo", "Option Explicit\r\nPublic Function Other() As String\r\n    Other = \"x\"\r\nEnd Function", Hash);

            StringAssert.DoesNotContain("VB_UserMemId", _editor.Raw);
            Assert.IsNull(info.Notes);
            CollectionAssert.AreEqual(new[] { "ReplaceCode" }, _editor.Calls);
        }

        [Test]
        public void ReplaceCode_WithNoMemberAttributes_WritesOnce()
        {
            _editor.Raw = Header + "Option Explicit\r\nPublic Function Hello() As String\r\n    Hello = \"hi\"\r\nEnd Function";

            var info = _api.ReplaceCode("agent", "P", "Todo", "Option Explicit\r\nPublic Function Hello() As String\r\n    Hello = \"x\"\r\nEnd Function", Hash);

            Assert.IsNull(info.Notes);
            CollectionAssert.AreEqual(new[] { "ReplaceCode" }, _editor.Calls);
        }

        [Test]
        public void EditLines_PutsBackTheAttributesToo()
        {
            var info = _api.EditLines("agent", "P", "Todo", 4, 4, "    Hello = \"edited\"", Hash);

            StringAssert.Contains("Attribute Hello.VB_UserMemId = 0", _editor.Raw);
            StringAssert.Contains("Hello.VB_UserMemId", info.Notes.Single());
        }

        [Test]
        public void Read_WithAttributes_GivesAnAttributesHash_ThatChangesWhenAnAttributeDoes()
        {
            var first = _api.Read("P", "Todo", false, true).AttributesHash;
            Assert.IsNotNull(first);
            Assert.AreEqual(first, _api.Read("P", "Todo", false, true).AttributesHash);

            _editor.Raw = _editor.Raw.Replace("VB_PredeclaredId = True", "VB_PredeclaredId = False");

            Assert.AreNotEqual(first, _api.Read("P", "Todo", false, true).AttributesHash);
            Assert.IsNull(_api.Read("P", "Todo", false, false).AttributesHash);
        }

        [Test]
        public void WriteAttributes_WithAStaleAttributesHash_IsRefusedBeforeAnythingIsWritten()
        {
            var stale = _api.Read("P", "Todo", false, true).AttributesHash;
            _editor.Raw = _editor.Raw.Replace("VB_PredeclaredId = True", "VB_PredeclaredId = False");

            var failure = Assert.Throws<ApiException>(() => _api.WriteAttributes("agent", "P", "Todo", JObject.Parse("{\"module\":{\"VB_Description\":\"x\"}}"), Hash, stale));

            Assert.AreEqual(412, failure.Status);
            Assert.IsEmpty(_editor.Calls);
        }

        [Test]
        public void WriteAttributes_WithTheCurrentAttributesHash_Writes()
        {
            var current = _api.Read("P", "Todo", false, true).AttributesHash;

            _api.WriteAttributes("agent", "P", "Todo", JObject.Parse("{\"module\":{\"VB_Description\":\"x\"}}"), Hash, current);

            CollectionAssert.AreEqual(new[] { "ReplaceText" }, _editor.Calls);
            StringAssert.Contains("VB_Description", _editor.Raw);
        }

        [Test]
        public void ImmediateRun_SaysWhereTheEditorIsAfterwards()
        {
            var text = string.Empty;
            var window = new Mock<IImmediateWindow>();
            window.Setup(w => w.Read()).Returns(() => text);
            window.Setup(w => w.Send(It.IsAny<string>())).Callback<string>(line => text += line + "\r\n");
            var now = TimeSpan.Zero;
            var api = new ImmediateApi(window.Object, new ApiExecuteApproval(_ => true), () => true, t => now += t, () => now, editorMode: () => "running_or_break");

            var run = api.Run("agent", "zApiStop.StopHere", 2000);

            Assert.AreEqual("running_or_break", run.VbeMode);
        }
    }
}
