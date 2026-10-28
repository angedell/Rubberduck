using NUnit.Framework;
using Rubberduck.Navigation.CodeExplorer;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    public class NoteModuleTests
    {
        [Test]
        public void IsNote_NeedsTheMarker_AndNothingButComments()
        {
            Assert.IsTrue(NoteModule.IsNote("'@Notes\r\n' # Title"));
            Assert.IsTrue(NoteModule.IsNote("'@Notes(\"CLAUDE.md\")\r\n'@Folder(\"Api\")\r\n\r\nOption Explicit\r\n' text"));
            Assert.IsTrue(NoteModule.IsNote("  '@notes"), "case and indentation do not matter");
            Assert.IsFalse(NoteModule.IsNote("' # Title"), "no marker");
            Assert.IsFalse(NoteModule.IsNote("'@Notes\r\nSub A()\r\nEnd Sub"), "code in it: an ordinary module");
            Assert.IsFalse(NoteModule.IsNote("' mentions '@Notes in a sentence"));
            Assert.IsFalse(NoteModule.IsNote(null));
        }

        [Test]
        public void HasMarker_IsTheQuickLook()
        {
            Assert.IsTrue(NoteModule.HasMarker("Sub A()\r\n'@Notes"));
            Assert.IsFalse(NoteModule.HasMarker("Sub A()\r\nEnd Sub"));
        }

        [Test]
        public void DisplayName_IsTheAnnotationsName_OrTheModuleNameWithMd()
        {
            Assert.AreEqual("CLAUDE.md", NoteModule.DisplayName("Whatever", "'@Notes(\"CLAUDE.md\")\r\n' x"));
            Assert.AreEqual("Readme.md", NoteModule.DisplayName("Readme", "'@Notes\r\n' x"));
            Assert.AreEqual("Readme.md", NoteModule.DisplayName("Readme", "'@Notes()\r\n' x"));
            Assert.AreEqual("CLAUDE.md", NoteModule.DisplayNameFromArgument("M", "\"CLAUDE.md\""));
            Assert.AreEqual("M.md", NoteModule.DisplayNameFromArgument("M", null));
        }

        [TestCase("CLAUDE.md", "CLAUDE_md")]
        [TestCase("claude.MD", "claude_md")]
        [TestCase("CLAUDE", null)]
        [TestCase("CLAUDE_md", null)]
        [TestCase(".md", null)]
        [TestCase("a b.md", null)]
        [TestCase("1a.md", null)]
        [TestCase("abcdefghijklmnopqrstuvwxyzab.md", "abcdefghijklmnopqrstuvwxyzab_md")]
        [TestCase("abcdefghijklmnopqrstuvwxyzabc.md", null)]
        public void ModuleNameFor_AcceptsOnlyWhatTheVbeAccepts(string name, string expected)
        {
            Assert.AreEqual(expected, NoteModule.ModuleNameFor(name));
        }

        [Test]
        public void ToModuleCode_MarksEveryLine()
        {
            Assert.AreEqual("' # Title\r\n'\r\n' text\r\n'     indented", NoteModule.ToModuleCode("# Title\r\n\r\ntext\n    indented"));
        }

        [Test]
        public void ToMarkdown_TakesTheMarksOff_AndKeepsUnmarkedLines()
        {
            var code = "Attribute VB_Name = \"X_md\"\r\n' # Title\r\n'\r\n'no space\r\n  ' indented mark\r\nstray line\r\n'\r\n'\r\n";
            Assert.AreEqual("# Title\r\n\r\nno space\r\nindented mark\r\nstray line", NoteModule.ToMarkdown(code));
        }

        [Test]
        public void RoundTrip_KeepsTheMarkdown()
        {
            var markdown = "# Title\r\n\r\n- one\r\n- two\r\n\r\n```vba\r\nDebug.Print \"'quoted'\"\r\n```";
            Assert.AreEqual(markdown, NoteModule.ToMarkdown(NoteModule.ToModuleCode(markdown)));
        }

        [Test]
        public void Annotations_StayOutOfTheMarkdown_AndComeBackFirst()
        {
            var code = "'@Notes(\"CLAUDE.md\")\r\n'@Folder(\"Api.Core\")\r\n' # Title\r\n' @not an annotation";
            Assert.AreEqual("# Title\r\n@not an annotation", NoteModule.ToMarkdown(code));
            CollectionAssert.AreEqual(new[] { "'@Notes(\"CLAUDE.md\")", "'@Folder(\"Api.Core\")" }, NoteModule.Annotations(code));
            Assert.AreEqual(code, NoteModule.ToModuleCode(NoteModule.ToMarkdown(code), NoteModule.Annotations(code)));
            Assert.IsTrue(NoteModule.IsNote(NoteModule.ToModuleCode("x", NoteModule.Annotations(code))), "still a note after a save");
        }

        [Test]
        public void Marker_IsSetOnce_AndFirst()
        {
            CollectionAssert.AreEqual(new[] { "'@Notes(\"New.md\")", "'@Folder(\"A\")" }, NoteModule.WithMarker(new[] { "'@Notes(\"Old.md\")", "'@Folder(\"A\")" }, "New.md"));
        }

        [Test]
        public void Folder_IsReadAndSet()
        {
            Assert.AreEqual("Api.Core", NoteModule.FolderOf(new[] { "'@IgnoreModule", "'@Folder(\"Api.Core\")" }));
            Assert.IsNull(NoteModule.FolderOf(new[] { "'@IgnoreModule" }));
            CollectionAssert.AreEqual(new[] { "'@IgnoreModule", "'@Folder(\"X\")" }, NoteModule.WithFolder(new[] { "'@Folder(\"Old\")", "'@IgnoreModule" }, "X"));
            CollectionAssert.AreEqual(new[] { "'@IgnoreModule" }, NoteModule.WithFolder(new[] { "'@Folder(\"Old\")", "'@IgnoreModule" }, ""));
            Assert.IsFalse(NoteModule.IsValidFolder("a\")b"));
        }

        [Test]
        public void FirstUnsupportedCharacter_FindsWhatTheCodePageLoses()
        {
            Assert.IsNull(NoteModule.FirstUnsupportedCharacter("plain text\r\n\twith tab"));
            Assert.AreEqual('☃', NoteModule.FirstUnsupportedCharacter("snow ☃ man"));
            Assert.IsNotNull(NoteModule.FirstUnsupportedCharacter("rocket \U0001F680"));
        }

        [Test]
        public void FirstLineTooLong_ReportsTheLine()
        {
            Assert.AreEqual(0, NoteModule.FirstLineTooLong("short\r\nlines"));
            Assert.AreEqual(2, NoteModule.FirstLineTooLong("short\r\n" + new string('x', NoteModule.MaxLineLength)));
        }
    }
}
