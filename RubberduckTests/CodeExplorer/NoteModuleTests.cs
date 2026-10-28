using System.Linq;
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

        [Test]
        public void NewModuleName_IsAValidVbaName_AndNeverOneTaken()
        {
            var taken = new[] { "Module1", "Note_00000000" };
            var names = Enumerable.Range(0, 200).Select(_ => NoteModule.NewModuleName(taken)).ToList();

            Assert.IsTrue(names.All(n => System.Text.RegularExpressions.Regex.IsMatch(n, "^[A-Za-z][A-Za-z0-9_]{0,30}$")));
            Assert.IsFalse(names.Intersect(taken, System.StringComparer.OrdinalIgnoreCase).Any());
            Assert.Greater(names.Distinct().Count(), 190, "random");
        }

        [TestCase(null, "CLAUDE.md", "CLAUDE.md")]
        [TestCase("", "CLAUDE.md", "CLAUDE.md")]
        [TestCase("Api", "CLAUDE.md", "Api/CLAUDE.md")]
        [TestCase("Api.Core", "CLAUDE.md", "Api/Core/CLAUDE.md")]
        public void PathOf_AndSplitPath_AreInverse(string folder, string name, string path)
        {
            Assert.AreEqual(path, NoteModule.PathOf(folder, name));

            NoteModule.SplitPath(path, out var splitFolder, out var splitName);
            Assert.AreEqual(name, splitName);
            Assert.IsTrue(NoteModule.SameFolder(folder, splitFolder));
        }

        [Test]
        public void SplitPath_TakesEitherSlash_AndTellsNoFolderFromTheRoot()
        {
            NoteModule.SplitPath(@"Api\Core/CLAUDE.md", out var folder, out var name);
            Assert.AreEqual("Api.Core", folder);
            Assert.AreEqual("CLAUDE.md", name);

            NoteModule.SplitPath("CLAUDE.md", out folder, out _);
            Assert.IsNull(folder, "no folder was named");

            NoteModule.SplitPath("/CLAUDE.md", out folder, out _);
            Assert.AreEqual(string.Empty, folder, "the root was named");
        }

        [TestCase("CLAUDE.md", true)]
        [TestCase("my notes \u00E9.md", true)]
        [TestCase("1st.md", true)]
        [TestCase("claude.MD", true)]
        [TestCase("CLAUDE", false)]
        [TestCase(".md", false)]
        [TestCase("a/b.md", false)]
        [TestCase("a\\b.md", false)]
        [TestCase("a\"b.md", false)]
        [TestCase("a(b).md", false)]
        [TestCase("a:b.md", false)]
        [TestCase(null, false)]
        public void IsValidName_IsWhatAMarkdownFileCanBeCalled(string name, bool expected)
        {
            Assert.AreEqual(expected, NoteModule.IsValidName(name));
        }

        [Test]
        public void ANameWithNonAscii_IsStoredEncoded_AndShownDecoded()
        {
            var annotations = NoteModule.WithMarker(null, "caf\u00E9.md");
            Assert.AreEqual("'@Notes(\"caf&#xE9;.md\")", annotations[0]);
            Assert.AreEqual("caf\u00E9.md", NoteModule.DisplayName("Note_1", annotations[0] + "\r\n' x"));
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
        public void NonAscii_IsStoredAsReferences_AndComesBack()
        {
            var markdown = "rocket \U0001F680 snowman ☃ café “quoted”\r\n\ttab";
            var code = NoteModule.ToModuleCode(markdown);

            Assert.IsTrue(code.All(c => c < 127 || c == '\r' || c == '\n'), "only ASCII in the module");
            StringAssert.Contains("&#x1F680;", code);
            StringAssert.Contains("&#xE9;", code);
            Assert.AreEqual(markdown, NoteModule.ToMarkdown(code));
        }

        [Test]
        public void Decode_LeavesWhatIsNotAValidReference()
        {
            Assert.AreEqual("&#x110000; &#xD800; &#1; &#; &#xZZ; é é", NoteModule.Decode("&#x110000; &#xD800; &#1; &#; &#xZZ; &#233; &#xe9;"));
        }

        [Test]
        public void FirstLineTooLong_ReportsTheLine()
        {
            Assert.AreEqual(0, NoteModule.FirstLineTooLong("short\r\nlines"));
            Assert.AreEqual(2, NoteModule.FirstLineTooLong("short\r\n" + new string('x', NoteModule.MaxLineLength)));
            Assert.AreEqual(1, NoteModule.FirstLineTooLong(string.Concat(System.Linq.Enumerable.Repeat("\U0001F680", 150))), "counted once encoded");
        }
    }
}
