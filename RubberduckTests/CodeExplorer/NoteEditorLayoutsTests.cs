using System.IO;
using System.IO.Abstractions;
using NUnit.Framework;
using Rubberduck.UI.CodeExplorer;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    public class NoteEditorLayoutsTests
    {
        private string _folder;
        private string Path => System.IO.Path.Combine(_folder, "note-editor.json");

        [SetUp]
        public void SetUp()
        {
            _folder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "rd-layouts-" + System.Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_folder);
        }

        [TearDown]
        public void TearDown() => Directory.Delete(_folder, true);

        [Test]
        public void AnUnknownNote_GetsTheDefaults()
        {
            var layout = new NoteEditorLayouts().Get("x|y");
            Assert.AreEqual(NoteEditorView.Split, layout.View);
            Assert.AreEqual(0.5, layout.SourceShare);
            Assert.AreEqual(100, layout.Zoom);
        }

        [Test]
        public void EachNoteKeepsItsOwnLayout_AcrossASaveAndALoad()
        {
            var files = new FileSystem();
            var layouts = new NoteEditorLayouts();
            layouts.Set(NoteEditorLayouts.KeyFor(@"C:\a\Book1.xlsm", "P", "CLAUDE_md"), new NoteEditorLayout { View = NoteEditorView.Preview, Zoom = 150 });
            layouts.Set(NoteEditorLayouts.KeyFor(@"C:\a\Book1.xlsm", "P", "Other_md"), new NoteEditorLayout { View = NoteEditorView.Source, SourceShare = 0.3 });
            layouts.Save(files, Path);

            var loaded = NoteEditorLayouts.Load(files, Path);
            var first = loaded.Get(NoteEditorLayouts.KeyFor(@"C:\A\book1.xlsm", "P", "claude_md"));
            Assert.AreEqual(NoteEditorView.Preview, first.View, "paths and names ignore case");
            Assert.AreEqual(150, first.Zoom);
            Assert.AreEqual(NoteEditorView.Source, loaded.Get(NoteEditorLayouts.KeyFor(@"C:\a\Book1.xlsm", "P", "Other_md")).View);
            Assert.AreEqual(NoteEditorView.Split, loaded.Get(NoteEditorLayouts.KeyFor(@"C:\a\Book2.xlsm", "P", "CLAUDE_md")).View);
        }

        [Test]
        public void Values_AreKeptInRange()
        {
            var layouts = new NoteEditorLayouts();
            layouts.Set("k", new NoteEditorLayout { Zoom = 9000, SourceShare = 7, View = (NoteEditorView)42 });
            var layout = layouts.Get("k");
            Assert.AreEqual(NoteEditorLayout.MaxZoom, layout.Zoom);
            Assert.AreEqual(0.9, layout.SourceShare);
            Assert.AreEqual(NoteEditorView.Split, layout.View);
        }

        [Test]
        public void ACorruptFile_GivesTheDefaults()
        {
            var files = new FileSystem();
            File.WriteAllText(Path, "{ not json");
            Assert.AreEqual(NoteEditorView.Split, NoteEditorLayouts.Load(files, Path).Get("k").View);
        }

        [Test]
        public void AProjectWithoutAFile_IsToldApartByName()
        {
            Assert.AreNotEqual(NoteEditorLayouts.KeyFor(null, "A", "M"), NoteEditorLayouts.KeyFor("", "B", "M"));
        }
    }
}
