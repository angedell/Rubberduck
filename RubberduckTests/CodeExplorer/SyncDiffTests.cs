using System.Linq;
using NUnit.Framework;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    [Category("GitSync")]
    public class SyncDiffTests
    {
        private static string Text(params string[] lines) => string.Join("\r\n", lines);

        [Test]
        public void SameLines_HaveNoHunks()
        {
            var diff = new SyncDiff(Text("A", "B"), Text("A", "B"));

            Assert.AreEqual(0, diff.Hunks.Count);
        }

        [Test]
        public void IndentationAndTrailingSpace_AreNotADifference()
        {
            var diff = new SyncDiff(Text("Sub A()", "    x = 1  ", "End Sub"), Text("Sub A()", "x = 1", "End Sub"));

            Assert.AreEqual(0, diff.Hunks.Count);
        }

        [Test]
        public void ChangedLine_IsOneHunk()
        {
            var diff = new SyncDiff(Text("A", "B", "C"), Text("A", "X", "C"));

            var hunk = diff.Hunks.Single();
            Assert.AreEqual((1, 1, 1, 1), (hunk.ModuleStart, hunk.ModuleCount, hunk.FileStart, hunk.FileCount));
        }

        [Test]
        public void AddedAndRemovedLines_HaveAnEmptySide()
        {
            var diff = new SyncDiff(Text("A", "B"), Text("A", "N", "B", "C"));

            Assert.AreEqual(2, diff.Hunks.Count);
            Assert.AreEqual(0, diff.Hunks[0].ModuleCount);
            Assert.AreEqual(1, diff.Hunks[0].FileCount);
            Assert.AreEqual(0, diff.Hunks[1].ModuleCount);
            Assert.AreEqual(1, diff.Hunks[1].FileCount);
        }

        [Test]
        public void Merge_TakesEachHunkFromItsChosenSide()
        {
            var diff = new SyncDiff(Text("A", "m1", "B", "m2", "C"), Text("A", "f1", "B", "f2", "C"));
            Assert.AreEqual(2, diff.Hunks.Count);

            diff.Hunks[0].Choice = HunkChoice.File;
            diff.Hunks[1].Choice = HunkChoice.Module;

            Assert.AreEqual(Text("A", "f1", "B", "m2", "C"), diff.Merge());
        }

        [Test]
        public void Merge_DropsLinesTheChosenSideLacks()
        {
            var diff = new SyncDiff(Text("A", "extra", "B"), Text("A", "B"));

            diff.ResolveAll(HunkChoice.File);

            Assert.AreEqual(Text("A", "B"), diff.Merge());
        }

        [Test]
        public void Merge_WithAllOneSide_GivesThatSide()
        {
            var module = Text("A", "m", "B", "C", "m2");
            var file = Text("A", "f", "B", "C");
            var diff = new SyncDiff(module, file);

            diff.ResolveAll(HunkChoice.Module);
            Assert.AreEqual(module, diff.Merge());

            diff.ResolveAll(HunkChoice.File);
            Assert.AreEqual(file, diff.Merge());
        }

        [Test]
        public void Merge_WithBoth_KeepsTheProjectsLinesThenTheFolders()
        {
            var diff = new SyncDiff(Text("A", "m", "B"), Text("A", "f", "B"));

            diff.ResolveAll(HunkChoice.Both);

            Assert.AreEqual(Text("A", "m", "f", "B"), diff.Merge());
        }

        [Test]
        public void Merge_WithUnresolvedHunks_Throws()
        {
            var diff = new SyncDiff("A", "B");

            Assert.IsFalse(diff.IsResolved);
            Assert.Throws<System.InvalidOperationException>(() => diff.Merge());
        }

        [Test]
        public void Compare_OfManyLines_FindsEachChange()
        {
            var module = Enumerable.Range(0, 5000).Select(i => "line" + i).ToList();
            var file = module.ToList();
            file.RemoveAt(2500);
            file.Insert(10, "new");
            file[4000] = "changed";

            var diff = new SyncDiff(string.Join("\r\n", module), string.Join("\r\n", file));

            Assert.AreEqual(3, diff.Hunks.Count);
        }

        [Test]
        public void WordDiff_MarksOnlyTheWordThatChanged()
        {
            const string module = "Dim x As Long, y As Long";
            const string file = "Dim x As Integer, y As Long";

            var diff = WinMergeWordDiff.Compute(module, file).Single();

            Assert.AreEqual("Long", module.Substring(diff.Begin0, diff.End0 - diff.Begin0 + 1));
            Assert.AreEqual("Integer", file.Substring(diff.Begin1, diff.End1 - diff.Begin1 + 1));
        }

        [Test]
        public void WordDiff_OfSameLines_IsEmpty()
        {
            Assert.AreEqual(0, WinMergeWordDiff.Compute("x = 1", "x   =  1").Count);
        }
    }
}
