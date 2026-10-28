using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    public class ParseFingerprintTests
    {
        private static ParseFingerprint Of(IEnumerable<(string Module, int Hash)> modules, params string[] references)
            => new ParseFingerprint(modules.Select(m => new KeyValuePair<string, int>(m.Module, m.Hash)), references);

        private static readonly (string, int)[] Two = { ("P.Module1", 1), ("P.Class1", 2) };

        [Test]
        public void TheSameModulesAndReferences_DifferInNothing()
        {
            Assert.IsEmpty(Of(Two, "a").DifferencesFrom(Of(Two.Reverse(), "a")));
        }

        [Test]
        public void ModuleNames_AreComparedIgnoringCase()
        {
            Assert.IsEmpty(Of(new[] { ("p.module1", 1) }).DifferencesFrom(Of(new[] { ("P.Module1", 1) })));
        }

        [Test]
        public void AChangedHash_IsReportedByModule()
        {
            var now = Of(new[] { ("P.Module1", 1), ("P.Class1", 9) });

            CollectionAssert.AreEqual(new[] { "changed: P.Class1" }, now.DifferencesFrom(Of(Two)));
        }

        [Test]
        public void AddedAndRemovedModules_AreReported()
        {
            var now = Of(new[] { ("P.Module1", 1), ("P.New", 3) });

            CollectionAssert.AreEqual(new[] { "added: P.New", "removed: P.Class1" }, now.DifferencesFrom(Of(Two)));
        }

        [Test]
        public void AChangedReference_IsReported()
        {
            CollectionAssert.AreEqual(new[] { "references of a project changed" }, Of(Two, "a", "b").DifferencesFrom(Of(Two, "a")));
        }

        [Test]
        public void OnlyAddsTo_IsTrueForNewModules_AndFalseForChangedOrRemovedOnes()
        {
            Assert.IsTrue(Of(new[] { ("P.Module1", 1), ("P.Class1", 2), ("P.New", 3) }).OnlyAddsTo(Of(Two)));
            Assert.IsTrue(Of(Two).OnlyAddsTo(Of(new (string, int)[0])));
            Assert.IsFalse(Of(new[] { ("P.Module1", 1), ("P.Class1", 9) }).OnlyAddsTo(Of(Two)));
            Assert.IsFalse(Of(new[] { ("P.Module1", 1) }).OnlyAddsTo(Of(Two)));
            Assert.IsTrue(Of(Two, "a").OnlyAddsTo(Of(Two)), "references of projects not listed before");
            Assert.IsTrue(Of(Two, "a", "b").OnlyAddsTo(Of(Two, "a")));
            Assert.IsFalse(Of(Two, "b").OnlyAddsTo(Of(Two, "a")), "a reference that changed or went away");
        }

        [Test]
        public void ManyChanges_ListOnlyTheFirstFew()
        {
            var before = Of(Enumerable.Range(0, 8).Select(i => ($"P.M{i}", 0)));
            var now = Of(Enumerable.Range(0, 8).Select(i => ($"P.M{i}", 1)));

            Assert.AreEqual("changed: P.M0, P.M1, P.M2, P.M3, P.M4 and 3 more", now.DifferencesFrom(before).Single());
        }
    }
}
