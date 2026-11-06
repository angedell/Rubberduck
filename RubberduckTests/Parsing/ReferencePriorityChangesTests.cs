using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.VBA.ComReferenceLoading;

namespace RubberduckTests.Parsing
{
    [TestFixture]
    public class ReferencePriorityChangesTests
    {
        private static IDictionary<string, ReferencePriorityMap> References(params (string library, string project, int priority)[] entries)
        {
            var references = new Dictionary<string, ReferencePriorityMap>();
            foreach (var (library, project, priority) in entries)
            {
                if (!references.TryGetValue(library, out var map))
                {
                    map = new ReferencePriorityMap(library);
                    references[library] = map;
                }

                map[project] = priority;
            }

            return references;
        }

        [Test]
        public void SameReferencesInTheSameOrder_ChangeNothing()
        {
            var old = References(("Excel", "P", 1), ("Office", "P", 2), ("Excel", "Q", 1));
            var current = References(("Excel", "P", 1), ("Office", "P", 2), ("Excel", "Q", 1));

            Assert.IsEmpty(COMReferenceSynchronizerBase.ReferencePriorityChanges(current, old));
        }

        [Test]
        public void AReferenceThatMovedDown_IsReportedForTheProjectThatHasIt()
        {
            var old = References(("Excel", "P", 1), ("Office", "P", 2), ("Office", "Q", 1));
            var current = References(("Excel", "P", 2), ("Office", "P", 1), ("Office", "Q", 1));

            var changes = COMReferenceSynchronizerBase.ReferencePriorityChanges(current, old);

            CollectionAssert.AreEquivalent(new[] { "P" }, changes.Keys);
            CollectionAssert.AreEquivalent(new[] { (2, 1), (1, 2) }, changes["P"]);
        }

        [Test]
        public void AReferenceThatWasRemovedFromOneProject_IsReportedForThatProjectOnly()
        {
            var old = References(("Excel", "P", 1), ("Excel", "Q", 1));
            var current = References(("Excel", "Q", 1));

            var changes = COMReferenceSynchronizerBase.ReferencePriorityChanges(current, old);

            CollectionAssert.AreEquivalent(new[] { "P" }, changes.Keys);
            Assert.AreEqual((int.MaxValue, 1), changes["P"].Single());
        }

        [Test]
        public void ALibraryThatIsNoLongerReferencedAtAll_IsReportedForEveryProjectThatHadIt()
        {
            var old = References(("Excel", "P", 1), ("Excel", "Q", 3));

            var changes = COMReferenceSynchronizerBase.ReferencePriorityChanges(References(), old);

            CollectionAssert.AreEquivalent(new[] { "P", "Q" }, changes.Keys);
        }

        [Test]
        public void ANewReference_IsReportedForTheProjectThatAddedIt()
        {
            var old = References(("Excel", "P", 1));
            var current = References(("Excel", "P", 1), ("Scripting", "P", 2));

            var changes = COMReferenceSynchronizerBase.ReferencePriorityChanges(current, old);

            Assert.AreEqual((2, int.MaxValue), changes["P"].Single());
        }
    }
}
