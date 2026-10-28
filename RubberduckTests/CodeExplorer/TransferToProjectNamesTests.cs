using System.Collections.Generic;
using NUnit.Framework;
using Rubberduck.UI.CodeExplorer.Commands;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    public class TransferToProjectNamesTests
    {
        [Test]
        public void UniqueName_AddsTheFirstFreeNumber_IgnoringCase()
        {
            Assert.AreEqual("Helpers2", TransferToProjectCommandBase.UniqueName("Helpers", new[] { "Helpers", "Other" }));
            Assert.AreEqual("Helpers3", TransferToProjectCommandBase.UniqueName("Helpers", new[] { "helpers", "HELPERS2" }));
        }

        [Test]
        public void UniqueNoteName_KeepsTheExtension_AndLooksOnlyInTheFolder()
        {
            var taken = new HashSet<string>(new[] { "Api/Notes.md", "Api/Notes2.md", "Notes.md" });

            Assert.AreEqual("Notes3.md", TransferToProjectCommandBase.UniqueNoteName("Notes.md", "Api", taken));
            Assert.AreEqual("Notes2.md", TransferToProjectCommandBase.UniqueNoteName("Notes.md", null, taken));
        }
    }
}
