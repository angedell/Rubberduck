using System.Threading;
using NUnit.Framework;
using Rubberduck.UI.CodeExplorer;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class FolderPickerDialogTests
    {
        private static readonly string[] Folders = { "Api.Core", "Api.Tests", "Models", "api.core" };

        [TestCase("Api.Tests", "Api.Tests")]
        [TestCase("models", "Models")]
        [TestCase("Api", "Api")]
        [TestCase("Brand.New.Deep", "Brand.New.Deep")]
        [TestCase("", "")]
        [TestCase(null, "")]
        public void StartsOnTheFolderTheNoteIsIn(string current, string expected)
        {
            using (var dialog = new FolderPickerDialog(Folders, current))
            {
                Assert.AreEqual(expected, dialog.SelectedFolder);
            }
        }

        [Test]
        public void FoldersThatOnlyDifferInCase_AreOneNode()
        {
            using (var dialog = new FolderPickerDialog(Folders, "API.CORE"))
            {
                Assert.AreEqual("Api.Core", dialog.SelectedFolder);
            }
        }
    }
}
