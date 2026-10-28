using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.PackageResources;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    [Category("PackageResources")]
    public class CodeExplorerResourceViewModelTests
    {
        private static PackageResource Ribbon(PackageResourceChange change = PackageResourceChange.None)
            => new PackageResource { Name = "customUI14.xml", Kind = PackageResourceKind.RibbonXml, Size = 300, Change = change };

        private static PackageResource Image(string name, PackageResourceChange change = PackageResourceChange.None)
            => new PackageResource { Name = name, Kind = PackageResourceKind.Image, Size = 2048, Change = change };

        private static CodeExplorerResourceFolderViewModel Folder(params PackageResource[] resources)
            => new CodeExplorerResourceFolderViewModel(null, @"C:\docs\Book.xlsm", resources);

        private static List<CodeExplorerResourceViewModel> Leaves(CodeExplorerResourceFolderViewModel folder)
            => folder.Children.OfType<CodeExplorerResourceViewModel>().ToList();

        [Test]
        public void Folder_ListsRibbonXmlFirst_ThenPicturesByName()
        {
            var folder = Folder(Image("b.png"), Image("A.png"), Ribbon());

            CollectionAssert.AreEqual(new[] { "customUI14.xml", "A.png", "b.png" }, Leaves(folder).Select(l => l.ResourceName).ToArray());
        }

        [Test]
        public void OnlyPicturesCanBeTicked_AndTickingTheFolderTicksThem()
        {
            var folder = Folder(Ribbon(), Image("a.png"), Image("b.png"));

            folder.IsMarked = true;

            Assert.IsTrue(folder.CanBeMarked);
            Assert.IsFalse(Leaves(folder).Single(l => l.IsRibbonXml).IsMarked);
            Assert.IsTrue(Leaves(folder).Where(l => l.IsImage).All(l => l.IsMarked));
        }

        [Test]
        public void ARemovedPicture_IsDimmed_AndCannotBeTicked()
        {
            var folder = Folder(Ribbon(), Image("a.png", PackageResourceChange.Removed));
            var removed = Leaves(folder).Single(l => l.IsImage);

            Assert.IsTrue(removed.IsDimmed);
            Assert.IsFalse(removed.CanBeMarked);
            StringAssert.Contains("a.png", removed.Name);
        }

        [Test]
        public void Synchronize_AddsRemovesAndUpdates_KeepingTheNodesThatRemain()
        {
            var folder = Folder(Ribbon(), Image("a.png"), Image("b.png"));
            var kept = Leaves(folder).Single(l => l.ResourceName == "a.png");
            kept.IsMarked = true;

            folder.Synchronize(new[] { Ribbon(PackageResourceChange.Modified), Image("a.png", PackageResourceChange.Removed), Image("c.png", PackageResourceChange.Added) });

            CollectionAssert.AreEqual(new[] { "customUI14.xml", "a.png", "c.png" }, Leaves(folder).Select(l => l.ResourceName).ToArray());
            Assert.AreSame(kept, Leaves(folder).Single(l => l.ResourceName == "a.png"));
            Assert.IsFalse(kept.IsMarked, "a picture that is now going away is no longer ticked");
            Assert.IsTrue(folder.HasPendingChanges);
            StringAssert.EndsWith("*", folder.Name);
        }

        [Test]
        public void HasRibbon_IgnoresRibbonXmlThatIsRemoved()
        {
            Assert.IsTrue(Folder(Ribbon()).HasRibbon);
            Assert.IsFalse(Folder(Image("a.png")).HasRibbon);
            Assert.IsFalse(Folder().HasRibbon);
        }

        //Stands in for any other kind of node: neither a reference nor a resource folder.
        private static ICodeExplorerNode Other() => new Mock<ICodeExplorerNode>().Object;

        [Test]
        public void ResourcesSortAfterReferencesAndBeforeFolders()
        {
            var resources = Folder();
            var comparer = CodeExplorerItemComparer.NodeType;

            Assert.Less(comparer.Compare(resources, Other()), 0);
            Assert.Greater(comparer.Compare(Other(), resources), 0);
        }
    }
}
