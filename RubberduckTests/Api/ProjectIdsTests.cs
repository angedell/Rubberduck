using System.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    public class ProjectIdsTests
    {
        private const string Xlsm = @"C:\Dropbox\2024\DevMacros.xlsm";
        private const string Other = @"C:\Dropbox\2024\Macro64.xlsm";

        private static readonly (string EditorId, string Name, string Path)[] Open =
        {
            ("㼿", "DevMacros", Xlsm),
            ("㼿", "Saro64", Other),
            ("-931300149", "DevMacros", @"C:\Dropbox\2024\DevMacros.xlam")
        };

        private static (string EditorId, string Name, string Path) Describe((string EditorId, string Name, string Path) p) => p;

        [Test]
        public void TheSameFileHasTheSameId_AndDifferentFilesHaveDifferentIds_EvenWhenTheEditorGivesThemOne()
        {
            Assert.AreEqual(ProjectIds.Of("㼿", "DevMacros", Xlsm), ProjectIds.Of("㼿", "DevMacros", Xlsm.ToUpperInvariant()));
            Assert.AreNotEqual(ProjectIds.Of("㼿", "DevMacros", Xlsm), ProjectIds.Of("㼿", "Saro64", Other));
        }

        [Test]
        public void TheIdIsPlainAscii()
        {
            StringAssert.IsMatch("^[pu][0-9a-f]{8}$", ProjectIds.Of("㼿", "DevMacros", Xlsm));
            StringAssert.IsMatch("^[pu][0-9a-f]{8}$", ProjectIds.Of("㼿", "Unsaved", null));
        }

        [Test]
        public void AnUnsavedProject_IsToldApartByNameAndEditorId()
        {
            Assert.AreNotEqual(ProjectIds.Of("1", "Book1", null), ProjectIds.Of("2", "Book1", null));
        }

        [Test]
        public void TheApiIdAddressesExactlyOneProject_WhereTheEditorsIdAddressesTwo()
        {
            var id = ProjectIds.Of("㼿", "Saro64", Other);

            var byApiId = ProjectIds.Match(id, Open, Describe);
            var byEditorId = ProjectIds.Match("㼿", Open, Describe);

            Assert.AreEqual(Other, byApiId.Single().Path);
            Assert.AreEqual(2, byEditorId.Count);
        }

        [Test]
        public void TheNameStillWorks_AndSaysWhenItIsNotEnough()
        {
            Assert.AreEqual(Other, ProjectIds.Match("Saro64", Open, Describe).Single().Path);
            Assert.AreEqual(2, ProjectIds.Match("DevMacros", Open, Describe).Count);
            Assert.AreEqual(0, ProjectIds.Match("Nothing", Open, Describe).Count);
        }
    }
}
