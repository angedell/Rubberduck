using System.Linq;
using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    [Category("GitSync")]
    public class SyncFixerTests
    {
        private static SyncConfig Fixing(params string[] rules)
        {
            var config = new SyncConfig { OnViolation = ViolationMode.Fix };
            config.Enforce.AddRange(rules.Length == 0 ? SyncConfig.Rules : rules);
            return config;
        }

        private static string[] Remaining(string text, SyncConfig config, bool folderFollowsPath = true) =>
            SyncHygiene.Check("Foo", text, SyncSide.Module, config, folderFollowsPath).Select(v => v.Rule).ToArray();

        [Test]
        public void NothingHappensUnlessFixing()
        {
            var config = Fixing();
            config.OnViolation = ViolationMode.Warn;
            var text = "Attribute VB_PredeclaredId = True\r\nOption Explicit";

            Assert.AreEqual(text, SyncFixer.Fix("Foo", text, config));
            Assert.AreEqual(text, SyncFixer.Fix("Foo", text, new SyncConfig { OnViolation = ViolationMode.Fix }));
        }

        [Test]
        public void MissingAnnotationsAreAddedForTrueAttributes()
        {
            var text = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1\r\nEND\r\nAttribute VB_Name = \"Foo\"\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\nOption Explicit";

            var fixedText = SyncFixer.Fix("Foo", text, Fixing());

            StringAssert.Contains("'@PredeclaredId", fixedText);
            StringAssert.Contains("'@Exposed", fixedText);
            CollectionAssert.IsEmpty(Remaining(fixedText, Fixing()));
        }

        [Test]
        public void AttributesFollowTheirAnnotations()
        {
            var text = "Attribute VB_Name = \"Foo\"\r\n'@PredeclaredId\r\n'@Exposed\r\nOption Explicit";

            var fixedText = SyncFixer.Fix("Foo", text, Fixing());

            StringAssert.Contains("Attribute VB_PredeclaredId = True", fixedText);
            StringAssert.Contains("Attribute VB_Exposed = True", fixedText);
            CollectionAssert.IsEmpty(Remaining(fixedText, Fixing()));
        }

        [Test]
        public void ModuleAttributeAnnotationsSetTheirAttributes()
        {
            var text = "Attribute VB_Name = \"Foo\"\r\nAttribute VB_Ext_Key = \"Key\", \"Old\"\r\n'@ModuleAttribute VB_Ext_Key, \"Key\", \"New\"\r\nOption Explicit";

            var fixedText = SyncFixer.Fix("Foo", text, Fixing());

            StringAssert.Contains("Attribute VB_Ext_Key = \"Key\", \"New\"", fixedText);
            StringAssert.DoesNotContain("Old", fixedText);
            CollectionAssert.IsEmpty(Remaining(fixedText, Fixing()));
        }

        [Test]
        public void VbNameIsAddedOrCorrected()
        {
            var missing = SyncFixer.Fix("Foo", "Option Explicit", Fixing("MissingAttribute"));
            var wrong = SyncFixer.Fix("Foo", "Attribute VB_Name = \"Bar\"\r\nOption Explicit", Fixing("MissingAttribute"));

            StringAssert.StartsWith("Attribute VB_Name = \"Foo\"", missing);
            StringAssert.Contains("Attribute VB_Name = \"Foo\"", wrong);
            StringAssert.DoesNotContain("Bar", wrong);
        }

        [Test]
        public void OnlyEnforcedRulesAreFixed()
        {
            var text = "Attribute VB_Name = \"Foo\"\r\nAttribute VB_Exposed = True\r\nOption Explicit";

            Assert.AreEqual(text, SyncFixer.Fix("Foo", text, Fixing("MissingAttribute")));
        }

        [Test]
        public void FixingTwiceChangesNothing()
        {
            var once = SyncFixer.Fix("Foo", "Attribute VB_Exposed = True\r\n'@PredeclaredId\r\nOption Explicit", Fixing());

            Assert.AreEqual(once, SyncFixer.Fix("Foo", once, Fixing()));
        }

        [Test]
        public void FolderComesFromTheFilesLocation()
        {
            var text = "Attribute VB_Name = \"Foo\"\r\nOption Explicit";

            StringAssert.Contains("'@Folder(\"Models.Orders\")", SyncFixer.FixFolder(text, "Models/Orders", Fixing()));
            Assert.AreEqual(text, SyncFixer.FixFolder(text, string.Empty, Fixing()));
        }
    }
}
