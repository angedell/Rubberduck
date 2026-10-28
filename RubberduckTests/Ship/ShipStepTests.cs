using System.IO;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources.Registration;
using Rubberduck.Ship;
using Rubberduck.VBEditor.SafeComWrappers;
using static RubberduckTests.Ship.ShipTestParser;

namespace RubberduckTests.Ship
{
    [TestFixture]
    [Category("Ship")]
    public class ShipStepTests
    {
        private static ShipModuleText Text() => new ShipModuleText(ShipTestParser.Create());

        private static ShipModule Std(string name, params string[] lines)
            => new ShipModule(name, ComponentType.StandardModule, Lines(new[] { "Attribute VB_Name = \"" + name + "\"" }.Concat(lines).ToArray()));

        private static string Note(string name)
            => NoteModule.ToModuleCode("# Notes\r\nSome words", new[] { "'@Folder(\"Docs\")", "'@Notes(\"" + name + ".md\")" }) + "\r\n";

        private static ShipWorkingSet Set(params ShipModule[] modules)
        {
            var set = new ShipWorkingSet("project-1");
            set.Modules.AddRange(modules);
            return set;
        }

        [Test]
        public void RemoveNotesRemovesWholeNoteModulesOnly()
        {
            var note = new ShipModule("Note_1", ComponentType.StandardModule, "Attribute VB_Name = \"Note_1\"\r\n" + Note("README"));
            var markedButCode = Std("Marked", "'@Notes(\"x.md\")", "Public Sub Foo()", "End Sub");
            var set = Set(note, markedButCode, Std("Plain", "Public Sub Bar()", "End Sub"));

            var report = new RemoveNotesStep().Run(set);

            Assert.AreEqual(ShipSteps.Notes, report.Step);
            Assert.AreEqual(1, report.Removed);
            CollectionAssert.AreEqual(new[] { "README.md" }, report.Details);
            Assert.IsTrue(note.Removed);
            Assert.IsFalse(markedButCode.Removed);
            Assert.AreEqual(2, set.Modules.Count(m => !m.Removed));
        }

        [Test]
        public void RemoveTestModulesRemovesTestModulesAndTheUnitTestingReference()
        {
            var tests = Std("TestMath", "'@TestModule", "'@Folder(\"Tests\")", "Private Assert As Object", "'@TestMethod", "Public Sub Adds()", "End Sub");
            var lowerCase = Std("TestOther", "'@testmodule", "Public Sub X()", "End Sub");
            var code = Std("Math", "Public Function Add() As Long", "End Function");
            var set = Set(tests, lowerCase, code);
            set.References.Add(new ShipReference { Name = "Rubberduck", Guid = "{" + RubberduckGuid.RubberduckTypeLibGuid + "}", Major = 2, Minor = 5 });
            set.References.Add(new ShipReference { Name = "Scripting", Guid = "{420B2830-E718-11CF-893D-00A0C9054228}", Major = 1, Minor = 0 });

            var report = new RemoveTestModulesStep().Run(set);

            Assert.AreEqual(ShipSteps.TestModules, report.Step);
            Assert.IsTrue(tests.Removed);
            Assert.IsTrue(lowerCase.Removed);
            Assert.IsFalse(code.Removed);
            Assert.IsTrue(set.References.Single(r => r.Name == "Rubberduck").Removed);
            Assert.IsFalse(set.References.Single(r => r.Name == "Scripting").Removed);
            CollectionAssert.AreEquivalent(new[] { "TestMath", "TestOther", "reference Rubberduck" }, report.Details);
            Assert.AreEqual(3, report.Removed);
        }

        [Test]
        public void RemoveTestModulesWarnsWhenRemainingCodeUsesTheLibrary()
        {
            var user = Std("Helper", "Public Function Make() As Rubberduck.AssertClass", "End Function");
            var set = Set(Std("TestMath", "'@TestModule", "Public Sub A()", "End Sub"), user);
            set.References.Add(new ShipReference { Name = "Rubberduck", Guid = "{" + RubberduckGuid.RubberduckTypeLibGuid + "}" });

            new RemoveTestModulesStep().Run(set);

            Assert.That(set.Warnings.Any(w => w.Contains("Helper")), "A warning should name the module that still uses the library.");
        }

        [Test]
        public void RemoveCommentsStripsEveryModuleAndCountsThem()
        {
            var a = Std("A", "' one", "Public Sub Foo() ' two", "End Sub");
            var b = Std("B", "Public Sub Bar()", "End Sub");
            var set = Set(a, b);

            var report = new RemoveCommentsStep(Text()).Run(set);

            Assert.AreEqual(ShipSteps.Comments, report.Step);
            Assert.AreEqual(2, report.Removed);
            Assert.AreEqual(Lines("Attribute VB_Name = \"A\"", "Public Sub Foo()", "End Sub"), a.Text);
            Assert.AreEqual(b.OriginalText, b.Text);
        }

        [Test]
        public void RemoveCommentsLeavesAModuleThatDoesNotParseAsItWasAndListsIt()
        {
            var broken = Std("Broken", "Public Sub Foo(", "' note", "End Sub");
            var set = Set(broken, Std("Fine", "' x", "Public Sub Bar()", "End Sub"));

            new RemoveCommentsStep(Text()).Run(set);

            Assert.AreEqual(broken.OriginalText, broken.Text);
            Assert.IsNotNull(broken.UnparsedReason);
            Assert.AreEqual(1, set.Modules.Count(m => m.UnparsedReason != null));
        }

        [Test]
        public void RemoveDescriptiveAttributesAppliesTheRuleToEveryModule()
        {
            var module = Std("A",
                "Attribute VB_Ext_KEY = \"Rubberduck\" ,\"Folder\"",
                "Public Sub Foo()",
                "Attribute Foo.VB_Description = \"shown\"",
                "End Sub");
            var set = Set(module);

            var report = new RemoveDescriptiveAttributesStep(Text()).Run(set);

            Assert.AreEqual(ShipSteps.DescriptiveAttributes, report.Step);
            Assert.AreEqual(1, report.Removed);
            StringAssert.Contains("Attribute Foo.VB_Description", module.Text);
            StringAssert.DoesNotContain("VB_Ext_KEY", module.Text);
        }

        [Test]
        public void RemoveConfigMarksTheConfigPartForRemoval()
        {
            var set = Set(Std("A", "Public Sub Foo()", "End Sub"));
            set.OpenPackage = () => ShipTestPackages.Package(("customUI/rubberduck.json", "{}"), ("customUI/customUI14.xml", "<customUI/>"));

            var report = new RemoveConfigStep().Run(set);

            Assert.AreEqual(ShipSteps.Config, report.Step);
            Assert.AreEqual(1, report.Removed);
            CollectionAssert.AreEqual(new[] { "rubberduck.json" }, set.RemovedResources);
        }

        [Test]
        public void RemoveConfigDoesNothingWhenThereIsNoConfig()
        {
            var set = Set();
            set.OpenPackage = () => ShipTestPackages.Package(("customUI/customUI14.xml", "<customUI/>"));

            var report = new RemoveConfigStep().Run(set);

            Assert.AreEqual(0, report.Removed);
            Assert.IsEmpty(set.RemovedResources);
        }

        [Test]
        public void ThePipelineRunsTheEnabledStepsInOrderAndSumsTheirReports()
        {
            var note = new ShipModule("Note_1", ComponentType.StandardModule, "Attribute VB_Name = \"Note_1\"\r\n" + Note("README"));
            var set = Set(note, Std("A", "' c", "Public Sub Foo()", "End Sub"));
            var settings = new ShipSettings();
            settings.Steps[ShipSteps.Comments] = false;

            var report = new ShipPipeline(ShipTestParser.Create()).Run(set, settings, dryRun: false);

            CollectionAssert.AreEqual(
                new[] { ShipSteps.Notes, ShipSteps.TestModules, ShipSteps.DescriptiveAttributes, ShipSteps.Config },
                report.Steps.Select(s => s.Step).ToArray());
            Assert.IsFalse(report.DryRun);
            StringAssert.Contains("' c", set.Modules.Single(m => m.Name == "A").Text, "A step that is switched off changes nothing.");
        }

        [Test]
        public void ADryRunReportsTheSameAsARealRunAndSaysSo()
        {
            ShipWorkingSet Fresh() => Set(Std("A", "' c", "Public Sub Foo() ' d", "End Sub"));

            var real = new ShipPipeline(ShipTestParser.Create()).Run(Fresh(), new ShipSettings(), dryRun: false);
            var dry = new ShipPipeline(ShipTestParser.Create()).Run(Fresh(), new ShipSettings(), dryRun: true);

            Assert.IsTrue(dry.DryRun);
            Assert.AreEqual(real.Steps.Select(s => (s.Step, s.Removed)), dry.Steps.Select(s => (s.Step, s.Removed)));
        }

        [Test]
        public void ThePipelineReportsModulesThatDidNotParse()
        {
            var set = Set(Std("Broken", "Public Sub Foo(", "' note", "End Sub"));

            var report = new ShipPipeline(ShipTestParser.Create()).Run(set, new ShipSettings(), dryRun: true);

            Assert.AreEqual(1, report.Unparsed.Count);
            StringAssert.StartsWith("Broken", report.Unparsed[0]);
        }
    }

    internal static class ShipTestPackages
    {
        public static Stream Package(params (string Path, string Content)[] parts)
        {
            var stream = new MemoryStream();
            using (var zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create, true))
            {
                foreach (var (path, content) in parts)
                {
                    var entry = zip.CreateEntry(path);
                    using (var writer = new StreamWriter(entry.Open()))
                    {
                        writer.Write(content);
                    }
                }
            }

            stream.Position = 0;
            return stream;
        }
    }
}
