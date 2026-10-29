using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using NUnit.Framework;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Ship;
using Rubberduck.VBEditor.SafeComWrappers;
using static RubberduckTests.Ship.ShipTestParser;

namespace RubberduckTests.Ship
{
    /// <summary>The step that takes the analysis' removals out of the exported text, by name, with the parser's spans. No analysis here: the removals are given.</summary>
    [TestFixture]
    [Category("Ship")]
    public class ShipTreeShakingStepTests
    {
        private static ShipRemoval R(string module, string name, string kind) => new ShipRemoval { Module = module, Name = name, Kind = kind };

        private static ShipModule Std(string name, params string[] lines)
            => new ShipModule(name, ComponentType.StandardModule, Lines(new[] { "Attribute VB_Name = \"" + name + "\"" }.Concat(lines).ToArray()));

        private static ShipWorkingSet Set(params ShipModule[] modules)
        {
            var set = new ShipWorkingSet("project-1");
            set.Modules.AddRange(modules);
            return set;
        }

        private static TreeShakingStep Step(IStringParser parser, params ShipRemoval[] removals) => new TreeShakingStep(new ShipDeclarationRemover(parser), removals);

        [Test]
        public void EveryKindOfDeclarationGoesWithItsLinesAndTheSpacingOfTheModuleIsKept()
        {
            var module = Std("M",
                "Option Explicit",
                "",
                "Private Const A As Long = 1",
                "Private Const B As Long = 2",
                "",
                "Private Type T1",
                "    X As Long",
                "End Type",
                "",
                "Public Enum E1",
                "    One = 1",
                "End Enum",
                "",
                "Private mA As Long",
                "Private mB As Long",
                "",
                "Private Declare PtrSafe Function Tick Lib \"kernel32\" () As Long",
                "",
                "Public Sub Keep()",
                "End Sub",
                "",
                "Public Property Get P() As Long",
                "End Property",
                "",
                "Public Property Let P(ByVal v As Long)",
                "End Property",
                "",
                "Private Sub Last()",
                "End Sub");
            var set = Set(module);

            var report = Step(Create(), R("M", "A", "constant"), R("M", "T1", "type"), R("M", "E1", "enum"), R("M", "mB", "variable"), R("M", "Tick", "declare"),
                R("M", "P", "procedure"), R("M", "Last", "procedure")).Run(set);

            Assert.AreEqual(Lines("Attribute VB_Name = \"M\"", "Option Explicit", "", "Private Const B As Long = 2", "", "Private mA As Long", "", "Public Sub Keep()", "End Sub"), module.Text);
            Assert.AreEqual(ShipSteps.TreeShaking, report.Step);
            Assert.AreEqual(7, report.Removed);
            CollectionAssert.AreEquivalent(new[] { "M.A", "M.T1", "M.E1", "M.mB", "M.Tick", "M.P", "M.Last" }, report.Details.Skip(1).ToList());
            StringAssert.Contains("7 declaration(s) and 0 module(s)", report.Details[0]);
            Assert.IsEmpty(set.Warnings);
        }

        [Test]
        public void RemovingTheLastThingsTakesTheBlankLinesInFrontOfThemToo()
        {
            var module = Std("M", "Public Sub Keep()", "End Sub", "", "Public Sub A()", "End Sub", "", "Public Sub B()", "End Sub");

            Step(Create(), R("M", "A", "procedure"), R("M", "B", "procedure")).Run(Set(module));

            Assert.AreEqual(Lines("Attribute VB_Name = \"M\"", "Public Sub Keep()", "End Sub"), module.Text);
        }

        [Test]
        public void ARemovedProcedureTakesItsAttributeLinesAndTheTextAfterItsEndOnTheSameLine()
        {
            var module = Std("M",
                "Public Sub Keep()",
                "End Sub",
                "",
                "Public Sub Gone()",
                "Attribute Gone.VB_Description = \"d\"",
                "    Debug.Print \"End Sub\" ' not the end",
                "End Sub ' trailing",
                "",
                "Public Sub Keep2()",
                "End Sub");

            Step(Create(), R("M", "Gone", "procedure")).Run(Set(module));

            Assert.AreEqual(Lines("Attribute VB_Name = \"M\"", "Public Sub Keep()", "End Sub", "", "Public Sub Keep2()", "End Sub"), module.Text);
        }

        [Test]
        public void ABlankRunIsNeverLongerThanTheLongestOneItJoined()
        {
            var module = Std("M", "Public Sub Keep()", "End Sub", "", "", "Public Sub Gone()", "End Sub", "", "Public Sub Keep2()", "End Sub");

            Step(Create(), R("M", "Gone", "procedure")).Run(Set(module));

            Assert.AreEqual(Lines("Attribute VB_Name = \"M\"", "Public Sub Keep()", "End Sub", "", "", "Public Sub Keep2()", "End Sub"), module.Text);
        }

        [Test]
        public void AModuleWhoseRemovalIsAskedForAndThatHasNoCodeLeftIsDropped()
        {
            var module = Std("M", "Option Explicit", "", "Private Sub A()", "End Sub");
            var other = Std("Other", "Public Sub Go()", "End Sub");
            var set = Set(module, other);

            var report = Step(Create(), R("M", "A", "procedure"), R("M", "M", "module")).Run(set);

            Assert.IsTrue(module.Removed);
            Assert.IsFalse(other.Removed);
            CollectionAssert.AreEqual(new[] { "M" }, report.Details.Skip(1).ToList());
            StringAssert.Contains("0 declaration(s) and 1 module(s)", report.Details[0]);
        }

        [Test]
        public void AClassThatOnlyImplementsAnInterfaceIsDroppedOnceItsMembersAreOut()
        {
            var module = new ShipModule("C", ComponentType.ClassModule, Lines("VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "END", "Attribute VB_Name = \"C\"",
                "Attribute VB_Exposed = False", "Implements IShape", "", "Private Function IShape_Area() As Double", "End Function"));
            var set = Set(module);

            Step(Create(), R("C", "IShape_Area", "procedure"), R("C", "C", "module")).Run(set);

            Assert.IsTrue(module.Removed);
        }

        [Test]
        public void AModuleThatStillHasCodeAfterTheRemovalsIsKeptWithAWarning()
        {
            var module = Std("M", "Public Sub Stays()", "End Sub", "", "Private Sub A()", "End Sub");
            var set = Set(module);

            Step(Create(), R("M", "A", "procedure"), R("M", "M", "module")).Run(set);

            Assert.IsFalse(module.Removed);
            Assert.That(set.Warnings, Has.Some.Contain("M").And.Contain("kept"));
            StringAssert.DoesNotContain("Private Sub A", module.Text);
        }

        [Test]
        public void UserFormsAndDocumentModulesAreNeverDropped()
        {
            var form = new ShipModule("F", ComponentType.UserForm, Lines("Attribute VB_Name = \"F\"", "Private Sub A()", "End Sub"));
            var document = new ShipModule("ThisWorkbook", ComponentType.Document, Lines("Attribute VB_Name = \"ThisWorkbook\"", "Private Sub A()", "End Sub"));
            var set = Set(form, document);

            Step(Create(), R("F", "A", "procedure"), R("F", "F", "module"), R("ThisWorkbook", "A", "procedure"), R("ThisWorkbook", "ThisWorkbook", "module")).Run(set);

            Assert.IsFalse(form.Removed);
            Assert.AreEqual(form.OriginalText, form.Text);
            Assert.IsFalse(document.Removed);
            StringAssert.DoesNotContain("Private Sub A", document.Text, "A document module is emptied of its unreachable members.");
        }

        [Test]
        public void ANameIsMatchedWithoutRegardToCaseAndBracketsAndTypeHints()
        {
            var module = Std("M", "Private Count%", "", "Private x As Long", "", "Public Sub Keep()", "End Sub", "", "Private Sub [Gone]()", "End Sub");

            Step(Create(), R("M", "gone", "procedure"), R("M", "count", "variable")).Run(Set(module));

            Assert.AreEqual(Lines("Attribute VB_Name = \"M\"", "", "Private x As Long", "", "Public Sub Keep()", "End Sub"), module.Text);
        }

        [Test]
        public void AStatementThatSharesItsLineWithOthersIsKeptWithAWarning()
        {
            var module = Std("M", "Private a As Long: Private b As Long", "", "Public Sub Keep()", "End Sub");
            var set = Set(module);

            Step(Create(), R("M", "a", "variable")).Run(set);

            Assert.AreEqual(module.OriginalText, module.Text);
            Assert.That(set.Warnings, Has.Some.Contain("shares its line"));
        }

        [Test]
        public void ARemovalOfOnlySomeNamesOfOneStatementIsNotApplied()
        {
            var module = Std("M", "Private a As Long, b As Long", "", "Public Sub Keep()", "End Sub");
            var set = Set(module);

            var report = Step(Create(), R("M", "a", "variable")).Run(set);

            Assert.AreEqual(module.OriginalText, module.Text);
            Assert.AreEqual(0, report.Removed);
            Assert.That(set.Warnings, Has.Some.Contain("only some of the names"));
        }

        [Test]
        public void AModuleThatDoesNotParseAfterTheEditShipsAsItWasAndIsListedWithTheReason()
        {
            var module = Std("M", "Public Sub Keep()", "End Sub", "", "Private Sub A()", "End Sub");
            var set = Set(module);

            Step(new FailsAfterTheFirstParse(Create()), R("M", "A", "procedure")).Run(set);

            Assert.AreEqual(module.OriginalText, module.Text);
            Assert.That(set.Warnings, Has.Some.Contain("skipped").And.Contain("M").And.Contain("does not parse after the removals"));
        }

        [Test]
        public void AModuleThatDoesNotParseToBeginWithIsLeftAloneAndListed()
        {
            var module = Std("M", "Public Sub Broken(", "End Sub");

            var set = Set(module);

            Step(Create(), R("M", "Broken", "procedure")).Run(set);

            Assert.AreEqual(module.OriginalText, module.Text);
            Assert.That(set.Warnings, Has.Some.Contain("skipped").And.Contain("M"));
        }

        [Test]
        public void IfOneModuleCannotBeShakenNothingIsRemovedFromAnyModuleAndTheWarningNamesIt()
        {
            var good = Std("Good", "Public Sub Keep()", "End Sub", "", "Private Sub Dead()", "End Sub");
            var gone = Std("Gone", "Private Sub OnlyOne()", "End Sub");
            var bad = Std("Bad", "Private a As Long: Private b As Long", "", "Public Sub Keep()", "End Sub");
            var set = Set(good, gone, bad);

            var report = Step(Create(), R("Good", "Dead", "procedure"), R("Gone", "OnlyOne", "procedure"), R("Gone", "Gone", "module"), R("Bad", "a", "variable")).Run(set);

            Assert.AreEqual(good.OriginalText, good.Text);
            Assert.AreEqual(gone.OriginalText, gone.Text);
            Assert.AreEqual(bad.OriginalText, bad.Text);
            Assert.IsFalse(good.Removed || gone.Removed || bad.Removed);
            Assert.AreEqual(0, report.Removed);
            Assert.That(set.Warnings, Has.Exactly(1).Contain("Tree shaking was skipped").And.Contain("Bad").And.Contain("shares its line"));
            Assert.That(report.Details[0], Does.StartWith("Skipped"));
        }

        [Test]
        public void ARemovalThatIsNotInTheTextSkipsTheWholeStep()
        {
            var good = Std("Good", "Public Sub Keep()", "End Sub", "", "Private Sub Dead()", "End Sub");
            var other = Std("Other", "Public Sub Keep()", "End Sub");
            var set = Set(good, other);

            Step(Create(), R("Good", "Dead", "procedure"), R("Other", "NotThere", "procedure")).Run(set);

            Assert.AreEqual(good.OriginalText, good.Text);
            Assert.That(set.Warnings, Has.Some.Contain("Other").And.Contain("not found"));
        }

        [Test]
        public void TheCommentAndAnnotationLinesDirectlyAboveARemovedMemberGoWithIt()
        {
            var module = Std("M",
                "Public Sub Keep()",
                "End Sub",
                "",
                "' Does the thing.",
                "'@Description(\"Does the thing\")",
                "Private Sub Dead()",
                "End Sub",
                "",
                "' Stays: a blank line separates it from the next member.",
                "",
                "Private Sub Dead2()",
                "End Sub");

            Step(Create(), R("M", "Dead", "procedure"), R("M", "Dead2", "procedure")).Run(Set(module));

            Assert.AreEqual(Lines("Attribute VB_Name = \"M\"", "Public Sub Keep()", "End Sub", "", "' Stays: a blank line separates it from the next member."), module.Text);
        }

        [Test]
        public void AModuleDescriptionAtTheTopOfTheCodeStaysAndANormalCommentAboveALaterDeadMemberGoes()
        {
            var module = Std("M",
                "' This module does the thing.",
                "' More about it.",
                "Private Sub First()",
                "End Sub",
                "",
                "Public Sub Keep()",
                "End Sub",
                "",
                "' Helper for something.",
                "Private Sub Later()",
                "End Sub");

            Step(Create(), R("M", "First", "procedure"), R("M", "Later", "procedure")).Run(Set(module));

            Assert.AreEqual(Lines("Attribute VB_Name = \"M\"", "' This module does the thing.", "' More about it.", "", "Public Sub Keep()", "End Sub"), module.Text);
        }

        [Test]
        public void AModuleAnnotationAboveTheFirstDeadMemberStaysEvenWithCodeAboveIt()
        {
            var module = Std("M",
                "Option Explicit",
                "'@Folder(\"Tools\")",
                "Private Sub First()",
                "End Sub",
                "",
                "Public Sub Keep()",
                "End Sub");

            Step(Create(), R("M", "First", "procedure")).Run(Set(module));

            Assert.AreEqual(Lines("Attribute VB_Name = \"M\"", "Option Explicit", "'@Folder(\"Tools\")", "", "Public Sub Keep()", "End Sub"), module.Text);
        }

        [Test]
        public void TheAttributeLinesOfARemovedWithEventsVariableGoWithIt()
        {
            var module = Std("M",
                "Private WithEvents mApp As Object",
                "Attribute mApp.VB_VarHelpID = -1",
                "Private mKeep As Long",
                "",
                "Public Sub Keep()",
                "End Sub");
            var set = Set(module);

            Step(Create(), R("M", "mApp", "variable")).Run(set);

            Assert.AreEqual(Lines("Attribute VB_Name = \"M\"", "Private mKeep As Long", "", "Public Sub Keep()", "End Sub"), module.Text);
            Assert.IsEmpty(set.Warnings);
        }

        [Test]
        public void ASkippedStepCarriesOnlyTheSkipWarning()
        {
            var unreachable = Std("Left", "Public Sub Stays()", "End Sub", "", "Private Sub A()", "End Sub");
            var bad = Std("Bad", "Private a As Long: Private b As Long", "", "Public Sub Keep()", "End Sub");
            var set = Set(unreachable, bad);

            Step(Create(), R("Left", "A", "procedure"), R("Left", "Left", "module"), R("Bad", "a", "variable")).Run(set);

            Assert.AreEqual(1, set.Warnings.Count, string.Join(" | ", set.Warnings));
            StringAssert.Contains("skipped", set.Warnings[0]);
        }

        [Test]
        public void ModulesAlreadyRemovedByEarlierStepsAreLeftAlone()
        {
            var module = Std("M", "Private Sub A()", "End Sub");
            module.Removed = true;

            var report = Step(Create(), R("M", "A", "procedure")).Run(Set(module));

            Assert.AreEqual(0, report.Removed);
            Assert.AreEqual(module.OriginalText, module.Text);
        }

        [Test]
        public void ThePipelineRunsTreeShakingAfterNotesAndTestModulesAndBeforeCommentsAndTheAnalysisSeesTheModulesEarlierStepsRemoved()
        {
            var tests = Std("TestMath", "'@TestModule", "Public Sub Adds()", "End Sub");
            var code = Std("Code", "'comment", "Public Sub Go()", "End Sub", "", "Private Sub Orphan()", "End Sub");
            var set = Set(tests, code);
            var analysis = new RecordingAnalysis(new ShipAnalysisResult { Removals = { R("Code", "Orphan", "procedure") }, Roots = { "Code.Go: test" } });
            var pipeline = new ShipPipeline(Create(), analysis);
            var settings = new ShipSettings();
            settings.Steps[ShipSteps.TreeShaking] = true;

            var report = pipeline.Run(set, settings, dryRun: false, project: new Moq.Mock<Rubberduck.VBEditor.SafeComWrappers.Abstract.IVBProject>().Object);

            CollectionAssert.AreEqual(new[] { ShipSteps.Notes, ShipSteps.TestModules, ShipSteps.TreeShaking, ShipSteps.Comments, ShipSteps.DescriptiveAttributes, ShipSteps.Config }, report.Steps.Select(s => s.Step).ToList());
            Assert.IsTrue(analysis.TestModuleWasGoneWhenCalled, "The analysis runs once the test modules are out, so what only they used is not kept alive.");
            Assert.IsTrue(analysis.CommentWasStillThereWhenCalled, "The analysis runs before the comments are out: the annotations are still in the text.");
            CollectionAssert.AreEqual(new[] { "Code.Go: test" }, report.Roots);
            StringAssert.DoesNotContain("Orphan", code.Text);
        }

        [Test]
        public void ThePipelineReportsTheDynamicCallsOfTheAnalysis()
        {
            var set = Set(Std("Code", "Public Sub Go()", "End Sub"));
            var analysis = new RecordingAnalysis(new ShipAnalysisResult { DynamicCalls = { "Code.Go: Application.Run with a computed target" } });
            var settings = new ShipSettings();
            settings.Steps[ShipSteps.TreeShaking] = true;

            var report = new ShipPipeline(Create(), analysis).Run(set, settings, dryRun: true, project: new Moq.Mock<Rubberduck.VBEditor.SafeComWrappers.Abstract.IVBProject>().Object);

            CollectionAssert.AreEqual(new[] { "Code.Go: Application.Run with a computed target" }, report.DynamicCalls);
        }

        [Test]
        public void WhenTheAnalysisSaysNoTheReasonIsAWarningAndTheOtherStepsStillRun()
        {
            var set = Set(Std("Code", "'comment", "Public Sub Go()", "End Sub"));
            var settings = new ShipSettings();
            settings.Steps[ShipSteps.TreeShaking] = true;

            var report = new ShipPipeline(Create(), new RecordingAnalysis(null) { WhyNot = "module Code does not parse" })
                .Run(set, settings, dryRun: false, project: new Moq.Mock<Rubberduck.VBEditor.SafeComWrappers.Abstract.IVBProject>().Object);

            Assert.That(report.Warnings, Has.Some.EqualTo("Tree shaking was skipped: module Code does not parse"));
            Assert.IsNull(report.Steps.SingleOrDefault(s => s.Step == ShipSteps.TreeShaking));
            Assert.IsNotNull(report.Steps.SingleOrDefault(s => s.Step == ShipSteps.Comments));
        }

        [Test]
        public void TreeShakingOffMeansTheAnalysisIsNeverAsked()
        {
            var analysis = new RecordingAnalysis(new ShipAnalysisResult());

            new ShipPipeline(Create(), analysis).Run(Set(Std("Code", "Public Sub Go()", "End Sub")), new ShipSettings(), dryRun: false, project: new Moq.Mock<Rubberduck.VBEditor.SafeComWrappers.Abstract.IVBProject>().Object);

            Assert.AreEqual(0, analysis.Calls);
        }

        private class RecordingAnalysis : IShipAnalysis
        {
            private readonly ShipAnalysisResult _result;

            public RecordingAnalysis(ShipAnalysisResult result)
            {
                _result = result;
            }

            public string WhyNot { get; set; }
            public int Calls { get; private set; }
            public bool TestModuleWasGoneWhenCalled { get; private set; }
            public bool CommentWasStillThereWhenCalled { get; private set; }

            public ShipAnalysisResult Analyze(Rubberduck.VBEditor.SafeComWrappers.Abstract.IVBProject project, ShipWorkingSet set, ShipTreeShakingSettings settings, out string whyNot)
            {
                Calls++;
                TestModuleWasGoneWhenCalled = set.Modules.Where(m => m.Name == "TestMath").All(m => m.Removed);
                CommentWasStillThereWhenCalled = set.Modules.Where(m => m.Name == "Code").All(m => m.Text.Contains("'comment"));
                whyNot = WhyNot;
                return _result;
            }
        }

        private class FailsAfterTheFirstParse : IStringParser
        {
            private readonly IStringParser _inner;
            private int _calls;

            public FailsAfterTheFirstParse(IStringParser inner)
            {
                _inner = inner;
            }

            public (IParseTree tree, ITokenStream tokenStream) Parse(string moduleName, string projectId, string code, CancellationToken token, CodeKind codeKind = CodeKind.SnippetCode, ParserMode parserMode = ParserMode.FallBackSllToLl)
            {
                if (_calls++ > 0)
                {
                    throw new InvalidOperationException("synthetic failure");
                }

                return _inner.Parse(moduleName, projectId, code, token, codeKind, parserMode);
            }
        }
    }
}
