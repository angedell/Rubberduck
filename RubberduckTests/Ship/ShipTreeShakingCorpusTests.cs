using System.Linq;
using NUnit.Framework;
using Rubberduck.Ship;
using Rubberduck.VBEditor.SafeComWrappers;
using static RubberduckTests.Ship.CorpusModule;

namespace RubberduckTests.Ship
{
    /// <summary>
    /// Small real projects, parsed by Rubberduck's own parser over a mock VBE, shaken, and parsed again. Every case asserts what goes and the invariant of ADR 0004:
    /// what is left parses, and every reference of every member that is left resolves to a declaration that is left.
    /// </summary>
    [TestFixture]
    [Category("Ship")]
    public class ShipTreeShakingCorpusTests
    {
        private class Outcome
        {
            public Corpus Corpus;
            public ShipWorkingSet Set;
            public ShipReport Report;
            public ShipStepReport Step => Report.Steps.SingleOrDefault(s => s.Step == ShipSteps.TreeShaking);
            public string Text(string module) => Set.Modules.Single(m => m.Name == module).Text;
            public bool Gone(string module) => Set.Modules.Single(m => m.Name == module).Removed;
            public string[] Names => Step.Details.Skip(1).OrderBy(n => n, System.StringComparer.OrdinalIgnoreCase).ToArray();
        }

        private static Outcome Shake(ShipSettings settings, params CorpusModule[] modules)
        {
            var corpus = Corpus.Parse(modules);
            var (set, report) = corpus.Shake(settings);
            Assert.IsEmpty(report.Warnings, "Tree shaking should have run: " + string.Join(" | ", report.Warnings));
            Assert.IsNotNull(report.Steps.SingleOrDefault(s => s.Step == ShipSteps.TreeShaking), "The step did not run.");
            Assert.IsEmpty(report.Unparsed, string.Join(" | ", report.Unparsed));
            corpus.AssertShakenIsSound(set);
            return new Outcome { Corpus = corpus, Set = set, Report = report };
        }

        private static Outcome Conservative(params CorpusModule[] modules) => Shake(Corpus.Only(), modules);
        private static Outcome EntryPointsOnly(params CorpusModule[] modules) => Shake(Corpus.Only(ShipTreeShakingMode.EntryPointsOnly), modules);

        [Test]
        public void AnUnreachablePrivateProcedureIsRemovedAndWhatItAloneUsedWithIt()
        {
            var outcome = Conservative(Std("Module1",
                "Option Explicit",
                "",
                "Public Sub Go()",
                "    Helper",
                "End Sub",
                "",
                "Private Sub Helper()",
                "    Debug.Print 1",
                "End Sub",
                "",
                "Private Sub Orphan()",
                "    OnlyForOrphan",
                "End Sub",
                "",
                "Private Sub OnlyForOrphan()",
                "End Sub"));

            CollectionAssert.AreEqual(new[] { "Module1.Orphan", "Module1.OnlyForOrphan" }.OrderBy(n => n).ToArray(), outcome.Names.OrderBy(n => n).ToArray());
            Assert.AreEqual(Lines("Attribute VB_Name = \"Module1\"", "Option Explicit", "", "Public Sub Go()", "    Helper", "End Sub", "", "Private Sub Helper()", "    Debug.Print 1", "End Sub"), outcome.Text("Module1"));
        }

        private static string[] Sorted(params string[] names) => names.OrderBy(n => n, System.StringComparer.OrdinalIgnoreCase).ToArray();

        [Test]
        public void ConservativeKeepsEveryPublicProcedureOfAStandardModuleAndEntryPointsOnlyDropsTheOnesNothingUses()
        {
            var module = Std("Module1",
                "Public Sub Used()",
                "    Helper",
                "End Sub",
                "",
                "Private Sub Helper()",
                "End Sub",
                "",
                "Public Sub Unused()",
                "End Sub");

            var conservative = Conservative(module);
            var entryPoints = EntryPointsOnly(module);

            CollectionAssert.IsEmpty(conservative.Names);
            Assert.AreEqual(module.Exported, conservative.Text("Module1"));
            CollectionAssert.AreEqual(Sorted("Module1"), entryPoints.Names, "Nothing is a root: the module is left empty and goes.");
            Assert.IsTrue(entryPoints.Gone("Module1"));
        }

        [Test]
        public void AnEntryPointAnnotationIsARootInEntryPointsOnlyMode()
        {
            var outcome = EntryPointsOnly(Std("Module1",
                "'@EntryPoint",
                "Public Sub Used()",
                "    Helper",
                "End Sub",
                "",
                "Private Sub Helper()",
                "End Sub",
                "",
                "Public Sub Unused()",
                "End Sub"));

            CollectionAssert.AreEqual(Sorted("Module1.Unused"), outcome.Names);
            Assert.AreEqual(Lines("Attribute VB_Name = \"Module1\"", "'@EntryPoint", "Public Sub Used()", "    Helper", "End Sub", "", "Private Sub Helper()", "End Sub"), outcome.Text("Module1"));
            Assert.That(outcome.Report.Roots, Has.Member("Module1.Used: @EntryPoint / @ExcelHotkey"));
        }

        [Test]
        public void EventHandlersAreRootsAndWhatTheyCallIsKept()
        {
            var outcome = EntryPointsOnly(
                Document("ThisWorkbook",
                    "Private Sub Workbook_Open()",
                    "    Module1.Setup",
                    "End Sub",
                    "",
                    "Private Sub NotAHandler()",
                    "End Sub"),
                Std("Module1",
                    "Public Sub Setup()",
                    "End Sub",
                    "",
                    "Public Sub Spare()",
                    "End Sub"));

            CollectionAssert.AreEqual(Sorted("ThisWorkbook.NotAHandler", "Module1.Spare"), outcome.Names);
            Assert.IsFalse(outcome.Gone("ThisWorkbook"), "A document module is only ever emptied of its unreachable members.");
        }

        [Test]
        public void WithEventsHandlersAndClassInitializeAreKeptWhileTheirClassIsInUse()
        {
            var outcome = EntryPointsOnly(
                Std("Module1",
                    "'@EntryPoint",
                    "Public Sub Go()",
                    "    Dim w As Watcher",
                    "    Set w = New Watcher",
                    "End Sub"),
                Class("Watcher",
                    "Private WithEvents mSource As Source",
                    "",
                    "Private Sub Class_Initialize()",
                    "    Set mSource = New Source",
                    "End Sub",
                    "",
                    "Private Sub mSource_Changed()",
                    "    Debug.Print 1",
                    "End Sub",
                    "",
                    "Private Sub Unused()",
                    "End Sub"),
                Class("Source",
                    "Public Event Changed()",
                    "",
                    "Public Sub Fire()",
                    "    RaiseEvent Changed",
                    "End Sub",
                    "",
                    "Public Sub NeverFired()",
                    "End Sub"));

            CollectionAssert.AreEqual(Sorted("Watcher.Unused", "Source.Fire", "Source.NeverFired"), outcome.Names);
        }

        [Test]
        public void AnImplementedInterfaceStaysWholeAndAnImplementationNothingUsesGoesWithItsClass()
        {
            var outcome = EntryPointsOnly(
                Std("Module1",
                    "'@EntryPoint",
                    "Public Sub Go()",
                    "    Dim s As IShape",
                    "    Set s = New Square",
                    "    Debug.Print s.Area",
                    "End Sub"),
                Class("IShape",
                    "Public Function Area() As Double",
                    "End Function",
                    "",
                    "Public Function Describe() As String",
                    "End Function"),
                Class("Square",
                    "Implements IShape",
                    "",
                    "Private Function IShape_Area() As Double",
                    "End Function",
                    "",
                    "Private Function IShape_Describe() As String",
                    "End Function"),
                Class("Circle",
                    "Implements IShape",
                    "",
                    "Private Function IShape_Area() As Double",
                    "End Function",
                    "",
                    "Private Function IShape_Describe() As String",
                    "End Function"));

            CollectionAssert.AreEqual(Sorted("Circle"), outcome.Names);
            Assert.IsTrue(outcome.Gone("Circle"));
            Assert.IsFalse(outcome.Gone("IShape"));
            StringAssert.Contains("Describe", outcome.Text("IShape"), "A class that implements an interface has to implement all of it, so all of it stays.");
            StringAssert.Contains("IShape_Describe", outcome.Text("Square"));
        }

        [Test]
        public void AnImplementationWhoseNameDiffersInCaseFromTheInterfaceMemberStays()
        {
            var outcome = EntryPointsOnly(
                Std("Module1",
                    "'@EntryPoint",
                    "Public Sub Go()",
                    "    Dim s As IShape",
                    "    Set s = New Square",
                    "    Debug.Print s.area",
                    "End Sub"),
                Class("IShape",
                    "Public Function area() As Double",
                    "End Function"),
                Class("Square",
                    "Implements IShape",
                    "",
                    "Private Function IShape_Area() As Double",
                    "End Function"));

            Assert.IsEmpty(outcome.Names);
            StringAssert.Contains("IShape_Area", outcome.Text("Square"), "The VBE does not compile a class that lacks a member of its interface, whatever the case.");
        }

        [Test]
        public void AnImplementationOfAnInterfaceFromATypeLibraryStays()
        {
            var outcome = EntryPointsOnly(
                Std("Module1",
                    "'@EntryPoint",
                    "Public Sub Go()",
                    "    Dim stub As RibbonButtonStub",
                    "    Set stub = New RibbonButtonStub",
                    "    stub.Init \"x\"",
                    "End Sub"),
                Class("RibbonButtonStub",
                    "Implements IRibbonControl",
                    "",
                    "Private mId As String",
                    "",
                    "Public Sub Init(ByVal buttonId As String)",
                    "    mId = buttonId",
                    "End Sub",
                    "",
                    "Private Property Get IRibbonControl_Context() As Object",
                    "End Property",
                    "",
                    "Private Property Get IRibbonControl_Id() As String",
                    "    IRibbonControl_Id = mId",
                    "End Property",
                    "",
                    "Private Property Get IRibbonControl_Tag() As String",
                    "End Property"));

            Assert.IsEmpty(outcome.Names);
            StringAssert.Contains("IRibbonControl_Id", outcome.Text("RibbonButtonStub"), "The VBE does not compile a class that lacks a member of its interface, library interface or not.");
            StringAssert.Contains("IRibbonControl_Tag", outcome.Text("RibbonButtonStub"));
        }

        [Test]
        public void ALiteralTargetOfApplicationRunAndOnTimeIsARootAnythingElseIsNot()
        {
            var outcome = EntryPointsOnly(
                Std("Module1",
                    "'@EntryPoint",
                    "Public Sub Go()",
                    "    Application.Run \"Module2.Target\"",
                    "    Application.OnTime Now, \"'Book.xlam'!Module2.Later\"",
                    "End Sub"),
                Std("Module2",
                    "Public Sub Target()",
                    "End Sub",
                    "",
                    "Public Sub Later()",
                    "End Sub",
                    "",
                    "Public Sub Other()",
                    "End Sub"));

            CollectionAssert.AreEqual(Sorted("Module2.Other"), outcome.Names);
            Assert.That(outcome.Report.Roots, Has.Member("Module2.Target: string literal in Module1"));
            Assert.That(outcome.Report.Roots, Has.Member("Module2.Later: string literal in Module1"));
        }

        [Test]
        public void ALateBoundCallAndACallByNameLiteralKeepTheMembersTheyCanReach()
        {
            var outcome = EntryPointsOnly(
                Std("Module1",
                    "'@EntryPoint",
                    "Public Sub Go()",
                    "    Dim w As Object",
                    "    Set w = New Worker",
                    "    w.Work",
                    "    CallByName w, \"ByName\", VbMethod",
                    "End Sub"),
                Class("Worker",
                    "Public Sub Work()",
                    "End Sub",
                    "",
                    "Public Sub ByName()",
                    "End Sub",
                    "",
                    "Public Sub Spare()",
                    "End Sub"));

            CollectionAssert.AreEqual(Sorted("Worker.Spare"), outcome.Names);
        }

        [Test]
        public void TheKeepListPinsByMemberOrByModuleAndMemberWithoutRegardToCase()
        {
            var outcome = Shake(Corpus.Only(ShipTreeShakingMode.EntryPointsOnly, "module1.KEPT", "plain"),
                Std("Module1", "Public Sub Kept()", "End Sub", "", "Public Sub Other()", "End Sub"),
                Std("Module2", "Public Sub Plain()", "End Sub", "", "Public Sub Gone()", "End Sub"));

            CollectionAssert.AreEqual(Sorted("Module1.Other", "Module2.Gone"), outcome.Names);
        }

        [Test]
        public void ThePropertiesOfOneNameAreOneUnit()
        {
            var outcome = EntryPointsOnly(
                Std("Module1",
                    "'@EntryPoint",
                    "Public Sub Go()",
                    "    Dim b As Box",
                    "    Set b = New Box",
                    "    Debug.Print b.Size",
                    "End Sub"),
                Class("Box",
                    "Private mSize As Long",
                    "Private mUnused As Long",
                    "",
                    "Public Property Get Size() As Long",
                    "    Size = mSize",
                    "End Property",
                    "",
                    "Public Property Let Size(ByVal value As Long)",
                    "    mSize = value",
                    "End Property",
                    "",
                    "Public Property Get Unused() As Long",
                    "    Unused = mUnused",
                    "End Property",
                    "",
                    "Public Property Let Unused(ByVal value As Long)",
                    "    mUnused = value",
                    "End Property"));

            CollectionAssert.AreEqual(Sorted("Box.Unused", "Box.mUnused"), outcome.Names);
            StringAssert.Contains("Property Let Size", outcome.Text("Box"), "Only the Get of Size is read, but Get and Let are one unit.");
            StringAssert.DoesNotContain("Unused", outcome.Text("Box"));
        }

        [Test]
        public void ConstantsTypesEnumsAndVariablesNothingReachesAreRemovedAndStatementsOfSeveralNamesStandOrFallTogether()
        {
            var outcome = EntryPointsOnly(Std("Module1",
                "Option Explicit",
                "",
                "Private Const KeptConst As Long = 3",
                "Private Const DeadConst As Long = 4",
                "",
                "Private Type TUsed",
                "    Id As Long",
                "End Type",
                "",
                "Private Type TDead",
                "    Id As Long",
                "End Type",
                "",
                "Public Enum EUsed",
                "    First = 1",
                "End Enum",
                "",
                "Public Enum EDead",
                "    Other = 2",
                "End Enum",
                "",
                "Private mCount As Long",
                "Private mDeadA As Long, mDeadB As Long",
                "Private mKept As Long, mAlsoKept As Long",
                "",
                "'@EntryPoint",
                "Public Sub Go()",
                "    Dim t As TUsed",
                "    Dim e As EUsed",
                "    t.Id = KeptConst + e",
                "    mCount = mCount + mKept",
                "End Sub"));

            CollectionAssert.AreEqual(Sorted("Module1.DeadConst", "Module1.TDead", "Module1.EDead", "Module1.mDeadA", "Module1.mDeadB"), outcome.Names);
            StringAssert.Contains("mAlsoKept", outcome.Text("Module1"));
            StringAssert.DoesNotContain("mDead", outcome.Text("Module1"));
        }

        [Test]
        public void ADeclareNothingCallsIsRemoved()
        {
            var outcome = EntryPointsOnly(Std("Module1",
                "Private Declare PtrSafe Function GetTickCount Lib \"kernel32\" () As Long",
                "Private Declare PtrSafe Function GetCurrentProcessId Lib \"kernel32\" () As Long",
                "",
                "'@EntryPoint",
                "Public Sub Go()",
                "    Debug.Print GetTickCount",
                "End Sub"));

            CollectionAssert.AreEqual(Sorted("Module1.GetCurrentProcessId"), outcome.Names);
        }

        [Test]
        public void AClassModuleNothingReferencesIsRemovedWhole()
        {
            var outcome = EntryPointsOnly(
                Std("Module1", "'@EntryPoint", "Public Sub Go()", "End Sub"),
                Class("Unused", "Public Sub Hello()", "End Sub"));

            CollectionAssert.AreEqual(Sorted("Unused"), outcome.Names);
            Assert.IsTrue(outcome.Gone("Unused"));
            Assert.IsFalse(outcome.Gone("Module1"));
        }

        [Test]
        public void AClassOnlyDeclaredAsIsKeptAndLosesTheMembersNothingCalls()
        {
            var outcome = EntryPointsOnly(
                Std("Module1", "'@EntryPoint", "Public Sub Go()", "    Dim c As Counter", "End Sub"),
                Class("Counter", "Public Sub Bump()", "End Sub"));

            CollectionAssert.AreEqual(Sorted("Counter.Bump"), outcome.Names);
            Assert.IsFalse(outcome.Gone("Counter"));
        }

        [Test]
        public void ReferencesAmongUnreachableMembersDoNotKeepThem()
        {
            var outcome = EntryPointsOnly(Std("Module1",
                "'@EntryPoint",
                "Public Sub Keep()",
                "End Sub",
                "",
                "Private Sub A()",
                "    B",
                "End Sub",
                "",
                "Private Sub B()",
                "    A",
                "End Sub",
                "",
                "Private Sub C()",
                "    C",
                "End Sub"));

            CollectionAssert.AreEqual(Sorted("Module1.A", "Module1.B", "Module1.C"), outcome.Names);
        }

        [Test]
        public void ExposedClassesKeepTheirPublicMembersAndDocumentModulesLoseTheUnreachableOnes()
        {
            var outcome = EntryPointsOnly(
                ExposedClass("Api", "Public Sub Hello()", "End Sub", "", "Private Sub Hidden()", "End Sub"),
                Document("Sheet1", "Public Sub Foo()", "End Sub"));

            CollectionAssert.AreEqual(Sorted("Api.Hidden", "Sheet1.Foo"), outcome.Names);
            Assert.IsFalse(outcome.Gone("Api"));
            Assert.IsFalse(outcome.Gone("Sheet1"));
        }

        [Test]
        public void AControlHandlerOnAWorksheetIsKeptEvenThoughNothingDeclaresTheControl()
        {
            var outcome = EntryPointsOnly(Document("Sheet1", "Private Sub CommandButton1_Click()", "End Sub", "", "Private Sub Helper()", "End Sub"));

            CollectionAssert.AreEqual(Sorted("Sheet1.Helper"), outcome.Names);
        }

        [Test]
        public void ForUserFormsNothingIsEverRemovedAndTheirCodeIsARoot()
        {
            var form = Form("UserForm1",
                "Private Sub CommandButton1_Click()",
                "    Helper",
                "End Sub",
                "",
                "Private Sub Helper()",
                "    Module1.UsedByTheForm",
                "End Sub",
                "",
                "Private Sub Spare()",
                "End Sub");
            var outcome = EntryPointsOnly(
                Std("Module1", "'@EntryPoint", "Public Sub Go()", "    Dim f As UserForm1", "End Sub", "", "Public Sub UsedByTheForm()", "End Sub", "", "Public Sub Spare()", "End Sub"),
                form);

            CollectionAssert.AreEqual(Sorted("Module1.Spare"), outcome.Names);
            Assert.AreEqual(form.Exported, outcome.Text("UserForm1"));
        }

        [Test]
        public void RibbonCallbacksAndShapeMacrosOfTheDocumentAreRoots()
        {
            using (var package = new System.IO.MemoryStream())
            {
                using (var zip = new System.IO.Compression.ZipArchive(package, System.IO.Compression.ZipArchiveMode.Create, true))
                using (var writer = new System.IO.StreamWriter(zip.CreateEntry("customUI/customUI14.xml").Open()))
                {
                    writer.Write("<customUI><ribbon><tabs><tab id=\"t\"><group id=\"g\"><button id=\"b\" onAction=\"OnRibbonClick\"/></group></tab></tabs></ribbon></customUI>");
                }

                var bytes = package.ToArray();
                var corpus = Corpus.Parse(Std("Module1", "Public Sub OnRibbonClick(control As Object)", "End Sub", "", "Public Sub Unreferenced()", "End Sub"));
                var (set, report) = corpus.Shake(Corpus.Only(ShipTreeShakingMode.EntryPointsOnly), () => new System.IO.MemoryStream(bytes));

                Assert.IsEmpty(report.Warnings);
                CollectionAssert.AreEqual(new[] { "Module1.Unreferenced" }, report.Steps.Single().Details.Skip(1).ToArray());
                Assert.That(report.Roots, Has.Member("Module1.OnRibbonClick: found in customUI/customUI14.xml onAction"));
                corpus.AssertShakenIsSound(set);
            }
        }

        [Test]
        public void ADynamicCallIsListedWhenItsMemberIsReachableAndNotWhenItIsNot()
        {
            var corpus = Corpus.Parse(Std("Module1",
                "'@EntryPoint",
                "Public Sub Go(ByVal name As String, ByVal o As Object)",
                "    Application.Run name",
                "    Application.Run \"Literal\"",
                "    Application.OnTime Now, name",
                "    Application.OnKey \"^q\", name",
                "    CallByName o, name, VbMethod",
                "End Sub",
                "",
                "Public Sub Literal()",
                "End Sub",
                "",
                "Private Sub Dead(ByVal n As String)",
                "    Application.Run n",
                "End Sub"));

            var result = corpus.Analyze(new ShipTreeShakingSettings { Mode = ShipTreeShakingMode.EntryPointsOnly }, out var why);

            Assert.IsNotNull(result, why);
            Assert.AreEqual(4, result.DynamicCalls.Count, string.Join("\r\n", result.DynamicCalls));
            Assert.That(result.DynamicCalls, Has.All.StartsWith("Module1.Go: "));
            Assert.That(result.DynamicCalls, Has.Some.Contain("Application.Run").And.Contain("(name)"));
            Assert.That(result.DynamicCalls, Has.Some.Contain("Application.OnTime"));
            Assert.That(result.DynamicCalls, Has.Some.Contain("Application.OnKey"));
            Assert.That(result.DynamicCalls, Has.Some.Contain("CallByName"));
        }

        [Test]
        public void AParserThatIsNotReadySkipsTheAnalysisWithTheReason()
        {
            var corpus = Corpus.Parse(Std("Module1", "Public Sub Go()", "End Sub"));
            var analysis = new Rubberduck.Ship.ParserStateShipAnalysis(corpus.State, ShipTestParser.CreateFor(null, corpus.ProjectId), () => Rubberduck.Parsing.VBA.ParserState.Parsing);

            var result = analysis.Analyze(corpus.Project, corpus.Set(), new ShipTreeShakingSettings(), out var why);

            Assert.IsNull(result);
            StringAssert.Contains("Parsing", why);
        }

        [Test]
        public void AParserThatCannotCatchUpSkipsTheAnalysisWithTheReason()
        {
            var corpus = Corpus.Parse(Std("Module1", "Public Sub Go()", "End Sub"));
            var analysis = new Rubberduck.Ship.ParserStateShipAnalysis(corpus.State, ShipTestParser.CreateFor(null, corpus.ProjectId), () => throw Rubberduck.Api.ApiException.NotInDesignMode());

            var result = analysis.Analyze(corpus.Project, corpus.Set(), new ShipTreeShakingSettings(), out var why);

            Assert.IsNull(result);
            StringAssert.Contains("up to date", why);
        }

        [Test]
        public void AModuleThatDoesNotParseSkipsTheWholeStepAndTheOtherStepsStillRun()
        {
            var corpus = Corpus.Parse(
                Std("Module1", "'a comment", "Public Sub Go()", "End Sub", "", "Private Sub Orphan()", "End Sub"),
                Std("Broken", "Public Sub Oops(", "End Sub"));
            var settings = Corpus.Only(ShipTreeShakingMode.EntryPointsOnly);
            settings.Steps[ShipSteps.Comments] = true;

            var (set, report) = corpus.Shake(settings);

            Assert.That(report.Warnings, Has.Some.StartsWith("Tree shaking was skipped"));
            Assert.IsNull(report.Steps.SingleOrDefault(s => s.Step == ShipSteps.TreeShaking));
            Assert.IsNotNull(report.Steps.SingleOrDefault(s => s.Step == ShipSteps.Comments), "The other steps still run.");
            StringAssert.Contains("Orphan", set.Modules.Single(m => m.Name == "Module1").Text, "Nothing is removed when the project cannot be analysed.");
            StringAssert.DoesNotContain("a comment", set.Modules.Single(m => m.Name == "Module1").Text);
            CollectionAssert.IsEmpty(report.Roots);
        }

        [Test]
        public void WithoutAnAnalysisTheStepIsSkippedWithAWarningAndNothingElseChanges()
        {
            var corpus = Corpus.Parse(Std("Module1", "Public Sub Go()", "End Sub"));
            var set = corpus.Set();
            var pipeline = new ShipPipeline(ShipTestParser.CreateFor(null, corpus.ProjectId));

            var report = pipeline.Run(set, Corpus.Only(), dryRun: true, project: corpus.Project);

            Assert.That(report.Warnings, Has.Some.StartsWith("Tree shaking was skipped"));
            CollectionAssert.IsEmpty(report.Steps);
            Assert.IsTrue(report.DryRun);
        }

        [Test]
        public void AnAnalysisThatFailsSkipsTheStepWithAWarningInsteadOfCrashing()
        {
            var corpus = Corpus.Parse(Std("Module1", "Public Sub Go()", "End Sub"));
            var pipeline = new ShipPipeline(ShipTestParser.CreateFor(null, corpus.ProjectId), new ThrowingAnalysis());

            var report = pipeline.Run(corpus.Set(), Corpus.Only(), dryRun: false, project: corpus.Project);

            Assert.That(report.Warnings, Has.Some.Contain("analysis failed").And.Contain("boom"));
        }

        private class ThrowingAnalysis : IShipAnalysis
        {
            public ShipAnalysisResult Analyze(Rubberduck.VBEditor.SafeComWrappers.Abstract.IVBProject project, ShipWorkingSet set, ShipTreeShakingSettings settings, out string whyNot)
                => throw new System.InvalidOperationException("boom");
        }

        [Test]
        public void ADryRunAnalysesExactlyAsARealRunDoes()
        {
            var modules = new[] { Std("Module1", "'@EntryPoint", "Public Sub Go()", "End Sub", "", "Public Sub Orphan()", "End Sub") };
            var corpus = Corpus.Parse(modules);
            var pipeline = new ShipPipeline(ShipTestParser.CreateFor(null, corpus.ProjectId), corpus.Analysis());
            var settings = Corpus.Only(ShipTreeShakingMode.EntryPointsOnly);

            var dry = pipeline.Run(corpus.Set(), settings, dryRun: true, project: corpus.Project);
            var real = pipeline.Run(corpus.Set(), settings, dryRun: false, project: corpus.Project);

            CollectionAssert.AreEqual(real.Steps.Single().Details, dry.Steps.Single().Details);
            CollectionAssert.AreEqual(real.Roots, dry.Roots);
        }

        private static CorpusModule Entry(string name, params string[] body)
            => Std(name, new[] { "'@EntryPoint", "Public Sub Go()" }.Concat(body).Concat(new[] { "End Sub" }).ToArray());

        [Test]
        public void TheEnumeratorMemberOfACollectionClassIsKeptEvenThoughOnlyForEachUsesIt()
        {
            var outcome = EntryPointsOnly(
                Entry("Module1",
                    "    Dim c As Items",
                    "    Set c = New Items",
                    "    c.Add 1",
                    "    Dim x As Variant",
                    "    For Each x In c",
                    "        Debug.Print x",
                    "    Next"),
                Class("Items",
                    "Public Sub Add(ByVal v As Variant)",
                    "End Sub",
                    "",
                    "Public Function NewEnum() As Object",
                    "Attribute NewEnum.VB_UserMemId = -4",
                    "    Set NewEnum = Nothing",
                    "End Function",
                    "",
                    "Public Sub Spare()",
                    "End Sub"));

            CollectionAssert.AreEqual(Sorted("Items.Spare"), outcome.Names);
            Assert.That(outcome.Report.Roots, Has.Some.Contain("Items.NewEnum").And.Contain("VB_UserMemId"));
        }

        [Test]
        public void TheDefaultMemberOfAClassIsKeptEvenThoughOnlyCollParenthesesOneUsesIt()
        {
            var outcome = EntryPointsOnly(
                Entry("Module1",
                    "    Dim c As Items",
                    "    Set c = New Items",
                    "    Debug.Print c(1)"),
                Class("Items",
                    "Public Function Item(ByVal i As Long) As Long",
                    "Attribute Item.VB_UserMemId = 0",
                    "    Item = i",
                    "End Function",
                    "",
                    "Public Sub Spare()",
                    "End Sub"));

            CollectionAssert.AreEqual(Sorted("Items.Spare"), outcome.Names);
        }

        [Test]
        public void LateBoundAccessAndCallByNameKeepFieldsConstantsAndPropertiesToo()
        {
            var outcome = EntryPointsOnly(
                Entry("Module1",
                    "    Dim o As Object",
                    "    Set o = New Bag",
                    "    o.Total = 1",
                    "    Debug.Print o.Size",
                    "    Debug.Print CallByName(o, \"Limit\", VbGet)",
                    "    Debug.Print CallByName(o, \"Colour\", VbGet)"),
                Class("Bag",
                    "Public Total As Long",
                    "Public Spare As Long",
                    "Public Const Limit As Long = 5",
                    "Public Const SpareLimit As Long = 6",
                    "Private mColour As Long",
                    "",
                    "Public Property Get Size() As Long",
                    "End Property",
                    "",
                    "Public Property Get Colour() As Long",
                    "    Colour = mColour",
                    "End Property",
                    "",
                    "Public Property Get SpareSize() As Long",
                    "End Property"));

            CollectionAssert.AreEqual(Sorted("Bag.Spare", "Bag.SpareLimit", "Bag.SpareSize"), outcome.Names);
        }

        [Test]
        public void AMemberNameAsAWholeWordInsideAnyStringLiteralIsARoot()
        {
            var outcome = EntryPointsOnly(
                Entry("Module1",
                    "    Dim s As String",
                    "    s = \"=MyUdf(A1)+1\"",
                    "    s = \"Evaluate(OtherUdf(2))\"",
                    "    s = \"'Module2.OnClick 1'\""),
                Std("Module2",
                    "Public Function MyUdf(ByVal x As Variant) As Variant",
                    "End Function",
                    "",
                    "Public Function OtherUdf(ByVal x As Variant) As Variant",
                    "End Function",
                    "",
                    "Public Sub OnClick(ByVal n As Long)",
                    "End Sub",
                    "",
                    "Public Function MyUdfs() As Long",
                    "End Function",
                    "",
                    "Public Function Unrelated() As Long",
                    "End Function"));

            CollectionAssert.AreEqual(Sorted("Module2.MyUdfs", "Module2.Unrelated"), outcome.Names, "Whole words only: MyUdf is not MyUdfs.");
        }

        [Test]
        public void ACallInAnInactiveConditionalBranchStillKeepsItsTarget()
        {
            var outcome = EntryPointsOnly(
                Std("Module1",
                    "'@EntryPoint",
                    "Public Sub Go()",
                    "#If Mac Then",
                    "    Call MacHelper",
                    "#End If",
                    "End Sub",
                    "",
                    "Private Sub MacHelper()",
                    "End Sub",
                    "",
                    "Private Sub Orphan()",
                    "End Sub"));

            CollectionAssert.AreEqual(Sorted("Module1.Orphan"), outcome.Names);
        }

        [Test]
        public void ADeclareTwinInAnInactiveBranchIsNotLeftBehindWhenItsActiveTwinGoes()
        {
            var outcome = EntryPointsOnly(
                Std("Module1",
                    "#If Win64 Then",
                    "Private Declare PtrSafe Function GetTick Lib \"kernel32\" Alias \"GetTickCount\" () As Long",
                    "#Else",
                    "Private Declare Function GetTick Lib \"kernel32\" Alias \"GetTickCount\" () As Long",
                    "#End If",
                    "",
                    "'@EntryPoint",
                    "Public Sub Go()",
                    "End Sub",
                    "",
                    "Private Sub Orphan()",
                    "End Sub"));

            CollectionAssert.AreEqual(Sorted("Module1.Orphan"), outcome.Names);
            StringAssert.Contains("PtrSafe Function GetTick", outcome.Text("Module1"));
            StringAssert.Contains("Declare Function GetTick", outcome.Text("Module1"));
        }

        [Test]
        public void RemovingAMemberBetweenConditionalBlocksLeavesTheDirectivesBalanced()
        {
            var outcome = EntryPointsOnly(
                Std("Module1",
                    "'@EntryPoint",
                    "Public Sub Go()",
                    "End Sub",
                    "#If Mac Then",
                    "Private Sub MacOne()",
                    "End Sub",
                    "#End If",
                    "Private Sub Orphan()",
                    "End Sub",
                    "#If Mac Then",
                    "Private Sub MacTwo()",
                    "End Sub",
                    "#End If"));

            CollectionAssert.AreEqual(Sorted("Module1.Orphan"), outcome.Names);
            var text = outcome.Text("Module1");
            Assert.AreEqual(2, System.Text.RegularExpressions.Regex.Matches(text, @"^#If ", System.Text.RegularExpressions.RegexOptions.Multiline).Count);
            Assert.AreEqual(2, System.Text.RegularExpressions.Regex.Matches(text, @"^#End If", System.Text.RegularExpressions.RegexOptions.Multiline).Count);
        }

        [Test]
        public void AKeepEntryThatMatchesNothingIsAWarning()
        {
            var corpus = Corpus.Parse(Std("Module1", "Public Sub Kept()", "End Sub"));

            var (set, report) = corpus.Shake(Corpus.Only(ShipTreeShakingMode.EntryPointsOnly, "Module1.Kept", "Module1.Typo", "Nowhere"));

            Assert.That(report.Warnings, Has.Exactly(1).Contain("Module1.Typo").And.Contain("matches nothing"));
            Assert.That(report.Warnings, Has.Exactly(1).Contain("Nowhere").And.Contain("matches nothing"));
            Assert.That(report.Warnings, Has.None.Contain("Module1.Kept"));
        }

        [TestCase("Logger.Logger")]
        [TestCase("logger")]
        public void AKeepEntryThatNamesAModuleKeepsTheModuleItself(string entry)
        {
            var outcome = Shake(Corpus.Only(ShipTreeShakingMode.EntryPointsOnly, entry),
                Entry("Module1", "    Debug.Print 1"),
                Class("Logger", "Public Sub Log()", "End Sub"));

            Assert.IsFalse(outcome.Gone("Logger"));
        }

        [Test]
        public void AKeepEntryForAModuleWithNoMembersKeepsIt()
        {
            var outcome = Shake(Corpus.Only(ShipTreeShakingMode.EntryPointsOnly, "Marker.Marker"),
                Entry("Module1", "    Debug.Print 1"),
                Class("Marker", "Option Explicit"));

            Assert.IsFalse(outcome.Gone("Marker"));
            Assert.That(outcome.Report.Warnings, Is.Empty);
        }

        [Test]
        public void AClassUsedOnlyInAnInactiveConditionalBranchIsNotDropped()
        {
            var outcome = EntryPointsOnly(
                Entry("Module1",
                    "#If Mac Then",
                    "    Dim l As Logger",
                    "#End If"),
                Class("Logger", "Public Sub Log()", "End Sub"),
                Class("Unused", "Public Sub Log()", "End Sub"));

            Assert.IsFalse(outcome.Gone("Logger"));
            Assert.IsTrue(outcome.Gone("Unused"));
            Assert.That(outcome.Report.Roots, Has.Some.Contain("Logger: name inside a conditional compilation block in Module1"));
        }

        [Test]
        public void AClassNamedOnlyInAStringIsNotDropped()
        {
            var outcome = EntryPointsOnly(
                Entry("Module1", "    Debug.Print CreateObject(\"Logger\")"),
                Class("Logger", "Public Sub Log()", "End Sub"),
                Class("Unused", "Public Sub Log()", "End Sub"));

            Assert.IsFalse(outcome.Gone("Logger"));
            Assert.IsTrue(outcome.Gone("Unused"));
        }

        [Test]
        public void ADefaultMemberDoesNotKeepAnUnreachableStandardModule()
        {
            var outcome = EntryPointsOnly(
                Entry("Module1", "    Debug.Print 1"),
                Std("Orphan",
                    "Public Function Item(ByVal i As Long) As Long",
                    "Attribute Item.VB_UserMemId = 0",
                    "    Item = i",
                    "End Function"));

            Assert.IsTrue(outcome.Gone("Orphan"));
        }

        [Test]
        public void IfTheParserStateChangesWhileTheAnalysisRunsTheStepIsSkippedWithAClearReason()
        {
            var corpus = Corpus.Parse(Std("Module1", "Public Sub Go()", "End Sub"));
            var generation = 0;
            var analysis = new Rubberduck.Ship.ParserStateShipAnalysis(corpus.State, ShipTestParser.CreateFor(null, corpus.ProjectId), () => corpus.State.Status, () => generation++);

            var result = analysis.Analyze(corpus.Project, corpus.Set(), new ShipTreeShakingSettings(), out var why);

            Assert.IsNull(result);
            Assert.AreEqual("the project changed while Ship was working; try again", why);
        }

        [Test]
        public void AnEnumerationThatBreaksBecauseOfAReparseIsTheSameSkipAndNotAnAnalysisFailure()
        {
            var corpus = Corpus.Parse(Std("Module1", "Public Sub Go()", "End Sub"));
            var calls = 0;
            var analysis = new Rubberduck.Ship.ParserStateShipAnalysis(corpus.State, ShipTestParser.CreateFor(null, corpus.ProjectId), () => corpus.State.Status,
                () => ++calls == 1 ? 0 : throw new System.InvalidOperationException("Collection was modified"));
            var pipeline = new ShipPipeline(ShipTestParser.CreateFor(null, corpus.ProjectId), analysis);

            var report = pipeline.Run(corpus.Set(), Corpus.Only(), dryRun: false, project: corpus.Project);

            Assert.That(report.Warnings, Has.Some.EqualTo("Tree shaking was skipped: the project changed while Ship was working; try again"));
            Assert.That(report.Warnings, Has.None.Contain("analysis failed"));
        }

        private static string Lines(params string[] lines) => string.Join("\r\n", lines) + "\r\n";
    }
}
