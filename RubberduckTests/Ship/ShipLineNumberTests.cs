using System.Linq;
using NUnit.Framework;
using Rubberduck.Ship;
using Rubberduck.VBEditor.SafeComWrappers;
using static RubberduckTests.Ship.ShipTestParser;

namespace RubberduckTests.Ship
{
    /// <summary>The line numbers step: <c>Erl</c> in the shipped project says the line of the user's own module.</summary>
    [TestFixture]
    [Category("Ship")]
    public class ShipLineNumberTests
    {
        private static ShipModuleText Text() => new ShipModuleText(ShipTestParser.Create());

        private static ShipTextResult Number(string text) => Text().NumberLines("Module1", "project-1", text);

        [Test]
        public void ExecutableLinesGetTheirLineInTheCodePane()
        {
            var result = Number(Lines(
                "Attribute VB_Name = \"Module1\"",
                "Option Explicit",
                "",
                "Public Sub Run()",
                "Attribute Run.VB_Description = \"Runs it\"",
                "    Dim x As Long",
                "    x = 1",
                "    If x > 0 Then",
                "        Debug.Print x",
                "    Else",
                "        Debug.Print 0",
                "    End If",
                "End Sub"));

            Assert.IsTrue(result.Parsed, result.Problem);
            //The attribute lines are not in the code pane: Option Explicit is its line 1.
            Assert.AreEqual(Lines(
                "Attribute VB_Name = \"Module1\"",
                "Option Explicit",
                "",
                "Public Sub Run()",
                "Attribute Run.VB_Description = \"Runs it\"",
                "    Dim x As Long",
                "5     x = 1",
                "6     If x > 0 Then",
                "7         Debug.Print x",
                "    Else",
                "9         Debug.Print 0",
                "    End If",
                "End Sub"), result.Text);
            Assert.AreEqual(4, result.Numbered);
        }

        [Test]
        public void ContinuationsColonsCommentsLabelsAndDeclarationsAreNotNumbered()
        {
            var result = Number(Lines(
                "Attribute VB_Name = \"Module1\"",
                "Public Function F(ByVal n As Long) As Long",
                "    Const k = 1",
                "    Dim s As String",
                "    s = \"a\" & _",
                "        \"b\"",
                "    If n = 1 Then F = 1 Else F = 2",
                "    n = n + 1: n = n + 2",
                "    ' a comment",
                "    Select Case n",
                "        Case 1",
                "            F = 3",
                "        Case Else",
                "            F = 4",
                "    End Select",
                "Done:",
                "    F = F + k",
                "End Function"));

            Assert.IsTrue(result.Parsed, result.Problem);
            Assert.AreEqual(Lines(
                "Attribute VB_Name = \"Module1\"",
                "Public Function F(ByVal n As Long) As Long",
                "    Const k = 1",
                "    Dim s As String",
                "4     s = \"a\" & _",
                "        \"b\"",
                "6     If n = 1 Then F = 1 Else F = 2",
                "7     n = n + 1: n = n + 2",
                "    ' a comment",
                "9     Select Case n",
                "        Case 1",
                "11             F = 3",
                "        Case Else",
                "13             F = 4",
                "    End Select",
                "Done:",
                "16     F = F + k",
                "End Function"), result.Text);
        }

        [Test]
        public void TheHeaderOfAClassIsNotCounted()
        {
            var result = Text().NumberLines("Thing", "project-1", Lines(
                "VERSION 1.0 CLASS",
                "BEGIN",
                "  MultiUse = -1  'True",
                "END",
                "Attribute VB_Name = \"Thing\"",
                "Attribute VB_Exposed = False",
                "Option Explicit",
                "",
                "Public Sub Go()",
                "    Debug.Print 1",
                "End Sub"));

            Assert.IsTrue(result.Parsed, result.Problem);
            StringAssert.Contains("4     Debug.Print 1", result.Text);
        }

        [Test]
        public void AProcedureWithLineNumbersOfItsOwnIsLeftAlone()
        {
            var result = Number(Lines(
                "Attribute VB_Name = \"Module1\"",
                "Public Sub Mine()",
                "10  Debug.Print 1",
                "20  Debug.Print 2",
                "End Sub",
                "Public Sub Plain()",
                "    Debug.Print 3",
                "End Sub"));

            Assert.IsTrue(result.Parsed, result.Problem);
            Assert.AreEqual(Lines(
                "Attribute VB_Name = \"Module1\"",
                "Public Sub Mine()",
                "10  Debug.Print 1",
                "20  Debug.Print 2",
                "End Sub",
                "Public Sub Plain()",
                "6     Debug.Print 3",
                "End Sub"), result.Text);
            CollectionAssert.AreEqual(new[] { "Mine" }, result.Skipped);
        }

        [Test]
        public void AModuleThatDoesNotParseComesBackUnchanged()
        {
            var text = Lines("Attribute VB_Name = \"Module1\"", "Public Sub Broken(", "End Sub");

            var result = Number(text);

            Assert.IsFalse(result.Parsed);
            Assert.AreEqual(text, result.Text);
        }

        [Test]
        public void TheNumbersStayTheLinesOfTheSourceWhenTheCommentsAreTakenOut()
        {
            var module = new ShipModule("Module1", ComponentType.StandardModule, Lines(
                "Attribute VB_Name = \"Module1\"",
                "' a comment line",
                "'@Folder(\"X\")",
                "Public Sub Run()",
                "    ' another",
                "    Debug.Print 1",
                "End Sub"));
            var set = new ShipWorkingSet("project-1");
            set.Modules.Add(module);

            var numbers = new AddLineNumbersStep(Text()).Run(set);
            new RemoveCommentsStep(Text()).Run(set);

            Assert.AreEqual(1, numbers.Removed);
            Assert.AreEqual(Lines(
                "Attribute VB_Name = \"Module1\"",
                "Public Sub Run()",
                "5     Debug.Print 1",
                "End Sub"), module.Text);
        }

        [Test]
        public void TheStepIsOffUntilSwitchedOnAndRunsBeforeTheStepsThatTakeLinesOut()
        {
            Assert.IsFalse(new ShipSettings().IsOn(ShipSteps.LineNumbers));
            var order = ShipSteps.All.ToList();
            Assert.Less(order.IndexOf(ShipSteps.LineNumbers), order.IndexOf(ShipSteps.TreeShaking));
            Assert.Less(order.IndexOf(ShipSteps.LineNumbers), order.IndexOf(ShipSteps.Comments));
        }
    }
}
