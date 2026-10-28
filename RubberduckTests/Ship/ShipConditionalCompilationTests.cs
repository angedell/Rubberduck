using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using NUnit.Framework;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Ship;
using static RubberduckTests.Ship.ShipTestParser;

namespace RubberduckTests.Ship
{
    /// <summary>
    /// What the parser does with #If: the preprocessor hides the tokens of the inactive branch and of every directive line, so the parse tree has no comments there.
    /// These tests pin down what the stripper does with each of them.
    /// </summary>
    [TestFixture]
    [Category("Ship")]
    public class ShipConditionalCompilationTests
    {
        private static ShipTextResult Strip(string text, short dev = 1)
            => new ShipModuleText(ShipTestParser.Create(new Dictionary<string, short> { { "Dev", dev } })).StripComments("Module1", "project-1", text);

        private static readonly string Code = Lines(
            "Public Sub Foo()",
            "#If Dev Then ' why dev",
            "    Debug.Print 1 ' live",
            "    ' live line",
            "#ElseIf Other Then ' other",
            "    Debug.Print 2 ' other branch",
            "#Else ' not dev",
            "    Debug.Print \"it's dead\" ' dead",
            "    Rem dead rem",
            "    ' dead line",
            "#End If ' done",
            "End Sub");

        private static readonly string Expected = Lines(
            "Public Sub Foo()",
            "#If Dev Then",
            "    Debug.Print 1",
            "#ElseIf Other Then",
            "    Debug.Print 2",
            "#Else",
            "    Debug.Print \"it's dead\"",
            "#End If",
            "End Sub");

        [Test]
        public void CommentsInTheActiveBranchTheInactiveBranchesAndOnTheDirectiveLinesAreAllRemoved()
        {
            var result = Strip(Code, dev: 1);

            Assert.IsTrue(result.Parsed, result.Problem);
            Assert.AreEqual(Expected, result.Text);
        }

        [Test]
        public void ItDoesNotMatterWhichBranchIsActive()
        {
            Assert.AreEqual(Expected, Strip(Code, dev: 0).Text);
        }

        [Test]
        public void ThePreprocessorHidesTheCommentsOfInactiveBranchesAndDirectiveLinesButNotOfTheActiveBranch()
        {
            // The reason the stripper looks for comment-mark tokens as well as at the tree's comment contexts: lines 2, 5, 7 and 11 are directive lines, 6, 8 and 10 are in inactive branches.
            // Lines 3 and 4 are in the active branch (Dev is 1) and are ordinary comments in the parse tree.
            var tokens = ShipTestParser.Create(new Dictionary<string, short> { { "Dev", 1 } })
                .Parse("Module1", "project-1", Code, System.Threading.CancellationToken.None, CodeKind.AttributesCode).tokenStream;
            var hiddenMarks = ((BufferedTokenStream)tokens).GetTokens()
                .Where(t => t.Type == VBALexer.SINGLEQUOTE && t.Channel == TokenConstants.HiddenChannel)
                .Select(t => t.Line)
                .ToList();

            CollectionAssert.AreEqual(new[] { 2, 5, 6, 7, 8, 10, 11 }, hiddenMarks);
        }

        [Test]
        public void ADirectiveThatDoesNotParseMakesTheModuleShipUnchanged()
        {
            var code = Lines("Public Sub Foo()", "#If Then ' broken", "End Sub");
            var result = Strip(code);

            Assert.IsFalse(result.Parsed);
            Assert.AreEqual(code, result.Text);
        }
    }
}
