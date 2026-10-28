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
    /// The stripped code must be the original code and nothing else. The expected tokens come from a rule that does not use the stripper or the parse tree: in the lexer's tokens of the
    /// original, a comment mark (or Rem at the start of a statement) runs to the end of its line. What is left, ignoring spaces, line breaks and line continuations (whitespace to the grammar),
    /// has to be what the stripped text lexes to, token for token, type and text.
    /// </summary>
    [TestFixture]
    [Category("Ship")]
    public class ShipRoundTripTests
    {
        private static readonly string Module = Lines(
            "'@Folder(\"Core\")",
            "Attribute VB_Name = \"Module1\"",
            "Option Explicit",
            "' header comment",
            "Private Const Quote As String = \"it's ' not a comment\" ' but this is",
            "Private Const RemText As String = \"Rem also not\"",
            "Public Declare PtrSafe Function GetTickCount Lib \"kernel32\" () As Long ' win api",
            "#If VBA7 Then",
            "    Private Declare PtrSafe Sub Sleep Lib \"kernel32\" (ByVal ms As Long) ' modern",
            "#Else",
            "    Private Declare Sub Sleep Lib \"kernel32\" (ByVal ms As Long) ' legacy",
            "#End If",
            "",
            "'@Description(\"Adds\")",
            "Public Function Add(ByVal a As Long, _",
            "                    ByVal b As Long) As Long",
            "Attribute Add.VB_Description = \"Adds\"",
            "    Rem old style",
            "    Add = a + _",
            "          b ' sum",
            "End Function",
            "",
            "Public Sub Labels()",
            "10  Debug.Print \"one\" ' ten",
            "Again: ' label",
            "    If Len(\"x\") = 1 Then Debug.Print \"'\": Rem trailing",
            "    GoTo Again ' loop",
            "End Sub",
            "' last line");

        private static IEnumerable<(int Type, string Text)> Significant(IEnumerable<IToken> tokens)
            => tokens.Where(t => t.Type != VBALexer.WS && t.Type != VBALexer.NEWLINE && t.Type != VBALexer.LINE_CONTINUATION && t.Type != TokenConstants.Eof)
                .Select(t => (t.Type, t.Text));

        private static List<IToken> Lex(string code)
        {
            var stream = new SimpleVBAModuleTokenStreamProvider().Tokens(code);
            stream.Fill();
            return stream.GetTokens().ToList();
        }

        private static IEnumerable<IToken> WithoutComments(IList<IToken> tokens)
        {
            var atStatementStart = true;
            for (var i = 0; i < tokens.Count; i++)
            {
                var token = tokens[i];
                if (token.Type == VBALexer.SINGLEQUOTE || (token.Type == VBALexer.REM && atStatementStart))
                {
                    while (i + 1 < tokens.Count && tokens[i + 1].Type != VBALexer.NEWLINE)
                    {
                        i++;
                    }

                    continue;
                }

                if (token.Type == VBALexer.NEWLINE || token.Type == VBALexer.COLON)
                {
                    atStatementStart = true;
                }
                else if (token.Type != VBALexer.WS && token.Type != VBALexer.LINE_CONTINUATION)
                {
                    atStatementStart = false;
                }

                yield return token;
            }
        }

        [TestCase(0)]
        [TestCase(1)]
        public void TheStrippedCodeHasTheSameTokensAsTheOriginalWithoutItsComments(short vba7)
        {
            var parser = ShipTestParser.Create(new Dictionary<string, short> { { "VBA7", vba7 } });
            var result = new ShipModuleText(parser).StripComments("Module1", "project-1", Module);
            Assert.IsTrue(result.Parsed, result.Problem);

            var expected = Significant(WithoutComments(Lex(Module))).ToList();
            var actual = Significant(Lex(result.Text)).ToList();

            CollectionAssert.AreEqual(expected, actual);
            Assert.That(result.Removed, Is.GreaterThan(10));
            StringAssert.DoesNotContain("' ", result.Text.Replace("\"it's ' not a comment\"", string.Empty).Replace("\"'\"", string.Empty));
        }

        [Test]
        public void TheStrippedCodeStillParses()
        {
            var parser = ShipTestParser.Create();
            var result = new ShipModuleText(parser).StripComments("Module1", "project-1", Module);

            var again = new ShipModuleText(parser).StripComments("Module1", "project-1", result.Text);

            Assert.IsTrue(again.Parsed, again.Problem);
            Assert.AreEqual(0, again.Removed, "Nothing is left to strip the second time.");
            Assert.AreEqual(result.Text, again.Text);
        }
    }
}
