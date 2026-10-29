using System.Linq;
using NUnit.Framework;
using Rubberduck.Ship;

namespace RubberduckTests.Ship
{
    /// <summary>
    /// The names Tree shaking keeps because they stand inside a conditional compilation block (ADR 0004 item 3). Found with Rubberduck's lexer and its conditional compilation grammar,
    /// so what is a directive is decided by the grammar and not by the text of a line.
    /// </summary>
    [TestFixture]
    [Category("Ship")]
    public class ShipConditionalWordsTests
    {
        private static string Code(params string[] lines) => string.Join("\r\n", lines) + "\r\n";

        private static string[] Words(params string[] lines)
            => ParserStateShipAnalysis.ConditionalWords(Code(lines)).Distinct().OrderBy(w => w).ToArray();

        [Test]
        public void NamesInAllBranchesOfABlockAreFoundAndNamesOutsideAreNot()
        {
            var words = Words("Sub Outside()", "End Sub", "#If Win64 Then", "Call InWin64", "#ElseIf Mac Then", "Call InMac", "#Else", "Call InElse", "#End If", "Call AlsoOutside");

            CollectionAssert.IsSubsetOf(new[] { "InWin64", "InMac", "InElse" }, words);
            CollectionAssert.DoesNotContain(words, "Outside");
            CollectionAssert.DoesNotContain(words, "AlsoOutside");
        }

        [Test]
        public void NestedBlocksAreOneBlockAndTheCodeAfterTheInnerEndIfIsStillInside()
        {
            var words = Words("#If A Then", "#If B Then", "Call Inner", "#End If", "Call StillOuter", "#End If", "Call After");

            CollectionAssert.IsSubsetOf(new[] { "A", "B", "Inner", "StillOuter" }, words);
            CollectionAssert.DoesNotContain(words, "After");
        }

        [Test]
        public void ExtraSpacesAndALineContinuationInTheDirectiveAreNotAProblem()
        {
            var words = Words("#If   Win64   Then", "Call Spaced", "#End   If", "#If Win64 And _", "    Mac Then", "Call Continued", "#End If", "Call After");

            CollectionAssert.IsSubsetOf(new[] { "Spaced", "Continued" }, words);
            CollectionAssert.DoesNotContain(words, "After");
        }

        [Test]
        public void ConstLinesAreNotBlocksButTheyCountInsideOne()
        {
            var words = Words("#Const Debugging = 1", "Call Before", "#If Debugging Then", "#Const Inner = 2", "Call During", "#End If");

            CollectionAssert.DoesNotContain(words, "Before");
            CollectionAssert.IsSubsetOf(new[] { "During", "Inner" }, words);
        }

        [Test]
        public void AnIfDirectiveInAStringOrInACommentDoesNotOpenABlock()
        {
            var words = Words("s = \"#If X Then\"", "Call AfterString", "' #If Y Then", "Call AfterComment", "Call X ' #If Z Then", "Call AfterTrailingComment");

            CollectionAssert.IsEmpty(words);
        }

        [Test]
        public void ADirectiveSpelledWithASpaceAfterTheHashIsNotOneForTheGrammar()
        {
            //The grammar of the preprocessor wants '#If' with nothing between the two tokens, as the VBE writes it; what is not a directive for it is not a block here.
            var words = Words("# If Mac Then", "Call Spelled", "#End If");

            CollectionAssert.DoesNotContain(words, "Spelled");
        }

        [Test]
        public void CommentsInsideABlockAreNotCountedButStringLiteralsAre()
        {
            var words = Words("#If Mac Then", "' Commented", "Call Real ' trailing", "s = \"=Formula(1)\"", "#End If");

            CollectionAssert.Contains(words, "Real");
            CollectionAssert.Contains(words, "Formula");
            CollectionAssert.DoesNotContain(words, "Commented");
            CollectionAssert.DoesNotContain(words, "trailing");
        }
    }
}
