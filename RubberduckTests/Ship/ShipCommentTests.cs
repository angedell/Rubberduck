using NUnit.Framework;
using Rubberduck.Ship;
using static RubberduckTests.Ship.ShipTestParser;

namespace RubberduckTests.Ship
{
    [TestFixture]
    [Category("Ship")]
    public class ShipCommentTests
    {
        private static ShipTextResult Strip(string text)
            => new ShipModuleText(ShipTestParser.Create()).StripComments("Module1", "project-1", text);

        [Test]
        public void ACommentOnALineOfItsOwnIsRemovedWithItsLine()
        {
            var result = Strip(Lines("Option Explicit", "' about Foo", "Public Sub Foo()", "End Sub"));

            Assert.IsTrue(result.Parsed, result.Problem);
            Assert.AreEqual(Lines("Option Explicit", "Public Sub Foo()", "End Sub"), result.Text);
            Assert.AreEqual(1, result.Removed);
        }

        [Test]
        public void ATrailingCommentIsRemovedWithTheSpaceBeforeIt()
        {
            var result = Strip(Lines("Public Sub Foo()", "    Dim x As Long ' counter", "End Sub"));

            Assert.AreEqual(Lines("Public Sub Foo()", "    Dim x As Long", "End Sub"), result.Text);
        }

        [Test]
        public void IndentedCommentLinesAreRemoved()
        {
            var result = Strip(Lines("Public Sub Foo()", "    ' one", "\t' two", "    Debug.Print 1", "End Sub"));

            Assert.AreEqual(Lines("Public Sub Foo()", "    Debug.Print 1", "End Sub"), result.Text);
        }

        [Test]
        public void RemCommentsAreRemoved()
        {
            var result = Strip(Lines("Public Sub Foo()", "    Rem the old way", "    Debug.Print 1: Rem trailing", "    REM", "End Sub"));

            Assert.AreEqual(Lines("Public Sub Foo()", "    Debug.Print 1:", "End Sub"), result.Text);
            Assert.AreEqual(3, result.Removed);
        }

        // The grammar reads "Rem" after a colon in a one-line If as a call of a procedure named Rem, so the parse tree has no comment there.
        [Test]
        public void ARemAfterAColonInAOneLineIfIsStillAComment()
        {
            var result = Strip(Lines("Public Sub Foo(ByVal x As Boolean)", "    If x Then Debug.Print 1: Rem why", "    If x Then Debug.Print 2 Else Debug.Print 3 ' either", "End Sub"));

            Assert.AreEqual(Lines("Public Sub Foo(ByVal x As Boolean)", "    If x Then Debug.Print 1:", "    If x Then Debug.Print 2 Else Debug.Print 3", "End Sub"), result.Text);
        }

        [Test]
        public void ACommentMarkInAStringIsNotAComment()
        {
            var code = Lines("Public Sub Foo()", "    Debug.Print \"it's here ' still\" ' but this is a comment", "    Debug.Print \"Rem not a comment\"", "End Sub");
            var result = Strip(code);

            Assert.AreEqual(Lines("Public Sub Foo()", "    Debug.Print \"it's here ' still\"", "    Debug.Print \"Rem not a comment\"", "End Sub"), result.Text);
        }

        [Test]
        public void AnAnnotationIsAComment()
        {
            var result = Strip(Lines("'@Folder(\"Core.Util\")", "'@IgnoreModule ProcedureNotUsed", "Option Explicit", "'@EntryPoint", "Public Sub Foo()", "End Sub"));

            Assert.AreEqual(Lines("Option Explicit", "Public Sub Foo()", "End Sub"), result.Text);
        }

        [Test]
        public void ACommentAfterAContinuedStatementIsRemoved()
        {
            var result = Strip(Lines("Public Sub Foo()", "    Debug.Print 1 + _", "        2 ' sum", "End Sub"));

            Assert.AreEqual(Lines("Public Sub Foo()", "    Debug.Print 1 + _", "        2", "End Sub"), result.Text);
        }

        // The grammar follows the language specification: a comment is part of a logical line, so a comment that ends in " _" goes on into the next line.
        [Test]
        public void ACommentThatEndsInAContinuationTakesTheNextLineWithIt()
        {
            var result = Strip(Lines("Public Sub Foo()", "    Debug.Print 1 ' note _", "    this line is still the comment", "    Debug.Print 2", "End Sub"));

            Assert.IsTrue(result.Parsed, result.Problem);
            Assert.AreEqual(Lines("Public Sub Foo()", "    Debug.Print 1", "    Debug.Print 2", "End Sub"), result.Text);
        }

        // "_ ' note" is not a line continuation: the mark has to be the last thing on the line. The grammar rejects it, as VBA does, so such a module ships unchanged.
        [Test]
        public void ACommentAfterTheContinuationMarkIsNotValidVbaSoTheModuleShipsUnchanged()
        {
            var code = Lines("Public Sub Foo()", "    Debug.Print 1 + _ ' not allowed", "        2", "End Sub");
            var result = Strip(code);

            Assert.IsFalse(result.Parsed);
            Assert.AreEqual(code, result.Text);
        }

        [Test]
        public void AttributesLabelsAndLineNumbersStay()
        {
            var code = Lines(
                "Attribute VB_Name = \"Module1\"",
                "Option Explicit",
                "Public Sub Foo()",
                "Attribute Foo.VB_Description = \"it's described\"",
                "10  Debug.Print 1 ' numbered",
                "Again: ' label",
                "20  Debug.Print 2",
                "    GoTo Again",
                "End Sub");
            var result = Strip(code);

            Assert.AreEqual(Lines(
                "Attribute VB_Name = \"Module1\"",
                "Option Explicit",
                "Public Sub Foo()",
                "Attribute Foo.VB_Description = \"it's described\"",
                "10  Debug.Print 1",
                "Again:",
                "20  Debug.Print 2",
                "    GoTo Again",
                "End Sub"), result.Text);
        }

        [Test]
        public void ABlankLineStaysABlankLine()
        {
            var result = Strip(Lines("Option Explicit", "", "' a", "", "Public Sub Foo()", "End Sub"));

            Assert.AreEqual(Lines("Option Explicit", "", "", "Public Sub Foo()", "End Sub"), result.Text);
        }

        [Test]
        public void ACommentOnTheLastLineWithoutANewlineIsRemoved()
        {
            var result = Strip("Option Explicit\r\n' the end");

            Assert.AreEqual("Option Explicit\r\n", result.Text);
        }

        [Test]
        public void AClassHeaderIsLeftAlone()
        {
            var code = Lines("VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "END", "Attribute VB_Name = \"Thing\"", "Option Explicit", "' about Thing", "Public x As Long");
            var result = Strip(code);

            Assert.AreEqual(Lines("VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "END", "Attribute VB_Name = \"Thing\"", "Option Explicit", "Public x As Long"), result.Text);
        }

        [Test]
        public void AModuleWithoutCommentsIsReturnedAsItWas()
        {
            var code = Lines("Option Explicit", "Public Sub Foo()", "End Sub");
            var result = Strip(code);

            Assert.AreEqual(code, result.Text);
            Assert.AreEqual(0, result.Removed);
        }

        [Test]
        public void AModuleThatDoesNotParseIsReturnedAsItWasAndSaysWhy()
        {
            var code = Lines("Public Sub Foo(", "    ' a comment", "End Sub");
            var result = Strip(code);

            Assert.IsFalse(result.Parsed);
            Assert.AreEqual(code, result.Text);
            Assert.IsNotEmpty(result.Problem);
        }
    }
}
