using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.SplitParameters;

namespace RubberduckTests.Refactoring.SplitParameters
{
    [TestFixture]
    public class SplitParametersRefactoringActionTests : RefactoringActionTestBase<SplitParametersModel>
    {
        [Test]
        [Category("Refactorings")]
        [Category("Split Parameters")]
        public void SplitsEachParameterOntoItsOwnLine()
        {
            const string inputCode =
@"Public Function Create(ByVal appKey As String, ByVal appSecret As String, ByVal refreshToken As String, _
                       ByVal RootPath As String, Optional ByVal timeoutMs As Long = 30000) As DropboxApiStore
End Function";

            const string expected =
@"Public Function Create( _
       ByVal appKey As String _
     , ByVal appSecret As String _
     , ByVal refreshToken As String _
     , ByVal RootPath As String _
     , Optional ByVal timeoutMs As Long = 30000 _
       ) As DropboxApiStore
End Function";

            Assert.AreEqual(Lf(expected), Lf(RefactoredCode(inputCode, "Create")));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Split Parameters")]
        public void ResultIsStableWhenAlreadySplit()
        {
            const string inputCode =
@"Private Sub Foo( _
       ByVal a As Long _
     , ByVal b As Long _
       )
End Sub";

            Assert.AreEqual(Lf(inputCode), Lf(RefactoredCode(inputCode, "Foo")));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Split Parameters")]
        public void SingleParameterIsLeftAlone()
        {
            const string inputCode =
@"Private Sub Foo(ByVal a As Long)
End Sub";

            Assert.AreEqual(Lf(inputCode), Lf(RefactoredCode(inputCode, "Foo")));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Split Parameters")]
        public void SplitsEveryGivenMemberInOneGo()
        {
            const string inputCode =
@"Private Sub One(ByVal a As Long, ByVal b As Long)
End Sub

Private Sub Two(ByVal c As Long, ByVal d As Long)
End Sub";

            var actual = Lf(base.RefactoredCode(Crlf(inputCode), state => new SplitParametersModel(
                state.DeclarationFinder.UserDeclarations(DeclarationType.Procedure))));

            StringAssert.Contains("One( _\n       ByVal a As Long _\n     , ByVal b As Long _\n       )", actual);
            StringAssert.Contains("Two( _\n       ByVal c As Long _\n     , ByVal d As Long _\n       )", actual);
        }

        // The code pane uses CRLF; the string literals above use whatever line ends the source file has.
        private static string Lf(string code) => code.Replace("\r\n", "\n");

        private static string Crlf(string code) => Lf(code).Replace("\n", "\r\n");

        private string RefactoredCode(string inputCode, string memberName)
        {
            return base.RefactoredCode(Crlf(inputCode), state => new SplitParametersModel(
                state.DeclarationFinder.UserDeclarations(DeclarationType.Function)
                    .Concat(state.DeclarationFinder.UserDeclarations(DeclarationType.Procedure))
                    .Single(d => d.IdentifierName == memberName)));
        }

        protected override IRefactoringAction<SplitParametersModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            return new SplitParametersRefactoringAction(rewritingManager);
        }
    }
}
