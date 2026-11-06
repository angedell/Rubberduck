using System.Text;
using System.Text.RegularExpressions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Refactorings.SplitParameters
{
    /// <summary>
    /// Puts each parameter of a member's signature on a line of its own, with leading commas, so that adding, removing
    /// or reordering a parameter changes as few lines as possible in a diff.
    /// </summary>
    public class SplitParametersRefactoringAction : CodeOnlyRefactoringActionBase<SplitParametersModel>
    {
        private const int CommaIndent = 5;
        private const int ParameterIndent = 7;

        private static readonly Regex LineContinuationRegex = new Regex(@"\s*_\r?\n\s*", RegexOptions.Compiled);

        public SplitParametersRefactoringAction(IRewritingManager rewritingManager)
            : base(rewritingManager)
        {}

        public override void Refactor(SplitParametersModel model, IRewriteSession rewriteSession)
        {
            foreach (var target in model.Targets)
            {
                var text = SplitText(target);
                if (text == null)
                {
                    continue;
                }

                var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);
                rewriter.Replace(target.Context.GetChild<VBAParser.ArgListContext>(), text);
            }
        }

        /// <summary>True if the member has at least two parameters and they are not laid out one to a line already.</summary>
        public static bool NeedsSplit(Declaration member)
        {
            var text = SplitText(member);
            return text != null && text != member.Context.GetChild<VBAParser.ArgListContext>().GetText();
        }

        /// <summary>The parameter list as it should read, from the opening parenthesis to the closing one, or null if the member has fewer than two parameters.</summary>
        private static string SplitText(Declaration member)
        {
            var parameters = member.Context.GetChild<VBAParser.ArgListContext>()?.arg();
            if (parameters == null || parameters.Length < 2)
            {
                return null;
            }

            var baseIndent = member.Context.Start.Column;
            var text = new StringBuilder("( _");
            for (var i = 0; i < parameters.Length; i++)
            {
                var parameter = LineContinuationRegex.Replace(parameters[i].GetText(), " ").Trim();
                text.AppendLine()
                    .Append(new string(' ', baseIndent + (i == 0 ? ParameterIndent : CommaIndent)))
                    .Append(i == 0 ? string.Empty : ", ")
                    .Append(parameter)
                    .Append(" _");
            }
            text.AppendLine()
                .Append(new string(' ', baseIndent + ParameterIndent))
                .Append(")");

            return text.ToString();
        }
    }
}
