using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.SplitParameters
{
    public class SplitParametersModel : IRefactoringModel
    {
        public SplitParametersModel(Declaration target)
            : this(new[] { target })
        {}

        public SplitParametersModel(IEnumerable<Declaration> targets)
        {
            Targets = targets.ToList();
        }

        /// <summary>The members whose parameter lists to split, in one pass: a whole project at once is a single rewrite.</summary>
        public IReadOnlyList<Declaration> Targets { get; }
    }
}
