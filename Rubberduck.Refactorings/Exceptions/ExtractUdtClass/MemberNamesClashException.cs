using System.Collections.Generic;

namespace Rubberduck.Refactorings.Exceptions.ExtractUdtClass
{
    /// <summary>A member of the type has a name that the generated class needs for something of its own.</summary>
    public class MemberNamesClashException : RefactoringException
    {
        public MemberNamesClashException(IReadOnlyList<string> problems)
        {
            Problems = problems;
        }

        public IReadOnlyList<string> Problems { get; }

        public override string Message => string.Join(" ", Problems);
    }
}
