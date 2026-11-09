using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.ExtractInterface;

namespace Rubberduck.Refactorings.ExtractUdtClass
{
    public enum ExtractUdtClassStyle
    {
        /// <summary>Get/Let/Set properties and a <c>Self</c> property.</summary>
        Plain,
        /// <summary>The plain members, plus a predeclared instance with a <c>Create</c> factory and chainable <c>WithX</c> methods.</summary>
        Fluent
    }

    public class ExtractUdtClassModel : IRefactoringModel
    {
        public ExtractUdtClassModel(Declaration target, IReadOnlyList<UdtMemberInfo> members, IExtractInterfaceConflictFinder conflictFinder)
        {
            Target = target;
            Members = members;
            ConflictFinder = conflictFinder;
        }

        /// <summary>The user defined type to build a class for. It is left as it is.</summary>
        public Declaration Target { get; }
        public IReadOnlyList<UdtMemberInfo> Members { get; }
        public IExtractInterfaceConflictFinder ConflictFinder { get; }

        public string ClassName { get; set; }
        public ExtractUdtClassStyle Style { get; set; } = ExtractUdtClassStyle.Plain;

        /// <summary>Member names that collide with what the class adds, for the chosen style. Empty if there are none.</summary>
        public IReadOnlyList<string> Clashes => ExtractUdtClassCodeBuilder.Clashes(Members, Style);

        /// <summary>Why the class cannot have the name it has now, or null if it can.</summary>
        public string ClassNameProblem
        {
            get
            {
                if (!VBAIdentifierValidator.IsValidIdentifier(ClassName, DeclarationType.ClassModule))
                {
                    return $"'{ClassName}' is not a valid class name.";
                }

                return ConflictFinder.IsConflictingModuleName(ClassName)
                    ? $"The project already has a module or public type named '{ClassName}'."
                    : null;
            }
        }

        /// <summary>Everything that stops the class from being written as asked: the name, an empty type, clashing members.</summary>
        public IReadOnlyList<string> Problems
        {
            get
            {
                var problems = new List<string>();
                if (ClassNameProblem != null)
                {
                    problems.Add(ClassNameProblem);
                }
                if (Members.Count == 0)
                {
                    problems.Add($"The type '{Target.IdentifierName}' has no members.");
                }
                problems.AddRange(Clashes);
                return problems;
            }
        }

        public static IReadOnlyList<UdtMemberInfo> MembersOf(Declaration udt, IDeclarationFinderProvider declarationFinderProvider)
        {
            return declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.UserDefinedTypeMember)
                .Where(member => udt.Equals(member.ParentDeclaration))
                .OrderBy(member => member.Selection.StartLine)
                .ThenBy(member => member.Selection.StartColumn)
                .Select(member => UdtMemberInfo.From(member))
                .ToList();
        }

        /// <summary>The class name suggested for a type: a leading T or UDT prefix is dropped (TPerson becomes Person).</summary>
        public static string SuggestedClassName(string typeName)
        {
            foreach (var prefix in new[] { "UDT", "T" })
            {
                if (typeName.Length > prefix.Length
                    && typeName.StartsWith(prefix)
                    && (char.IsUpper(typeName[prefix.Length]) || typeName[prefix.Length] == '_'))
                {
                    return typeName.Substring(prefix.Length).TrimStart('_');
                }
            }

            return typeName;
        }
    }
}
