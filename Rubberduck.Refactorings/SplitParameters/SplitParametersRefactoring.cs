using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.SplitParameters
{
    public class SplitParametersRefactoring : RefactoringBase
    {
        private readonly SplitParametersRefactoringAction _refactoringAction;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public SplitParametersRefactoring(
            SplitParametersRefactoringAction refactoringAction,
            IDeclarationFinderProvider declarationFinderProvider,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider)
            : base(selectionProvider)
        {
            _refactoringAction = refactoringAction;
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var selectedDeclaration = _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
            if (selectedDeclaration == null || ValidDeclarationTypes.Contains(selectedDeclaration.DeclarationType))
            {
                return selectedDeclaration;
            }

            return selectedDeclaration.DeclarationType == DeclarationType.Parameter
                ? _selectedDeclarationProvider.SelectedMember(targetSelection)
                : null;
        }

        public override void Refactor(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (!ValidDeclarationTypes.Contains(target.DeclarationType))
            {
                throw new InvalidDeclarationTypeException(target);
            }

            _refactoringAction.Refactor(new SplitParametersModel(target));
        }

        /// <summary>The members of the project whose parameter lists are not laid out one to a line yet, in module order.</summary>
        public IReadOnlyList<Declaration> MembersToSplit(string projectId)
        {
            return ValidDeclarationTypes
                .SelectMany(type => _declarationFinderProvider.DeclarationFinder.UserDeclarations(type))
                .Where(member => member.ProjectId == projectId && SplitParametersRefactoringAction.NeedsSplit(member))
                .OrderBy(member => member.QualifiedModuleName.ComponentName)
                .ThenBy(member => member.Selection.StartLine)
                .ToList();
        }

        public int ParameterCount(Declaration member)
            => member.Context.GetChild<Rubberduck.Parsing.Grammar.VBAParser.ArgListContext>()?.arg().Length ?? 0;

        /// <summary>Splits all the given members' parameter lists in one go.</summary>
        public void Refactor(IReadOnlyCollection<Declaration> members)
        {
            if (members.Count > 0)
            {
                _refactoringAction.Refactor(new SplitParametersModel(members));
            }
        }

        public static readonly DeclarationType[] ValidDeclarationTypes =
        {
            DeclarationType.Event,
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };
    }
}
