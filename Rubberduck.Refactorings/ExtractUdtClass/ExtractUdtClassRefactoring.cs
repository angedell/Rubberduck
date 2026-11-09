using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.ExtractUdtClass
{
    public class ExtractUdtClassRefactoring : InteractiveRefactoringBase<ExtractUdtClassModel>
    {
        private readonly ExtractUdtClassRefactoringAction _refactoringAction;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IExtractInterfaceConflictFinderFactory _conflictFinderFactory;

        public ExtractUdtClassRefactoring(
            ExtractUdtClassRefactoringAction refactoringAction,
            IDeclarationFinderProvider declarationFinderProvider,
            RefactoringUserInteraction<IExtractUdtClassPresenter, ExtractUdtClassModel> userInteraction,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            IExtractInterfaceConflictFinderFactory conflictFinderFactory)
            : base(selectionProvider, userInteraction)
        {
            _refactoringAction = refactoringAction;
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _conflictFinderFactory = conflictFinderFactory;
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            return UserDefinedTypeOf(_selectedDeclarationProvider.SelectedDeclaration(targetSelection));
        }

        /// <summary>The type itself, or the type that a selected member belongs to; null for anything else.</summary>
        public static Declaration UserDefinedTypeOf(Declaration selected)
        {
            switch (selected?.DeclarationType)
            {
                case DeclarationType.UserDefinedType:
                    return selected;
                case DeclarationType.UserDefinedTypeMember:
                    return selected.ParentDeclaration;
                default:
                    return null;
            }
        }

        /// <summary>The model the dialog starts from; for a caller that answers the dialog's questions itself.</summary>
        public ExtractUdtClassModel CreateModel(Declaration target) => InitializeModel(target);

        protected override ExtractUdtClassModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (target.DeclarationType != DeclarationType.UserDefinedType)
            {
                throw new InvalidDeclarationTypeException(target);
            }

            var conflictFinder = _conflictFinderFactory.Create(_declarationFinderProvider, target.ProjectId);
            var className = ExtractUdtClassModel.SuggestedClassName(target.IdentifierName);
            if (conflictFinder.IsConflictingModuleName(className))
            {
                className = conflictFinder.GenerateNoConflictModuleName(className);
            }

            var members = ExtractUdtClassModel.MembersOf(target, _declarationFinderProvider);
            return new ExtractUdtClassModel(target, members, conflictFinder) { ClassName = className };
        }

        protected override void RefactorImpl(ExtractUdtClassModel model)
        {
            _refactoringAction.Refactor(model);
        }
    }
}
