using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.SplitParameters;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class RefactorSplitParametersCommand : RefactorCodePaneCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public RefactorSplitParametersCommand(
            SplitParametersRefactoring refactoring,
            SplitParametersFailedNotifier failedNotifier,
            RubberduckParserState state,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider)
            : base(refactoring, failedNotifier, selectionProvider, state)
        {
            _state = state;
            _selectedDeclarationProvider = selectedDeclarationProvider;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var member = GetTarget();
            if (member == null || _state.IsNewOrModified(member.QualifiedModuleName))
            {
                return false;
            }

            return _state.DeclarationFinder
                .UserDeclarations(DeclarationType.Parameter)
                .Count(item => member.Equals(item.ParentScopeDeclaration)) > 1;
        }

        private Declaration GetTarget()
        {
            var selectedDeclaration = _selectedDeclarationProvider.SelectedDeclaration();
            if (selectedDeclaration == null || SplitParametersRefactoring.ValidDeclarationTypes.Contains(selectedDeclaration.DeclarationType))
            {
                return selectedDeclaration;
            }

            return selectedDeclaration.DeclarationType == DeclarationType.Parameter
                ? _selectedDeclarationProvider.SelectedMember()
                : null;
        }
    }
}
