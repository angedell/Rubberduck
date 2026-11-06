using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.SplitParameters;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.Command.Refactorings
{
    /// <summary>Splits the parameters of every member in the project of the active code pane, in one go.</summary>
    [ComVisible(false)]
    public class RefactorSplitParametersInProjectCommand : RefactorCodePaneCommandBase
    {
        private readonly SplitParametersRefactoring _refactoring;
        private readonly RubberduckParserState _state;

        public RefactorSplitParametersInProjectCommand(
            SplitParametersRefactoring refactoring,
            SplitParametersFailedNotifier failedNotifier,
            RubberduckParserState state,
            ISelectionProvider selectionProvider)
            : base(refactoring, failedNotifier, selectionProvider, state)
        {
            _refactoring = refactoring;
            _state = state;

            AddToCanExecuteEvaluation(SpecializedEvaluateCanExecute);
        }

        private bool SpecializedEvaluateCanExecute(object parameter)
        {
            var selection = SelectionProvider.ActiveSelection();
            if (!selection.HasValue || _state.IsNewOrModified(selection.Value.QualifiedName))
            {
                return false;
            }

            return _refactoring.MembersToSplit(selection.Value.QualifiedName.ProjectId).Any();
        }

        protected override void OnExecute(object parameter)
        {
            var selection = SelectionProvider.ActiveSelection();
            if (!selection.HasValue)
            {
                return;
            }

            try
            {
                _refactoring.Refactor(_refactoring.MembersToSplit(selection.Value.QualifiedName.ProjectId));
            }
            catch (RefactoringAbortedException)
            {}
            catch (RefactoringException exception)
            {
                FailureNotifier.Notify(exception);
            }
        }
    }
}
