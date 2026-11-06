using Rubberduck.Interaction;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class SplitParametersFailedNotifier : RefactoringFailureNotifierBase
    {
        public SplitParametersFailedNotifier(IMessageBox messageBox)
            : base(messageBox)
        { }

        protected override string Caption => RefactoringsUI.SplitParameters_Caption;
    }
}
