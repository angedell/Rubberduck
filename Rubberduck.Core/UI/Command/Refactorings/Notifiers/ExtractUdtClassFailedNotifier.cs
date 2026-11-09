using Rubberduck.Interaction;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class ExtractUdtClassFailedNotifier : RefactoringFailureNotifierBase
    {
        public ExtractUdtClassFailedNotifier(IMessageBox messageBox)
            : base(messageBox)
        { }

        protected override string Caption => RefactoringsUI.ExtractUdtClass_Caption;
    }
}
