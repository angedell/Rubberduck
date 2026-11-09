using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractUdtClass;

namespace Rubberduck.UI.Refactorings.ExtractUdtClass
{
    internal class ExtractUdtClassPresenter : RefactoringPresenterBase<ExtractUdtClassModel>, IExtractUdtClassPresenter
    {
        private static readonly DialogData DialogData = DialogData.Create(RefactoringsUI.ExtractUdtClass_Caption, 430, 420);

        public ExtractUdtClassPresenter(ExtractUdtClassModel model,
            IRefactoringDialogFactory dialogFactory) : base(DialogData, model, dialogFactory) { }
    }
}
