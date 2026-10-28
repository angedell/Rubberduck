using System.Windows;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveFolder;

namespace Rubberduck.UI.Refactorings.MoveFolder
{
    public partial class MoveMultipleFoldersView : IRefactoringView<MoveMultipleFoldersModel>
    {
        public MoveMultipleFoldersView()
        {
            InitializeComponent();

            Loaded += AfterLoadHandler;
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            UI.CodeExplorer.FolderPickerDialog.Browse(this);
            MoveToFolderTextBox.Focus();
        }

        private void AfterLoadHandler(object sender, RoutedEventArgs e)
        {
            MoveToFolderTextBox.Focus();
            MoveToFolderTextBox.SelectAll();
            Loaded -= AfterLoadHandler;
        }
    }
}
