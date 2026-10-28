using System.Windows;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveToFolder;

namespace Rubberduck.UI.Refactorings.MoveToFolder
{
    public partial class MoveMultipleToFolderView : IRefactoringView<MoveMultipleToFolderModel>
    {
        public MoveMultipleToFolderView()
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
