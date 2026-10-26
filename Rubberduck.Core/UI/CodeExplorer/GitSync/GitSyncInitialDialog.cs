using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// A module that exists in the project and in the folder, differs, and has no stamp or sync history to tell which side is newer.
    /// </summary>
    public class InitialConflict
    {
        public InitialConflict(string moduleName, DateTime fileModified)
        {
            ModuleName = moduleName;
            FileModified = fileModified;
        }

        public string ModuleName { get; }

        /// <summary>When the file was last written, in local time. The project does not know when its module was last modified.</summary>
        public DateTime FileModified { get; }
    }

    public enum InitialChoice
    {
        Cancel,
        /// <summary>The files in the folder win: they are imported into the project.</summary>
        Folder,
        /// <summary>The modules in the project win: they are exported over the files.</summary>
        Project
    }

    /// <summary>
    /// Asks who wins the first time modules that exist on both sides are synced.
    /// </summary>
    public sealed class GitSyncInitialDialog : Form
    {
        private readonly RadioButton _folderWins;

        private GitSyncInitialDialog(IReadOnlyList<InitialConflict> conflicts)
        {
            Text = CodeExplorerUI.GitSync_DialogTitle;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(520, 400);

            var explanation = new Label
            {
                Text = string.Format(CodeExplorerUI.GitSync_InitialExplanation, conflicts.Count),
                Location = new Point(12, 12),
                Size = new Size(496, 60)
            };

            var list = new ListView
            {
                View = View.Details,
                FullRowSelect = true,
                Location = new Point(12, 78),
                Size = new Size(496, 190)
            };
            list.Columns.Add(CodeExplorerUI.GitSync_InitialModuleColumn, 220);
            list.Columns.Add(CodeExplorerUI.GitSync_InitialFileModifiedColumn, 250);
            foreach (var conflict in conflicts)
            {
                var item = new ListViewItem(conflict.ModuleName);
                item.SubItems.Add(conflict.FileModified.ToString("g"));
                list.Items.Add(item);
            }

            _folderWins = new RadioButton { Text = CodeExplorerUI.GitSync_InitialFolderWins, Location = new Point(12, 280), Size = new Size(496, 22), Checked = true };
            var projectWins = new RadioButton { Text = CodeExplorerUI.GitSync_InitialProjectWins, Location = new Point(12, 306), Size = new Size(496, 22) };

            var ok = new Button { Text = CodeExplorerUI.GitSync_InitialContinue, Location = new Point(334, 356), Size = new Size(86, 28), DialogResult = DialogResult.OK };
            var cancel = new Button { Text = CodeExplorerUI.BulkRemove_Cancel, Location = new Point(426, 356), Size = new Size(82, 28), DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[] { explanation, list, _folderWins, projectWins, ok, cancel });
            AcceptButton = ok;
            CancelButton = cancel;
        }

        /// <summary>
        /// Shows the dialog; the answer applies to all the listed modules.
        /// </summary>
        public static InitialChoice Ask(IReadOnlyList<InitialConflict> conflicts)
        {
            using (var dialog = new GitSyncInitialDialog(conflicts))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return InitialChoice.Cancel;
                }

                return dialog._folderWins.Checked ? InitialChoice.Folder : InitialChoice.Project;
            }
        }
    }
}
