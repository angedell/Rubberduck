using System;
using System.Drawing;
using System.Windows.Forms;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer
{
    public class BulkRemoveOptions
    {
        public bool Export { get; set; } = true;
        public bool PreserveFolderStructure { get; set; } = true;
        public string ExportFolder { get; set; } = string.Empty;
    }

    /// <summary>
    /// The one dialog of the bulk removal: whether to export first, how, and where to.
    /// </summary>
    public sealed class BulkRemoveDialog : Form
    {
        private readonly Func<string, string> _browse;

        private readonly RadioButton _exportFirst;
        private readonly RadioButton _noExport;
        private readonly GroupBox _layoutGroup;
        private readonly RadioButton _preserveStructure;
        private readonly RadioButton _singleFolder;
        private readonly TextBox _folder;
        private readonly Button _browseButton;
        private readonly Button _remove;

        /// <param name="moduleCount">The number of modules about to be removed.</param>
        /// <param name="initial">The options to start with, e.g. those remembered from last time.</param>
        /// <param name="browse">Shows a folder browser starting at the given folder; returns the chosen folder or null.</param>
        public BulkRemoveDialog(int moduleCount, BulkRemoveOptions initial, Func<string, string> browse)
        {
            _browse = browse;

            Text = CodeExplorerUI.BulkRemove_Caption;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(480, 330);

            var summary = new Label
            {
                Text = string.Format(CodeExplorerUI.BulkRemove_Summary, moduleCount),
                Location = new Point(12, 12),
                Size = new Size(456, 20),
                Font = new Font(Font, FontStyle.Bold)
            };

            _exportFirst = new RadioButton { Text = CodeExplorerUI.BulkRemove_ExportFirst, Location = new Point(12, 22), Size = new Size(420, 20), Checked = initial.Export };
            _noExport = new RadioButton { Text = CodeExplorerUI.BulkRemove_NoExport, Location = new Point(12, 46), Size = new Size(420, 20), Checked = !initial.Export };
            var beforeGroup = new GroupBox
            {
                Text = CodeExplorerUI.BulkRemove_BeforeRemoving,
                Location = new Point(12, 44),
                Size = new Size(456, 76)
            };
            beforeGroup.Controls.AddRange(new Control[] { _exportFirst, _noExport });

            _preserveStructure = new RadioButton { Text = CodeExplorerUI.BulkRemove_PreserveStructure, Location = new Point(12, 22), Size = new Size(420, 20), Checked = initial.PreserveFolderStructure };
            _singleFolder = new RadioButton { Text = CodeExplorerUI.BulkRemove_SingleFolder, Location = new Point(12, 46), Size = new Size(420, 20), Checked = !initial.PreserveFolderStructure };
            _layoutGroup = new GroupBox
            {
                Text = CodeExplorerUI.BulkRemove_Layout,
                Location = new Point(12, 128),
                Size = new Size(456, 76)
            };
            _layoutGroup.Controls.AddRange(new Control[] { _preserveStructure, _singleFolder });

            var folderLabel = new Label
            {
                Text = CodeExplorerUI.BulkRemove_FolderLabel,
                Location = new Point(12, 214),
                Size = new Size(456, 18)
            };
            _folder = new TextBox { Text = initial.ExportFolder ?? string.Empty, Location = new Point(12, 236), Size = new Size(366, 23) };
            _browseButton = new Button { Text = CodeExplorerUI.BulkRemove_Browse, Location = new Point(386, 235), Size = new Size(82, 25) };
            _browseButton.Click += (sender, e) => OnBrowse();

            _remove = new Button { Text = CodeExplorerUI.BulkRemove_RemoveButton, Location = new Point(298, 286), Size = new Size(82, 28) };
            _remove.Click += (sender, e) => OnRemove();
            var cancel = new Button { Text = CodeExplorerUI.BulkRemove_Cancel, Location = new Point(386, 286), Size = new Size(82, 28), DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[] { summary, beforeGroup, _layoutGroup, folderLabel, _folder, _browseButton, _remove, cancel });
            AcceptButton = _remove;
            CancelButton = cancel;

            _exportFirst.CheckedChanged += (sender, e) => UpdateExportControls();
            UpdateExportControls();
        }

        public BulkRemoveOptions Result => new BulkRemoveOptions
        {
            Export = _exportFirst.Checked,
            PreserveFolderStructure = _preserveStructure.Checked,
            ExportFolder = _folder.Text.Trim()
        };

        private void UpdateExportControls()
        {
            var exporting = _exportFirst.Checked;
            _layoutGroup.Enabled = exporting;
            _folder.Enabled = exporting;
            _browseButton.Enabled = exporting;
        }

        private void OnBrowse()
        {
            var selected = _browse(_folder.Text.Trim());
            if (!string.IsNullOrEmpty(selected))
            {
                _folder.Text = selected;
            }
        }

        private void OnRemove()
        {
            if (_exportFirst.Checked && !IsUsableFolder(_folder.Text.Trim()))
            {
                MessageBox.Show(this, CodeExplorerUI.BulkRemove_InvalidFolder, CodeExplorerUI.BulkRemove_Caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _folder.Focus();
                return;
            }

            DialogResult = DialogResult.OK;
        }

        private static bool IsUsableFolder(string path)
        {
            try
            {
                return path.Length > 0 && System.IO.Path.IsPathRooted(path) && System.IO.Path.GetFullPath(path).Length > 0;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
