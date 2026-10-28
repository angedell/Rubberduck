using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>Picks one of the open projects: the target of a copy or a move.</summary>
    public class ProjectPickerDialog : Form
    {
        private readonly ListBox _list;
        private readonly Button _ok;

        public ProjectPickerDialog(string title, IEnumerable<KeyValuePair<string, string>> projects)
        {
            Text = title;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(420, 280);
            MinimumSize = new Size(300, 200);
            ShowIcon = false;
            ShowInTaskbar = false;

            _list = new ListBox
            {
                Dock = DockStyle.Fill,
                DisplayMember = nameof(KeyValuePair<string, string>.Value),
                IntegralHeight = false
            };
            foreach (var project in projects)
            {
                _list.Items.Add(project);
            }

            _list.DoubleClick += (sender, args) =>
            {
                if (_list.SelectedItem != null)
                {
                    _ok.PerformClick();
                }
            };
            _list.SelectedIndexChanged += (sender, args) => _ok.Enabled = _list.SelectedItem != null;

            _ok = new Button { Text = CodeExplorerUI.ProjectPicker_Ok, DialogResult = DialogResult.OK, Enabled = false, AutoSize = true, MinimumSize = new Size(80, 26) };
            var cancel = new Button { Text = CodeExplorerUI.ProjectPicker_Cancel, DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(80, 26) };
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(8) };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(_ok);

            var prompt = new Label { Dock = DockStyle.Top, Text = CodeExplorerUI.ProjectPicker_Prompt, AutoSize = true, Padding = new Padding(8, 8, 8, 4) };

            Controls.Add(_list);
            Controls.Add(prompt);
            Controls.Add(buttons);
            AcceptButton = _ok;
            CancelButton = cancel;

            if (_list.Items.Count > 0)
            {
                _list.SelectedIndex = 0;
            }
        }

        /// <summary>The id of the project chosen.</summary>
        public string SelectedProjectId => _list.SelectedItem is KeyValuePair<string, string> item ? item.Key : null;

        public static string Pick(string title, IEnumerable<KeyValuePair<string, string>> projects)
        {
            using (var dialog = new ProjectPickerDialog(title, projects))
            {
                return dialog.ShowDialog() == DialogResult.OK ? dialog.SelectedProjectId : null;
            }
        }
    }
}
