using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Git;

namespace Rubberduck.UI.CodeExplorer.GitClient
{
    /// <summary>Shows what changed since the last commit and asks for the commit message.</summary>
    public sealed class GitCommitDialog : Form
    {
        private readonly TextBox _message;

        private GitCommitDialog(string projectName, IReadOnlyList<GitChange> changes)
        {
            Text = "Git: commit '" + projectName + "'";
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(560, 420);

            Controls.Add(new Label { Text = $"{changes.Count} changed file(s) since the last commit:", Location = new Point(12, 10), Size = new Size(536, 18) });

            var list = new ListView { View = View.Details, FullRowSelect = true, HeaderStyle = ColumnHeaderStyle.None, Location = new Point(12, 32), Size = new Size(536, 190) };
            list.Columns.Add(string.Empty, 80);
            list.Columns.Add(string.Empty, 430);
            foreach (var change in changes.OrderBy(c => c.Path))
            {
                list.Items.Add(new ListViewItem(new[] { change.Kind.ToString(), change.Path }));
            }

            Controls.Add(list);

            Controls.Add(new Label { Text = "Message:", Location = new Point(12, 232), Size = new Size(536, 18) });
            _message = new TextBox { Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical, Location = new Point(12, 254), Size = new Size(536, 110) };
            Controls.Add(_message);

            var commit = new Button { Text = "Commit", DialogResult = DialogResult.OK, Size = new Size(90, 28), Location = new Point(ClientSize.Width - 12 - 90 - 8 - 90, ClientSize.Height - 40) };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Size = new Size(90, 28), Location = new Point(ClientSize.Width - 12 - 90, ClientSize.Height - 40) };
            Controls.Add(commit);
            Controls.Add(cancel);
            AcceptButton = commit;
            CancelButton = cancel;
            Shown += (sender, args) => _message.Focus();
        }

        /// <returns>The message, or null if the user cancelled. An empty message is not accepted.</returns>
        public static string Ask(string projectName, IReadOnlyList<GitChange> changes)
        {
            using (var dialog = new GitCommitDialog(projectName, changes))
            {
                while (dialog.ShowDialog() == DialogResult.OK)
                {
                    var message = dialog._message.Text.Trim();
                    if (message.Length > 0)
                    {
                        return message;
                    }

                    MessageBox.Show(dialog, "Write a message that says what changed.", dialog.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                return null;
            }
        }
    }

    /// <summary>The commits made for a project, newest first.</summary>
    public sealed class GitHistoryDialog : Form
    {
        private GitHistoryDialog(string projectName, IReadOnlyList<GitCommit> commits)
        {
            Text = "Git: history of '" + projectName + "'";
            Font = new Font("Segoe UI", 9f);
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(720, 420);

            var list = new ListView { View = View.Details, FullRowSelect = true, Dock = DockStyle.Fill };
            list.Columns.Add("Commit", 70);
            list.Columns.Add("Date", 130);
            list.Columns.Add("Author", 130);
            list.Columns.Add("Message", 370);
            foreach (var commit in commits)
            {
                list.Items.Add(new ListViewItem(new[] { commit.Id.Substring(0, 7), commit.Date.LocalDateTime.ToString("yyyy-MM-dd HH:mm"), commit.Author, commit.Subject }));
            }

            Controls.Add(list);
            var close = new Button { Text = "Close", DialogResult = DialogResult.Cancel, Dock = DockStyle.Bottom, Height = 30 };
            Controls.Add(close);
            CancelButton = close;
        }

        public static void Show(string projectName, IReadOnlyList<GitCommit> commits)
        {
            using (var dialog = new GitHistoryDialog(projectName, commits))
            {
                dialog.ShowDialog();
            }
        }
    }
}
