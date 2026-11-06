using System;
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

        private GitCommitDialog(string projectName, IReadOnlyList<GitChange> changes, Func<IWin32Window, string> suggest)
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
            if (suggest != null)
            {
                var ask = new Button { Text = "Suggest message (AI)", Size = new Size(150, 28), Location = new Point(12, ClientSize.Height - 40) };
                ask.Click += (sender, args) =>
                {
                    var suggestion = suggest(this);
                    if (!string.IsNullOrWhiteSpace(suggestion))
                    {
                        _message.Text = suggestion;
                    }
                };
                Controls.Add(ask);
            }

            AcceptButton = commit;
            CancelButton = cancel;
            Shown += (sender, args) => _message.Focus();
        }

        /// <returns>The message, or null if the user cancelled. An empty message is not accepted.</returns>
        /// <param name="suggest">Asked for a message when the user presses the AI button, which is only there if this is given. Returns null for none.</param>
        public static string Ask(string projectName, IReadOnlyList<GitChange> changes, Func<IWin32Window, string> suggest = null)
        {
            using (var dialog = new GitCommitDialog(projectName, changes, suggest))
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

    /// <summary>The commits in which a module changed; the user picks the one to go back to.</summary>
    public sealed class GitRestoreDialog : Form
    {
        private readonly ListView _list;
        private readonly IReadOnlyList<GitCommit> _commits;

        private GitRestoreDialog(string moduleName, IReadOnlyList<GitCommit> commits)
        {
            _commits = commits;
            Text = "Git: restore '" + moduleName + "'";
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(640, 360);

            Controls.Add(new Label { Text = "Choose the version to put back. The module as it is now is backed up first.", Location = new Point(12, 10), Size = new Size(616, 18) });

            _list = new ListView { View = View.Details, FullRowSelect = true, MultiSelect = false, Location = new Point(12, 32), Size = new Size(616, 270) };
            _list.Columns.Add("Commit", 70);
            _list.Columns.Add("Date", 130);
            _list.Columns.Add("Author", 120);
            _list.Columns.Add("Message", 270);
            foreach (var commit in commits)
            {
                _list.Items.Add(new ListViewItem(new[] { commit.Id.Substring(0, 7), commit.Date.LocalDateTime.ToString("yyyy-MM-dd HH:mm"), commit.Author, commit.Subject }));
            }

            _list.Items[0].Selected = true;
            Controls.Add(_list);

            var restore = new Button { Text = "Restore", DialogResult = DialogResult.OK, Size = new Size(90, 28), Location = new Point(ClientSize.Width - 12 - 90 - 8 - 90, ClientSize.Height - 40) };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Size = new Size(90, 28), Location = new Point(ClientSize.Width - 12 - 90, ClientSize.Height - 40) };
            Controls.Add(restore);
            Controls.Add(cancel);
            AcceptButton = restore;
            CancelButton = cancel;
            _list.DoubleClick += (sender, args) => restore.PerformClick();
        }

        /// <returns>The chosen commit, or null if the user cancelled.</returns>
        public static GitCommit Ask(string moduleName, IReadOnlyList<GitCommit> commits)
        {
            using (var dialog = new GitRestoreDialog(moduleName, commits))
            {
                return dialog.ShowDialog() == DialogResult.OK && dialog._list.SelectedIndices.Count == 1
                    ? commits[dialog._list.SelectedIndices[0]]
                    : null;
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
