using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Rubberduck.Git;

namespace Rubberduck.UI.CodeExplorer.GitClient
{
    /// <summary>
    /// Shows the changes cut into hunks and the commits they are grouped into. The user moves hunks between commits, edits the messages,
    /// and can ask the AI to group them again.
    /// </summary>
    public sealed class GitSplitDialog : Form
    {
        private class Group
        {
            public string Message = string.Empty;
            public readonly List<GitHunk> Hunks = new List<GitHunk>();
        }

        private readonly IReadOnlyList<GitHunk> _hunks;
        private readonly Func<IWin32Window, GitGroupSuggestion> _group;
        private readonly List<Group> _groups = new List<Group>();
        private readonly ListBox _groupList;
        private readonly TextBox _message;
        private readonly ListView _hunkList;
        private readonly TextBox _preview;
        private readonly Label _note;
        private bool _loading;

        private GitSplitDialog(string projectName, IReadOnlyList<GitHunk> hunks, IReadOnlyList<GitCommitPart> initial, Func<IWin32Window, GitGroupSuggestion> group)
        {
            _hunks = hunks;
            _group = group;
            Text = "Git: commit '" + projectName + "' in parts";
            Font = new Font("Segoe UI", 9f);
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(900, 640);

            Controls.Add(new Label { Text = "Commits (in this order):", Location = new Point(12, 10), Size = new Size(260, 18) });
            _groupList = new ListBox { Location = new Point(12, 30), Size = new Size(260, 200), IntegralHeight = false };
            Controls.Add(_groupList);

            Controls.Add(new Label { Text = "Message of the selected commit:", Location = new Point(284, 10), Size = new Size(600, 18) });
            _message = new TextBox { Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical, Location = new Point(284, 30), Size = new Size(604, 200) };
            Controls.Add(_message);

            var add = new Button { Text = "New commit", Location = new Point(12, 236), Size = new Size(125, 26) };
            var up = new Button { Text = "Up", Location = new Point(143, 236), Size = new Size(60, 26) };
            var down = new Button { Text = "Down", Location = new Point(209, 236), Size = new Size(63, 26) };
            Controls.AddRange(new Control[] { add, up, down });

            var move = new Button { Text = "Move the selected hunks to the selected commit", Location = new Point(284, 236), Size = new Size(320, 26) };
            var created = new Button { Text = "Make them a new commit", Location = new Point(610, 236), Size = new Size(180, 26) };
            Controls.AddRange(new Control[] { move, created });

            _hunkList = new ListView { View = View.Details, FullRowSelect = true, MultiSelect = true, HideSelection = false, ShowGroups = true, Location = new Point(12, 270), Size = new Size(876, 200) };
            _hunkList.Columns.Add("File", 220);
            _hunkList.Columns.Add("Change", 640);
            Controls.Add(_hunkList);

            _preview = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false, Font = new Font("Consolas", 9f), Location = new Point(12, 478), Size = new Size(876, 110) };
            Controls.Add(_preview);

            _note = new Label { Location = new Point(12, 594), Size = new Size(876, 18), ForeColor = Color.DarkRed };
            Controls.Add(_note);

            var ai = new Button { Text = "Group with AI", Location = new Point(12, ClientSize.Height - 34), Size = new Size(130, 28), Visible = group != null };
            var commit = new Button { Text = "Commit", DialogResult = DialogResult.OK, Size = new Size(100, 28), Location = new Point(ClientSize.Width - 12 - 100 - 8 - 100, ClientSize.Height - 34) };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Size = new Size(100, 28), Location = new Point(ClientSize.Width - 12 - 100, ClientSize.Height - 34) };
            Controls.AddRange(new Control[] { ai, commit, cancel });
            CancelButton = cancel;

            foreach (var part in initial)
            {
                var g = new Group { Message = part.Message };
                g.Hunks.AddRange(part.Hunks);
                _groups.Add(g);
            }

            if (_groups.Count == 0)
            {
                _groups.Add(new Group());
                _groups[0].Hunks.AddRange(hunks);
            }

            _groupList.SelectedIndexChanged += (sender, args) => ShowSelectedGroup();
            _message.TextChanged += (sender, args) =>
            {
                if (!_loading && Selected() != null)
                {
                    Selected().Message = _message.Text;
                    RefreshLists(keepGroup: true);
                }
            };
            _hunkList.SelectedIndexChanged += (sender, args) => ShowPreview();
            add.Click += (sender, args) => { _groups.Add(new Group()); RefreshLists(_groups.Count - 1); };
            up.Click += (sender, args) => Shift(-1);
            down.Click += (sender, args) => Shift(1);
            move.Click += (sender, args) => MoveSelectedHunks(Selected());
            created.Click += (sender, args) =>
            {
                var g = new Group();
                _groups.Add(g);
                MoveSelectedHunks(g);
                RefreshLists(_groups.IndexOf(g));
            };
            ai.Click += (sender, args) => AskTheAi();
            commit.Click += (sender, args) =>
            {
                if (!Valid())
                {
                    DialogResult = DialogResult.None;
                }
            };

            RefreshLists(0);
        }

        private Group Selected() => _groupList.SelectedIndex >= 0 && _groupList.SelectedIndex < _groups.Count ? _groups[_groupList.SelectedIndex] : null;

        private static string Subject(Group g) => string.IsNullOrWhiteSpace(g.Message) ? "(no message yet)" : g.Message.Trim().Split('\n')[0].TrimEnd('\r');

        private void RefreshLists(int? select = null, bool keepGroup = false)
        {
            var index = select ?? Math.Max(0, _groupList.SelectedIndex);
            var wasLoading = _loading;
            _loading = true;
            _groupList.BeginUpdate();
            _groupList.Items.Clear();
            for (var i = 0; i < _groups.Count; i++)
            {
                _groupList.Items.Add($"{i + 1}. {Subject(_groups[i])}  ({_groups[i].Hunks.Count})");
            }

            _groupList.SelectedIndex = Math.Min(index, _groups.Count - 1);
            _groupList.EndUpdate();

            _hunkList.BeginUpdate();
            _hunkList.Items.Clear();
            _hunkList.Groups.Clear();
            for (var i = 0; i < _groups.Count; i++)
            {
                var listGroup = new ListViewGroup($"{i + 1}. {Subject(_groups[i])}");
                _hunkList.Groups.Add(listGroup);
                foreach (var hunk in _groups[i].Hunks.OrderBy(h => int.Parse(h.Id.Substring(1))))
                {
                    _hunkList.Items.Add(new ListViewItem(new[] { hunk.Path, hunk.Summary }, listGroup) { Tag = hunk });
                }
            }

            _hunkList.EndUpdate();
            _loading = wasLoading;
            if (!keepGroup)
            {
                ShowSelectedGroup();
            }
        }

        private void ShowSelectedGroup()
        {
            _loading = true;
            _message.Text = Selected()?.Message ?? string.Empty;
            _loading = false;
        }

        private void ShowPreview()
        {
            var hunk = _hunkList.SelectedItems.Count > 0 ? _hunkList.SelectedItems[0].Tag as GitHunk : null;
            _preview.Text = hunk == null ? string.Empty
                : hunk.IsBinary ? "(binary file)"
                : Encoding.GetEncoding(1252).GetString(GitHunkParser.Bytes.GetBytes(hunk.Body)).Replace("\r\n", "\n").Replace("\n", Environment.NewLine);
        }

        private void Shift(int by)
        {
            var i = _groupList.SelectedIndex;
            var j = i + by;
            if (i < 0 || j < 0 || j >= _groups.Count)
            {
                return;
            }

            var g = _groups[i];
            _groups[i] = _groups[j];
            _groups[j] = g;
            RefreshLists(j);
        }

        private void MoveSelectedHunks(Group target)
        {
            if (target == null)
            {
                return;
            }

            var moving = _hunkList.SelectedItems.Cast<ListViewItem>().Select(item => (GitHunk)item.Tag).ToList();
            foreach (var g in _groups)
            {
                g.Hunks.RemoveAll(h => moving.Contains(h));
            }

            target.Hunks.AddRange(moving);
            RefreshLists(_groups.IndexOf(target));
        }

        private void AskTheAi()
        {
            var suggestion = _group(this);
            if (suggestion == null)
            {
                return;
            }

            _groups.Clear();
            foreach (var part in suggestion.Parts)
            {
                var g = new Group { Message = part.Message };
                g.Hunks.AddRange(part.Hunks);
                _groups.Add(g);
            }

            _note.Text = suggestion.Note ?? string.Empty;
            RefreshLists(0);
        }

        private bool Valid()
        {
            _groups.RemoveAll(g => g.Hunks.Count == 0);
            if (_groups.Count == 0)
            {
                MessageBox.Show(this, "There is nothing to commit.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            var unnamed = _groups.FindIndex(g => string.IsNullOrWhiteSpace(g.Message));
            if (unnamed >= 0)
            {
                RefreshLists(unnamed);
                MessageBox.Show(this, "Write a message for every commit.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            return true;
        }

        /// <param name="group">Asked to group the hunks when the user presses the AI button, which is only there if this is given. Returns null for no change.</param>
        /// <returns>The commits to make, in order, or null if the user cancelled.</returns>
        public static IReadOnlyList<GitCommitPart> Ask(string projectName, IReadOnlyList<GitHunk> hunks, Func<IWin32Window, GitGroupSuggestion> group)
        {
            var shown = hunks.Where(hunk => !hunk.IsCompanion).ToList();
            using (var dialog = new GitSplitDialog(projectName, shown, new[] { new GitCommitPart(string.Empty, shown) }, group))
            {
                return dialog.ShowDialog() == DialogResult.OK
                    ? dialog._groups.Select(g => new GitCommitPart(g.Message.Trim(), g.Hunks.OrderBy(h => int.Parse(h.Id.Substring(1))).ToList())).ToList()
                    : null;
            }
        }
    }
}
