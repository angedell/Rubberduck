using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    public enum ReviewOutcome
    {
        Cancel,
        /// <summary>The checked items carry their action in <see cref="SyncItem.Chosen"/>.</summary>
        Apply,
        ChangeFolder
    }

    /// <summary>
    /// Shows what is out of sync and lets the user pick what to sync, which side wins where the history cannot say,
    /// and what to ignore for good.
    /// </summary>
    public sealed class GitSyncReviewDialog : Form
    {
        private readonly GitSyncScan _scan;
        private readonly Action<IReadOnlyList<SyncItem>> _ignore;
        private readonly Dictionary<SyncItem, SyncAction> _actions;
        private readonly ListView _list;
        private readonly Label _header;
        private readonly Label _hint;
        private bool _updating;

        private GitSyncReviewDialog(string projectName, GitSyncScan scan, Action<IReadOnlyList<SyncItem>> ignore)
        {
            _scan = scan;
            _ignore = ignore;
            _actions = scan.Plan.Items.ToDictionary(item => item, item => item.Suggested);

            Text = CodeExplorerUI.GitSync_DialogTitle;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(820, 520);
            MinimumSize = new Size(640, 400);

            _header = new Label
            {
                Location = new Point(12, 10),
                Size = new Size(796, 60),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            _list = new ListView
            {
                View = View.Details,
                FullRowSelect = true,
                CheckBoxes = true,
                HideSelection = false,
                Location = new Point(12, 76),
                Size = new Size(796, 354),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            _list.Columns.Add(CodeExplorerUI.GitSync_Col_Module, 170);
            _list.Columns.Add(CodeExplorerUI.GitSync_Col_Status, 140);
            _list.Columns.Add(CodeExplorerUI.GitSync_Col_Action, 150);
            _list.Columns.Add(CodeExplorerUI.GitSync_Col_Notes, 320);
            _list.ItemCheck += OnItemCheck;

            _hint = new Label
            {
                Location = new Point(12, 434),
                Size = new Size(796, 28),
                ForeColor = Color.Firebrick,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            var bottom = 466;
            var selectAll = Button(CodeExplorerUI.GitSync_SelectAll, 12, bottom, 90, (s, e) => CheckAll(true));
            var selectNone = Button(CodeExplorerUI.GitSync_SelectNone, 106, bottom, 90, (s, e) => CheckAll(false));
            var useModule = Button(CodeExplorerUI.GitSync_UseModule, 202, bottom, 130, (s, e) => ChooseForSelected(SyncAction.Export));
            var useFile = Button(CodeExplorerUI.GitSync_UseFile, 336, bottom, 130, (s, e) => ChooseForSelected(SyncAction.Import));
            var ignoreButton = Button(CodeExplorerUI.GitSync_IgnoreSelected, 470, bottom, 120, (s, e) => IgnoreSelected());
            foreach (var button in new[] { selectAll, selectNone, useModule, useFile, ignoreButton })
            {
                button.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            }

            var changeFolder = new Button { Text = CodeExplorerUI.GitSync_ChangeFolder, Location = new Point(12, 490), Size = new Size(120, 26), Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            changeFolder.Click += (s, e) => { Outcome = ReviewOutcome.ChangeFolder; Close(); };

            var apply = new Button { Text = CodeExplorerUI.GitSync_Apply, Location = new Point(618, 490), Size = new Size(110, 26), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            apply.Click += (s, e) => Apply();
            var cancel = new Button { Text = CodeExplorerUI.BulkRemove_Cancel, Location = new Point(734, 490), Size = new Size(74, 26), Anchor = AnchorStyles.Bottom | AnchorStyles.Right, DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[] { _header, _list, _hint, selectAll, selectNone, useModule, useFile, ignoreButton, changeFolder, apply, cancel });
            CancelButton = cancel;

            _header.Tag = projectName;
            Fill();
        }

        private ReviewOutcome Outcome { get; set; } = ReviewOutcome.Cancel;

        private static Button Button(string text, int x, int y, int width, EventHandler click)
        {
            var button = new Button { Text = text, Location = new Point(x, y), Size = new Size(width, 24) };
            button.Click += click;
            return button;
        }

        /// <summary>Shows the dialog. For <see cref="ReviewOutcome.Apply"/>, <see cref="SyncItem.Chosen"/> holds the picks.</summary>
        /// <param name="ignore">Called with the items the user ignores, so that the ignore list can be saved right away.</param>
        public static ReviewOutcome Ask(string projectName, GitSyncScan scan, Action<IReadOnlyList<SyncItem>> ignore)
        {
            using (var dialog = new GitSyncReviewDialog(projectName, scan, ignore))
            {
                dialog.ShowDialog();
                return dialog.Outcome;
            }
        }

        private void Fill()
        {
            _updating = true;
            try
            {
                _list.Items.Clear();
                foreach (var item in _scan.Plan.Items)
                {
                    var row = new ListViewItem(item.Name) { Tag = item };
                    row.SubItems.Add(StatusText(item.Status));
                    row.SubItems.Add(string.Empty);
                    row.SubItems.Add(string.Empty);
                    _list.Items.Add(row);
                    RefreshRow(row);
                    row.Checked = _actions[item] != SyncAction.None && !item.IsBlockedFor(_actions[item]);
                }
            }
            finally
            {
                _updating = false;
            }

            ShowHint();
            _header.Text = string.Format(
                CodeExplorerUI.GitSync_Review_Header,
                _header.Tag, _scan.Folder, _scan.Plan.Items.Count, _scan.Plan.InSyncCount, _scan.Plan.Ignored.Count);
        }

        private void RefreshRow(ListViewItem row)
        {
            var item = (SyncItem)row.Tag;
            var action = _actions[item];
            row.SubItems[2].Text = ActionText(action);

            var notes = new List<string>();
            if (item.MovesFile)
            {
                notes.Add(CodeExplorerUI.GitSync_MovesFile);
            }

            foreach (var violation in item.Violations)
            {
                var blocks = violation.Blocking && action != SyncAction.None && item.IsBlockedFor(action) && violation.Side == (action == SyncAction.Export ? SyncSide.Module : SyncSide.File);
                var side = violation.Side == SyncSide.Module ? "project" : "folder";
                var text = $"[{violation.Rule}, {side}] {violation.Message}";
                notes.Add(blocks ? string.Format(CodeExplorerUI.GitSync_Blocked, text) : text);
            }

            row.SubItems[3].Text = string.Join("  ", notes);
            row.ForeColor = action != SyncAction.None && item.IsBlockedFor(action) ? Color.Firebrick : SystemColors.WindowText;
        }

        private static string StatusText(SyncStatus status)
        {
            switch (status)
            {
                case SyncStatus.ModuleOnly: return CodeExplorerUI.GitSync_Status_ModuleOnly;
                case SyncStatus.FileOnly: return CodeExplorerUI.GitSync_Status_FileOnly;
                case SyncStatus.ModuleChanged: return CodeExplorerUI.GitSync_Status_ModuleChanged;
                case SyncStatus.FileChanged: return CodeExplorerUI.GitSync_Status_FileChanged;
                case SyncStatus.BothChanged: return CodeExplorerUI.GitSync_Status_BothChanged;
                default: return CodeExplorerUI.GitSync_Status_Differs;
            }
        }

        private static string ActionText(SyncAction action)
        {
            switch (action)
            {
                case SyncAction.Export: return CodeExplorerUI.GitSync_Action_Export;
                case SyncAction.Import: return CodeExplorerUI.GitSync_Action_Import;
                default: return CodeExplorerUI.GitSync_Action_None;
            }
        }

        //A row can only be ticked once it has a side to sync from, and that side is not blocked by an enforced rule.
        private void OnItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (_updating || e.NewValue != CheckState.Checked)
            {
                return;
            }

            var item = (SyncItem)_list.Items[e.Index].Tag;
            var action = _actions[item];
            if (action == SyncAction.None || item.IsBlockedFor(action))
            {
                e.NewValue = CheckState.Unchecked;
            }
        }

        private void CheckAll(bool check)
        {
            foreach (ListViewItem row in _list.Items)
            {
                row.Checked = check; //Rows that cannot be synced refuse in OnItemCheck.
                row.Selected = check; //So that "Use project version" and "Use folder version" apply to all of them.
            }

            ShowHint();
        }

        //Rows with no side to sync from refuse to be ticked; say so, or it looks as if nothing happened.
        private void ShowHint()
        {
            var undecided = _list.Items.Cast<ListViewItem>().Count(row => _actions[(SyncItem)row.Tag] == SyncAction.None);
            var blocked = _list.Items.Cast<ListViewItem>().Count(row =>
            {
                var item = (SyncItem)row.Tag;
                return _actions[item] != SyncAction.None && item.IsBlockedFor(_actions[item]);
            });

            var lines = new List<string>();
            if (undecided > 0)
            {
                lines.Add(string.Format(CodeExplorerUI.GitSync_Hint_NeedSide, undecided));
            }

            if (blocked > 0)
            {
                lines.Add(string.Format(CodeExplorerUI.GitSync_Hint_Blocked, blocked));
            }

            _hint.Text = string.Join(Environment.NewLine, lines);
        }

        private void ChooseForSelected(SyncAction action)
        {
            //With nothing selected the choice is for every row.
            var rows = _list.SelectedItems.Count > 0 ? _list.SelectedItems.Cast<ListViewItem>().ToList() : _list.Items.Cast<ListViewItem>().ToList();
            foreach (var row in rows)
            {
                var item = (SyncItem)row.Tag;
                if ((action == SyncAction.Export && !item.ModuleExists) || (action == SyncAction.Import && !item.FileExists))
                {
                    continue;
                }

                _actions[item] = action;
                RefreshRow(row);
                row.Checked = false;
                row.Checked = !item.IsBlockedFor(action);
            }

            ShowHint();
        }

        private void IgnoreSelected()
        {
            var rows = _list.SelectedItems.Cast<ListViewItem>().ToList();
            if (rows.Count == 0)
            {
                return;
            }

            var message = string.Format(CodeExplorerUI.GitSync_IgnoreConfirm, rows.Count);
            if (MessageBox.Show(this, message, CodeExplorerUI.GitSync_DialogTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            var items = rows.Select(row => (SyncItem)row.Tag).ToList();
            _ignore(items);
            foreach (var item in items)
            {
                _scan.Plan.Items.Remove(item);
                _scan.Plan.Ignored.Add(item.Name);
                _actions.Remove(item);
            }

            Fill();
        }

        private void Apply()
        {
            foreach (ListViewItem row in _list.Items)
            {
                var item = (SyncItem)row.Tag;
                item.Chosen = row.Checked ? _actions[item] : SyncAction.None;
            }

            Outcome = ReviewOutcome.Apply;
            Close();
        }
    }
}
