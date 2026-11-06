using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Resources.Ship;
using Rubberduck.Ship;

namespace Rubberduck.UI.CodeExplorer.Ship
{
    /// <summary>
    /// The Ship dialog: choices on the first page, a confirmation (unless the user ships without asking), progress and the result.
    /// All decisions are in <see cref="ShipViewModel"/>; this class only shows it.
    /// </summary>
    public sealed class ShipDialog : Form
    {
        private static readonly Color Accent = Color.FromArgb(0, 102, 204);
        private static readonly Color AccentBack = Color.FromArgb(230, 240, 251);

        private readonly ShipViewModel _model;
        private readonly Func<string, string> _browse;
        private bool _updating;

        private readonly Panel _settingsPage = Page();
        private readonly Panel _confirmPage = Page();
        private readonly Panel _workingPage = Page();
        private readonly Panel _donePage = Page();

        private ChoiceCard _xlsmCard;
        private ChoiceCard _xlamCard;
        private TextBox _folder;
        private Label _resolved;
        private Label _folderError;
        private TextBox _fileName;
        private Label _willWrite;
        private readonly Dictionary<string, CheckBox> _steps = new Dictionary<string, CheckBox>();
        private readonly Dictionary<ShipOnExisting, RadioButton> _existing = new Dictionary<ShipOnExisting, RadioButton>();
        private TableLayoutPanel _treeShakingPanel;
        private readonly Dictionary<ShipTreeShakingMode, RadioButton> _modes = new Dictionary<ShipTreeShakingMode, RadioButton>();
        private TextBox _keep;
        private CheckBox _remember;
        private Label _previewState;
        private TextBox _previewText;
        private Button _refresh;
        private CheckBox _unattended;

        private Label _confirmText;
        private Label _existsPrompt;
        private Button _replace;
        private Button _keepBoth;
        private Label _workingTitle;
        private TextBox _workingLog;
        private Label _doneText;
        private TextBox _doneDetails;
        private LinkLabel _openFolder;

        private Button _ship;
        private Button _back;
        private Button _write;
        private Button _cancel;

        /// <param name="browse">Asks the user for a folder, starting at the given one; null if cancelled.</param>
        public ShipDialog(ShipViewModel model, Func<string, string> browse)
        {
            _model = model;
            _browse = browse;

            Text = string.Format(ShipUI.Ship_Title, model.ProjectName);
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(700, 700);
            MinimumSize = new Size(620, 560);
            ShowIcon = false;
            ShowInTaskbar = false;
            Font = SystemFonts.MessageBoxFont;

            BuildSettingsPage();
            BuildConfirmPage();
            BuildWorkingPage();
            BuildDonePage();

            var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, Height = 46, Padding = new Padding(10, 8, 10, 0) };
            _cancel = Btn(ShipUI.Ship_Button_Cancel, DialogResult.Cancel);
            _ship = Btn(ShipUI.Ship_Button_Ship, DialogResult.None);
            _back = Btn(ShipUI.Ship_Button_Back, DialogResult.None);
            _write = Btn(ShipUI.Ship_Confirm_Go, DialogResult.None);
            _write.MinimumSize = new Size(120, 28);
            _replace = Btn(ShipUI.Ship_Confirm_Replace, DialogResult.None);
            _keepBoth = Btn(ShipUI.Ship_Confirm_KeepBoth, DialogResult.None);
            bar.Controls.AddRange(new Control[] { _cancel, _ship, _write, _replace, _keepBoth, _back });

            _ship.Click += async (s, e) => await _model.ProceedAsync();
            _back.Click += (s, e) => _model.Back();
            _write.Click += async (s, e) => await _model.ConfirmAsync();
            _replace.Click += async (s, e) => await _model.ConfirmAsync(ShipOnExisting.Replace);
            _keepBoth.Click += async (s, e) => await _model.ConfirmAsync(ShipOnExisting.KeepBoth);

            Controls.Add(_settingsPage);
            Controls.Add(_confirmPage);
            Controls.Add(_workingPage);
            Controls.Add(_donePage);
            Controls.Add(bar);
            AcceptButton = _ship;
            CancelButton = _cancel;

            _model.Changed += OnModelChanged;
            FormClosed += (s, e) => _model.Changed -= OnModelChanged;
            Shown += async (s, e) => await _model.RefreshPreviewAsync();

            Render();
        }

        public static void Open(IWin32Window owner, ShipViewModel model, Func<string, string> browse)
        {
            using (var dialog = new ShipDialog(model, browse))
            {
                dialog.ShowDialog(owner);
            }
        }

        private static Panel Page() => new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(14, 10, 14, 6), Visible = false };

        private static Button Btn(string text, DialogResult result) =>
            new Button { Text = text, DialogResult = result, AutoSize = true, MinimumSize = new Size(88, 28), Margin = new Padding(4, 0, 4, 0) };

        // --- pages

        private void BuildSettingsPage()
        {
            var flow = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, GrowStyle = TableLayoutPanelGrowStyle.AddRows };
            flow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _settingsPage.Controls.Add(flow);

            flow.Controls.Add(Text14(string.Format(ShipUI.Ship_Title, _model.ProjectName)));
            flow.Controls.Add(Wrap(ShipUI.Ship_Intro, 0));

            // format: two cards side by side
            flow.Controls.Add(Heading(ShipUI.Ship_FormatGroup));
            var cards = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, Height = 80 };
            cards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            cards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            _xlsmCard = new ChoiceCard(ShipUI.Ship_Format_Xlsm_Title, ShipUI.Ship_Format_Xlsm_Description);
            _xlamCard = new ChoiceCard(ShipUI.Ship_Format_Xlam_Title, ShipUI.Ship_Format_Xlam_Description);
            cards.Controls.Add(_xlsmCard, 0, 0);
            cards.Controls.Add(_xlamCard, 1, 0);
            _xlsmCard.Chosen += (s, e) => { if (!_updating) _model.SetFormat(ShipFormat.Xlsm); };
            _xlamCard.Chosen += (s, e) => { if (!_updating) _model.SetFormat(ShipFormat.Xlam); };
            flow.Controls.Add(cards);

            // folder
            flow.Controls.Add(Heading(ShipUI.Ship_FolderGroup));
            flow.Controls.Add(Wrap(ShipUI.Ship_FolderWritten, 0));
            var row = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2 };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _folder = new TextBox { Dock = DockStyle.Fill };
            _folder.TextChanged += (s, e) => { if (!_updating) _model.SetOutputFolder(_folder.Text); };
            var browse = new Button { Text = ShipUI.Ship_Browse, AutoSize = true, MinimumSize = new Size(88, 26) };
            browse.Click += (s, e) =>
            {
                var chosen = _browse?.Invoke(_model.ResolvedFolder ?? _model.OutputFolder);
                if (!string.IsNullOrEmpty(chosen))
                {
                    _model.SetOutputFolder(chosen);
                }
            };
            row.Controls.Add(_folder, 0, 0);
            row.Controls.Add(browse, 1, 0);
            flow.Controls.Add(row);
            _resolved = Wrap(string.Empty, 0);
            _folderError = Wrap(string.Empty, 0);
            _folderError.ForeColor = Color.Firebrick;
            flow.Controls.Add(_resolved);
            flow.Controls.Add(_folderError);

            // file name
            flow.Controls.Add(Heading(ShipUI.Ship_FileNameGroup));
            _fileName = new TextBox { Dock = DockStyle.Top };
            _fileName.TextChanged += (s, e) => { if (!_updating) _model.SetFileName(_fileName.Text); };
            flow.Controls.Add(_fileName);
            flow.Controls.Add(Wrap(string.Format(ShipUI.Ship_FileNameHint, _model.ProjectName), 0));
            _willWrite = Wrap(string.Empty, 0);
            flow.Controls.Add(_willWrite);

            // steps
            flow.Controls.Add(Heading(ShipUI.Ship_StepsGroup));
            foreach (var step in ShipSteps.All)
            {
                var box = new CheckBox { AutoSize = true, Text = ShipViewModel.StepLabel(step), Margin = new Padding(3, 2, 3, 2) };
                var captured = step;
                box.CheckedChanged += (s, e) => { if (!_updating) _model.SetStep(captured, box.Checked); };
                _steps[step] = box;
                flow.Controls.Add(box);
                if (step == ShipSteps.TreeShaking)
                {
                    BuildTreeShakingPanel(flow);
                }
            }

            // existing file
            flow.Controls.Add(Heading(ShipUI.Ship_ExistingGroup));
            AddExisting(flow, ShipOnExisting.Ask, ShipUI.Ship_Existing_Ask);
            AddExisting(flow, ShipOnExisting.Replace, ShipUI.Ship_Existing_Replace);
            AddExisting(flow, ShipOnExisting.KeepBoth, ShipUI.Ship_Existing_KeepBoth);
            _remember = new CheckBox { AutoSize = true, Text = ShipUI.Ship_RememberChoice, Margin = new Padding(3, 6, 3, 2) };
            _remember.CheckedChanged += (s, e) => { if (!_updating) _model.SetRememberExistingChoice(_remember.Checked); };
            flow.Controls.Add(_remember);

            // preview
            flow.Controls.Add(Heading(ShipUI.Ship_PreviewGroup));
            var previewHead = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2 };
            previewHead.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            previewHead.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _previewState = new Label { AutoSize = true, Anchor = AnchorStyles.Left, MaximumSize = new Size(520, 0) };
            _refresh = new Button { Text = ShipUI.Ship_Preview_Refresh, AutoSize = true, MinimumSize = new Size(88, 26) };
            _refresh.Click += async (s, e) => await _model.RefreshPreviewAsync();
            previewHead.Controls.Add(_previewState, 0, 0);
            previewHead.Controls.Add(_refresh, 1, 0);
            flow.Controls.Add(previewHead);
            _previewText = new TextBox { Dock = DockStyle.Top, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Height = 170, BackColor = SystemColors.Window };
            flow.Controls.Add(_previewText);

            // unattended
            flow.Controls.Add(Heading(string.Empty));
            _unattended = new CheckBox { AutoSize = true, Text = ShipUI.Ship_UnattendedCheck };
            _unattended.CheckedChanged += (s, e) => { if (!_updating) _model.SetUnattended(_unattended.Checked); };
            flow.Controls.Add(_unattended);
            var warning = Wrap(ShipUI.Ship_UnattendedWarning, 0);
            warning.ForeColor = Color.Firebrick;
            flow.Controls.Add(warning);
        }

        /// <summary>The choices that belong to tree shaking; shown only while the step is on.</summary>
        private void BuildTreeShakingPanel(Control parent)
        {
            _treeShakingPanel = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, Margin = new Padding(24, 0, 0, 4), Visible = false };
            _treeShakingPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            AddMode(ShipTreeShakingMode.Conservative, ShipUI.Ship_TreeShaking_Mode_Conservative);
            AddMode(ShipTreeShakingMode.EntryPointsOnly, ShipUI.Ship_TreeShaking_Mode_EntryPointsOnly);
            _treeShakingPanel.Controls.Add(new Label { Text = ShipUI.Ship_TreeShaking_ModeHint, AutoSize = true, MaximumSize = new Size(600, 0), Margin = new Padding(3, 2, 0, 6) });
            _treeShakingPanel.Controls.Add(new Label { Text = ShipUI.Ship_TreeShaking_KeepLabel, AutoSize = true, MaximumSize = new Size(600, 0), Margin = new Padding(3, 2, 0, 2), Font = new Font(Font, FontStyle.Bold) });
            _keep = new TextBox { Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical, Height = 80, Width = 560, Margin = new Padding(3, 0, 3, 2) };
            _keep.Leave += (s, e) => { if (!_updating) _model.SetTreeShakingKeepText(_keep.Text); };
            _keep.TextChanged += (s, e) => { if (!_updating) _model.SetTreeShakingKeepText(_keep.Text); };
            _treeShakingPanel.Controls.Add(_keep);
            _treeShakingPanel.Controls.Add(new Label { Text = ShipUI.Ship_TreeShaking_KeepHint, AutoSize = true, MaximumSize = new Size(600, 0), Margin = new Padding(3, 2, 0, 2) });
            parent.Controls.Add(_treeShakingPanel);
        }

        private void AddMode(ShipTreeShakingMode mode, string text)
        {
            var radio = new RadioButton { AutoSize = true, Text = text, Margin = new Padding(3, 2, 3, 2) };
            radio.CheckedChanged += (s, e) => { if (!_updating && radio.Checked) _model.SetTreeShakingMode(mode); };
            _modes[mode] = radio;
            _treeShakingPanel.Controls.Add(radio);
        }

        private void AddExisting(Control parent, ShipOnExisting value, string text)
        {
            var radio = new RadioButton { AutoSize = true, Text = text, Margin = new Padding(3, 2, 3, 2) };
            radio.CheckedChanged += (s, e) => { if (!_updating && radio.Checked) _model.SetOnExisting(value); };
            _existing[value] = radio;
            parent.Controls.Add(radio);
        }

        private void BuildConfirmPage()
        {
            var flow = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1 };
            flow.Controls.Add(Text14(ShipUI.Ship_Confirm_Title));
            _confirmText = Wrap(string.Empty, 0);
            flow.Controls.Add(_confirmText);
            _existsPrompt = Wrap(ShipUI.Ship_Confirm_Exists, 0);
            _existsPrompt.Font = new Font(Font, FontStyle.Bold);
            flow.Controls.Add(_existsPrompt);
            _confirmPage.Controls.Add(flow);
        }

        private void BuildWorkingPage()
        {
            _workingTitle = Text14(string.Empty);
            _workingTitle.Dock = DockStyle.Top;
            var bar = new ProgressBar { Dock = DockStyle.Top, Style = ProgressBarStyle.Marquee, Height = 18 };
            _workingLog = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, BackColor = SystemColors.Window };
            _workingPage.Controls.Add(_workingLog);
            _workingPage.Controls.Add(bar);
            _workingPage.Controls.Add(_workingTitle);
        }

        private void BuildDonePage()
        {
            _doneText = Text14(string.Empty);
            _doneText.Dock = DockStyle.Top;
            _openFolder = new LinkLabel { Text = ShipUI.Ship_Done_Open, Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(0, 4, 0, 8) };
            _openFolder.LinkClicked += (s, e) =>
            {
                var folder = _model.ResultFolder;
                if (folder != null)
                {
                    Process.Start(new ProcessStartInfo("explorer.exe", "\"" + folder + "\"") { UseShellExecute = true });
                }
            };
            _doneDetails = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, BackColor = SystemColors.Window };
            _donePage.Controls.Add(_doneDetails);
            _donePage.Controls.Add(_openFolder);
            _donePage.Controls.Add(_doneText);
        }

        // --- rendering the model

        private void OnModelChanged(object sender, EventArgs e)
        {
            if (!IsHandleCreated || IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                try { BeginInvoke((Action)Render); } catch (InvalidOperationException) { }
            }
            else
            {
                Render();
            }
        }

        public void Render()
        {
            _updating = true;
            try
            {
                _settingsPage.Visible = _model.Page == ShipPage.Settings;
                _confirmPage.Visible = _model.Page == ShipPage.Confirm;
                _workingPage.Visible = _model.Page == ShipPage.Working;
                _donePage.Visible = _model.Page == ShipPage.Done;

                RenderSettings();
                RenderConfirm();
                RenderRun();

                var settings = _model.Page == ShipPage.Settings;
                var confirm = _model.Page == ShipPage.Confirm;
                _ship.Visible = settings;
                _ship.Enabled = _model.CanShip;
                _back.Visible = confirm;
                _write.Visible = confirm && !_model.NeedsExistingAnswer;
                _replace.Visible = _keepBoth.Visible = confirm && _model.NeedsExistingAnswer;
                _cancel.Visible = _model.Page != ShipPage.Working;
                _cancel.Text = _model.Page == ShipPage.Done ? ShipUI.Ship_Button_Close : ShipUI.Ship_Button_Cancel;
                AcceptButton = settings ? _ship : confirm && _write.Visible ? _write : (IButtonControl)_cancel;
            }
            finally
            {
                _updating = false;
            }
        }

        private void RenderSettings()
        {
            _xlsmCard.Selected = _model.Format == ShipFormat.Xlsm;
            _xlamCard.Selected = _model.Format == ShipFormat.Xlam;
            SetText(_folder, _model.OutputFolder);
            SetText(_fileName, _model.FileName);

            _resolved.Text = _model.ResolvedFolder == null ? string.Empty : string.Format(ShipUI.Ship_FolderResolved, _model.ResolvedFolder);
            _folderError.Text = _model.FolderMessage ?? string.Empty;
            _willWrite.Text = _model.FileNameValid
                ? (_model.TargetPath == null ? string.Empty : string.Format(ShipUI.Ship_WillWrite, _model.TargetPath))
                : ShipUI.Ship_FileNameInvalid;
            _willWrite.ForeColor = _model.FileNameValid ? ForeColor : Color.Firebrick;

            foreach (var step in _steps)
            {
                step.Value.Checked = _model.IsStepOn(step.Key);
            }

            foreach (var radio in _existing)
            {
                radio.Value.Checked = _model.OnExisting == radio.Key;
            }

            _treeShakingPanel.Visible = _model.IsStepOn(ShipSteps.TreeShaking);
            foreach (var mode in _modes)
            {
                mode.Value.Checked = _model.TreeShakingMode == mode.Key;
            }

            //Do not rewrite the box while the user types: the model's text drops blank lines and would eat the Enter key.
            if (!_keep.Focused)
            {
                SetText(_keep, _model.TreeShakingKeepText);
            }

            _remember.Checked = _model.RememberExistingChoice;
            _remember.Enabled = _model.OnExisting != ShipOnExisting.Ask || _model.NeedsExistingAnswer;
            _unattended.Checked = _model.Unattended;

            _previewState.Text = _model.BlockedMessage != null
                ? string.Format(ShipUI.Ship_Blocked, _model.BlockedMessage)
                : string.Empty;
            _previewState.Visible = _previewState.Text.Length > 0;
            _previewText.Text = string.Join(Environment.NewLine, _model.PreviewLines());
            _refresh.Enabled = _model.PreviewState != ShipPreviewState.Running;
        }

        private void RenderConfirm()
        {
            _confirmText.Text = string.Join(Environment.NewLine + Environment.NewLine, _model.ConfirmationLines());
            _existsPrompt.Visible = _model.NeedsExistingAnswer;
        }

        private void RenderRun()
        {
            _workingTitle.Text = string.Format(ShipUI.Ship_Working, _model.ProjectName);
            _workingLog.Text = string.Join(Environment.NewLine, _model.ProgressLines);
            _workingLog.SelectionStart = _workingLog.TextLength;
            _workingLog.ScrollToCaret();

            if (_model.Page != ShipPage.Done)
            {
                return;
            }

            var lines = new List<string>();
            if (_model.ResultError != null)
            {
                _doneText.Text = string.Format(ShipUI.Ship_Done_Failed, _model.ResultError);
                _doneText.ForeColor = Color.Firebrick;
                _openFolder.Visible = false;
                if (_model.BlockedReason == "unsaved")
                {
                    lines.Add(ShipUI.Ship_Refused_Unsaved_Hint);
                }
            }
            else
            {
                _doneText.Text = string.Format(ShipUI.Ship_Done_Written, _model.Result?.OutputPath);
                _doneText.ForeColor = ForeColor;
                _openFolder.Visible = _model.ResultFolder != null;
                lines.AddRange(ShipViewModel.DescribeReport(_model.Result));
            }

            _doneDetails.Text = string.Join(Environment.NewLine, lines);
        }

        private static void SetText(TextBox box, string text)
        {
            if (box.Text != text)
            {
                var start = box.SelectionStart;
                box.Text = text;
                box.SelectionStart = Math.Min(start, text.Length);
            }
        }

        // --- small helpers

        private Label Text14(string text) =>
            new Label { Text = text, AutoSize = true, Font = new Font(Font.FontFamily, 13f, FontStyle.Bold), Margin = new Padding(0, 0, 0, 6), MaximumSize = new Size(640, 0) };

        private Label Heading(string text) =>
            new Label { Text = text, AutoSize = true, Font = new Font(Font, FontStyle.Bold), Margin = new Padding(0, 12, 0, 3) };

        private static Label Wrap(string text, int _) =>
            new Label { Text = text, AutoSize = true, MaximumSize = new Size(640, 0), Margin = new Padding(0, 2, 0, 2) };

        /// <summary>A card with a radio button as title and a line of plain description. Clicking anywhere on it chooses it.</summary>
        private sealed class ChoiceCard : Panel
        {
            private readonly RadioButton _radio;
            private readonly Label _description;
            private bool _selected;

            public ChoiceCard(string title, string description)
            {
                Dock = DockStyle.Fill;
                Height = 76;
                Padding = new Padding(10, 8, 10, 8);
                Margin = new Padding(0, 0, 8, 0);
                BorderStyle = BorderStyle.FixedSingle;
                Cursor = Cursors.Hand;

                _radio = new RadioButton { Text = title, Dock = DockStyle.Top, AutoSize = false, Height = 24, Font = new Font(Font, FontStyle.Bold) };
                _description = new Label { Text = description, Dock = DockStyle.Fill, Padding = new Padding(20, 2, 0, 0) };
                Controls.Add(_description);
                Controls.Add(_radio);
                _radio.Click += (s, e) => Chosen?.Invoke(this, EventArgs.Empty);
                Click += (s, e) => Chosen?.Invoke(this, EventArgs.Empty);
                _description.Click += (s, e) => Chosen?.Invoke(this, EventArgs.Empty);
                Selected = false;
            }

            public event EventHandler Chosen;

            public bool Selected
            {
                get => _selected;
                set
                {
                    _selected = value;
                    _radio.Checked = value;
                    BackColor = value ? AccentBack : SystemColors.Control;
                    _radio.BackColor = BackColor;
                    _description.BackColor = BackColor;
                    _description.ForeColor = value ? Accent : ForeColor;
                }
            }
        }
    }
}
