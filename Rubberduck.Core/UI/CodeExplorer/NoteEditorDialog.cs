using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using ICSharpCode.AvalonEdit;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Resources.CodeExplorer;
using Button = System.Windows.Forms.Button;
using Label = System.Windows.Forms.Label;
using Panel = System.Windows.Forms.Panel;
using Size = System.Drawing.Size;
using TextBox = System.Windows.Forms.TextBox;
using FlowDirection = System.Windows.Forms.FlowDirection;
using GridLength = System.Windows.GridLength;
using GridUnitType = System.Windows.GridUnitType;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>
    /// Edits a note: the markdown on the left, as it will be shown on the right, and its name and folder on top.
    /// Either side can be hidden and the room shared by dragging; the preview zooms. Ctrl + S saves; Ctrl + mouse wheel changes the font size of the markdown or the zoom of the preview.
    /// </summary>
    public class NoteEditorDialog : Form
    {
        private const double PointsToWpf = 96d / 72d;
        private const double ZoomStep = 10;

        private readonly string _original;
        private readonly string _originalName;
        private readonly string _originalFolder;
        private readonly TextEditor _editor;
        private readonly FlowDocumentScrollViewer _preview;
        private readonly Grid _grid;
        private readonly GridSplitter _splitter;
        private readonly TextBox _name;
        private readonly TextBox _folder;
        private readonly Label _problem;
        private readonly Button _save;
        private readonly Button _viewSource;
        private readonly Button _viewSplit;
        private readonly Button _viewPreview;
        private readonly Button _zoomLabel;
        private readonly Timer _debounce;
        private bool _previewStale = true;

        private readonly IReadOnlyList<string> _knownFolders;

        /// <param name="knownFolders">The folders the project has, for the folder picker.</param>
        public NoteEditorDialog(string title, string markdown, string name, string folder, NoteEditorLayout layout = null, ResourceEditorPreferences preferences = null, IReadOnlyList<string> knownFolders = null)
        {
            _knownFolders = knownFolders ?? new string[0];
            _original = Crlf(markdown);
            _originalName = name ?? string.Empty;
            _originalFolder = folder ?? string.Empty;
            Layout = (layout ?? new NoteEditorLayout()).Sanitized();
            Preferences = preferences ?? new ResourceEditorPreferences();

            Text = title;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(1100, 700);
            MinimumSize = new Size(960, 360);
            ShowIcon = false;
            ShowInTaskbar = false;
            KeyPreview = true;

            _editor = new TextEditor
            {
                ShowLineNumbers = true,
                WordWrap = true,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Text = _original
            };
            _editor.Options.EnableHyperlinks = false;
            _editor.Options.EnableEmailHyperlinks = false;
            ApplyFont();
            _editor.TextChanged += (sender, args) => { _debounce.Stop(); _debounce.Start(); };
            _editor.PreviewMouseWheel += EditorMouseWheel;

            _preview = new FlowDocumentScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, IsToolBarVisible = false, MinZoom = NoteEditorLayout.MinZoom, MaxZoom = NoteEditorLayout.MaxZoom };
            _preview.PreviewMouseWheel += PreviewMouseWheel;

            _splitter = new GridSplitter { Width = 6, HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch, ToolTip = CodeExplorerUI.NoteEditor_SplitterTip };
            _splitter.DragCompleted += (sender, args) => RememberShare();
            _splitter.MouseDoubleClick += (sender, args) => { Layout.SourceShare = NoteEditorLayout.DefaultSourceShare; ApplyView(); };

            _grid = new Grid();
            _grid.ColumnDefinitions.Add(new ColumnDefinition());
            _grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            _grid.ColumnDefinitions.Add(new ColumnDefinition());
            Grid.SetColumn(_editor, 0);
            Grid.SetColumn(_splitter, 1);
            Grid.SetColumn(_preview, 2);
            _grid.Children.Add(_editor);
            _grid.Children.Add(_splitter);
            _grid.Children.Add(_preview);
            _grid.PreviewKeyDown += GridKeyDown;
            var host = new ElementHost { Dock = DockStyle.Fill, Child = _grid };

            var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 38, WrapContents = false, Padding = new Padding(6, 5, 6, 0) };
            _name = new TextBox { Width = 190, Text = _originalName };
            _folder = new TextBox { Width = 200, Text = _originalFolder };
            _name.TextChanged += (sender, args) => RefreshProblem();
            _folder.TextChanged += (sender, args) => RefreshProblem();
            _viewSource = ViewButton(CodeExplorerUI.NoteEditor_ViewSource, CodeExplorerUI.NoteEditor_ViewSourceTip, NoteEditorView.Source);
            _viewSplit = ViewButton(CodeExplorerUI.NoteEditor_ViewSplit, CodeExplorerUI.NoteEditor_ViewSplitTip, NoteEditorView.Split);
            _viewPreview = ViewButton(CodeExplorerUI.NoteEditor_ViewPreview, CodeExplorerUI.NoteEditor_ViewPreviewTip, NoteEditorView.Preview);
            var zoomOut = SmallButton("−", CodeExplorerUI.NoteEditor_ZoomOutTip, (sender, args) => ZoomPreview(Layout.Zoom - ZoomStep));
            var zoomIn = SmallButton("+", CodeExplorerUI.NoteEditor_ZoomInTip, (sender, args) => ZoomPreview(Layout.Zoom + ZoomStep));
            _zoomLabel = SmallButton(string.Empty, CodeExplorerUI.NoteEditor_ZoomResetTip, (sender, args) => ZoomPreview(NoteEditorLayout.DefaultZoom));
            _zoomLabel.Width = 52;
            top.Controls.Add(Caption(CodeExplorerUI.NoteEditor_Name));
            top.Controls.Add(_name);
            top.Controls.Add(Caption(CodeExplorerUI.NoteEditor_Folder));
            top.Controls.Add(_folder);
            var browse = SmallButton("…", CodeExplorerUI.NoteEditor_FolderBrowse, (sender, args) =>
            {
                var chosen = FolderPickerDialog.Pick(this, _knownFolders, _folder.Text.Trim());
                if (chosen != null)
                {
                    _folder.Text = chosen;
                }
            });
            browse.Margin = new Padding(0, 0, 1, 0);
            top.Controls.Add(browse);
            top.Controls.Add(new Label { Width = 18 });
            top.Controls.Add(_viewSource);
            top.Controls.Add(_viewSplit);
            top.Controls.Add(_viewPreview);
            top.Controls.Add(new Label { Width = 12 });
            top.Controls.Add(zoomOut);
            top.Controls.Add(_zoomLabel);
            top.Controls.Add(zoomIn);

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 40 };
            _problem = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(8, 0, 0, 0), AutoEllipsis = true, ForeColor = Color.Firebrick };
            _save = new Button { Text = CodeExplorerUI.ResourceEditor_Save, Size = new Size(84, 26), Margin = new Padding(4, 7, 4, 7) };
            _save.Click += (sender, args) => Save();
            var cancel = new Button { Text = CodeExplorerUI.ResourceEditor_Cancel, Size = new Size(84, 26), Margin = new Padding(4, 7, 8, 7), DialogResult = DialogResult.Cancel };

            var buttons = new FlowLayoutPanel { Dock = DockStyle.Right, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.RightToLeft, WrapContents = false };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(_save);
            bottom.Controls.Add(_problem);
            bottom.Controls.Add(buttons);

            Controls.Add(host);
            Controls.Add(top);
            Controls.Add(bottom);
            CancelButton = cancel;

            _debounce = new Timer { Interval = 300 };
            _debounce.Tick += (sender, args) => { _debounce.Stop(); _previewStale = true; RefreshPreview(); };

            Shown += (sender, args) => (Layout.View == NoteEditorView.Preview ? (System.Windows.UIElement)_preview : _editor).Focus();
            ApplyView();
            RefreshProblem();
        }

        /// <summary>The font of the markdown as it is now; changed by the user while the window is open.</summary>
        public ResourceEditorPreferences Preferences { get; }

        /// <summary>The layout as the user left it: to be remembered for this note.</summary>
        public NoteEditorLayout Layout { get; }

        /// <summary>The markdown as saved. Only meaningful when the dialog ended with OK.</summary>
        public string EditedText => Crlf(_editor.Text);

        public string EditedName => _name.Text.Trim();

        public string EditedFolder => _folder.Text.Trim();

        public bool IsModified => Crlf(_editor.Text) != _original || EditedName != _originalName.Trim() || EditedFolder != _originalFolder.Trim();

        private static string Crlf(string text) => (text ?? string.Empty).Replace("\r\n", "\n").Replace("\n", "\r\n");

        private static Label Caption(string text) => new Label { Text = text, AutoSize = true, Margin = new Padding(4, 5, 2, 0) };

        private Button ViewButton(string text, string tip, NoteEditorView view)
        {
            var button = SmallButton(text, tip, (sender, args) => { Layout.View = view; ApplyView(); });
            button.Width = 64;
            return button;
        }

        private static Button SmallButton(string text, string tip, EventHandler click)
        {
            var button = new Button { Text = text, Size = new Size(30, 24), Margin = new Padding(1, 0, 1, 0), FlatStyle = FlatStyle.Standard, UseVisualStyleBackColor = true, TabStop = false };
            new System.Windows.Forms.ToolTip().SetToolTip(button, tip);
            button.Click += click;
            return button;
        }

        /// <summary>Shows the sides the layout asks for. The columns of a hidden side shrink to nothing so that the other one takes the whole width.</summary>
        private void ApplyView()
        {
            var source = Layout.View != NoteEditorView.Preview;
            var preview = Layout.View != NoteEditorView.Source;

            _editor.Visibility = source ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            _preview.Visibility = preview ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            _splitter.Visibility = source && preview ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

            _grid.ColumnDefinitions[0].Width = source ? new GridLength(preview ? Layout.SourceShare : 1, GridUnitType.Star) : new GridLength(0);
            _grid.ColumnDefinitions[2].Width = preview ? new GridLength(source ? 1 - Layout.SourceShare : 1, GridUnitType.Star) : new GridLength(0);

            _viewSource.Font = new Font(_viewSource.Font, Layout.View == NoteEditorView.Source ? FontStyle.Bold : FontStyle.Regular);
            _viewSplit.Font = new Font(_viewSplit.Font, Layout.View == NoteEditorView.Split ? FontStyle.Bold : FontStyle.Regular);
            _viewPreview.Font = new Font(_viewPreview.Font, Layout.View == NoteEditorView.Preview ? FontStyle.Bold : FontStyle.Regular);

            _preview.Zoom = Layout.Zoom;
            _zoomLabel.Text = Layout.Zoom.ToString("0") + "%";

            if (preview)
            {
                RefreshPreview();
            }
        }

        private void RememberShare()
        {
            var source = _grid.ColumnDefinitions[0].ActualWidth;
            var preview = _grid.ColumnDefinitions[2].ActualWidth;
            if (source + preview > 0)
            {
                Layout.SourceShare = NoteEditorLayout.ClampShare(source / (source + preview));
                //The splitter has set pixel widths: go back to shares, which survive a resize of the window.
                ApplyView();
            }
        }

        private void ZoomPreview(double percent)
        {
            Layout.Zoom = NoteEditorLayout.ClampZoom(percent);
            _preview.Zoom = Layout.Zoom;
            _zoomLabel.Text = Layout.Zoom.ToString("0") + "%";
        }

        private void RefreshPreview()
        {
            //A hidden preview is not rendered; it catches up when it is shown again.
            if (Layout.View == NoteEditorView.Source || !_previewStale)
            {
                RefreshProblem();
                return;
            }

            try
            {
                _preview.Document = MarkdownRenderer.ToDocument(_editor.Text);
                _previewStale = false;
            }
            catch (Exception exception)
            {
                //A half-typed table or the like must not take the window down; the last good preview stays.
                _problem.Text = exception.Message;
                return;
            }

            RefreshProblem();
        }

        private void RefreshProblem()
        {
            var tooLong = NoteModule.FirstLineTooLong(_editor.Text);
            _problem.Text = !NoteModule.IsValidName(_name.Text)
                ? CodeExplorerUI.NoteEditor_NameInvalid
                : !NoteModule.IsValidFolder(_folder.Text)
                    ? CodeExplorerUI.NoteEditor_FolderInvalid
                    : tooLong > 0
                        ? string.Format(CodeExplorerUI.NoteEditor_LineTooLong, tooLong, NoteModule.MaxLineLength - 2)
                        : string.Empty;
            _save.Enabled = _problem.Text.Length == 0;
        }

        private void ApplyFont()
        {
            _editor.FontFamily = new System.Windows.Media.FontFamily(Preferences.FontName);
            _editor.FontSize = Preferences.FontSize * PointsToWpf;
        }

        private static bool CtrlDown => (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) != 0;

        private void EditorMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs args)
        {
            if (CtrlDown)
            {
                Preferences.FontSize = ResourceEditorPreferences.ClampSize(Preferences.FontSize + (args.Delta > 0 ? 1 : -1));
                ApplyFont();
                args.Handled = true;
            }
        }

        private void PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs args)
        {
            if (CtrlDown)
            {
                ZoomPreview(Layout.Zoom + (args.Delta > 0 ? ZoomStep : -ZoomStep));
                args.Handled = true;
            }
        }

        private void GridKeyDown(object sender, System.Windows.Input.KeyEventArgs args)
        {
            if (!CtrlDown)
            {
                return;
            }

            switch (args.Key)
            {
                case System.Windows.Input.Key.S:
                    args.Handled = true;
                    Save();
                    break;
                //The zoom keys belong to the preview, which is shown or not; the markdown has the mouse wheel.
                case System.Windows.Input.Key.OemPlus:
                case System.Windows.Input.Key.Add:
                    args.Handled = true;
                    ZoomPreview(Layout.Zoom + ZoomStep);
                    break;
                case System.Windows.Input.Key.OemMinus:
                case System.Windows.Input.Key.Subtract:
                    args.Handled = true;
                    ZoomPreview(Layout.Zoom - ZoomStep);
                    break;
                case System.Windows.Input.Key.D0:
                case System.Windows.Input.Key.NumPad0:
                    args.Handled = true;
                    ZoomPreview(NoteEditorLayout.DefaultZoom);
                    break;
            }
        }

        private void Save()
        {
            if (!_save.Enabled)
            {
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult != DialogResult.OK && IsModified && e.CloseReason == CloseReason.UserClosing)
            {
                var answer = System.Windows.Forms.MessageBox.Show(this, CodeExplorerUI.ResourceEditor_DiscardEdits, Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (answer != DialogResult.Yes)
                {
                    e.Cancel = true;
                    return;
                }
            }

            base.OnFormClosing(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _debounce?.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
