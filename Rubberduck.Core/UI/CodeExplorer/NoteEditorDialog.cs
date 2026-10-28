using System;
using System.Drawing;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using ICSharpCode.AvalonEdit;
using Rubberduck.Resources.CodeExplorer;
using Button = System.Windows.Forms.Button;
using Label = System.Windows.Forms.Label;
using Panel = System.Windows.Forms.Panel;
using Size = System.Drawing.Size;
using FlowDirection = System.Windows.Forms.FlowDirection;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>Edits a note: the markdown on the left, as it will be shown on the right. Ctrl + S saves; Ctrl + mouse wheel changes the font size.</summary>
    public class NoteEditorDialog : Form
    {
        private const double PointsToWpf = 96d / 72d;

        private readonly string _original;
        private readonly TextEditor _editor;
        private readonly FlowDocumentScrollViewer _preview;
        private readonly Label _problem;
        private readonly Button _save;
        private readonly Timer _debounce;

        public NoteEditorDialog(string title, string markdown, ResourceEditorPreferences preferences = null)
        {
            _original = Crlf(markdown);
            Preferences = preferences ?? new ResourceEditorPreferences();

            Text = title;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(1100, 700);
            MinimumSize = new Size(600, 360);
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
            _editor.PreviewKeyDown += EditorKeyDown;
            _editor.PreviewMouseWheel += EditorMouseWheel;

            _preview = new FlowDocumentScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = System.Windows.GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });
            Grid.SetColumn(_editor, 0);
            var splitter = new GridSplitter { Width = 5, HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch };
            Grid.SetColumn(splitter, 1);
            Grid.SetColumn(_preview, 2);
            grid.Children.Add(_editor);
            grid.Children.Add(splitter);
            grid.Children.Add(_preview);
            var host = new ElementHost { Dock = DockStyle.Fill, Child = grid };

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
            Controls.Add(bottom);
            CancelButton = cancel;

            _debounce = new Timer { Interval = 300 };
            _debounce.Tick += (sender, args) => { _debounce.Stop(); RefreshPreview(); };

            Shown += (sender, args) => _editor.Focus();
            RefreshPreview();
        }

        /// <summary>The font as it is now; changed by the user while the window is open.</summary>
        public ResourceEditorPreferences Preferences { get; }

        /// <summary>The markdown as saved. Only meaningful when the dialog ended with OK.</summary>
        public string EditedText => Crlf(_editor.Text);

        public bool IsModified => Crlf(_editor.Text) != _original;

        private static string Crlf(string text) => (text ?? string.Empty).Replace("\r\n", "\n").Replace("\n", "\r\n");

        private void RefreshPreview()
        {
            try
            {
                _preview.Document = MarkdownRenderer.ToDocument(_editor.Text);
            }
            catch (Exception exception)
            {
                //A half-typed table or the like must not take the window down; the last good preview stays.
                _problem.Text = exception.Message;
                return;
            }

            var tooLong = Navigation.CodeExplorer.NoteModule.FirstLineTooLong(_editor.Text);
            var unsupported = Navigation.CodeExplorer.NoteModule.FirstUnsupportedCharacter(_editor.Text);
            _problem.Text = tooLong > 0
                ? string.Format(CodeExplorerUI.NoteEditor_LineTooLong, tooLong, Navigation.CodeExplorer.NoteModule.MaxLineLength - 2)
                : unsupported.HasValue
                    ? string.Format(CodeExplorerUI.NoteEditor_CharacterNotKept, unsupported.Value, ((int)unsupported.Value).ToString("X4"))
                    : string.Empty;
            _save.Enabled = _problem.Text.Length == 0;
        }

        private void ApplyFont()
        {
            _editor.FontFamily = new System.Windows.Media.FontFamily(Preferences.FontName);
            _editor.FontSize = Preferences.FontSize * PointsToWpf;
        }

        private void Zoom(float points)
        {
            Preferences.FontSize = ResourceEditorPreferences.ClampSize(points);
            ApplyFont();
        }

        private void EditorMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs args)
        {
            if ((System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) != 0)
            {
                Zoom(Preferences.FontSize + (args.Delta > 0 ? 1 : -1));
                args.Handled = true;
            }
        }

        private void EditorKeyDown(object sender, System.Windows.Input.KeyEventArgs args)
        {
            if ((System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) != 0 && args.Key == System.Windows.Input.Key.S)
            {
                args.Handled = true;
                Save();
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
