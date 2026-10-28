using System;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Windows.Input;
using System.Xml;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Highlighting;
using Newtonsoft.Json;
using Rubberduck.PackageResources;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>
    /// Edits the text of a resource: ribbon XML or rubberduck.json. The text is coloured and checked as it is typed and the problem is shown where it is;
    /// saving is only possible while there is no blocking problem (a warning does not stop it). Format indents the text and Linearize puts it on one line.
    /// The font can be changed with the Font button or Ctrl + mouse wheel.
    /// </summary>
    public class ResourceTextEditorDialog : Form
    {
        //WPF measures in 1/96 inch, fonts in points.
        private const double PointsToWpf = 96d / 72d;

        private readonly Func<string, TextProblem> _check;
        private readonly Func<string, string> _format;
        private readonly Func<string, string> _linearize;
        private readonly string _original;
        private readonly TextEditor _editor;
        private readonly Label _status;
        private readonly Label _position;
        private readonly Button _save;
        private readonly Timer _debounce;
        private bool _valid;

        /// <param name="check">Null if the text is fine; see <see cref="TextProblem"/>.</param>
        /// <param name="format">Indents the text; throws <see cref="XmlException"/> or <see cref="JsonException"/> if it cannot (not well-formed, say); the message is shown. Null if there is no such thing; the button is left out.</param>
        /// <param name="linearize">The same, to a single line.</param>
        /// <param name="highlighting">How the text is coloured, or null for none.</param>
        /// <param name="preferences">The font to start with. Changes are made to this object; saving it is up to the caller.</param>
        public ResourceTextEditorDialog(string title, string text, Func<string, TextProblem> check, Func<string, string> format, Func<string, string> linearize = null, IHighlightingDefinition highlighting = null, ResourceEditorPreferences preferences = null)
        {
            _check = check;
            _format = format;
            _linearize = linearize;
            _original = text ?? string.Empty;
            Preferences = preferences ?? new ResourceEditorPreferences();

            Text = title;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(900, 640);
            MinimumSize = new Size(520, 360);
            ShowIcon = false;
            ShowInTaskbar = false;
            KeyPreview = true;

            var toolbar = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Padding = new Padding(4, 4, 4, 2) };
            if (_format != null)
            {
                toolbar.Controls.Add(ToolButton(CodeExplorerUI.ResourceEditor_Format, () => Rewrite(_format)));
            }

            if (_linearize != null)
            {
                toolbar.Controls.Add(ToolButton(CodeExplorerUI.ResourceEditor_Linearize, () => Rewrite(_linearize)));
            }

            toolbar.Controls.Add(ToolButton(CodeExplorerUI.ResourceEditor_Font, ChooseFont));
            toolbar.Controls.Add(new Label
            {
                Text = CodeExplorerUI.ResourceEditor_Note,
                AutoSize = true,
                Margin = new Padding(8, 7, 0, 0),
                ForeColor = SystemColors.GrayText
            });

            _editor = new TextEditor
            {
                ShowLineNumbers = true,
                WordWrap = false,
                HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto,
                SyntaxHighlighting = highlighting,
                Text = Crlf(_original)
            };
            _editor.Options.EnableHyperlinks = false;
            _editor.Options.EnableEmailHyperlinks = false;
            ApplyFont();
            _editor.TextChanged += (sender, args) => { _debounce.Stop(); _debounce.Start(); };
            _editor.TextArea.Caret.PositionChanged += (sender, args) => ShowPosition();
            _editor.PreviewKeyDown += EditorKeyDown;
            _editor.PreviewMouseWheel += EditorMouseWheel;
            var host = new ElementHost { Dock = DockStyle.Fill, Child = _editor };

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 40 };
            _status = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(8, 0, 0, 0), AutoEllipsis = true };
            _position = new Label { Dock = DockStyle.Right, Width = 110, TextAlign = ContentAlignment.MiddleRight };
            _save = new Button { Text = CodeExplorerUI.ResourceEditor_Save, Size = new Size(84, 26), Margin = new Padding(4, 7, 4, 7) };
            _save.Click += (sender, args) => Save();
            var cancel = new Button { Text = CodeExplorerUI.ResourceEditor_Cancel, Size = new Size(84, 26), Margin = new Padding(4, 7, 8, 7), DialogResult = DialogResult.Cancel };

            //Docked, not placed by coordinates: the buttons stay at the right edge whatever size or scaling the window gets. First in a right-to-left flow is the rightmost.
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Right, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.RightToLeft, WrapContents = false };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(_save);

            //Docking is laid out last-added first, so the text that takes what is left goes in first.
            bottom.Controls.Add(_status);
            bottom.Controls.Add(_position);
            bottom.Controls.Add(buttons);

            //Fill goes last so it takes what the docked panels leave.
            Controls.Add(host);
            Controls.Add(bottom);
            Controls.Add(toolbar);
            CancelButton = cancel;

            _debounce = new Timer { Interval = 300 };
            _debounce.Tick += (sender, args) => { _debounce.Stop(); CheckText(); };

            KeyDown += (sender, args) =>
            {
                if (args.Control && args.KeyCode == Keys.S)
                {
                    args.SuppressKeyPress = true;
                    Save();
                }
            };

            Shown += (sender, args) => { _editor.Focus(); ShowPosition(); };
            CheckText();
        }

        /// <summary>The font as it is now; changed by the user while the window is open.</summary>
        public ResourceEditorPreferences Preferences { get; }

        /// <summary>The text as saved. Only meaningful when the dialog ended with OK.</summary>
        public string EditedText => Crlf(_editor.Text);

        public bool IsModified => Crlf(_editor.Text) != Crlf(_original);

        private static string Crlf(string text) => text.Replace("\r\n", "\n").Replace("\n", "\r\n");

        private Button ToolButton(string text, Action click)
        {
            var button = new Button { Text = text, AutoSize = true, MinimumSize = new Size(84, 26), Margin = new Padding(4, 2, 0, 2) };
            button.Click += (sender, args) => click();
            return button;
        }

        private void ApplyFont()
        {
            _editor.FontFamily = new System.Windows.Media.FontFamily(Preferences.FontName);
            _editor.FontSize = Preferences.FontSize * PointsToWpf;
        }

        private void ChooseFont()
        {
            using (var dialog = new FontDialog
            {
                Font = new Font(Preferences.FontName, Preferences.FontSize),
                FontMustExist = true,
                ShowEffects = false,
                MinSize = (int)ResourceEditorPreferences.MinFontSize,
                MaxSize = (int)ResourceEditorPreferences.MaxFontSize
            })
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    Preferences.FontName = dialog.Font.FontFamily.Name;
                    Preferences.FontSize = ResourceEditorPreferences.ClampSize(dialog.Font.SizeInPoints);
                    ApplyFont();
                }
            }
        }

        private void Zoom(float points)
        {
            Preferences.FontSize = ResourceEditorPreferences.ClampSize(points);
            ApplyFont();
        }

        private void EditorMouseWheel(object sender, MouseWheelEventArgs args)
        {
            if ((Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) != 0)
            {
                Zoom(Preferences.FontSize + (args.Delta > 0 ? 1 : -1));
                args.Handled = true;
            }
        }

        private void EditorKeyDown(object sender, System.Windows.Input.KeyEventArgs args)
        {
            if ((Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) == 0)
            {
                return;
            }

            switch (args.Key)
            {
                case Key.S:
                    args.Handled = true;
                    Save();
                    break;
                case Key.OemPlus:
                case Key.Add:
                    args.Handled = true;
                    Zoom(Preferences.FontSize + 1);
                    break;
                case Key.OemMinus:
                case Key.Subtract:
                    args.Handled = true;
                    Zoom(Preferences.FontSize - 1);
                    break;
                case Key.D0:
                case Key.NumPad0:
                    args.Handled = true;
                    Zoom(ResourceEditorPreferences.DefaultFontSize);
                    break;
            }
        }

        private void CheckText()
        {
            var problem = _check(_editor.Text);
            _valid = problem == null || !problem.Blocking;
            _save.Enabled = _valid;

            if (problem == null)
            {
                _status.ForeColor = Color.ForestGreen;
                _status.Text = CodeExplorerUI.ResourceEditor_WellFormed;
            }
            else
            {
                _status.ForeColor = problem.Blocking ? Color.Firebrick : Color.DarkOrange;
                _status.Text = problem.Line > 0
                    ? string.Format(CodeExplorerUI.ResourceEditor_NotWellFormed, problem.Line, problem.Column, problem.Message)
                    : problem.Message;
            }
        }

        private void ShowPosition()
        {
            var caret = _editor.TextArea.Caret;
            _position.Text = string.Format(CodeExplorerUI.ResourceEditor_Position, caret.Line, caret.Column);
        }

        //Replaced as an edit of the document, so Ctrl+Z brings the old layout back.
        private void Rewrite(Func<string, string> layout)
        {
            try
            {
                var document = _editor.Document;
                document.Replace(0, document.TextLength, Crlf(layout(_editor.Text)));
                _editor.CaretOffset = 0;
                _editor.ScrollToHome();
            }
            catch (Exception exception) when (exception is XmlException || exception is JsonException)
            {
                //Not well-formed: the status line already says where. Otherwise it is something else that stops the layout, and the status line says that.
                CheckText();
                if (_valid)
                {
                    _status.ForeColor = Color.DarkOrange;
                    _status.Text = exception.Message;
                }
            }
        }

        private void Save()
        {
            if (!_valid)
            {
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            //Asked only when the user walks away from edits; saving and the Cancel button after "nothing changed" go straight through.
            if (DialogResult != DialogResult.OK && IsModified && e.CloseReason == CloseReason.UserClosing)
            {
                var answer = MessageBox.Show(this, CodeExplorerUI.ResourceEditor_DiscardEdits, Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
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
