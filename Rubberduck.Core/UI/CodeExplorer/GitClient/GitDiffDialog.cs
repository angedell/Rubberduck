using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Rubberduck.Git;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace Rubberduck.UI.CodeExplorer.GitClient
{
    /// <summary>
    /// One file before and after, side by side and read-only, with the same line comparison as Git sync's compare window: what was removed is red on the left,
    /// what was added is green on the right, and within a changed line the words that differ are marked more strongly.
    /// </summary>
    public sealed class GitDiffDialog : Form
    {
        private static readonly Color RemovedBack = Color.FromArgb(255, 220, 220);
        private static readonly Color RemovedWord = Color.FromArgb(255, 160, 160);
        private static readonly Color AddedBack = Color.FromArgb(210, 240, 210);
        private static readonly Color AddedWord = Color.FromArgb(120, 210, 140);
        private static readonly Color GhostBack = Color.FromArgb(240, 240, 240);
        private static readonly Color CurrentBack = Color.FromArgb(255, 200, 120);

        private const int EmGetFirstVisibleLine = 0xCE;
        private const int EmLineScroll = 0xB6;
        private const int WmSetRedraw = 0x000B;

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);

        private readonly SyncDiff _diff;
        private readonly RichTextBox _before;
        private readonly RichTextBox _after;
        private readonly Label _status;
        private readonly int[] _firstRow;
        private int _selected;
        private bool _scrolling;

        private GitDiffDialog(string title, string beforeCaption, string afterCaption, string beforeText, string afterText)
        {
            _diff = new SyncDiff(beforeText, afterText);
            _firstRow = new int[_diff.Hunks.Count];

            Text = title;
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(1100, 640);
            MinimumSize = new Size(700, 400);

            var captionBefore = new Label { Text = beforeCaption, Font = new Font(Font, FontStyle.Bold), Location = new Point(12, 10), Size = new Size(400, 18) };
            var captionAfter = new Label { Text = afterCaption, Font = new Font(Font, FontStyle.Bold), Location = new Point(12, 10), Size = new Size(400, 18) };
            _before = Pane();
            _after = Pane();
            _status = new Label { Location = new Point(300, 610), Size = new Size(600, 20), Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right };

            var previous = new Button { Text = "Previous change", Location = new Point(12, 606), Size = new Size(130, 26), Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            var next = new Button { Text = "Next change", Location = new Point(148, 606), Size = new Size(130, 26), Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            var close = new Button { Text = "Close", DialogResult = DialogResult.Cancel, Location = new Point(ClientSize.Width - 12 - 90, 606), Size = new Size(90, 26), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            previous.Click += (s, e) => Goto(_selected - 1);
            next.Click += (s, e) => Goto(_selected + 1);
            CancelButton = close;

            Controls.AddRange(new Control[] { captionBefore, captionAfter, _before, _after, _status, previous, next, close });
            Fill();
            PaintAll();

            void LayoutPanes()
            {
                var width = (ClientSize.Width - 36) / 2;
                var height = Math.Max(100, ClientSize.Height - 30 - 46);
                _before.SetBounds(12, 30, width, height);
                _after.SetBounds(24 + width, 30, width, height);
                captionBefore.SetBounds(12, 10, width, 18);
                captionAfter.SetBounds(24 + width, 10, width, 18);
            }

            LayoutPanes();
            Resize += (s, e) => LayoutPanes();
            _before.VScroll += (s, e) => Scrolled(_before, _after);
            _after.VScroll += (s, e) => Scrolled(_after, _before);
            UpdateStatus(beforeText, afterText);
            Shown += (s, e) => Goto(0);
        }

        /// <summary>Shows the comparison. Text that is not text (a form's binary part) is not compared, only said to differ.</summary>
        /// <param name="before">The earlier content, or null if the file did not exist then.</param>
        /// <param name="after">The later content, or null if the file is gone.</param>
        public static void Show(IWin32Window owner, string path, string beforeCaption, string afterCaption, byte[] before, byte[] after)
        {
            if (IsBinary(before) || IsBinary(after))
            {
                MessageBox.Show(owner, "'" + path + "' is not a text file, so its changes cannot be shown.", "Git: changes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var dialog = new GitDiffDialog("Git: changes in '" + path + "'", beforeCaption, afterCaption, Decode(before), Decode(after)))
            {
                dialog.ShowDialog(owner);
            }
        }

        private static bool IsBinary(byte[] bytes) => bytes != null && Array.IndexOf(bytes, (byte)0) >= 0;

        //Git sync writes modules in the system's code page; other text (notes) may be UTF-8, which is tried first because it is rarely mistaken.
        private static string Decode(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
            {
                return string.Empty;
            }

            try
            {
                return new UTF8Encoding(false, true).GetString(bytes);
            }
            catch (ArgumentException)
            {
                return GitSyncFolderProcessor.SourceEncoding.GetString(bytes);
            }
        }

        private static RichTextBox Pane() => new RichTextBox
        {
            ReadOnly = true,
            WordWrap = false,
            DetectUrls = false,
            HideSelection = false,
            ScrollBars = RichTextBoxScrollBars.Both,
            Font = new Font(FontFamily.GenericMonospace, 9f),
            BackColor = Color.White,
            ForeColor = Color.Black
        };

        //One row per line on both sides, with blank rows where one side has fewer lines in a change, so that the rows line up.
        private void Fill()
        {
            var left = new StringBuilder();
            var right = new StringBuilder();
            var row = 0;
            var position = 0;
            var other = 0;

            void Add(string l, string r)
            {
                left.Append(l).Append('\n');
                right.Append(r).Append('\n');
                row++;
            }

            for (var h = 0; h < _diff.Hunks.Count; h++)
            {
                var hunk = _diff.Hunks[h];
                while (position < hunk.ModuleStart)
                {
                    Add(_diff.ModuleLines[position++], _diff.FileLines[other++]);
                }

                _firstRow[h] = row;
                var rows = Math.Max(hunk.ModuleCount, hunk.FileCount);
                for (var i = 0; i < rows; i++)
                {
                    Add(i < hunk.ModuleCount ? _diff.ModuleLines[hunk.ModuleStart + i] : string.Empty,
                        i < hunk.FileCount ? _diff.FileLines[hunk.FileStart + i] : string.Empty);
                }

                position = hunk.ModuleStart + hunk.ModuleCount;
                other = hunk.FileStart + hunk.FileCount;
            }

            while (position < _diff.ModuleLines.Count)
            {
                Add(_diff.ModuleLines[position++], _diff.FileLines[other++]);
            }

            _before.Text = left.ToString();
            _after.Text = right.ToString();
        }

        private void PaintAll()
        {
            foreach (var box in new[] { _before, _after })
            {
                SendMessage(box.Handle, WmSetRedraw, IntPtr.Zero, IntPtr.Zero);
            }

            try
            {
                for (var h = 0; h < _diff.Hunks.Count; h++)
                {
                    PaintHunk(h);
                }
            }
            finally
            {
                foreach (var box in new[] { _before, _after })
                {
                    SendMessage(box.Handle, WmSetRedraw, (IntPtr)1, IntPtr.Zero);
                    box.Invalidate();
                }
            }
        }

        private void PaintHunk(int h)
        {
            var hunk = _diff.Hunks[h];
            var rows = Math.Max(hunk.ModuleCount, hunk.FileCount);
            PaintRows(_before, h, hunk.ModuleCount, rows, RemovedBack);
            PaintRows(_after, h, hunk.FileCount, rows, AddedBack);

            for (var i = 0; i < Math.Min(hunk.ModuleCount, hunk.FileCount); i++)
            {
                foreach (var word in WinMergeWordDiff.Compute(_diff.ModuleLines[hunk.ModuleStart + i], _diff.FileLines[hunk.FileStart + i]))
                {
                    Mark(_before, _firstRow[h] + i, word.Begin0, word.End0, RemovedWord);
                    Mark(_after, _firstRow[h] + i, word.Begin1, word.End1, AddedWord);
                }
            }
        }

        private void PaintRows(RichTextBox box, int h, int lines, int rows, Color color)
        {
            var start = box.GetFirstCharIndexFromLine(_firstRow[h]);
            var end = _firstRow[h] + rows < box.GetLineFromCharIndex(box.TextLength) + 1 ? box.GetFirstCharIndexFromLine(_firstRow[h] + rows) : box.TextLength;
            box.Select(start, Math.Max(0, end - start));
            box.SelectionBackColor = lines == 0 ? GhostBack : h == _selected ? Blend(color, CurrentBack) : color;
            box.SelectionLength = 0;
        }

        private static Color Blend(Color a, Color b) => Color.FromArgb((a.R + b.R) / 2, (a.G + b.G) / 2, (a.B + b.B) / 2);

        private static void Mark(RichTextBox box, int row, int first, int last, Color color)
        {
            if (last < first)
            {
                return;
            }

            box.Select(box.GetFirstCharIndexFromLine(row) + first, last - first + 1);
            box.SelectionBackColor = color;
            box.SelectionLength = 0;
        }

        private void UpdateStatus(string beforeText = null, string afterText = null)
        {
            if (_diff.Hunks.Count == 0)
            {
                //Indentation and trailing space do not count in the comparison, so a file can have changed without a change to show.
                _status.Text = beforeText == afterText ? "No differences." : "Only indentation or spacing differs.";
            }
            else
            {
                _status.Text = $"Change {_selected + 1} of {_diff.Hunks.Count}";
            }
        }

        private void Goto(int index)
        {
            if (index < 0 || index >= _diff.Hunks.Count)
            {
                return;
            }

            var previous = _selected;
            _selected = index;
            PaintHunk(previous);
            PaintHunk(index);
            _status.Text = $"Change {index + 1} of {_diff.Hunks.Count}";

            _scrolling = true;
            try
            {
                var row = Math.Max(0, _firstRow[index] - 3);
                ScrollTo(_before, row);
                ScrollTo(_after, row);
            }
            finally
            {
                _scrolling = false;
            }
        }

        private static int FirstVisible(RichTextBox box) => SendMessage(box.Handle, EmGetFirstVisibleLine, IntPtr.Zero, IntPtr.Zero).ToInt32();

        private static void ScrollTo(RichTextBox box, int line)
            => SendMessage(box.Handle, EmLineScroll, IntPtr.Zero, (IntPtr)(Math.Max(0, line) - FirstVisible(box)));

        private void Scrolled(RichTextBox source, RichTextBox target)
        {
            if (_scrolling)
            {
                return;
            }

            _scrolling = true;
            try
            {
                ScrollTo(target, FirstVisible(source));
            }
            finally
            {
                _scrolling = false;
            }
        }
    }
}
