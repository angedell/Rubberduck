using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// Shows the project's module, the result and the folder's file side by side, in the manner of WinMerge's three-way merge view.
    /// What is the same on all three is black; what differs is colored. Clicking a differing block in the project or the folder puts it in
    /// the result, clicking it again takes it back, Ctrl+click takes both. The result can be edited by hand.
    /// </summary>
    public sealed class GitSyncCompareDialog : Form
    {
        //Differing blocks that still wait for a decision.
        private static readonly Color PendingBack = Color.FromArgb(255, 238, 170);
        private static readonly Color PendingText = Color.FromArgb(150, 60, 0);
        //Blocks that are in the result.
        private static readonly Color KeptBack = Color.FromArgb(198, 239, 206);
        private static readonly Color KeptText = Color.FromArgb(0, 100, 0);
        //The other side of a block that was decided.
        private static readonly Color DroppedBack = Color.FromArgb(225, 225, 225);
        private static readonly Color DroppedText = Color.FromArgb(110, 110, 110);
        //What the user wrote.
        private static readonly Color EditedBack = Color.FromArgb(214, 230, 255);
        private static readonly Color EditedText = Color.FromArgb(0, 40, 160);
        private static readonly Color GhostBack = Color.FromArgb(240, 240, 240);
        private static readonly Color CurrentBack = Color.FromArgb(255, 200, 120);
        private static readonly Color WordPending = Color.FromArgb(255, 190, 90);
        private static readonly Color WordKept = Color.FromArgb(120, 210, 140);

        private const int EmGetFirstVisibleLine = 0xCE;
        private const int EmLineScroll = 0xB6;
        private const int WmSetRedraw = 0x000B;

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);

        /// <summary>A stretch of the result: either lines all three share, or the place of one differing block.</summary>
        private class Segment
        {
            public int Hunk = -1;      //-1 for lines that are the same.
            public int ModuleLine;     //Where the shared lines start in the module.
            public int Count;          //How many shared lines.
            public int SideRow;        //Where the segment starts in the aligned project and folder panes,
            public int SideRows;       //and how many rows it takes there.
            public int Length;         //The characters it holds in the result right now.
            public bool Edited;        //Shared lines the user changed.
        }

        private readonly SyncDiff _diff;
        private readonly RichTextBox _module;
        private readonly RichTextBox _result;
        private readonly RichTextBox _file;
        private readonly Label _status;
        private readonly Label _captionResult;
        private readonly Label _captionFile;
        private readonly Label _legend;
        private readonly List<Segment> _segments = new List<Segment>();
        private readonly int[] _hunkSegment;
        private readonly int[] _firstRow;
        private readonly string _placeholder = CodeExplorerUI.GitSync_Compare_Placeholder + "\n";
        private int _selected;
        private string _lastText = string.Empty;
        private bool _programmatic;
        private bool _scrolling;

        private GitSyncCompareDialog(string name, SyncDiff diff, HunkChoice preferred)
        {
            _diff = diff;
            if (preferred != HunkChoice.Unresolved)
            {
                diff.ResolveAll(preferred);
            }

            Text = string.Format(CodeExplorerUI.GitSync_Compare_Title, name);
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(1280, 700);
            MinimumSize = new Size(900, 480);

            var captionModule = Caption(CodeExplorerUI.GitSync_Compare_Module);
            _captionResult = Caption(CodeExplorerUI.GitSync_Compare_Result);
            _captionFile = Caption(CodeExplorerUI.GitSync_Compare_File);
            _module = Pane(true);
            _result = Pane(false);
            _file = Pane(true);
            _result.AcceptsTab = true;

            _legend = new Label { Location = new Point(12, 612), Size = new Size(1256, 32), Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right, Text = CodeExplorerUI.GitSync_Compare_Legend };
            _status = new Label { Location = new Point(12, 646), Size = new Size(1256, 18), Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right };

            var bottom = 668;
            var x = 12;
            var buttons = new List<Control>();
            void Add(string text, int width, EventHandler click)
            {
                var button = Button(text, x, bottom, width, click);
                button.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
                buttons.Add(button);
                x += width + 4;
            }

            Add(CodeExplorerUI.GitSync_Compare_Prev, 130, (s, e) => Goto(_selected - 1, true));
            Add(CodeExplorerUI.GitSync_Compare_Next, 130, (s, e) => Goto(_selected + 1, true));
            x += 16;
            Add(CodeExplorerUI.GitSync_Compare_UseModule, 96, (s, e) => Resolve(HunkChoice.Module));
            Add(CodeExplorerUI.GitSync_Compare_UseFile, 96, (s, e) => Resolve(HunkChoice.File));
            Add(CodeExplorerUI.GitSync_Compare_UseBoth, 96, (s, e) => Resolve(HunkChoice.Both));
            x += 16;
            Add(CodeExplorerUI.GitSync_Compare_AllModule, 130, (s, e) => ResolveAll(HunkChoice.Module));
            Add(CodeExplorerUI.GitSync_Compare_AllFile, 130, (s, e) => ResolveAll(HunkChoice.File));

            var accept = new Button { Text = CodeExplorerUI.GitSync_Compare_Accept, Location = new Point(1100, bottom), Size = new Size(80, 24), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            accept.Click += (s, e) => Accept();
            var cancel = new Button { Text = CodeExplorerUI.BulkRemove_Cancel, Location = new Point(1188, bottom), Size = new Size(80, 24), Anchor = AnchorStyles.Bottom | AnchorStyles.Right, DialogResult = DialogResult.Cancel };
            CancelButton = cancel;

            Controls.AddRange(new Control[] { captionModule, _captionResult, _captionFile, _module, _result, _file, _legend, _status, accept, cancel });
            Controls.AddRange(buttons.ToArray());

            _hunkSegment = new int[diff.Hunks.Count];
            _firstRow = new int[diff.Hunks.Count];
            FillSides();
            BuildResult();
            LayoutPanes(captionModule);
            Resize += (s, e) => LayoutPanes(captionModule);

            _module.VScroll += (s, e) => SideScrolled(_module);
            _file.VScroll += (s, e) => SideScrolled(_file);
            _result.VScroll += (s, e) => ResultScrolled();
            _module.MouseMove += (s, e) => Hover(_module, e.Location);
            _file.MouseMove += (s, e) => Hover(_file, e.Location);
            _module.MouseClick += (s, e) => Clicked(_module, e, HunkChoice.Module);
            _file.MouseClick += (s, e) => Clicked(_file, e, HunkChoice.File);
            _result.TextChanged += (s, e) => ResultChanged();
            _result.MouseUp += (s, e) => CaretMoved();
            _result.KeyUp += (s, e) => CaretMoved();

            RepaintAll();
            UpdateStatus();
            Shown += (s, e) => Goto(FirstUnresolvedOr(0), true);
        }

        /// <summary>The text the user merged, or null if they cancelled.</summary>
        public string Merged { get; private set; }

        /// <summary>
        /// Opens the comparison. Every difference starts on the side the history favors (<paramref name="preferred"/>), or is left for the user
        /// when that is <see cref="HunkChoice.Unresolved"/>.
        /// </summary>
        public static string Ask(IWin32Window owner, string name, string moduleText, string fileText, HunkChoice preferred, bool ignoreCaseOutsideStrings = false)
        {
            using (var dialog = new GitSyncCompareDialog(name, new SyncDiff(moduleText, fileText, ignoreCaseOutsideStrings), preferred))
            {
                if (dialog._diff.Hunks.Count == 0)
                {
                    MessageBox.Show(owner, CodeExplorerUI.GitSync_Compare_NoDifferences, dialog.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return null;
                }

                dialog.ShowDialog(owner);
                return dialog.Merged;
            }
        }

        #region Building

        private static Label Caption(string text) => new Label { Text = text, Location = new Point(12, 10), Size = new Size(400, 18), Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold) };

        private static RichTextBox Pane(bool readOnly) => new RichTextBox
        {
            Location = new Point(12, 30),
            Size = new Size(400, 560),
            ReadOnly = readOnly,
            WordWrap = false,
            DetectUrls = false,
            HideSelection = false,
            ScrollBars = RichTextBoxScrollBars.Both,
            Font = new Font(FontFamily.GenericMonospace, 9f),
            BackColor = Color.White,
            ForeColor = Color.Black
        };

        private static Button Button(string text, int x, int y, int width, EventHandler click)
        {
            var button = new Button { Text = text, Location = new Point(x, y), Size = new Size(width, 24) };
            button.Click += click;
            return button;
        }

        private void LayoutPanes(Label captionModule)
        {
            var width = (ClientSize.Width - 48) / 3;
            var height = Math.Max(100, ClientSize.Height - 30 - 96);
            var boxes = new[] { _module, _result, _file };
            var captions = new[] { captionModule, _captionResult, _captionFile };
            for (var i = 0; i < 3; i++)
            {
                boxes[i].SetBounds(12 + i * (width + 12), 30, width, height);
                captions[i].SetBounds(boxes[i].Left, 10, width, 18);
            }
        }

        //The project and the folder get one row per line, with blank rows where one side has fewer lines in a block, so that the rows line up.
        private void FillSides()
        {
            var left = new StringBuilder();
            var right = new StringBuilder();
            var row = 0;
            var module = 0;
            var file = 0;

            void Add(string l, string r)
            {
                left.Append(l).Append('\n');
                right.Append(r).Append('\n');
                row++;
            }

            void Shared(int upTo)
            {
                if (module >= upTo)
                {
                    return;
                }

                _segments.Add(new Segment { ModuleLine = module, Count = upTo - module, SideRow = row, SideRows = upTo - module });
                while (module < upTo)
                {
                    Add(_diff.ModuleLines[module++], _diff.FileLines[file++]);
                }
            }

            for (var h = 0; h < _diff.Hunks.Count; h++)
            {
                var hunk = _diff.Hunks[h];
                Shared(hunk.ModuleStart);

                _firstRow[h] = row;
                var rows = Math.Max(hunk.ModuleCount, hunk.FileCount);
                _hunkSegment[h] = _segments.Count;
                _segments.Add(new Segment { Hunk = h, SideRow = row, SideRows = rows });
                for (var i = 0; i < rows; i++)
                {
                    Add(i < hunk.ModuleCount ? _diff.ModuleLines[hunk.ModuleStart + i] : string.Empty,
                        i < hunk.FileCount ? _diff.FileLines[hunk.FileStart + i] : string.Empty);
                }

                module = hunk.ModuleStart + hunk.ModuleCount;
                file = hunk.FileStart + hunk.FileCount;
            }

            Shared(_diff.ModuleLines.Count);

            _module.Text = left.ToString();
            _file.Text = right.ToString();
        }

        //What the result holds for a segment, line by line, each with its line break.
        private string TextOf(Segment segment)
        {
            string Lines(IList<string> lines, int start, int count) => string.Concat(lines.Skip(start).Take(count).Select(line => line + "\n"));

            if (segment.Hunk < 0)
            {
                return Lines(_diff.ModuleLines, segment.ModuleLine, segment.Count);
            }

            var hunk = _diff.Hunks[segment.Hunk];
            switch (hunk.Choice)
            {
                case HunkChoice.Module: return Lines(_diff.ModuleLines, hunk.ModuleStart, hunk.ModuleCount);
                case HunkChoice.File: return Lines(_diff.FileLines, hunk.FileStart, hunk.FileCount);
                case HunkChoice.Both: return Lines(_diff.ModuleLines, hunk.ModuleStart, hunk.ModuleCount) + Lines(_diff.FileLines, hunk.FileStart, hunk.FileCount);
                default: return _placeholder;
            }
        }

        //Fills the whole result from the choices made so far, which also drops anything typed by hand.
        private void BuildResult()
        {
            var text = new StringBuilder();
            foreach (var segment in _segments)
            {
                var part = TextOf(segment);
                segment.Length = part.Length;
                segment.Edited = false;
                text.Append(part);
            }

            _programmatic = true;
            try
            {
                _result.Clear();
                _result.Text = text.ToString();
                _lastText = _result.Text;
            }
            finally
            {
                _programmatic = false;
            }

            //The box may count line breaks differently from the builder; the lengths follow the box.
            var excess = _lastText.Length - _segments.Sum(s => s.Length);
            if (excess != 0 && _segments.Count > 0)
            {
                _segments[_segments.Count - 1].Length += excess;
            }
        }

        private int[] Starts()
        {
            var starts = new int[_segments.Count];
            var at = 0;
            for (var i = 0; i < starts.Length; i++)
            {
                starts[i] = at;
                at += _segments[i].Length;
            }

            return starts;
        }

        #endregion

        #region Painting

        private static void SetRedraw(Control box, bool on) => SendMessage(box.Handle, WmSetRedraw, (IntPtr)(on ? 1 : 0), IntPtr.Zero);

        private static Color Blend(Color a, Color b) => Color.FromArgb((a.R + b.R) / 2, (a.G + b.G) / 2, (a.B + b.B) / 2);

        private static int LineCount(RichTextBox box) => box.GetLineFromCharIndex(box.TextLength) + 1;

        private RichTextBox[] Boxes => new[] { _module, _result, _file };

        private void RepaintAll()
        {
            foreach (var box in Boxes)
            {
                SetRedraw(box, false);
            }

            try
            {
                var starts = Starts();
                for (var h = 0; h < _diff.Hunks.Count; h++)
                {
                    PaintSides(h);
                    PaintSegment(_hunkSegment[h], starts);
                }
            }
            finally
            {
                foreach (var box in Boxes)
                {
                    SetRedraw(box, true);
                    box.Invalidate();
                }
            }
        }

        //Paints the given segments of the result, and the matching blocks of the project and the folder. The cursor stays where it was.
        private void Repaint(IEnumerable<int> segments)
        {
            var saved = _result.SelectionStart;
            var savedLength = _result.SelectionLength;
            foreach (var box in Boxes)
            {
                SetRedraw(box, false);
            }

            _programmatic = true;
            try
            {
                var starts = Starts();
                foreach (var index in segments.Distinct())
                {
                    if (index < 0 || index >= _segments.Count)
                    {
                        continue;
                    }

                    if (_segments[index].Hunk >= 0)
                    {
                        PaintSides(_segments[index].Hunk);
                    }

                    PaintSegment(index, starts);
                }
            }
            finally
            {
                _result.Select(saved, savedLength);
                _programmatic = false;
                foreach (var box in Boxes)
                {
                    SetRedraw(box, true);
                    box.Invalidate();
                }
            }
        }

        private void RepaintHunks(params int[] hunks) => Repaint(hunks.Where(h => h >= 0 && h < _hunkSegment.Length).Select(h => _hunkSegment[h]));

        private static void ApplyColors(RichTextBox box, int start, int length, Color text, Color back)
        {
            box.Select(start, Math.Max(0, length));
            box.SelectionColor = text;
            box.SelectionBackColor = back;
            box.SelectionLength = 0;
        }

        //The rows of one block in the project and folder panes.
        private void PaintSides(int h)
        {
            var hunk = _diff.Hunks[h];
            var rows = Math.Max(hunk.ModuleCount, hunk.FileCount);
            PaintSide(_module, h, hunk.ModuleCount, rows, hunk.Choice == HunkChoice.Module || hunk.Choice == HunkChoice.Both, hunk.Choice == HunkChoice.File);
            PaintSide(_file, h, hunk.FileCount, rows, hunk.Choice == HunkChoice.File || hunk.Choice == HunkChoice.Both, hunk.Choice == HunkChoice.Module);
            HighlightWords(h, hunk);
        }

        private void PaintSide(RichTextBox box, int h, int lines, int rows, bool kept, bool dropped)
        {
            var start = box.GetFirstCharIndexFromLine(_firstRow[h]);
            var end = _firstRow[h] + rows < LineCount(box) ? box.GetFirstCharIndexFromLine(_firstRow[h] + rows) : box.TextLength;

            var text = kept ? KeptText : dropped ? DroppedText : PendingText;
            var back = kept ? KeptBack : dropped ? DroppedBack : PendingBack;
            if (lines == 0)
            {
                back = GhostBack;
            }
            else if (h == _selected)
            {
                back = Blend(back, CurrentBack);
            }

            ApplyColors(box, start, end - start, text, back);
        }

        private void PaintSegment(int index, int[] starts)
        {
            var segment = _segments[index];
            Color text, back;
            if (segment.Hunk < 0)
            {
                text = segment.Edited ? EditedText : Color.Black;
                back = segment.Edited ? EditedBack : Color.White;
            }
            else
            {
                switch (_diff.Hunks[segment.Hunk].Choice)
                {
                    case HunkChoice.Custom:
                        text = EditedText;
                        back = EditedBack;
                        break;
                    case HunkChoice.Unresolved:
                        text = DroppedText;
                        back = PendingBack;
                        break;
                    default:
                        text = KeptText;
                        back = KeptBack;
                        break;
                }

                if (segment.Hunk == _selected)
                {
                    back = Blend(back, CurrentBack);
                }
            }

            ApplyColors(_result, starts[index], segment.Length, text, back);
        }

        //Within lines that were changed (rather than added or removed) the words that differ are marked more strongly, as WinMerge does.
        private void HighlightWords(int index, DiffHunk hunk)
        {
            for (var i = 0; i < Math.Min(hunk.ModuleCount, hunk.FileCount); i++)
            {
                var row = _firstRow[index] + i;
                foreach (var diff in WinMergeWordDiff.Compute(_diff.Fold(_diff.ModuleLines[hunk.ModuleStart + i]), _diff.Fold(_diff.FileLines[hunk.FileStart + i])))
                {
                    if (hunk.Choice != HunkChoice.File)
                    {
                        Mark(_module, row, diff.Begin0, diff.End0, hunk.Choice == HunkChoice.Module || hunk.Choice == HunkChoice.Both ? WordKept : WordPending);
                    }

                    if (hunk.Choice != HunkChoice.Module)
                    {
                        Mark(_file, row, diff.Begin1, diff.End1, hunk.Choice == HunkChoice.File || hunk.Choice == HunkChoice.Both ? WordKept : WordPending);
                    }
                }
            }
        }

        private static void Mark(RichTextBox box, int row, int first, int last, Color color)
        {
            if (last < first)
            {
                return; //Nothing on this side.
            }

            box.Select(box.GetFirstCharIndexFromLine(row) + first, last - first + 1);
            box.SelectionBackColor = color;
            box.SelectionLength = 0;
        }

        private void UpdateStatus()
        {
            var resolved = _diff.Hunks.Count(h => h.Choice != HunkChoice.Unresolved);
            _status.Text = string.Format(CodeExplorerUI.GitSync_Compare_Status, _selected + 1, _diff.Hunks.Count, resolved);
        }

        #endregion

        #region Choosing

        private int FirstUnresolvedOr(int fallback)
        {
            var first = _diff.Hunks.Select((h, i) => new { h, i }).FirstOrDefault(x => x.h.Choice == HunkChoice.Unresolved);
            return first?.i ?? fallback;
        }

        private int HunkAtRow(int row)
        {
            for (var h = 0; h < _diff.Hunks.Count; h++)
            {
                if (row >= _firstRow[h] && row < _firstRow[h] + Math.Max(_diff.Hunks[h].ModuleCount, _diff.Hunks[h].FileCount))
                {
                    return h;
                }
            }

            return -1;
        }

        private int HunkAt(RichTextBox box, Point point)
        {
            if (box.TextLength == 0)
            {
                return -1;
            }

            return HunkAtRow(box.GetLineFromCharIndex(box.GetCharIndexFromPosition(point)));
        }

        //The hand marks what can be clicked.
        private void Hover(RichTextBox box, Point point)
        {
            box.Cursor = HunkAt(box, point) >= 0 ? Cursors.Hand : Cursors.IBeam;
        }

        private void Clicked(RichTextBox box, MouseEventArgs e, HunkChoice side)
        {
            if (e.Button != MouseButtons.Left || box.SelectionLength > 0)
            {
                return; //Selecting text to copy it is not a pick.
            }

            var h = HunkAt(box, e.Location);
            if (h >= 0)
            {
                Pick(h, (ModifierKeys & Keys.Control) != 0 ? HunkChoice.Both : side);
            }
        }

        //Puts a block in the result, or takes it out again if that is what is there already.
        private void Pick(int h, HunkChoice choice)
        {
            var hunk = _diff.Hunks[h];
            if (hunk.Choice == HunkChoice.Custom && !ConfirmDiscard())
            {
                return;
            }

            var previous = _selected;
            hunk.Choice = hunk.Choice == choice ? HunkChoice.Unresolved : choice;
            _selected = h;
            ReplaceSegment(_hunkSegment[h]);
            RepaintHunks(h, previous);
            UpdateStatus();
        }

        private bool ConfirmDiscard() =>
            MessageBox.Show(this, CodeExplorerUI.GitSync_Compare_DiscardEdits, Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;

        private void ReplaceSegment(int index)
        {
            var segment = _segments[index];
            var start = Starts()[index];
            var before = _result.TextLength;
            _programmatic = true;
            try
            {
                _result.Select(start, segment.Length);
                _result.SelectedText = TextOf(segment);
                _lastText = _result.Text;
                segment.Length += _result.TextLength - before;
                segment.Edited = false;
            }
            finally
            {
                _programmatic = false;
            }
        }

        //Decides the block that is current, then moves on to the next one that is open.
        private void Resolve(HunkChoice choice)
        {
            if (_selected < 0 || _selected >= _diff.Hunks.Count || _diff.Hunks[_selected].Choice == choice)
            {
                return; //Unlike a click, a button does not take a choice back.
            }

            Pick(_selected, choice);
            var next = _diff.Hunks.Select((h, i) => new { h, i }).FirstOrDefault(x => x.i > _selected && x.h.Choice == HunkChoice.Unresolved);
            if (next != null)
            {
                Goto(next.i, true);
            }
        }

        private void ResolveAll(HunkChoice choice)
        {
            var edited = _diff.Hunks.Any(h => h.Choice == HunkChoice.Custom) || _segments.Any(s => s.Edited);
            if (edited && !ConfirmDiscard())
            {
                return;
            }

            _diff.ResolveAll(choice);
            BuildResult();
            RepaintAll();
            UpdateStatus();
        }

        private void Goto(int index, bool scroll)
        {
            if (index < 0 || index >= _diff.Hunks.Count)
            {
                return;
            }

            var previous = _selected;
            _selected = index;
            RepaintHunks(previous, index);
            UpdateStatus();

            if (scroll)
            {
                _scrolling = true;
                try
                {
                    var row = Math.Max(0, _firstRow[index] - 3);
                    ScrollTo(_module, row);
                    ScrollTo(_file, row);
                    ScrollTo(_result, ResultLineForRow(_firstRow[index]) - Math.Min(3, _firstRow[index]));
                }
                finally
                {
                    _scrolling = false;
                }
            }
        }

        private void Accept()
        {
            if (_diff.Hunks.Any(h => h.Choice == HunkChoice.Unresolved) || _result.Text.Contains(_placeholder.TrimEnd('\n')))
            {
                MessageBox.Show(this, CodeExplorerUI.GitSync_Compare_Unresolved, Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                Goto(FirstUnresolvedOr(_selected), true);
                return;
            }

            //What is in the result is what is applied; every line had its break added when it was put there.
            var lines = _result.Text.Split('\n').ToList();
            if (lines.Count > 1 && lines[lines.Count - 1].Length == 0)
            {
                lines.RemoveAt(lines.Count - 1);
            }

            Merged = string.Join("\r\n", lines);
            Close();
        }

        #endregion

        #region Editing the result

        //Works out what the user changed from the text before and after, and keeps track of which block it was in.
        private void ResultChanged()
        {
            if (_programmatic)
            {
                return;
            }

            var text = _result.Text;
            if (text == _lastText)
            {
                return;
            }

            var old = _lastText;
            var prefix = 0;
            var limit = Math.Min(old.Length, text.Length);
            while (prefix < limit && old[prefix] == text[prefix])
            {
                prefix++;
            }

            var suffix = 0;
            limit -= prefix;
            while (suffix < limit && old[old.Length - 1 - suffix] == text[text.Length - 1 - suffix])
            {
                suffix++;
            }

            var removedEnd = old.Length - suffix; //The old text from prefix to here is gone,
            var inserted = text.Length - suffix - prefix; //and this many characters took its place.
            _lastText = text;

            var starts = Starts();
            var owner = -1;
            for (var i = 0; i < _segments.Count; i++)
            {
                if (starts[i] <= prefix && prefix < starts[i] + _segments[i].Length)
                {
                    owner = i;
                    break;
                }
            }

            for (var i = _segments.Count - 1; owner < 0 && i >= 0; i--)
            {
                if (starts[i] <= prefix)
                {
                    owner = i;
                }
            }

            if (owner < 0)
            {
                return;
            }

            var touched = new List<int>();
            var wasUnresolved = _segments[owner].Hunk >= 0 && _diff.Hunks[_segments[owner].Hunk].Choice == HunkChoice.Unresolved;
            for (var i = 0; i < _segments.Count; i++)
            {
                var segment = _segments[i];
                var overlap = Math.Max(0, Math.Min(starts[i] + segment.Length, removedEnd) - Math.Max(starts[i], prefix));
                segment.Length -= overlap;
                if (i == owner)
                {
                    segment.Length += inserted;
                }

                if (overlap > 0 || i == owner)
                {
                    touched.Add(i);
                    if (segment.Hunk >= 0)
                    {
                        _diff.Hunks[segment.Hunk].Choice = HunkChoice.Custom;
                    }
                    else
                    {
                        segment.Edited = true;
                    }
                }
            }

            //Typing into a block that was not decided yet replaces the reminder that stood there.
            if (wasUnresolved)
            {
                DropPlaceholder(owner);
            }

            Repaint(touched);
            UpdateStatus();
        }

        private void DropPlaceholder(int index)
        {
            var segment = _segments[index];
            var start = Starts()[index];
            var current = _result.Text.Substring(start, segment.Length);
            var cleaned = current.Replace(_placeholder, string.Empty);
            if (cleaned == current)
            {
                return;
            }

            var caret = _result.SelectionStart;
            var before = _result.TextLength;
            _programmatic = true;
            try
            {
                _result.Select(start, segment.Length);
                _result.SelectedText = cleaned;
                segment.Length += _result.TextLength - before;
                _lastText = _result.Text;
                _result.Select(Math.Min(Math.Max(start, caret - _placeholder.Length), _result.TextLength), 0);
            }
            finally
            {
                _programmatic = false;
            }
        }

        //The block the cursor sits in becomes the current one.
        private void CaretMoved()
        {
            var starts = Starts();
            var caret = _result.SelectionStart;
            for (var i = 0; i < _segments.Count; i++)
            {
                if (starts[i] <= caret && caret < starts[i] + _segments[i].Length)
                {
                    if (_segments[i].Hunk >= 0 && _segments[i].Hunk != _selected)
                    {
                        Goto(_segments[i].Hunk, false);
                    }

                    return;
                }
            }
        }

        #endregion

        #region Scrolling

        private static int FirstVisible(RichTextBox box) => SendMessage(box.Handle, EmGetFirstVisibleLine, IntPtr.Zero, IntPtr.Zero).ToInt32();

        private static void ScrollTo(RichTextBox box, int line)
        {
            SendMessage(box.Handle, EmLineScroll, IntPtr.Zero, (IntPtr)(Math.Max(0, line) - FirstVisible(box)));
        }

        //The project and the folder move together row by row; the result follows block by block, because it may have other line counts.
        private void SideScrolled(RichTextBox source)
        {
            if (_scrolling)
            {
                return;
            }

            _scrolling = true;
            try
            {
                var row = FirstVisible(source);
                ScrollTo(source == _module ? _file : _module, row);
                ScrollTo(_result, ResultLineForRow(row));
            }
            finally
            {
                _scrolling = false;
            }
        }

        private void ResultScrolled()
        {
            if (_scrolling)
            {
                return;
            }

            _scrolling = true;
            try
            {
                var row = RowForResultLine(FirstVisible(_result));
                ScrollTo(_module, row);
                ScrollTo(_file, row);
            }
            finally
            {
                _scrolling = false;
            }
        }

        private int ResultLineOf(int character) => _result.GetLineFromCharIndex(Math.Min(character, _result.TextLength));

        private int ResultLineForRow(int row)
        {
            var low = 0;
            var high = _segments.Count - 1;
            var found = 0;
            while (low <= high)
            {
                var middle = (low + high) / 2;
                if (_segments[middle].SideRow <= row)
                {
                    found = middle;
                    low = middle + 1;
                }
                else
                {
                    high = middle - 1;
                }
            }

            var starts = Starts();
            var first = ResultLineOf(starts[found]);
            var lines = ResultLineOf(starts[found] + _segments[found].Length) - first;
            return first + Math.Min(row - _segments[found].SideRow, Math.Max(0, lines - 1));
        }

        private int RowForResultLine(int line)
        {
            var starts = Starts();
            var low = 0;
            var high = _segments.Count - 1;
            var found = 0;
            while (low <= high)
            {
                var middle = (low + high) / 2;
                if (ResultLineOf(starts[middle]) <= line)
                {
                    found = middle;
                    low = middle + 1;
                }
                else
                {
                    high = middle - 1;
                }
            }

            var segment = _segments[found];
            return segment.SideRow + Math.Min(Math.Max(0, line - ResultLineOf(starts[found])), Math.Max(0, segment.SideRows - 1));
        }

        #endregion
    }
}
