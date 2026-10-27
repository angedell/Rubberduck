using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>A stretch of two lines that differs. The ends are inclusive, so a side where nothing is there has End = Begin - 1.</summary>
    public struct WordDiff
    {
        public int Begin0;
        public int End0;
        public int Begin1;
        public int End1;
    }

    public enum WhitespaceMode
    {
        /// <summary>Every difference in white space counts.</summary>
        CompareAll,
        /// <summary>A run of white space equals any other run.</summary>
        IgnoreChange,
        /// <summary>White space does not count at all.</summary>
        IgnoreAll
    }

    /// <summary>
    /// The way WinMerge finds what differs within a changed line (Src/stringdiffs.cpp, ComputeWordDiffs): the lines are cut into words,
    /// white space and punctuation, the words are compared with the O(NP) algorithm of Sun Wu, Udi Manber and Gene Myers, and neighbouring
    /// differences are joined. Ported from C++ to C#; the case, end-of-line, number and byte level options are left out.
    /// </summary>
    /// <remarks>
    /// Derived from WinMerge (https://github.com/winmerge/winmerge, Src/stringdiffs.cpp). WinMerge is licensed under the GNU General Public
    /// License, version 2 or (at your option) any later version.
    /// </remarks>
    public class WinMergeWordDiff
    {
        private const string BreakChars = ",.;:";
        private const int TimeoutMilliseconds = 500;
        private const int CountMax = 100000;

        private enum Kind
        {
            Word,
            Space,
            Break
        }

        private struct Word
        {
            public int Start;
            public int End;
            public Kind Kind;
            public uint Hash;
            public int Length => End + 1 - Start;
        }

        private struct EditStep
        {
            public char Op;
            public int Equal;
            public int PreviousK;
            public int PreviousIndex;
        }

        private readonly string _first;
        private readonly string _second;
        private readonly WhitespaceMode _whitespace;
        private readonly bool _breakOnPunctuation;
        private Word[] _words1;
        private Word[] _words2;
        private readonly List<WordDiff> _diffs = new List<WordDiff>();

        private WinMergeWordDiff(string first, string second, WhitespaceMode whitespace, bool breakOnPunctuation)
        {
            _first = first ?? string.Empty;
            _second = second ?? string.Empty;
            _whitespace = whitespace;
            _breakOnPunctuation = breakOnPunctuation;
        }

        /// <summary>The differences between two lines, with neighbouring ones joined.</summary>
        public static List<WordDiff> Compute(string first, string second, WhitespaceMode whitespace = WhitespaceMode.IgnoreChange, bool breakOnPunctuation = true)
        {
            var worker = new WinMergeWordDiff(first, second, whitespace, breakOnPunctuation);
            worker.BuildWordDiffList();
            return worker.Populate();
        }

        private bool IsSpace(Word word) => word.Kind == Kind.Space;

        private void BuildWordDiffList()
        {
            _words1 = BuildWords(_first);
            _words2 = BuildWords(_second);

            if (!(_words1.Length < 20480 && _words2.Length < 20480 && BuildFromEditScript()))
            {
                //Too much to compare: the whole line counts as one difference.
                _diffs.Clear();
                _diffs.Add(new WordDiff
                {
                    Begin0 = _words1[0].Start,
                    End0 = _words1[_words1.Length - 1].End,
                    Begin1 = _words2[0].Start,
                    End1 = _words2[_words2.Length - 1].End
                });
            }
        }

        private bool BuildFromEditScript()
        {
            var script = new List<char>();
            if (Onp(script) < 0)
            {
                return false;
            }

            int i = 1, j = 1;
            foreach (var op in script)
            {
                if (op == '-')
                {
                    if (_whitespace == WhitespaceMode.IgnoreAll && IsSpace(_words1[i]))
                    {
                        i++;
                        continue;
                    }

                    var s2 = _words2[j - 1].End + 1;
                    _diffs.Add(new WordDiff { Begin0 = _words1[i].Start, End0 = _words1[i].End, Begin1 = s2, End1 = s2 - 1 });
                    i++;
                }
                else if (op == '+')
                {
                    if (_whitespace == WhitespaceMode.IgnoreAll && IsSpace(_words2[j]))
                    {
                        j++;
                        continue;
                    }

                    var s1 = _words1[i - 1].End + 1;
                    _diffs.Add(new WordDiff { Begin0 = s1, End0 = s1 - 1, Begin1 = _words2[j].Start, End1 = _words2[j].End });
                    j++;
                }
                else if (op == '!')
                {
                    if (_whitespace != WhitespaceMode.CompareAll && IsSpace(_words1[i]) && IsSpace(_words2[j]))
                    {
                        i++;
                        j++;
                        continue;
                    }

                    _diffs.Add(new WordDiff { Begin0 = _words1[i].Start, End0 = _words1[i].End, Begin1 = _words2[j].Start, End1 = _words2[j].End });
                    i++;
                    j++;
                }
                else
                {
                    i++;
                    j++;
                }
            }

            return true;
        }

        //Cuts a line into words, white space and punctuation. Index 0 is a dummy, as in the original, so that words count from 1.
        private Word[] BuildWords(string text)
        {
            var words = new List<Word> { new Word { Start = 0, End = -1, Kind = Kind.Word } };
            if (text.Length == 0)
            {
                return words.ToArray();
            }

            var begin = 0;
            var previous = Kind.Word;
            for (var i = 0; i < text.Length; i++)
            {
                var ch = text[i];
                var kind = char.IsWhiteSpace(ch) ? Kind.Space
                    : IsWordBreak(ch) ? Kind.Break
                    : Kind.Word;

                //A new word starts where the kind changes, and every punctuation mark is a word of its own.
                if (i > 0 && (kind != previous || kind == Kind.Break))
                {
                    words.Add(MakeWord(text, begin, i - 1, previous));
                    begin = i;
                }

                previous = kind;
            }

            words.Add(MakeWord(text, begin, text.Length - 1, previous));
            return words.ToArray();
        }

        private bool IsWordBreak(char ch)
        {
            if (ch <= 0xff)
            {
                return _breakOnPunctuation && BreakChars.IndexOf(ch) >= 0;
            }

            //Outside Latin-1 anything that is not a letter or a digit breaks.
            return !char.IsLetterOrDigit(ch);
        }

        private static Word MakeWord(string text, int begin, int end, Kind kind)
        {
            //The diffutils hash.
            uint h = 0;
            for (var i = begin; i <= end; i++)
            {
                h += text[i] + ((h << 7) | (h >> 25));
            }

            return new Word { Start = begin, End = end, Kind = kind, Hash = h };
        }

        private bool AreWordsSame(Word word1, Word word2)
        {
            if (_whitespace != WhitespaceMode.CompareAll && IsSpace(word1) && IsSpace(word2))
            {
                return true;
            }

            if (word1.Hash != word2.Hash || word1.Length != word2.Length)
            {
                return false;
            }

            return string.CompareOrdinal(_first, word1.Start, _second, word2.Start, word1.Length) == 0;
        }

        private int Snake(int k, int y, int m, int n, bool exchanged)
        {
            var x = y - k;
            if (exchanged)
            {
                while (x < m && y < n && AreWordsSame(_words1[y + 1], _words2[x + 1]))
                {
                    x++;
                    y++;
                }
            }
            else
            {
                while (x < m && y < n && AreWordsSame(_words1[x + 1], _words2[y + 1]))
                {
                    x++;
                    y++;
                }
            }

            return y;
        }

        /// <summary>An O(NP) sequence comparison, by Sun Wu, Udi Manber and Gene Myers. Returns the number of edits, or -1 when it takes too long.</summary>
        private int Onp(List<char> script)
        {
            var timer = Stopwatch.StartNew();

            var m = _words1.Length - 1;
            var n = _words2.Length - 1;
            var exchanged = m > n;
            if (exchanged)
            {
                var swap = m;
                m = n;
                n = swap;
            }

            //Diagonals run from -(M+1) to N+1; the arrays are shifted by M + 1.
            var shift = m + 1;
            var fp = new int[(m + 1) + 1 + (n + 1)];
            var es = new List<EditStep>[(m + 1) + 1 + (n + 1)];
            for (var i = 0; i < es.Length; i++)
            {
                es[i] = new List<EditStep>();
            }

            var delta = n - m;

            void AddEditStep(int k)
            {
                var step = new EditStep();
                if (fp[k - 1 + shift] + 1 > fp[k + 1 + shift])
                {
                    step.Op = '+';
                    step.Equal = fp[k + shift] - (fp[k - 1 + shift] + 1);
                    step.PreviousK = k - 1;
                }
                else
                {
                    step.Op = '-';
                    step.Equal = fp[k + shift] - fp[k + 1 + shift];
                    step.PreviousK = k + 1;
                }

                step.PreviousIndex = es[step.PreviousK + shift].Count - 1;
                es[k + shift].Add(step);
            }

            int Reach(int k) => Snake(k, Math.Max(fp[k - 1 + shift] + 1, fp[k + 1 + shift]), m, n, exchanged);

            var count = 0;
            for (var k = -(m + 1); k <= n + 1; k++)
            {
                fp[k + shift] = -1;
            }

            var p = -1;
            int last;
            do
            {
                p++;
                for (var k = -p; k <= delta - 1; k++)
                {
                    fp[k + shift] = Reach(k);
                    AddEditStep(k);
                    count++;
                }

                for (var k = delta + p; k >= delta + 1; k--)
                {
                    fp[k + shift] = Reach(k);
                    AddEditStep(k);
                    count++;
                }

                fp[delta + shift] = Reach(delta);
                AddEditStep(delta);
                count++;
                last = fp[delta + shift];

                if (count > CountMax)
                {
                    count = 0;
                    if (timer.ElapsedMilliseconds > TimeoutMilliseconds)
                    {
                        return -1;
                    }
                }
            }
            while (last != n);

            //Walk the steps back from the end to get the edit script.
            var ses = new List<char>();
            var at = delta;
            for (var i = es[delta + shift].Count - 1; i >= 0;)
            {
                var step = es[at + shift][i];
                for (var e = 0; e < step.Equal; e++)
                {
                    ses.Add('=');
                }

                ses.Add(step.Op);
                i = step.PreviousIndex;
                at = step.PreviousK;
            }

            ses.Reverse();

            //A deletion next to an insertion is a change.
            var edits = 0;
            for (int index = 1, total = ses.Count; index < total; index++)
            {
                var c = '!';
                var op = ses[index];
                var isPlus = op == '+';
                if (isPlus || op == '-')
                {
                    if (index != total - 1 && ses[index + 1] == (isPlus ? '-' : '+'))
                    {
                        index++;
                    }
                    else
                    {
                        c = exchanged == isPlus ? '-' : '+';
                    }

                    edits++;
                }
                else
                {
                    c = '=';
                }

                script.Add(c);
            }

            return edits;
        }

        //Joins differences that touch, so that "a b" changed to "c d" is one difference rather than three.
        private List<WordDiff> Populate()
        {
            var result = new List<WordDiff>();
            for (var i = 0; i < _diffs.Count; i++)
            {
                var diff = _diffs[i];
                if (i + 1 < _diffs.Count && diff.End0 + 1 == _diffs[i + 1].Begin0 && diff.End1 + 1 == _diffs[i + 1].Begin1)
                {
                    var next = _diffs[i + 1];
                    next.Begin0 = diff.Begin0;
                    next.Begin1 = diff.Begin1;
                    _diffs[i + 1] = next;
                    continue;
                }

                result.Add(diff);
            }

            return result;
        }
    }
}
