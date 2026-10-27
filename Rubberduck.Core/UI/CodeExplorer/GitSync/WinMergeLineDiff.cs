using System;
using System.Collections.Generic;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>One run of changed lines: DELETED lines of the first file starting at Line0 were replaced by INSERTED lines of the second starting at Line1.</summary>
    public struct LineChange
    {
        public int Line0;
        public int Line1;
        public int Deleted;
        public int Inserted;
    }

    /// <summary>
    /// The line comparison of WinMerge, which it has from GNU diffutils (Src/diffutils/src/analyze.c): Eugene Myers' O(ND) algorithm
    /// with Paul Eggert's TOO_EXPENSIVE heuristic, the removal of lines that have no counterpart ("confusing lines"), and the shifting of
    /// change boundaries so that neighbouring changes join. Ported from C to C#.
    /// </summary>
    /// <remarks>
    /// Derived from WinMerge (https://github.com/winmerge/winmerge, Src/diffutils/src/analyze.c), which is GNU diffutils:
    /// Copyright (C) 1988, 1989, 1992, 1993 Free Software Foundation, Inc. Licensed under the GNU General Public License,
    /// version 2 or (at your option) any later version.
    /// </remarks>
    public class WinMergeLineDiff
    {
        private const int SnakeLimit = 20; //Snakes bigger than this are considered 'big'.

        private class FileData
        {
            public int[] Equivs;        //The equivalence class of each line.
            public int Lines;
            public int[] Undiscarded;   //The lines that are really compared, as their classes.
            public int[] RealIndexes;   //From the undiscarded lines to the real ones.
            public int Nondiscarded;
            public bool[] Changed;      //Index + 1: there is a slot before the first line and one after the last.
        }

        private readonly FileData[] _files = { new FileData(), new FileData() };
        private int[] _xvec, _yvec;
        private int[] _fdiag, _bdiag;
        private int _diagOffset;
        private int _tooExpensive;
        private readonly bool _heuristic;
        private readonly bool _noDiscards;

        private WinMergeLineDiff(bool heuristic, bool noDiscards)
        {
            _heuristic = heuristic;
            _noDiscards = noDiscards;
        }

        /// <summary>
        /// Compares two lists of lines. Lines are the same when <paramref name="key"/> gives the same text for them.
        /// </summary>
        /// <param name="heuristic">Lets large files with many changes be compared faster, at the price of a longer script (diff's --speed-large-files).</param>
        /// <param name="noDiscards">Compares every line, even those without a counterpart (diff's --minimal behavior for discards).</param>
        public static List<LineChange> Compare(IList<string> first, IList<string> second, Func<string, string> key = null, bool heuristic = false, bool noDiscards = false)
        {
            key = key ?? (line => line);
            var worker = new WinMergeLineDiff(heuristic, noDiscards);
            return worker.Run(first, second, key);
        }

        private List<LineChange> Run(IList<string> first, IList<string> second, Func<string, string> key)
        {
            //Lines are given a number per kind (the equivalence classes of io.c); 0 is not used.
            var classes = new Dictionary<string, int>(StringComparer.Ordinal);
            int[] Number(IList<string> lines)
            {
                var result = new int[lines.Count];
                for (var i = 0; i < result.Length; i++)
                {
                    var text = key(lines[i]);
                    if (!classes.TryGetValue(text, out var id))
                    {
                        id = classes.Count + 1;
                        classes.Add(text, id);
                    }

                    result[i] = id;
                }

                return result;
            }

            _files[0].Equivs = Number(first);
            _files[1].Equivs = Number(second);
            _files[0].Lines = first.Count;
            _files[1].Lines = second.Count;
            var equivMax = classes.Count + 1;

            foreach (var file in _files)
            {
                file.Changed = new bool[file.Lines + 2];
            }

            DiscardConfusingLines(equivMax);

            _xvec = _files[0].Undiscarded;
            _yvec = _files[1].Undiscarded;
            var diags = _files[0].Nondiscarded + _files[1].Nondiscarded + 3;
            _fdiag = new int[diags];
            _bdiag = new int[diags];
            _diagOffset = _files[1].Nondiscarded + 1;

            //TOO_EXPENSIVE is the approximate square root of the input size, bounded below by 4096.
            _tooExpensive = 1;
            for (var i = _files[0].Nondiscarded + _files[1].Nondiscarded; i != 0; i >>= 2)
            {
                _tooExpensive <<= 1;
            }

            _tooExpensive = Math.Max(4096, _tooExpensive);

            CompareSeq(0, _files[0].Nondiscarded, 0, _files[1].Nondiscarded, _noDiscards);

            ShiftBoundaries();
            return BuildScript();
        }

        private bool Changed(int file, int index) => index >= -1 && _files[file].Changed[index + 1];
        private void SetChanged(int file, int index, bool value) => _files[file].Changed[index + 1] = value;

        /// <summary>
        /// Finds the midpoint of the shortest edit script for a portion of the two files by searching from both ends at once.
        /// Returns the approximate edit cost.
        /// </summary>
        private int Diag(int xoff, int xlim, int yoff, int ylim, bool minimal, out int xmid, out int ymid, out bool loMinimal, out bool hiMinimal)
        {
            var fd = _fdiag;
            var bd = _bdiag;
            var o = _diagOffset;
            var xv = _xvec;
            var yv = _yvec;
            var dmin = xoff - ylim; //Minimum valid diagonal.
            var dmax = xlim - yoff; //Maximum valid diagonal.
            var fmid = xoff - yoff; //Center diagonal of the top-down search.
            var bmid = xlim - ylim; //Center diagonal of the bottom-up search.
            int fmin = fmid, fmax = fmid;
            int bmin = bmid, bmax = bmid;
            var odd = ((fmid - bmid) & 1) != 0;

            fd[fmid + o] = xoff;
            bd[bmid + o] = xlim;

            for (var c = 1; ; ++c)
            {
                var bigSnake = false;

                //Extend the top-down search by an edit step in each diagonal.
                if (fmin > dmin) { fd[--fmin - 1 + o] = -1; } else { ++fmin; }
                if (fmax < dmax) { fd[++fmax + 1 + o] = -1; } else { --fmax; }
                for (var d = fmax; d >= fmin; d -= 2)
                {
                    int tlo = fd[d - 1 + o], thi = fd[d + 1 + o];
                    var x = tlo >= thi ? tlo + 1 : thi;
                    var oldx = x;
                    var y = x - d;
                    while (x < xlim && y < ylim && xv[x] == yv[y])
                    {
                        ++x;
                        ++y;
                    }

                    if (x - oldx > SnakeLimit)
                    {
                        bigSnake = true;
                    }

                    fd[d + o] = x;
                    if (odd && bmin <= d && d <= bmax && bd[d + o] <= x)
                    {
                        xmid = x;
                        ymid = y;
                        loMinimal = hiMinimal = true;
                        return 2 * c - 1;
                    }
                }

                //Similarly extend the bottom-up search.
                if (bmin > dmin) { bd[--bmin - 1 + o] = int.MaxValue; } else { ++bmin; }
                if (bmax < dmax) { bd[++bmax + 1 + o] = int.MaxValue; } else { --bmax; }
                for (var d = bmax; d >= bmin; d -= 2)
                {
                    int tlo = bd[d - 1 + o], thi = bd[d + 1 + o];
                    var x = tlo < thi ? tlo : thi - 1;
                    var oldx = x;
                    var y = x - d;
                    while (x > xoff && y > yoff && xv[x - 1] == yv[y - 1])
                    {
                        --x;
                        --y;
                    }

                    if (oldx - x > SnakeLimit)
                    {
                        bigSnake = true;
                    }

                    bd[d + o] = x;
                    if (!odd && fmin <= d && d <= fmax && x <= fd[d + o])
                    {
                        xmid = x;
                        ymid = y;
                        loMinimal = hiMinimal = true;
                        return 2 * c;
                    }
                }

                if (minimal)
                {
                    continue;
                }

                //Heuristic: check occasionally for a diagonal that has made lots of progress compared with the edit distance.
                if (c > 200 && bigSnake && _heuristic)
                {
                    var best = 0;
                    xmid = ymid = 0;
                    for (var d = fmax; d >= fmin; d -= 2)
                    {
                        var dd = d - fmid;
                        var x = fd[d + o];
                        var y = x - d;
                        var v = (x - xoff) * 2 - dd;
                        if (v > 12 * (c + (dd < 0 ? -dd : dd)))
                        {
                            if (v > best && xoff + SnakeLimit <= x && x < xlim && yoff + SnakeLimit <= y && y < ylim)
                            {
                                //A good enough best diagonal; now insist that it ends with a significant snake.
                                for (var k = 1; xv[x - k] == yv[y - k]; k++)
                                {
                                    if (k == SnakeLimit)
                                    {
                                        best = v;
                                        xmid = x;
                                        ymid = y;
                                        break;
                                    }
                                }
                            }
                        }
                    }

                    if (best > 0)
                    {
                        loMinimal = true;
                        hiMinimal = false;
                        return 2 * c - 1;
                    }

                    best = 0;
                    for (var d = bmax; d >= bmin; d -= 2)
                    {
                        var dd = d - bmid;
                        var x = bd[d + o];
                        var y = x - d;
                        var v = (xlim - x) * 2 + dd;
                        if (v > 12 * (c + (dd < 0 ? -dd : dd)))
                        {
                            if (v > best && xoff < x && x <= xlim - SnakeLimit && yoff < y && y <= ylim - SnakeLimit)
                            {
                                for (var k = 0; xv[x + k] == yv[y + k]; k++)
                                {
                                    if (k == SnakeLimit - 1)
                                    {
                                        best = v;
                                        xmid = x;
                                        ymid = y;
                                        break;
                                    }
                                }
                            }
                        }
                    }

                    if (best > 0)
                    {
                        loMinimal = false;
                        hiMinimal = true;
                        return 2 * c - 1;
                    }
                }

                //Heuristic: if we have gone well beyond the call of duty, give up and report halfway between the best results so far.
                if (c >= _tooExpensive)
                {
                    int fxbest = 0, bxbest = 0;

                    //The forward diagonal that maximizes X + Y.
                    var fxybest = -1;
                    for (var d = fmax; d >= fmin; d -= 2)
                    {
                        var x = Math.Min(fd[d + o], xlim);
                        var y = x - d;
                        if (ylim < y)
                        {
                            x = ylim + d;
                            y = ylim;
                        }

                        if (fxybest < x + y)
                        {
                            fxybest = x + y;
                            fxbest = x;
                        }
                    }

                    //The backward diagonal that minimizes X + Y.
                    var bxybest = int.MaxValue;
                    for (var d = bmax; d >= bmin; d -= 2)
                    {
                        var x = Math.Max(xoff, bd[d + o]);
                        var y = x - d;
                        if (y < yoff)
                        {
                            x = yoff + d;
                            y = yoff;
                        }

                        if (x + y < bxybest)
                        {
                            bxybest = x + y;
                            bxbest = x;
                        }
                    }

                    //Use the better of the two.
                    if ((xlim + ylim) - bxybest < fxybest - (xoff + yoff))
                    {
                        xmid = fxbest;
                        ymid = fxybest - fxbest;
                        loMinimal = true;
                        hiMinimal = false;
                    }
                    else
                    {
                        xmid = bxbest;
                        ymid = bxybest - bxbest;
                        loMinimal = false;
                        hiMinimal = true;
                    }

                    return 2 * c - 1;
                }
            }
        }

        /// <summary>Compares in detail contiguous subsequences of the two files which are known, as a whole, to match each other.</summary>
        private void CompareSeq(int xoff, int xlim, int yoff, int ylim, bool minimal)
        {
            //The recursion is on the second half in a loop, so that only the first half grows the stack.
            while (true)
            {
                //Slide down the bottom initial diagonal, and up the top one.
                while (xoff < xlim && yoff < ylim && _xvec[xoff] == _yvec[yoff])
                {
                    ++xoff;
                    ++yoff;
                }

                while (xlim > xoff && ylim > yoff && _xvec[xlim - 1] == _yvec[ylim - 1])
                {
                    --xlim;
                    --ylim;
                }

                if (xoff == xlim)
                {
                    while (yoff < ylim)
                    {
                        SetChanged(1, _files[1].RealIndexes[yoff++], true);
                    }

                    return;
                }

                if (yoff == ylim)
                {
                    while (xoff < xlim)
                    {
                        SetChanged(0, _files[0].RealIndexes[xoff++], true);
                    }

                    return;
                }

                //Find a point of correspondence in the middle of the files, and split into two problems.
                Diag(xoff, xlim, yoff, ylim, minimal, out var xmid, out var ymid, out var loMinimal, out var hiMinimal);
                CompareSeq(xoff, xmid, yoff, ymid, loMinimal);
                xoff = xmid;
                yoff = ymid;
                minimal = hiMinimal;
            }
        }

        /// <summary>
        /// Discards lines of one file that have no matches in the other, as if they were not there; they are marked as changed.
        /// Lines with very many matches are only discarded in a run of discardable lines.
        /// </summary>
        private void DiscardConfusingLines(int equivMax)
        {
            var equivCount = new[] { new int[equivMax], new int[equivMax] };
            var discarded = new[] { new byte[_files[0].Lines], new byte[_files[1].Lines] };

            for (var f = 0; f < 2; f++)
            {
                foreach (var equiv in _files[f].Equivs)
                {
                    ++equivCount[f][equiv];
                }
            }

            //Mark to be discarded each line that matches no line of the other file. If a line matches many lines, it is provisionally discardable.
            for (var f = 0; f < 2; f++)
            {
                var end = _files[f].Lines;
                var discards = discarded[f];
                var counts = equivCount[1 - f];
                var equivs = _files[f].Equivs;
                var many = 5;
                var tem = end / 64;

                //MANY is multiplied by the approximate square root of the number of lines.
                while ((tem >>= 2) > 0)
                {
                    many *= 2;
                }

                for (var i = 0; i < end; i++)
                {
                    if (equivs[i] == 0)
                    {
                        continue;
                    }

                    var nmatch = counts[equivs[i]];
                    if (nmatch == 0)
                    {
                        discards[i] = 1;
                    }
                    else if (nmatch > many)
                    {
                        discards[i] = 2;
                    }
                }
            }

            //Provisional lines are only really discarded in a run of discardables with nonprovisionals at the beginning and end.
            for (var f = 0; f < 2; f++)
            {
                var end = _files[f].Lines;
                var discards = discarded[f];

                for (var i = 0; i < end; i++)
                {
                    if (discards[i] == 2)
                    {
                        discards[i] = 0; //Not in the middle of a run of discards.
                    }
                    else if (discards[i] != 0)
                    {
                        //A nonprovisional discard: find the end of this run of discardable lines, counting the provisional ones.
                        int j;
                        var provisional = 0;
                        for (j = i; j < end; j++)
                        {
                            if (discards[j] == 0)
                            {
                                break;
                            }

                            if (discards[j] == 2)
                            {
                                ++provisional;
                            }
                        }

                        //Cancel provisional discards at the end, and shrink the run.
                        while (j > i && discards[j - 1] == 2)
                        {
                            discards[--j] = 0;
                            --provisional;
                        }

                        var length = j - i;

                        //If 1/4 of the lines in the run are provisional, cancel discarding of all of them.
                        if (provisional * 4 > length)
                        {
                            while (j > i)
                            {
                                if (discards[--j] == 2)
                                {
                                    discards[j] = 0;
                                }
                            }
                        }
                        else
                        {
                            var minimum = 1;
                            var quarter = length / 4;

                            //MINIMUM is the approximate square root of LENGTH/4. A subrun of two or more provisionals can stand when LENGTH
                            //is at least 16, and one of four or more when LENGTH >= 64.
                            while ((quarter >>= 2) > 0)
                            {
                                minimum *= 2;
                            }

                            minimum++;

                            //Cancel any subrun of MINIMUM or more provisionals within the larger run.
                            int consec;
                            for (j = 0, consec = 0; j < length; j++)
                            {
                                if (discards[i + j] != 2)
                                {
                                    consec = 0;
                                }
                                else if (minimum == ++consec)
                                {
                                    j -= consec; //Back up to the start of the subrun, to cancel it all.
                                }
                                else if (minimum < consec)
                                {
                                    discards[i + j] = 0;
                                }
                            }

                            //Scan from the beginning of the run until 3 or more nonprovisionals in a row, or until the first nonprovisional
                            //at least 8 lines in. Until that point, cancel any provisionals.
                            for (j = 0, consec = 0; j < length; j++)
                            {
                                if (j >= 8 && discards[i + j] == 1)
                                {
                                    break;
                                }

                                if (discards[i + j] == 2)
                                {
                                    consec = 0;
                                    discards[i + j] = 0;
                                }
                                else if (discards[i + j] == 0)
                                {
                                    consec = 0;
                                }
                                else
                                {
                                    consec++;
                                }

                                if (consec == 3)
                                {
                                    break;
                                }
                            }

                            //I advances to the last line of the run.
                            i += length - 1;

                            //Same thing, from the end.
                            for (j = 0, consec = 0; j < length; j++)
                            {
                                if (j >= 8 && discards[i - j] == 1)
                                {
                                    break;
                                }

                                if (discards[i - j] == 2)
                                {
                                    consec = 0;
                                    discards[i - j] = 0;
                                }
                                else if (discards[i - j] == 0)
                                {
                                    consec = 0;
                                }
                                else
                                {
                                    consec++;
                                }

                                if (consec == 3)
                                {
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            //Actually discard the lines.
            for (var f = 0; f < 2; f++)
            {
                var file = _files[f];
                var discards = discarded[f];
                file.Undiscarded = new int[file.Lines];
                file.RealIndexes = new int[file.Lines];
                var j = 0;
                for (var i = 0; i < file.Lines; ++i)
                {
                    if (_noDiscards || discards[i] == 0)
                    {
                        file.Undiscarded[j] = file.Equivs[i];
                        file.RealIndexes[j++] = i;
                    }
                    else
                    {
                        SetChanged(f, i, true);
                    }
                }

                file.Nondiscarded = j;
            }
        }

        /// <summary>
        /// Adjusts inserts/deletes of identical lines to join changes as much as possible: when a run of changed lines has an identical,
        /// unchanged line at the other end, either may be taken as the change, and the one that merges with a neighbour is.
        /// </summary>
        private void ShiftBoundaries()
        {
            for (var f = 0; f < 2; f++)
            {
                var changed = _files[f].Changed; //Index + 1.
                var otherChanged = _files[1 - f].Changed;
                var equivs = _files[f].Equivs;
                var i = 0;
                var j = 0;
                var iEnd = _files[f].Lines;

                //Each changed[x] of the C code is changed[x + 1] here, and so is each otherChanged[x].
                bool C(int x) => changed[x + 1];
                bool O(int x) => otherChanged[x + 1];

                while (true)
                {
                    int runLength, start, corresponding;

                    //Scan forwards to find the beginning of another run of changes, keeping track of the corresponding point in the other file.
                    while (i < iEnd && !C(i))
                    {
                        while (O(j++))
                        {
                        }

                        i++;
                    }

                    if (i == iEnd)
                    {
                        break;
                    }

                    start = i;

                    //Find the end of this run of changes.
                    while (C(++i))
                    {
                    }

                    while (O(j))
                    {
                        j++;
                    }

                    do
                    {
                        //Record the length of this run of changes, so that it can be told whether the run has grown.
                        runLength = i - start;

                        //Move the changed region back, so long as the previous unchanged line matches the last changed one.
                        while (start != 0 && equivs[start - 1] == equivs[i - 1])
                        {
                            changed[--start + 1] = true;
                            changed[--i + 1] = false;
                            while (C(start - 1))
                            {
                                start--;
                            }

                            while (O(--j))
                            {
                            }
                        }

                        //The end of the changed run, at the last point where it corresponds to a changed run in the other file.
                        corresponding = O(j - 1) ? i : iEnd;

                        //Move the changed region forward, so long as the first changed line matches the following unchanged one.
                        while (i != iEnd && equivs[start] == equivs[i])
                        {
                            changed[start++ + 1] = false;
                            changed[i++ + 1] = true;
                            while (C(i))
                            {
                                i++;
                            }

                            while (O(++j))
                            {
                                corresponding = i;
                            }
                        }
                    }
                    while (runLength != i - start);

                    //If possible, move the fully-merged run of changes back to a corresponding run in the other file.
                    while (corresponding < i)
                    {
                        changed[--start + 1] = true;
                        changed[--i + 1] = false;
                        while (O(--j))
                        {
                        }
                    }
                }
            }
        }

        /// <summary>Scans the tables of which lines are inserted and deleted, producing the edit script in forward order.</summary>
        private List<LineChange> BuildScript()
        {
            var script = new List<LineChange>();
            int i0 = _files[0].Lines, i1 = _files[1].Lines;

            //Changed[-1] exists, and is false.
            while (i0 >= 0 || i1 >= 0)
            {
                if (Changed(0, i0 - 1) || Changed(1, i1 - 1))
                {
                    int line0 = i0, line1 = i1;

                    while (Changed(0, i0 - 1))
                    {
                        --i0;
                    }

                    while (Changed(1, i1 - 1))
                    {
                        --i1;
                    }

                    script.Add(new LineChange { Line0 = i0, Line1 = i1, Deleted = line0 - i0, Inserted = line1 - i1 });
                }

                //Lines that match each other have been reached.
                i0--;
                i1--;
            }

            script.Reverse(); //Built from the end.
            return script;
        }
    }
}
