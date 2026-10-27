using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    public enum HunkChoice
    {
        Unresolved,
        /// <summary>Keep the project's lines.</summary>
        Module,
        /// <summary>Keep the folder's lines.</summary>
        File,
        /// <summary>Keep both: the project's lines, then the folder's.</summary>
        Both,
        /// <summary>The user wrote the result by hand. Only the compare dialog knows the text; <see cref="SyncDiff.Merge"/> cannot use it.</summary>
        Custom
    }

    /// <summary>A stretch where the two sides differ. Either side may have no lines in it: those were added or removed.</summary>
    public class DiffHunk
    {
        public int ModuleStart { get; set; }
        public int ModuleCount { get; set; }
        public int FileStart { get; set; }
        public int FileCount { get; set; }

        public HunkChoice Choice { get; set; }
    }

    /// <summary>
    /// A line by line comparison of the module and the file, done by WinMerge's own engine (<see cref="WinMergeLineDiff"/>): the lines both
    /// sides share, and between them the hunks that differ, each of which can be resolved to either side. Indentation and trailing space do not count as a difference,
    /// and the merged text is made of the lines of the side picked in each hunk, exactly as they were written.
    /// </summary>
    public class SyncDiff
    {
        private readonly IList<string> _module;
        private readonly IList<string> _file;

        public SyncDiff(string moduleText, string fileText)
        {
            _module = SyncText.Lines(moduleText).ToList();
            _file = SyncText.Lines(fileText).ToList();
            Hunks = Compare(_module, _file);
        }

        public IList<string> ModuleLines => _module;
        public IList<string> FileLines => _file;
        public IReadOnlyList<DiffHunk> Hunks { get; }

        public bool IsResolved => Hunks.All(hunk => hunk.Choice != HunkChoice.Unresolved);

        public void ResolveAll(HunkChoice choice)
        {
            foreach (var hunk in Hunks)
            {
                hunk.Choice = choice;
            }
        }

        /// <summary>The text with each hunk taken from the side chosen for it. Unresolved hunks are an error.</summary>
        public string Merge()
        {
            if (!IsResolved || Hunks.Any(hunk => hunk.Choice == HunkChoice.Custom))
            {
                throw new InvalidOperationException("Some differences have not been resolved.");
            }

            var result = new List<string>();
            var position = 0; //In the module's lines; outside the hunks both sides are the same.
            foreach (var hunk in Hunks)
            {
                result.AddRange(_module.Skip(position).Take(hunk.ModuleStart - position));
                if (hunk.Choice != HunkChoice.File)
                {
                    result.AddRange(_module.Skip(hunk.ModuleStart).Take(hunk.ModuleCount));
                }

                if (hunk.Choice != HunkChoice.Module)
                {
                    result.AddRange(_file.Skip(hunk.FileStart).Take(hunk.FileCount));
                }

                position = hunk.ModuleStart + hunk.ModuleCount;
            }

            result.AddRange(_module.Skip(position));
            return string.Join("\r\n", result);
        }

        //What counts as the same line.
        private static string Key(string line) => line.Trim();

        private static IReadOnlyList<DiffHunk> Compare(IList<string> module, IList<string> file)
        {
            return WinMergeLineDiff.Compare(module, file, Key)
                .Select(change => new DiffHunk
                {
                    ModuleStart = change.Line0,
                    ModuleCount = change.Deleted,
                    FileStart = change.Line1,
                    FileCount = change.Inserted
                })
                .ToList();
        }
    }
}
