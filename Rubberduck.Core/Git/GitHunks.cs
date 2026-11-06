using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Rubberduck.Git
{
    /// <summary>
    /// One piece of a patch that can be committed on its own: a block of changed lines in a file, or a whole file where it cannot be cut
    /// (a binary one, or one with no lines to cut, like a new empty file).
    /// </summary>
    public class GitHunk
    {
        public string Id { get; set; }
        public string Path { get; set; }

        /// <summary>The file's lines before its first block (<c>diff --git</c>, mode, index, ---, +++). For a whole file, the file's entire patch.</summary>
        public string Header { get; set; }

        /// <summary>The block, from its <c>@@</c> line on. Empty for a whole file.</summary>
        public string Body { get; set; }

        public bool IsWholeFile { get; set; }
        public bool IsBinary { get; set; }

        /// <summary>A form's binary half. It is never grouped by itself; it goes with its form.</summary>
        public bool IsCompanion => Path.EndsWith(".frx", StringComparison.OrdinalIgnoreCase);

        /// <summary>The block's <c>@@</c> line and the first line it changes, for a list the user reads.</summary>
        public string Summary
        {
            get
            {
                if (IsBinary)
                {
                    return "binary file";
                }

                var lines = (Body ?? string.Empty).Split('\n');
                var first = lines.FirstOrDefault(line => (line.StartsWith("+") || line.StartsWith("-")) && line.Trim().Length > 1) ?? string.Empty;
                var at = lines.Length > 0 ? Regex.Match(lines[0], @"^@@ [^@]*@@").Value : string.Empty;
                return (at + " " + first.TrimEnd('\r')).Trim();
            }
        }
    }

    /// <summary>One commit of a split: its message and the hunks in it.</summary>
    public class GitCommitPart
    {
        public GitCommitPart(string message, IReadOnlyList<GitHunk> hunks)
        {
            Message = message;
            Hunks = hunks;
        }

        public string Message { get; }
        public IReadOnlyList<GitHunk> Hunks { get; }
    }

    public static class GitHunkParser
    {
        /// <summary>
        /// The text of a patch is read byte for byte (as Latin-1, which maps every byte to a character and back),
        /// because a module is in the ANSI code page and has to reach <c>git apply</c> exactly as it left git.
        /// </summary>
        public static readonly Encoding Bytes = Encoding.GetEncoding("ISO-8859-1");

        public static IReadOnlyList<GitHunk> Parse(string patch)
        {
            var hunks = new List<GitHunk>();
            foreach (var section in Regex.Split(patch, @"(?m)^(?=diff --git )").Where(section => section.StartsWith("diff --git ")))
            {
                var firstLine = section.Substring(0, section.IndexOf('\n') < 0 ? section.Length : section.IndexOf('\n')).TrimEnd('\r');
                var split = firstLine.LastIndexOf(" b/", StringComparison.Ordinal);
                var path = split >= 0 ? firstLine.Substring(split + 3) : firstLine;

                var at = Regex.Match(section, @"(?m)^@@ ");
                var binary = section.Contains("\nGIT binary patch") || Regex.IsMatch(section, @"(?m)^Binary files ");
                if (binary || !at.Success)
                {
                    hunks.Add(new GitHunk { Path = path, Header = section, Body = string.Empty, IsWholeFile = true, IsBinary = binary });
                    continue;
                }

                var header = section.Substring(0, at.Index);
                var starts = Regex.Matches(section, @"(?m)^@@ ").Cast<Match>().Select(match => match.Index).ToList();
                for (var i = 0; i < starts.Count; i++)
                {
                    var end = i + 1 < starts.Count ? starts[i + 1] : section.Length;
                    hunks.Add(new GitHunk { Path = path, Header = header, Body = section.Substring(starts[i], end - starts[i]) });
                }
            }

            for (var i = 0; i < hunks.Count; i++)
            {
                hunks[i].Id = "h" + (i + 1);
            }

            return hunks;
        }

        /// <summary>The patch that makes exactly these hunks happen. Hunks of one file keep their order.</summary>
        public static string BuildPatch(IEnumerable<GitHunk> hunks)
        {
            var patch = new StringBuilder();
            foreach (var file in hunks.OrderBy(hunk => int.Parse(hunk.Id.Substring(1))).GroupBy(hunk => hunk.Header))
            {
                patch.Append(file.Key);
                foreach (var hunk in file)
                {
                    patch.Append(hunk.Body);
                }
            }

            return patch.ToString();
        }
    }
}
