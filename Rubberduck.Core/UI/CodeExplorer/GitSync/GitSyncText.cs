using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// Pure text helpers for the git-sync feature: line ending normalization and @Folder annotation handling.
    /// </summary>
    public static class GitSyncText
    {
        private const string Crlf = "\r\n";
        private static readonly Regex NonWindowsLineEnding = new Regex("(?<!\r)\n|\r(?!\n)", RegexOptions.Compiled);

        //Matches both '@Folder("A.B") and '@Folder "A.B"; group 1 is the (still escaped) folder path.
        private static readonly Regex FolderAnnotation = new Regex(
            "'@Folder(?:\\s*\\(\\s*|\\s+)\"((?:[^\"]|\"\")*)\"(?:\\s*\\))?",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        /// <summary>
        /// Converts all LF-only and CR-only line endings to CRLF.
        /// </summary>
        public static string NormalizeLineEndings(string text)
        {
            return NonWindowsLineEnding.Replace(text, Crlf);
        }

        /// <summary>
        /// Converts the directory of a file, relative to the sync root, to a @Folder path, e.g. "Models\Orders" to "Models.Orders".
        /// Returns an empty string for files located directly in the root.
        /// </summary>
        public static string FolderPath(string relativeDirectory)
        {
            if (string.IsNullOrWhiteSpace(relativeDirectory))
            {
                return string.Empty;
            }

            //The dot is the folder separator in the annotation, so it cannot be part of a segment.
            var segments = relativeDirectory
                .Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(segment => segment.Replace('.', '_'));

            return string.Join(".", segments);
        }

        /// <summary>
        /// Converts a @Folder path to a relative directory, e.g. "Models.Orders" to "Models\Orders".
        /// </summary>
        public static string RelativeDirectory(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                return string.Empty;
            }

            var segments = folderPath
                .Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(SafeSegment);

            return string.Join("\\", segments);
        }

        private static string SafeSegment(string segment)
        {
            var invalid = System.IO.Path.GetInvalidFileNameChars();
            var safe = invalid.Aggregate(segment.Trim(), (current, c) => current.Replace(c, '_'));
            return safe.Length == 0 ? "_" : safe;
        }

        /// <summary>
        /// Returns the path in the first @Folder annotation of the source, or null if there is none.
        /// </summary>
        public static string GetFolderAnnotation(string source)
        {
            var match = FolderAnnotation.Match(source);
            return match.Success
                ? match.Groups[1].Value.Replace("\"\"", "\"")
                : null;
        }

        /// <summary>
        /// Makes sure the source carries a @Folder annotation with the given path, replacing an existing one
        /// or inserting a new one after the leading Attribute lines.
        /// </summary>
        public static string SetFolderAnnotation(string source, string folderPath)
        {
            var annotation = $"'@Folder(\"{folderPath.Replace("\"", "\"\"")}\")";

            var existing = FolderAnnotation.Match(source);
            if (existing.Success)
            {
                return existing.Value == annotation
                    ? source
                    : source.Substring(0, existing.Index) + annotation + source.Substring(existing.Index + existing.Length);
            }

            var lines = source.Split(new[] { Crlf }, StringSplitOptions.None).ToList();
            lines.Insert(IndexAfterHeader(lines), annotation);
            return string.Join(Crlf, lines);
        }

        //The exported header ends after the contiguous block of Attribute lines that starts at "Attribute VB_Name".
        private static int IndexAfterHeader(IList<string> lines)
        {
            var firstAttribute = -1;
            for (var i = 0; i < lines.Count; i++)
            {
                if (lines[i].StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
                {
                    firstAttribute = i;
                    break;
                }
            }

            if (firstAttribute < 0)
            {
                return 0;
            }

            var index = firstAttribute;
            while (index < lines.Count && lines[index].StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
            {
                index++;
            }

            return index;
        }
    }
}
