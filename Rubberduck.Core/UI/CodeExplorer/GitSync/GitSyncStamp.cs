using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// The revision and content hash a sync left behind, as read from a stamp line.
    /// </summary>
    public class SyncStamp
    {
        public SyncStamp(int revision, string hash, DateTime? utc = null)
        {
            Revision = revision;
            Hash = hash;
            Utc = utc;
        }

        /// <summary>When the code was last modified, as far as the sync could tell. Null in stamps that predate this field.</summary>
        public DateTime? Utc { get; }

        public int Revision { get; }
        public string Hash { get; }
    }

    /// <summary>
    /// What a side (module or file) looks like in terms of its stamp.
    /// </summary>
    public class StampState
    {
        public StampState(SyncStamp stamp, string bodyHash)
        {
            Revision = stamp.Revision;
            Hash = stamp.Hash;
            Utc = stamp.Utc;
            BodyHash = bodyHash;
        }

        /// <summary>When the code was last modified, according to the stamp.</summary>
        public DateTime? Utc { get; }

        public int Revision { get; }

        /// <summary>The hash recorded in the stamp.</summary>
        public string Hash { get; }

        /// <summary>The hash of the code as it is now.</summary>
        public string BodyHash { get; }

        /// <summary>Whether the code was edited since the stamp was written.</summary>
        public bool Dirty => BodyHash != Hash;
    }

    /// <summary>
    /// The stamp is a single comment line, written into both the module and its file whenever they are synced:
    /// <c>' rdsync rev=12 hash=3fa9c1d4e07b22aa utc=2025-06-01T12:34:56Z</c>. The revision grows with every change that goes through a sync and the
    /// hash covers the module's code without the stamp line itself. Since both sides carry it, the sync direction
    /// can be told on any machine, without any state outside the workbook and the repository.
    /// </summary>
    public static class GitSyncStamp
    {
        private static readonly Regex StampLine = new Regex(
            @"^[ \t]*'[ \t]*rdsync[ \t]+rev=(\d+)[ \t]+hash=([0-9a-f]+)(?:[ \t]+utc=(\S+))?[ \t]*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex BeginBlock = new Regex(@"^\s*Begin\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex EndBlock = new Regex(@"^\s*End\s*$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static string Format(int revision, string hash, DateTime? utc = null)
        {
            var line = $"' rdsync rev={revision} hash={hash}";
            return utc.HasValue
                ? line + " utc=" + utc.Value.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture)
                : line;
        }

        public static bool IsStampLine(string line) => StampLine.IsMatch(line.TrimEnd('\r'));

        /// <summary>
        /// Reads the first stamp in the text, which can be a file or the code of a module. Returns null if there is none.
        /// </summary>
        public static SyncStamp Parse(string text)
        {
            foreach (var line in Lines(text))
            {
                var match = StampLine.Match(line);
                if (match.Success && int.TryParse(match.Groups[1].Value, out var revision))
                {
                    DateTime? utc = null;
                    if (match.Groups[3].Success
                        && DateTime.TryParse(match.Groups[3].Value, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var parsed))
                    {
                        utc = parsed;
                    }

                    return new SyncStamp(revision, match.Groups[2].Value.ToLowerInvariant(), utc);
                }
            }

            return null;
        }

        /// <summary>
        /// The state of a module's code in terms of its stamp, or null if it has no stamp.
        /// </summary>
        public static StampState StateOfCode(string code)
        {
            var stamp = code == null ? null : Parse(code);
            return stamp == null ? null : new StampState(stamp, HashOfCode(code));
        }

        /// <summary>
        /// The state of an exported file in terms of its stamp, or null if it has no stamp.
        /// </summary>
        public static StampState StateOfFile(string fileText)
        {
            return StateOfCode(fileText == null ? null : CodeOfFile(fileText));
        }

        /// <summary>
        /// Hash of the code, ignoring the stamp line, the kind of line endings and blank space around the code.
        /// </summary>
        public static string HashOfCode(string code)
        {
            var normalized = string.Join("\r\n", Lines(code).Where(line => !StampLine.IsMatch(line))).Trim();
            using (var sha = SHA1.Create())
            {
                var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(normalized));
                return string.Concat(hash.Take(8).Select(b => b.ToString("x2")));
            }
        }

        /// <summary>
        /// Returns the code of an exported file, which is the file without the header the VBE puts in front
        /// (the VERSION and BEGIN...END block and the Attribute lines).
        /// </summary>
        public static string CodeOfFile(string fileText)
        {
            var lines = Lines(fileText).ToList();
            var index = 0;

            if (lines.Count > 0 && lines[0].StartsWith("VERSION ", StringComparison.OrdinalIgnoreCase))
            {
                index = 1;
                var depth = 0;
                for (; index < lines.Count; index++)
                {
                    if (BeginBlock.IsMatch(lines[index]))
                    {
                        depth++;
                    }
                    else if (EndBlock.IsMatch(lines[index]))
                    {
                        depth--;
                        if (depth <= 0)
                        {
                            index++;
                            break;
                        }
                    }
                }
            }

            while (index < lines.Count && lines[index].StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
            {
                index++;
            }

            return string.Join("\r\n", lines.Skip(index));
        }

        /// <summary>
        /// Puts the stamp in a file's text: replaces the existing one, or adds it after the header.
        /// </summary>
        public static string WithStamp(string fileText, int revision, string hash, DateTime? utc = null)
        {
            var stamp = Format(revision, hash, utc);
            var lines = Lines(fileText).ToList();

            var existing = lines.FindIndex(line => StampLine.IsMatch(line));
            if (existing >= 0)
            {
                lines[existing] = stamp;
            }
            else
            {
                lines.Insert(GitSyncText.IndexAfterHeader(lines), stamp);
            }

            return string.Join("\r\n", lines);
        }

        private static IEnumerable<string> Lines(string text)
        {
            return (text ?? string.Empty).Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
        }
    }
}
