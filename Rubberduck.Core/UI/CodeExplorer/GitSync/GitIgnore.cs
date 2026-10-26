using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// The rules of one ignore file (.rdsyncignore), written like a .gitignore and with the semantics git uses: comments, negation with "!",
    /// directory-only patterns with a trailing "/", anchoring with a "/" in the pattern, "*", "?", "[...]" and "**".
    /// Paths are relative to the folder containing the ignore file and use "/" as separator.
    /// </summary>
    public class GitIgnore
    {
        private class Rule
        {
            public Regex Pattern;
            public bool Negated;
            public bool DirectoryOnly;
        }

        private readonly List<Rule> _rules = new List<Rule>();

        public GitIgnore(IEnumerable<string> lines)
        {
            foreach (var line in lines)
            {
                var rule = Parse(line);
                if (rule != null)
                {
                    _rules.Add(rule);
                }
            }
        }

        public bool IsEmpty => _rules.Count == 0;

        /// <summary>
        /// Returns true if the last matching rule ignores the path, false if it re-includes it, null if no rule matches.
        /// </summary>
        public bool? Match(string relativePath, bool isDirectory)
        {
            bool? result = null;
            foreach (var rule in _rules)
            {
                if ((!rule.DirectoryOnly || isDirectory) && rule.Pattern.IsMatch(relativePath))
                {
                    result = !rule.Negated;
                }
            }

            return result;
        }

        private static Rule Parse(string line)
        {
            //Trailing spaces are ignored unless escaped with a backslash.
            line = Regex.Replace(line, @"(?<!\\)\s+$", string.Empty);
            if (line.Length == 0 || line[0] == '#')
            {
                return null;
            }

            var negated = false;
            if (line[0] == '!')
            {
                negated = true;
                line = line.Substring(1);
            }
            else if (line.StartsWith(@"\!") || line.StartsWith(@"\#"))
            {
                line = line.Substring(1);
            }

            var directoryOnly = line.EndsWith("/");
            line = line.TrimEnd('/');
            if (line.Length == 0)
            {
                return null;
            }

            //A separator at the beginning or in the middle anchors the pattern to the ignore file's folder.
            var anchored = line.Contains("/");
            line = line.TrimStart('/');

            var regex = "^" + (anchored ? string.Empty : "(?:.*/)?") + Translate(line) + "$";
            return new Rule
            {
                Pattern = new Regex(regex, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant),
                Negated = negated,
                DirectoryOnly = directoryOnly
            };
        }

        private static string Translate(string glob)
        {
            var regex = new StringBuilder();
            for (var i = 0; i < glob.Length; i++)
            {
                var c = glob[i];
                switch (c)
                {
                    case '*':
                        if (i + 1 < glob.Length && glob[i + 1] == '*')
                        {
                            var atStart = i == 0 || glob[i - 1] == '/';
                            var followedBySlash = i + 2 < glob.Length && glob[i + 2] == '/';
                            var atEnd = i + 2 == glob.Length;

                            if (atStart && followedBySlash)
                            {
                                regex.Append("(?:.*/)?");   // "**/" matches zero or more directories
                                i += 2;
                            }
                            else if (atStart && atEnd)
                            {
                                regex.Append(".*");         // trailing "/**" matches everything inside
                                i += 1;
                            }
                            else
                            {
                                regex.Append("[^/]*");      // otherwise "**" acts like "*"
                                i += 1;
                            }
                        }
                        else
                        {
                            regex.Append("[^/]*");
                        }
                        break;
                    case '?':
                        regex.Append("[^/]");
                        break;
                    case '[':
                        var close = glob.IndexOf(']', i + 2 <= glob.Length ? i + 2 : glob.Length);
                        if (close < 0)
                        {
                            regex.Append(@"\[");
                            break;
                        }

                        var set = glob.Substring(i + 1, close - i - 1);
                        if (set.StartsWith("!"))
                        {
                            set = "^" + set.Substring(1);
                        }
                        regex.Append('[').Append(set.Replace(@"\", @"\\")).Append(']');
                        i = close;
                        break;
                    case '\\':
                        if (i + 1 < glob.Length)
                        {
                            regex.Append(Regex.Escape(glob[++i].ToString()));
                        }
                        break;
                    default:
                        regex.Append(Regex.Escape(c.ToString()));
                        break;
                }
            }

            return regex.ToString();
        }
    }

    /// <summary>
    /// All .rdsyncignore files found below a sync root. Like in git, a deeper ignore file takes precedence over a higher one,
    /// and nothing inside an ignored directory can be re-included.
    /// </summary>
    public class GitIgnoreSet
    {
        //Folder of the ignore file relative to the root, using "/", mapped to its rules; the root's is "".
        private readonly List<KeyValuePair<string, GitIgnore>> _files;

        public GitIgnoreSet(IEnumerable<KeyValuePair<string, GitIgnore>> files)
        {
            _files = files
                .Where(file => !file.Value.IsEmpty)
                .OrderBy(file => file.Key.Count(c => c == '/') + (file.Key.Length == 0 ? -1 : 0))
                .ToList();
        }

        public static GitIgnoreSet Empty { get; } = new GitIgnoreSet(Enumerable.Empty<KeyValuePair<string, GitIgnore>>());

        /// <param name="relativePath">The path relative to the sync root, with either kind of separator.</param>
        public bool IsIgnored(string relativePath, bool isDirectory)
        {
            if (_files.Count == 0)
            {
                return false;
            }

            var segments = relativePath.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
            for (var length = 1; length <= segments.Length; length++)
            {
                var path = string.Join("/", segments.Take(length));
                var directory = length < segments.Length || isDirectory;
                if (IsIgnoredByRules(path, directory))
                {
                    return true;
                }
            }

            return false;
        }

        private bool IsIgnoredByRules(string path, bool isDirectory)
        {
            bool? result = null;
            foreach (var file in _files)
            {
                string relativeToFile;
                if (file.Key.Length == 0)
                {
                    relativeToFile = path;
                }
                else if (path.StartsWith(file.Key + "/", StringComparison.OrdinalIgnoreCase))
                {
                    relativeToFile = path.Substring(file.Key.Length + 1);
                }
                else
                {
                    continue;
                }

                result = file.Value.Match(relativeToFile, isDirectory) ?? result;
            }

            return result ?? false;
        }
    }
}
