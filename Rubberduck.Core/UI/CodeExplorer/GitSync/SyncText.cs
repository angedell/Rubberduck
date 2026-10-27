using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// Reads exported module files (or the same text exported from a module) and boils them down to a signature that
    /// does not change with line endings, indentation or the like, so that only real differences count.
    /// </summary>
    public static class SyncText
    {
        //Left behind by earlier versions of the sync, which marked the code with a revision and a hash.
        private static readonly Regex LegacyStampLine = new Regex(
            @"^[ \t]*'[ \t]*rdsync[ \t]+rev=\d+[ \t]+hash=[0-9a-f]+(?:[ \t]+utc=\S+)?[ \t]*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex BeginBlock = new Regex(@"^\s*Begin\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex EndBlock = new Regex(@"^\s*End\s*$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex NoIndentAnnotation = new Regex(@"^\s*'\s*@NoIndent\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex AttributeLine = new Regex(@"^Attribute\s+(\w+)\s*=\s*(.*?)\s*$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static IList<string> Lines(string text)
        {
            return (text ?? string.Empty).Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
        }

        private static bool IsAttribute(string line) => line.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// The number of lines at the top of an exported file that are the header the VBE puts in front of the code:
        /// the VERSION line and the BEGIN...END block of forms and classes, then the module's Attribute lines.
        /// </summary>
        public static int CodeStart(IList<string> lines)
        {
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

            while (index < lines.Count && IsAttribute(lines[index]))
            {
                index++;
            }

            return index;
        }

        /// <summary>The module attributes in the header of an exported file, by name.</summary>
        public static IDictionary<string, string> ModuleAttributes(string fileText)
        {
            var lines = Lines(fileText);
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < CodeStart(lines); i++)
            {
                var match = AttributeLine.Match(lines[i]);
                if (match.Success)
                {
                    result[match.Groups[1].Value] = match.Groups[2].Value;
                }
            }

            return result;
        }

        /// <summary>The code of an exported file, without the header and without legacy stamps.</summary>
        public static string CodeOf(string fileText)
        {
            var lines = Lines(fileText);
            return string.Join("\r\n", lines.Skip(CodeStart(lines)).Where(line => !LegacyStampLine.IsMatch(line)));
        }

        /// <summary>
        /// Makes a file ready for the VBE to import: Windows line endings, and a @Folder annotation that matches where the file is.
        /// Files directly in the sync folder keep whatever annotation they have.
        /// </summary>
        public static string PrepareForImport(string fileText, string relativeDirectory)
        {
            var text = GitSyncText.NormalizeLineEndings(fileText);
            var folder = GitSyncText.FolderPath(relativeDirectory);
            return folder.Length == 0 ? text : GitSyncText.SetFolderAnnotation(text, folder);
        }

        /// <summary>
        /// The text in lower case except for what is inside string literals. VBA does not tell upper and lower case apart, apart from in strings.
        /// A string runs from a quote to the next one, and a doubled quote inside it is a quote that belongs to it; strings do not span lines.
        /// The result is as long as the text, so positions in it are positions in the original.
        /// </summary>
        public static string FoldCase(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            var folded = text.ToCharArray();
            var inString = false;
            for (var i = 0; i < folded.Length; i++)
            {
                var ch = folded[i];
                if (ch == '\n' || ch == '\r')
                {
                    inString = false;
                }
                else if (ch == '"')
                {
                    inString = !inString;
                }
                else if (!inString)
                {
                    folded[i] = char.ToLowerInvariant(ch);
                }
            }

            return new string(folded);
        }

        /// <summary>
        /// A short hash of everything that matters about the file: its header and attributes as they are, and its code once it
        /// is indented the way the settings say. Two files that only differ in line endings, indentation, stamps or trailing space have the same signature.
        /// </summary>
        public static string Signature(string fileText, IIndenterSettings settings, ISimpleIndenter indenter, bool ignoreCaseOutsideStrings = false)
        {
            var lines = Lines(fileText);
            var start = CodeStart(lines);

            var header = lines.Take(start).Select(line => line.TrimEnd());
            var attributes = lines.Where(IsAttribute).Select(line => line.TrimEnd());
            var code = lines.Skip(start)
                .Where(line => !LegacyStampLine.IsMatch(line) && !IsAttribute(line))
                .ToList();

            if (indenter != null && !code.Any(line => NoIndentAnnotation.IsMatch(line)))
            {
                try
                {
                    code = indenter.Indent(code, true, settings).ToList();
                }
                catch (Exception)
                {
                    //Code the indenter cannot make sense of is compared as it is.
                }
            }

            var normalized = string.Join("\n", header)
                             + "\u0001" + string.Join("\n", attributes)
                             + "\u0001" + string.Join("\n", code.SelectMany(Lines).Select(line => line.TrimEnd())).Trim('\n');
            if (ignoreCaseOutsideStrings)
            {
                normalized = FoldCase(normalized);
            }

            using (var sha = SHA1.Create())
            {
                return string.Concat(sha.ComputeHash(Encoding.UTF8.GetBytes(normalized)).Take(8).Select(b => b.ToString("x2")));
            }
        }
    }
}
