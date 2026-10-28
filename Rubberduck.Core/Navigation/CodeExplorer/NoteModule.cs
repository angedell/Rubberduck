using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.Navigation.CodeExplorer
{
    /// <summary>
    /// A note is a standard module marked with the @Notes annotation whose code is nothing but comments: markdown with a comment mark in front of every line,
    /// so that it compiles. The Code Explorer shows it under its name (the annotation's argument, e.g. CLAUDE.md) and edits the text without the marks.
    /// This class works on the module's text, so the API, which reads the editor directly and must not wait for the parser, and the Code Explorer agree.
    /// </summary>
    public static class NoteModule
    {
        private const string Extension = ".md";
        private const string DefaultSuffix = "_md";
        private const string NewLine = "\r\n";

        /// <summary>The longest line the VBE accepts.</summary>
        public const int MaxLineLength = 1023;

        //What the VBE accepts as the name of a module.
        private static readonly Regex ValidModuleName = new Regex(@"^[A-Za-z][A-Za-z0-9_]{0,30}$");

        private static readonly Regex FolderAnnotation = new Regex(@"^\s*'@Folder\b\s*\(?\s*""?(?<folder>[^""\)]*)", RegexOptions.IgnoreCase);
        private static readonly Regex NotesAnnotation = new Regex(@"^\s*'@Notes\b\s*\(?\s*(""(?<name>[^""\)]*)"")?", RegexOptions.IgnoreCase);

        /// <summary>Whether the module carries the @Notes annotation. A quick look before <see cref="IsNote"/> reads everything.</summary>
        public static bool HasMarker(string moduleCode) => Lines(moduleCode).Any(line => NotesAnnotation.IsMatch(line));

        /// <summary>A note: marked with @Notes, and nothing in it but comments, blank lines and the Option Explicit the editor adds to a new module. A marked module with code in it is an ordinary module.</summary>
        public static bool IsNote(string moduleCode)
        {
            var lines = Lines(moduleCode).ToList();
            return lines.Any(line => NotesAnnotation.IsMatch(line)) && lines.All(IsNoteLine);
        }

        /// <summary>The name the annotation gives the note, or null.</summary>
        public static string NameOf(string moduleCode)
        {
            foreach (var line in Lines(moduleCode))
            {
                var match = NotesAnnotation.Match(line);
                if (match.Success)
                {
                    return NameFromArgument(match.Groups["name"].Value);
                }
            }

            return null;
        }

        /// <summary>What the note is called: the annotation's name, or the module name with ".md" after it.</summary>
        public static string DisplayName(string moduleName, string moduleCode) => NameOf(moduleCode) ?? moduleName + Extension;

        /// <summary>The same from a parsed annotation's argument, which still has its quotes (null if there was none).</summary>
        public static string DisplayNameFromArgument(string moduleName, string argument) => NameFromArgument(argument?.Trim().Trim('"')) ?? moduleName + Extension;

        /// <summary>The module name for a note name: "CLAUDE.md" gives "CLAUDE_md". Null if the name does not end in .md or is not one the VBE accepts for a module.</summary>
        public static string ModuleNameFor(string noteName)
        {
            var name = (noteName ?? string.Empty).Trim();
            if (!name.EndsWith(Extension, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            name = name.Substring(0, name.Length - Extension.Length) + DefaultSuffix;
            return ValidModuleName.IsMatch(name) ? name : null;
        }

        /// <summary>The annotations with the @Notes marker for <paramref name="displayName"/> first, every other annotation as it was.</summary>
        public static IReadOnlyList<string> WithMarker(IEnumerable<string> annotations, string displayName)
        {
            var result = (annotations ?? Enumerable.Empty<string>()).Where(line => !NotesAnnotation.IsMatch(line)).ToList();
            result.Insert(0, "'@Notes(\"" + displayName + "\")");
            return result;
        }

        /// <summary>The folder an annotation line list puts the note in ('@Folder("Api.Core")), or null for the root.</summary>
        public static string FolderOf(IEnumerable<string> annotations)
        {
            foreach (var line in annotations ?? Enumerable.Empty<string>())
            {
                var match = FolderAnnotation.Match(line);
                if (match.Success)
                {
                    var folder = match.Groups["folder"].Value.Trim();
                    return folder.Length == 0 ? null : folder;
                }
            }

            return null;
        }

        /// <summary>The annotations with the folder set (null or empty: taken out), every other annotation as it was.</summary>
        public static IReadOnlyList<string> WithFolder(IEnumerable<string> annotations, string folder)
        {
            var result = (annotations ?? Enumerable.Empty<string>()).Where(line => !FolderAnnotation.IsMatch(line)).ToList();
            if (!string.IsNullOrWhiteSpace(folder))
            {
                result.Add("'@Folder(\"" + folder.Trim() + "\")");
            }

            return result;
        }

        /// <summary>Whether the text can be an '@Folder argument: dot-separated names, nothing that would end the annotation.</summary>
        public static bool IsValidFolder(string folder)
            => string.IsNullOrEmpty(folder) || (folder.Trim().Length > 0 && folder.IndexOfAny(new[] { '"', '\'', '(', ')', '\r', '\n' }) < 0);

        /// <summary>The first character of the text that the VBE cannot keep (it stores code in the system's ANSI code page, and anything else becomes a "?"), or null.</summary>
        public static char? FirstUnsupportedCharacter(string text)
        {
            var encoding = System.Text.Encoding.GetEncoding(System.Globalization.CultureInfo.CurrentCulture.TextInfo.ANSICodePage);
            foreach (var character in text ?? string.Empty)
            {
                if (character == '\r' || character == '\n' || character == '\t')
                {
                    continue;
                }

                if (char.IsSurrogate(character) || encoding.GetString(encoding.GetBytes(new[] { character })) != character.ToString())
                {
                    return character;
                }
            }

            return null;
        }

        /// <summary>The markdown in the code of a note: the comment mark and the space after it are taken off; a line without a mark is kept as it is. Annotation lines ('@Notes, '@Folder...) are not text; see <see cref="Annotations"/>.</summary>
        public static string ToMarkdown(string moduleCode)
        {
            var lines = Lines(moduleCode)
                .Where(line => !line.StartsWith("Attribute VB_", StringComparison.OrdinalIgnoreCase) && !IsAnnotation(line))
                .Select(Unmark)
                .ToList();

            while (lines.Count > 0 && lines[lines.Count - 1].Length == 0)
            {
                lines.RemoveAt(lines.Count - 1);
            }

            return string.Join(NewLine, lines);
        }

        /// <summary>The annotation lines of the module, as they are. They make it a note and put it in its folder, so saving the text must not lose them.</summary>
        public static IReadOnlyList<string> Annotations(string moduleCode)
            => Lines(moduleCode).Where(IsAnnotation).Select(line => line.Trim()).ToList();

        /// <summary>The code for the markdown: the annotations first, then a comment mark in front of every line and a space where the line has text.</summary>
        public static string ToModuleCode(string markdown, IEnumerable<string> annotations = null)
            => string.Join(NewLine, (annotations ?? Enumerable.Empty<string>()).Concat(Lines(markdown).Select(line => line.Length == 0 ? "'" : "' " + line)));

        /// <summary>The 1-based number of the first line that is too long for the VBE once marked, or 0.</summary>
        public static int FirstLineTooLong(string markdown)
        {
            var number = 0;
            foreach (var line in Lines(markdown))
            {
                number++;
                if (line.Length + 2 > MaxLineLength)
                {
                    return number;
                }
            }

            return 0;
        }

        private static string NameFromArgument(string argument)
        {
            var name = argument?.Trim();
            return string.IsNullOrEmpty(name) ? null : name;
        }

        private static bool IsNoteLine(string line)
        {
            var trimmed = line.Trim();
            return trimmed.Length == 0
                   || trimmed.StartsWith("'", StringComparison.Ordinal)
                   || trimmed.Equals("Option Explicit", StringComparison.OrdinalIgnoreCase)
                   || trimmed.StartsWith("Attribute VB_", StringComparison.OrdinalIgnoreCase);
        }

        //Text lines get a space after the mark, so a line of markdown that starts with @ is never taken for an annotation.
        private static bool IsAnnotation(string line) => line.TrimStart().StartsWith("'@", StringComparison.Ordinal);

        private static string Unmark(string line)
        {
            var trimmed = line.TrimStart();
            if (!trimmed.StartsWith("'", StringComparison.Ordinal))
            {
                return line;
            }

            trimmed = trimmed.Substring(1);
            return trimmed.StartsWith(" ", StringComparison.Ordinal) ? trimmed.Substring(1) : trimmed;
        }

        private static IEnumerable<string> Lines(string text)
            => string.IsNullOrEmpty(text) ? Enumerable.Empty<string>() : text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
    }
}
