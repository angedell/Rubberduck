using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Rubberduck.Settings;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    /// <summary>
    /// Repairs what <see cref="SyncHygiene"/> finds, the way Rubberduck's own quick fixes do (AddAttributeAnnotation, AddMissingAttribute,
    /// AdjustAttributeValues, AdjustAttributeAnnotation) but on the exported text. Those quick fixes need a parsed module, and the file
    /// side of a sync has none, so the same changes are made here to either side before it is imported or written.
    /// The @Folder annotation is set from the file's location by <see cref="GitSyncText.SetFolderAnnotation"/>.
    /// </summary>
    public static class SyncFixer
    {
        private static readonly Regex AttributeLine = new Regex(@"^Attribute\s+(\w+)\s*=", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ModuleAttribute = new Regex(@"^\s*'@ModuleAttribute\b\s*\(?\s*(\w+)\s*,\s*(.*?)\s*\)?\s*$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static bool Applies(SyncConfig config) => config != null && config.OnViolation == ViolationMode.Fix && config.Enforce.Count > 0;

        /// <summary>The text with every enforced rule that can be repaired repaired. The text itself if nothing needed it or fixing is not switched on.</summary>
        public static string Fix(string name, string exportedText, SyncConfig config)
        {
            if (!Applies(config))
            {
                return exportedText;
            }

            bool Enforced(string rule) => config.Enforce.Contains(rule, StringComparer.OrdinalIgnoreCase);

            var lines = SyncText.Lines(exportedText).ToList();
            var changed = false;

            if (Enforced("MissingAttribute"))
            {
                changed |= SetAttribute(lines, "VB_Name", "\"" + name + "\"", onlyIf: current => current == null || !current.Trim('"').Equals(name, StringComparison.OrdinalIgnoreCase));
            }

            if (Enforced("MissingModuleAnnotation") || Enforced("AttributeValueOutOfSync"))
            {
                changed |= FixPair(lines, "VB_PredeclaredId", "PredeclaredId", Enforced("MissingModuleAnnotation"), Enforced("AttributeValueOutOfSync"));
                changed |= FixPair(lines, "VB_Exposed", "Exposed", Enforced("MissingModuleAnnotation"), Enforced("AttributeValueOutOfSync"));
            }

            if (Enforced("AttributeValueOutOfSync"))
            {
                foreach (var line in lines.Skip(SyncText.CodeStart(lines)).ToList())
                {
                    var match = ModuleAttribute.Match(line);
                    if (match.Success)
                    {
                        changed |= SetAttribute(lines, match.Groups[1].Value, match.Groups[2].Value, onlyIf: current => current == null || !Same(current, match.Groups[2].Value));
                    }
                }
            }

            return changed ? string.Join("\r\n", lines) : exportedText;
        }

        /// <summary>Sets the @Folder annotation of a text that has none from the folder its file is in. Text that has one, or a file in the root, is left alone.</summary>
        public static string FixFolder(string exportedText, string relativeDirectory, SyncConfig config)
        {
            if (!Applies(config) || !config.Enforce.Contains("ModuleWithoutFolder", StringComparer.OrdinalIgnoreCase))
            {
                return exportedText;
            }

            var folder = GitSyncText.FolderPath(relativeDirectory);
            return folder.Length == 0 || GitSyncText.GetFolderAnnotation(SyncText.CodeOf(exportedText)) != null
                ? exportedText
                : GitSyncText.SetFolderAnnotation(exportedText, folder);
        }

        //The attribute and its annotation have to agree. A True attribute gets the annotation; an annotation sets its attribute to True.
        private static bool FixPair(List<string> lines, string attribute, string annotation, bool addAnnotation, bool adjustAttribute)
        {
            var current = ValueOf(lines, attribute);
            var isTrue = current != null && current.Equals("True", StringComparison.OrdinalIgnoreCase);
            var code = lines.Skip(SyncText.CodeStart(lines));
            var annotated = code.Any(line => Regex.IsMatch(line, @"^\s*'@" + annotation + @"\b", RegexOptions.IgnoreCase));

            if (isTrue && !annotated && addAnnotation)
            {
                lines.Insert(SyncText.CodeStart(lines), "'@" + annotation);
                return true;
            }

            if (annotated && !isTrue && adjustAttribute)
            {
                return SetAttribute(lines, attribute, "True", onlyIf: _ => true);
            }

            return false;
        }

        private static string ValueOf(IList<string> lines, string attribute)
        {
            return SyncText.ModuleAttributes(string.Join("\r\n", lines)).TryGetValue(attribute, out var value) ? value : null;
        }

        private static bool Same(string a, string b) =>
            string.Equals(Regex.Replace(a, @"\s*,\s*", ", ").Trim(), Regex.Replace(b, @"\s*,\s*", ", ").Trim(), StringComparison.OrdinalIgnoreCase);

        //Changes the attribute line, or adds one after the last attribute of the header.
        private static bool SetAttribute(List<string> lines, string attribute, string value, Func<string, bool> onlyIf)
        {
            var current = ValueOf(lines, attribute);
            if (!onlyIf(current))
            {
                return false;
            }

            var text = $"Attribute {attribute} = {value}";
            var end = SyncText.CodeStart(lines);
            for (var i = 0; i < end; i++)
            {
                var match = AttributeLine.Match(lines[i]);
                if (match.Success && match.Groups[1].Value.Equals(attribute, StringComparison.OrdinalIgnoreCase))
                {
                    lines[i] = text;
                    return true;
                }
            }

            //VB_Name leads the attributes; the others follow it.
            var index = attribute.Equals("VB_Name", StringComparison.OrdinalIgnoreCase)
                ? FirstAttribute(lines, end)
                : end;
            lines.Insert(index, text);
            return true;
        }

        private static int FirstAttribute(IList<string> lines, int end)
        {
            for (var i = 0; i < end; i++)
            {
                if (lines[i].StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return end;
        }
    }
}
