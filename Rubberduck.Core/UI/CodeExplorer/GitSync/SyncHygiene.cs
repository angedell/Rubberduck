using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Rubberduck.Settings;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    public enum SyncSide
    {
        Module,
        File
    }

    public class SyncViolation
    {
        public SyncViolation(string rule, string message, SyncSide side, bool blocking, bool fixable = false)
        {
            Rule = rule;
            Message = message;
            Side = side;
            Blocking = blocking;
            Fixable = fixable;
        }

        /// <summary>One of <see cref="SyncConfig.Rules"/>.</summary>
        public string Rule { get; }
        public string Message { get; }

        /// <summary>The side that breaks the rule.</summary>
        public SyncSide Side { get; }

        /// <summary>Whether the side cannot be used as the source of a sync until this is fixed.</summary>
        public bool Blocking { get; }

        /// <summary>Whether the sync repairs it by itself (<see cref="SyncFixer"/>).</summary>
        public bool Fixable { get; }
    }

    /// <summary>
    /// The rules that keep a module and its file from drifting apart: the hidden attributes and the annotations that mirror them,
    /// and the @Folder annotation that decides where the file goes. They work on the exported text, so they run on both sides and need no parser.
    /// </summary>
    public static class SyncHygiene
    {
        private static readonly Regex Annotation = new Regex(@"^\s*'@(\w+)\b(.*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <param name="name">The module's name.</param>
        /// <param name="exportedText">The module as the VBE exports it, or the file.</param>
        /// <param name="folderFollowsPath">Whether the @Folder annotation is set from the location of the file when importing, which makes a missing one harmless.</param>
        public static IEnumerable<SyncViolation> Check(string name, string exportedText, SyncSide side, SyncConfig config, bool folderFollowsPath = false)
        {
            if (config == null || config.Enforce.Count == 0)
            {
                return Enumerable.Empty<SyncViolation>();
            }

            var blocking = config.OnViolation == ViolationMode.Block;
            var violations = new List<SyncViolation>();
            void Add(string rule, string message)
            {
                if (config.Enforce.Contains(rule, StringComparer.OrdinalIgnoreCase))
                {
                    violations.Add(new SyncViolation(rule, message, side, blocking, config.OnViolation == ViolationMode.Fix && rule != "ModuleWithoutFolder"));
                }
            }

            var attributes = SyncText.ModuleAttributes(exportedText);
            var code = SyncText.CodeOf(exportedText);
            var annotations = SyncText.Lines(code)
                .Select(line => Annotation.Match(line))
                .Where(match => match.Success)
                .Select(match => new { Name = match.Groups[1].Value, Arguments = match.Groups[2].Value.Trim() })
                .ToList();

            bool HasAnnotation(string annotation) => annotations.Any(a => a.Name.Equals(annotation, StringComparison.OrdinalIgnoreCase));
            bool IsTrue(string attribute) => attributes.TryGetValue(attribute, out var value) && value.Equals("True", StringComparison.OrdinalIgnoreCase);

            if (!attributes.TryGetValue("VB_Name", out var declaredName))
            {
                Add("MissingAttribute", "Attribute VB_Name is missing.");
            }
            else if (!declaredName.Trim('"').Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                Add("MissingAttribute", $"Attribute VB_Name is {declaredName}, which is not the module's name '{name}'.");
            }

            if (!folderFollowsPath && GitSyncText.GetFolderAnnotation(code) == null)
            {
                Add("ModuleWithoutFolder", "The module has no @Folder annotation.");
            }

            if (IsTrue("VB_PredeclaredId") && !HasAnnotation("PredeclaredId"))
            {
                Add("MissingModuleAnnotation", "VB_PredeclaredId is True but there is no @PredeclaredId annotation.");
            }

            if (IsTrue("VB_Exposed") && !HasAnnotation("Exposed"))
            {
                Add("MissingModuleAnnotation", "VB_Exposed is True but there is no @Exposed annotation.");
            }

            if (HasAnnotation("PredeclaredId") && !IsTrue("VB_PredeclaredId"))
            {
                Add("AttributeValueOutOfSync", "@PredeclaredId is there but VB_PredeclaredId is not True.");
            }

            if (HasAnnotation("Exposed") && !IsTrue("VB_Exposed"))
            {
                Add("AttributeValueOutOfSync", "@Exposed is there but VB_Exposed is not True.");
            }

            foreach (var annotation in annotations.Where(a => a.Name.Equals("ModuleAttribute", StringComparison.OrdinalIgnoreCase)))
            {
                var arguments = SplitArguments(annotation.Arguments.Trim('(', ')'));
                if (arguments.Count < 2)
                {
                    continue;
                }

                var expected = string.Join(", ", arguments.Skip(1));
                if (!attributes.TryGetValue(arguments[0], out var actual))
                {
                    Add("AttributeValueOutOfSync", $"@ModuleAttribute {arguments[0]} has no matching attribute.");
                }
                else if (!string.Equals(Compact(actual), Compact(expected), StringComparison.OrdinalIgnoreCase))
                {
                    Add("AttributeValueOutOfSync", $"@ModuleAttribute {arguments[0]} says {expected} but the attribute is {actual}.");
                }
            }

            return violations;
        }

        private static string Compact(string value) => Regex.Replace(value, @"\s*,\s*", ", ").Trim();

        //Splits on the commas that are not inside a string literal.
        private static List<string> SplitArguments(string text)
        {
            var result = new List<string>();
            var current = new StringBuilder();
            var quoted = false;
            foreach (var c in text)
            {
                if (c == '"')
                {
                    quoted = !quoted;
                }

                if (c == ',' && !quoted)
                {
                    result.Add(current.ToString().Trim());
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }

            result.Add(current.ToString().Trim());
            return result;
        }
    }
}
