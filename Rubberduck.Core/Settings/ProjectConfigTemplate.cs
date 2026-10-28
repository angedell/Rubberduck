using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using Rubberduck.Ship;
using Rubberduck.SmartIndenter;

namespace Rubberduck.Settings
{
    /// <summary>
    /// The text of a new <c>rubberduck.json</c>: every setting with the value in effect and a comment saying what it does.
    /// The indenter section holds the user's own settings at the time, so that creating the file does not change how anything is indented.
    /// </summary>
    public static class ProjectConfigTemplate
    {
        /// <summary>The indenter settings in the order and with the comment they get in the file.</summary>
        public static readonly (string Name, string Comment)[] IndenterEntries =
        {
            (nameof(IIndenterSettings.IndentSpaces), "Spaces for each level of indentation."),
            (nameof(IIndenterSettings.IndentEntireProcedureBody), "Indent the code between the first line of a procedure and its End line."),
            (nameof(IIndenterSettings.IndentFirstCommentBlock), "Indent the comment block at the start of a procedure."),
            (nameof(IIndenterSettings.IndentFirstDeclarationBlock), "Indent the declarations at the start of a procedure."),
            (nameof(IIndenterSettings.IgnoreEmptyLinesInFirstBlocks), "Empty lines do not end the first comment or declaration block of a procedure."),
            (nameof(IIndenterSettings.IndentEnumTypeAsProcedure), "Indent the members of an Enum or Type the way a procedure body is."),
            (nameof(IIndenterSettings.IndentCase), "Indent the Case branches of a Select Case."),
            (nameof(IIndenterSettings.IndentCompilerDirectives), "Indent compiler directives (#If ... #End If) with the code."),
            (nameof(IIndenterSettings.ForceCompilerDirectivesInColumn1), "Keep compiler directives in the first column instead."),
            (nameof(IIndenterSettings.AlignCommentsWithCode), "Align full-line comments with the code that follows them."),
            (nameof(IIndenterSettings.AlignContinuations), "Align continued lines (the ones ending in _) with the line they continue."),
            (nameof(IIndenterSettings.IgnoreOperatorsInContinuations), "When aligning continued lines, do not count a leading operator."),
            (nameof(IIndenterSettings.AlignDims), "Align the variable names of Dim statements in a column."),
            (nameof(IIndenterSettings.AlignDimColumn), "The column Dim statements are aligned to, when alignDims is on."),
            (nameof(IIndenterSettings.EndOfLineCommentStyle), "Where comments after code go: \"Absolute\", \"SameGap\", \"StandardGap\" or \"AlignInColumn\"."),
            (nameof(IIndenterSettings.EndOfLineCommentColumnSpaceAlignment), "The column or the number of spaces the end-of-line comment style above uses."),
            (nameof(IIndenterSettings.ForceDebugStatementsInColumn1), "Keep all Debug statements and Stop in the first column. Set before the next three, which refine it."),
            (nameof(IIndenterSettings.ForceDebugPrintInColumn1), "Keep Debug.Print in the first column."),
            (nameof(IIndenterSettings.ForceDebugAssertInColumn1), "Keep Debug.Assert in the first column."),
            (nameof(IIndenterSettings.ForceStopInColumn1), "Keep Stop in the first column."),
            (nameof(IIndenterSettings.EmptyLineHandlingMethod), "What to do with the indentation of empty lines: \"Ignore\", \"Remove\" or \"Indent\"."),
            (nameof(IIndenterSettings.VerticallySpaceProcedures), "Put a fixed number of empty lines between procedures."),
            (nameof(IIndenterSettings.LinesBetweenProcedures), "How many empty lines (0 to 2) go between procedures, when verticallySpaceProcedures is on."),
            (nameof(IIndenterSettings.GroupRelatedProperties), "No empty lines between the Get, Let and Set of the same property."),
            (nameof(IIndenterSettings.MaximumConsecutiveEmptyLines), "Most empty lines allowed in a row anywhere in the code (0 to 10); extra ones are removed. 0 means no limit."),
            (nameof(IIndenterSettings.RemoveEmptyLinesAtProcedureEdges), "Remove empty lines right after the first line of a procedure and right before its End line."),
            (nameof(IIndenterSettings.RemoveLeadingEmptyLines), "Remove empty lines at the very top of a module.")
        };

        private static string CamelCase(string name) => char.ToLowerInvariant(name[0]) + name.Substring(1);

        public static string Create(IIndenterSettings settings)
        {
            var user = ProjectConfigReader.Copy(settings);
            var text = new StringBuilder();
            void Line(string line = "") => text.Append(line).Append("\r\n");

            Line("{");
            Line("  // Rubberduck settings for this repository. Rubberduck reads this file from the folder a project is synced with");
            Line("  // (Git Sync) or from a parent folder, and it applies to everyone working on the project.");
            Line("  // A setting that is left out keeps the value from each person's own Rubberduck options. Comments are allowed.");
            Line();
            Line("  // How code is indented, in the editor as much as by the sync. The values here were taken from the Rubberduck");
            Line("  // options of whoever created the file; change them to what the team agrees on.");
            Line("  \"indenter\": {");

            for (var i = 0; i < IndenterEntries.Length; i++)
            {
                var (name, comment) = IndenterEntries[i];
                Line("    // " + comment);
                Line("    \"" + CamelCase(name) + "\": " + ValueOf(user, name) + (i < IndenterEntries.Length - 1 ? "," : string.Empty));
            }

            Line("  },");
            Line();
            Line("  // How the project is kept in step with this folder.");
            Line("  \"sync\": {");
            Line("    // Files that are never synced, one pattern per entry, written like .gitignore lines and relative to this folder,");
            Line("    // for example \"/modScratch.bas\" or \"Sandbox/\". The Git Sync window adds to this list with \"Ignore permanently\".");
            Line("    \"ignore\": [],");
            Line();
            Line("    // Rules a module and its file have to follow before either can be used as the source of a sync. Available:");
            Line("    //   \"MissingAttribute\"          the module has no Attribute VB_Name, or it does not match its name");
            Line("    //   \"ModuleWithoutFolder\"       the module has no @Folder annotation");
            Line("    //   \"MissingModuleAnnotation\"   an attribute such as VB_PredeclaredId or VB_Exposed has no matching annotation");
            Line("    //   \"AttributeValueOutOfSync\"   an annotation such as @PredeclaredId or @ModuleAttribute disagrees with its attribute");
            Line("    \"enforce\": [],");
            Line();
            Line("    // What happens to a module or file that breaks an enforced rule: \"fix\" repairs it while it is synced (adds the missing");
            Line("    // annotation or attribute, sets an attribute to what its annotation says; a missing @Folder is set from the file's folder),");
            Line("    // \"block\" cannot be synced until it is fixed by hand, \"warn\" is only flagged in the Git Sync window.");
            Line("    \"onViolation\": \"fix\",");
            Line();
            Line("    // VBA does not care about case, so with true (the default) a difference in upper and lower case is not a difference,");
            Line("    // except inside a string literal, where it matters. Set to false to treat every change of case as a change.");
            Line("    \"ignoreCaseOutsideStrings\": true");
            Line("  },");
            Line();
            foreach (var comment in ShipSettingsJson.SectionComment)
            {
                Line("  // " + comment);
            }

            Line("  \"" + ShipSettingsJson.Key + "\": " + ShipSettingsJson.Section(new ShipSettings(), "  "));
            Line("}");
            return text.ToString();
        }

        private static string ValueOf(IIndenterSettings settings, string name)
        {
            var property = typeof(IIndenterSettings).GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
            var value = property.GetValue(settings);
            switch (value)
            {
                case bool flag:
                    return flag ? "true" : "false";
                case Enum named:
                    return "\"" + named + "\"";
                default:
                    return Convert.ToString(value, CultureInfo.InvariantCulture);
            }
        }

        /// <summary>The settings the indenter has that the template does not mention, so that a new setting is not forgotten.</summary>
        public static string[] Missing()
        {
            var listed = IndenterEntries.Select(entry => entry.Name).ToList();
            return typeof(IIndenterSettings).GetProperties()
                .Where(p => p.CanRead && p.CanWrite && !listed.Contains(p.Name))
                .Select(p => p.Name)
                .ToArray();
        }
    }
}
