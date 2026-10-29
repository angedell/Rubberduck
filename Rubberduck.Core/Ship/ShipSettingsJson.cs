using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Rubberduck.Ship
{
    /// <summary>The <c>ship</c> section of rubberduck.json: what it is read from and how it is written, with a comment for every setting.</summary>
    public static class ShipSettingsJson
    {
        public const string Key = "ship";

        public static readonly string[] FolderForms =
        {
            "Where the shipped file is written. Written as you like:",
            "  \"dist\"                      a folder relative to the folder of the workbook",
            "  \"../..\"                     two folders up from it (any number of ../)",
            "  \"%USERPROFILE%/Releases\"    these environment variables are replaced: USERPROFILE, HOMEDRIVE, HOMEPATH, APPDATA, LOCALAPPDATA, PROGRAMDATA, PUBLIC, TEMP, TMP, SYSTEMDRIVE, OneDrive, OneDriveCommercial, OneDriveConsumer",
            "  \"Documents\", \"Desktop\", \"Downloads\", \"Music\", \"Pictures\", \"Videos\"   a known Windows folder by its name",
            "  \"Documents/Releases\"        a folder inside a known folder",
            "  \"./Documents\"               a folder next to the workbook that has the name of a known folder",
            "  \"OneDrive\"                  your OneDrive folder (\"OneDrive/Releases\" for a folder in it)",
            "  \"Dropbox\"                   your main Dropbox folder (\"Dropbox/Releases\" for a folder in it)",
            "  \"C:/Releases\"               a full path",
            "Use / in paths, or write \\\\ for each backslash. Folders are resolved on each machine, so a team can share one setting."
        };

        private static readonly Dictionary<string, string> StepComments = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { ShipSteps.Notes, "Leave out the notes (the @Notes modules)." },
            { ShipSteps.TestModules, "Leave out the unit test modules (@TestModule) and the reference to Rubberduck's unit testing library." },
            { ShipSteps.LineNumbers, "Number the executable lines of every procedure with their line in your module (as the code pane shows it), so error handlers in the shipped file can report Erl. Off by default. A procedure that already has line numbers is left alone." },
            { ShipSteps.TreeShaking, "Remove the code nothing reachable uses (see \"treeShaking\" below for what counts as reachable). Off by default: it is the one step that can remove code you need. The preview lists everything it would remove." },
            { ShipSteps.Comments, "Remove every comment, the '@ annotations included." },
            { ShipSteps.DescriptiveAttributes, "Remove the attributes that only describe (such as VB_Ext_KEY). The ones that change how code runs stay, and so do the descriptions of public procedures in standard modules." },
            { ShipSteps.Config, "Leave out this file (customUI/rubberduck.json) from the shipped file. The ribbon and its pictures stay." }
        };

        /// <summary>Reads the section's body into <paramref name="ship"/>; what is wrong with it is added to <paramref name="errors"/>.</summary>
        public static void Read(JObject section, ShipSettings ship, ICollection<string> errors)
        {
            foreach (var property in section.Properties())
            {
                var isString = property.Value.Type == JTokenType.String;
                var text = isString ? property.Value.Value<string>() : null;
                switch (property.Name.ToLowerInvariant())
                {
                    case "format":
                        if (isString && Enum.TryParse(text, true, out ShipFormat format) && Enum.IsDefined(typeof(ShipFormat), format))
                        {
                            ship.Format = format;
                        }
                        else
                        {
                            errors.Add($"Invalid value for ship setting 'format': {property.Value}. Use \"xlsm\" or \"xlam\".");
                        }

                        break;
                    case "outputfolder":
                        if (isString)
                        {
                            ship.OutputFolder = string.IsNullOrWhiteSpace(text) ? null : text.Trim();
                        }
                        else
                        {
                            errors.Add($"Invalid value for ship setting 'outputFolder': {property.Value}. Use a folder as text.");
                        }

                        break;
                    case "filename":
                        if (isString)
                        {
                            ship.FileName = string.IsNullOrWhiteSpace(text) ? null : text.Trim();
                        }
                        else
                        {
                            errors.Add($"Invalid value for ship setting 'fileName': {property.Value}. Use a name as text, without extension.");
                        }

                        break;
                    case "onexisting":
                        if (isString && Enum.TryParse(text, true, out ShipOnExisting onExisting) && Enum.IsDefined(typeof(ShipOnExisting), onExisting))
                        {
                            ship.OnExisting = onExisting;
                        }
                        else
                        {
                            errors.Add($"Invalid value for ship setting 'onExisting': {property.Value}. Use \"ask\", \"replace\" or \"keepBoth\".");
                        }

                        break;
                    case "steps":
                        ReadSteps(property, ship, errors);
                        break;
                    case "treeshaking":
                        ReadTreeShaking(property, ship.TreeShaking, errors);
                        break;
                    case "$comment":
                        break;
                    default:
                        errors.Add($"Unknown ship setting '{property.Name}'.");
                        break;
                }
            }
        }

        private static void ReadSteps(JProperty property, ShipSettings ship, ICollection<string> errors)
        {
            if (!(property.Value is JObject steps))
            {
                errors.Add("Ship setting 'steps' must be an object of step names and true or false.");
                return;
            }

            foreach (var step in steps.Properties())
            {
                var known = ShipSteps.All.FirstOrDefault(name => name.Equals(step.Name, StringComparison.OrdinalIgnoreCase));
                if (step.Name.Equals("$comment", StringComparison.Ordinal))
                {
                    continue;
                }

                if (known == null)
                {
                    errors.Add($"Unknown ship step '{step.Name}'. Known steps: {string.Join(", ", ShipSteps.All)}.");
                }
                else if (step.Value.Type == JTokenType.Boolean)
                {
                    ship.Steps[known] = step.Value.Value<bool>();
                }
                else
                {
                    errors.Add($"Invalid value for ship step '{step.Name}': {step.Value}. Use true or false.");
                }
            }
        }

        private static readonly string[] ModeNames = { "conservative", "entryPointsOnly" };

        private static void ReadTreeShaking(JProperty property, ShipTreeShakingSettings settings, ICollection<string> errors)
        {
            if (!(property.Value is JObject body))
            {
                errors.Add("Ship setting 'treeShaking' must be an object with 'mode' and 'keep'.");
                return;
            }

            foreach (var setting in body.Properties())
            {
                switch (setting.Name.ToLowerInvariant())
                {
                    case "mode":
                        var (mode, modeError) = ShipKeepList.NormalizeMode(setting.Value);
                        if (mode == null)
                        {
                            errors.Add(modeError);
                        }
                        else
                        {
                            settings.Mode = mode.Value;
                        }

                        break;
                    case "keep":
                        var (valid, problems) = ShipKeepList.Normalize(setting.Value);
                        foreach (var problem in problems)
                        {
                            errors.Add(problem);
                        }

                        if (valid != null)
                        {
                            settings.Keep.Clear();
                            settings.Keep.AddRange(valid);
                        }

                        break;
                    case "$comment":
                        break;
                    default:
                        errors.Add($"Unknown ship setting 'treeShaking.{setting.Name}'.");
                        break;
                }
            }
        }

        /// <summary>The comment above the section.</summary>
        public static string[] SectionComment { get; } =
        {
            "Ship writes a clean copy of the saved workbook as .xlsm or .xlam, without what Rubberduck adds for development",
            "(notes, test modules, comments, this file). Everyone on the project shares these settings."
        };

        /// <summary>The section as text, from <c>{</c> to <c>}</c>, every setting with a comment above it. <paramref name="indent"/> is the indentation of the line the section starts on.</summary>
        public static string Section(ShipSettings ship, string indent, string newline = "\r\n")
        {
            ship = ship ?? new ShipSettings();
            var inner = indent + "  ";
            var text = new StringBuilder();
            void Line(string line = "") => text.Append(line).Append(newline);
            void Comment(string margin, params string[] lines)
            {
                foreach (var line in lines)
                {
                    Line(margin + "// " + line);
                }
            }

            Line("{");
            Comment(inner, "The kind of file: \"xlsm\" (macro-enabled workbook) or \"xlam\" (Excel add-in).");
            Line(inner + "\"format\": " + JsonConvert.ToString(ship.Format.ToString().ToLowerInvariant()) + ",");
            Line();
            Comment(inner, FolderForms);
            Line(inner + "\"outputFolder\": " + JsonConvert.ToString(ship.OutputFolder ?? string.Empty) + ",");
            Line();
            Comment(inner, "The name of the file, without extension. Empty: the name of the project.");
            Line(inner + "\"fileName\": " + JsonConvert.ToString(ship.FileName ?? string.Empty) + ",");
            Line();
            Comment(inner, "When the file is already there: \"ask\", \"replace\" (the old file is kept as .bak, or .bak2 and so on if that exists; agents can only replace when the user switched on \"Ship without asking\")",
                "or \"keepBoth\" (the new file is called \"Name (2)\").");
            Line(inner + "\"onExisting\": " + JsonConvert.ToString(OnExistingName(ship.OnExisting)) + ",");
            Line();
            Comment(inner, "What is taken out of the shipped file. true: removed, false: kept.");
            Line(inner + "\"steps\": {");
            for (var i = 0; i < ShipSteps.All.Length; i++)
            {
                var step = ShipSteps.All[i];
                Comment(inner + "  ", StepComments[step]);
                Line(inner + "  \"" + step + "\": " + (ship.IsOn(step) ? "true" : "false") + (i < ShipSteps.All.Length - 1 ? "," : string.Empty));
            }

            Line(inner + "},");
            Line();
            Comment(inner, "How the \"treeShaking\" step decides what is reachable. It reads what Rubberduck parsed, so the project has to parse without errors; if it does not, the step is skipped and the preview says why.",
                "Always kept: event handlers, interface implementations, Class_Initialize/Terminate, Auto_Open/Auto_Close, @EntryPoint and @ExcelHotkey members, public members of exposed classes,",
                "member names found in a string anywhere in the code, and in the ribbon, shape macros, defined names and formulas of the workbook.");
            Line(inner + "\"treeShaking\": {");
            Comment(inner + "  ", "\"conservative\": every public procedure of a standard module is kept (it may be a macro or a worksheet function).",
                "\"entryPointsOnly\": only what is listed above and in \"keep\" is kept, so a public procedure nothing calls is removed.");
            Line(inner + "  \"mode\": " + JsonConvert.ToString(ship.TreeShaking.Mode == ShipTreeShakingMode.EntryPointsOnly ? "entryPointsOnly" : "conservative") + ",");
            Comment(inner + "  ", "Names that are never removed: \"Name\" (a procedure, variable, constant, Type, Enum or module of that name anywhere) or \"Module.Name\". Not case sensitive.");
            Line(inner + "  \"keep\": [" + string.Join(", ", ship.TreeShaking.Keep.Select(k => JsonConvert.ToString(k))) + "]");
            Line(inner + "}");
            text.Append(indent).Append("}");
            return text.ToString();
        }

        private static string OnExistingName(ShipOnExisting value)
            => value == ShipOnExisting.KeepBoth ? "keepBoth" : value.ToString().ToLowerInvariant();
    }
}
