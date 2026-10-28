using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Rubberduck.Settings;
using Rubberduck.UI.CodeExplorer.GitSync;

namespace Rubberduck.Api
{
    /// <summary>One hidden attribute that differs from what the module's annotations say.</summary>
    public class AttributeChange
    {
        /// <summary>"module", or the name of the member.</summary>
        public string Scope { get; set; }

        public string Attribute { get; set; }

        /// <summary>What the attribute is now; null if it is missing.</summary>
        public JToken From { get; set; }

        public JToken To { get; set; }

        /// <summary>The annotation that asks for it, e.g. "@DefaultMember".</summary>
        public string Because { get; set; }
    }

    public class SyncAttributesReport
    {
        public string Project { get; set; }
        public string Module { get; set; }

        /// <summary>True if the attributes already agree with the annotations.</summary>
        public bool InSync { get; set; }

        /// <summary>True if the changes were written; false for a preview.</summary>
        public bool Applied { get; set; }

        public List<AttributeChange> Changes { get; set; } = new List<AttributeChange>();

        /// <summary>The module's hash as of the read (before any change). Send it as expectedHash to apply what the preview showed.</summary>
        public string Hash { get; set; }

        /// <summary>The module's hash after the change, if it was applied.</summary>
        public string NewHash { get; set; }
    }

    /// <summary>
    /// Brings the hidden attributes of a module in line with its annotations: the direction of the "attribute out of sync" inspections' quick fixes
    /// (AdjustAttributeValues, AddMissingAttribute), which need a parsed module and the user at the keyboard. The module-level part is the
    /// <see cref="SyncFixer"/> that git sync uses (VB_Name, @PredeclaredId, @Exposed, @ModuleAttribute); @ModuleDescription and the members' annotations
    /// (@DefaultMember, @Enumerator, @Description, @MemberAttribute) are read here from the code, and everything is written by <see cref="ModuleAttributes.Apply"/>.
    /// Annotations are never changed: where an attribute has no annotation, it is left as it is.
    /// </summary>
    public class SyncAttributesApi
    {
        private static readonly SyncConfig Rules = BuildRules();

        private static readonly Regex AnnotationLine = new Regex(@"^\s*'\s*@(?<name>\w+)\b[\s(]*(?<args>.*?)\s*\)?\s*$", RegexOptions.Compiled);
        private static readonly Regex DeclarationLine = new Regex(@"^\s*(?:(?:Public|Private|Friend|Global|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(?<name>\w+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private readonly IModuleStore _store;
        private readonly ApiWriteApproval _approval;
        private readonly Func<bool> _isInDesignMode;

        /// <param name="isInDesignMode">Where the editor is; applying is refused (409) unless it is in design mode. Null means no check here (the store has its own).</param>
        public SyncAttributesApi(IModuleStore store, ApiWriteApproval approval, Func<bool> isInDesignMode = null)
        {
            _store = store;
            _approval = approval;
            _isInDesignMode = isInDesignMode;
        }

        private static SyncConfig BuildRules()
        {
            var config = new SyncConfig { OnViolation = ViolationMode.Fix };
            config.Enforce.Add("MissingAttribute");
            config.Enforce.Add("AttributeValueOutOfSync");
            return config;
        }

        /// <summary>What would change, without changing anything and without approval.</summary>
        public SyncAttributesReport Preview(string project, string module)
        {
            var current = _store.Read(project, module, true);
            return Plan(current, out _);
        }

        /// <summary>Writes the changes <see cref="Preview"/> shows, as one whole-module write. Nothing is written (and nobody asked) if there are none.</summary>
        public SyncAttributesReport Apply(string client, string project, string module, string expectedHash)
        {
            if (string.IsNullOrWhiteSpace(expectedHash))
            {
                throw ApiException.PreconditionRequired("Send the hash from your last read or preview as 'expectedHash' (or the If-Match header), so edits made in the meantime are not overwritten.");
            }

            if (_isInDesignMode != null && !_isInDesignMode())
            {
                throw ApiException.NotInDesignMode();
            }

            var current = _store.Read(project, module, true);
            if (!string.Equals(current.Hash, expectedHash, StringComparison.OrdinalIgnoreCase))
            {
                throw ApiException.PreconditionFailed($"'{module}' changed since you read it. Read it again.");
            }

            var report = Plan(current, out var text);
            if (report.InSync)
            {
                return report;
            }

            _approval.Require(client);
            var written = _store.ReplaceText(project, module, text, expectedHash);
            report.Applied = true;
            report.NewHash = written.Hash;
            return report;
        }

        private static SyncAttributesReport Plan(ModuleContent current, out string newText)
        {
            var text = current.Text ?? string.Empty;
            var before = ModuleAttributes.Parse(text);
            var because = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            //Module level: what git sync repairs, plus the module's description.
            var fixedText = SyncFixer.Fix(current.Name, text, Rules);
            var wanted = new JObject();
            ReadAnnotations(fixedText, wanted, because);
            newText = wanted.HasValues ? ModuleAttributes.Apply(fixedText, wanted) : fixedText;

            var after = ModuleAttributes.Parse(newText);
            var report = new SyncAttributesReport { Project = current.Project, Module = current.Name, Hash = current.Hash };
            Diff(report, "module", before["module"] as JObject, after["module"] as JObject, because);
            foreach (var member in (after["members"] as JObject).Properties())
            {
                Diff(report, member.Name, (before["members"] as JObject)?[member.Name] as JObject, member.Value as JObject, because);
            }

            report.InSync = report.Changes.Count == 0;
            if (report.InSync)
            {
                newText = text;
            }

            return report;
        }

        private static void Diff(SyncAttributesReport report, string scope, JObject before, JObject after, IDictionary<string, string> because)
        {
            foreach (var attribute in after.Properties())
            {
                var old = before?[attribute.Name];
                if (old != null && JToken.DeepEquals(old, attribute.Value))
                {
                    continue;
                }

                because.TryGetValue(scope + "." + attribute.Name, out var reason);
                report.Changes.Add(new AttributeChange
                {
                    Scope = scope,
                    Attribute = attribute.Name,
                    From = old,
                    To = attribute.Value,
                    Because = reason ?? ModuleReason(attribute.Name)
                });
            }
        }

        private static string ModuleReason(string attribute)
        {
            switch (attribute)
            {
                case "VB_Name":
                    return "the module's name";
                case "VB_PredeclaredId":
                    return "@PredeclaredId";
                case "VB_Exposed":
                    return "@Exposed";
                default:
                    return "@ModuleAttribute";
            }
        }

        //Collects the attributes that the annotations in the code ask for: into 'wanted' (the shape ModuleAttributes.Apply takes) and, per "scope.attribute", the annotation.
        //An annotation belongs to the first declaration after it; blank lines and comments in between do not break that.
        private static void ReadAnnotations(string exportedText, JObject wanted, IDictionary<string, string> because)
        {
            var lines = SyncText.Lines(exportedText);
            var pending = new List<(string Name, string Args)>();
            var module = new JObject();
            var members = new JObject();

            foreach (var line in lines.Skip(SyncText.CodeStart(lines)))
            {
                var annotation = AnnotationLine.Match(line);
                if (annotation.Success)
                {
                    var name = annotation.Groups["name"].Value;
                    var args = annotation.Groups["args"].Value.Trim();
                    if (name.Equals("ModuleDescription", StringComparison.OrdinalIgnoreCase))
                    {
                        module["VB_Description"] = ValueOf(args);
                        because["module.VB_Description"] = "@ModuleDescription";
                    }
                    else
                    {
                        pending.Add((name, args));
                    }

                    continue;
                }

                if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("'", StringComparison.Ordinal))
                {
                    continue;
                }

                var declaration = DeclarationLine.Match(line);
                if (declaration.Success)
                {
                    var member = declaration.Groups["name"].Value;
                    foreach (var (name, args) in pending)
                    {
                        Member(members, because, member, name, args);
                    }
                }

                pending.Clear();
            }

            if (module.HasValues)
            {
                wanted["module"] = module;
            }

            if (members.HasValues)
            {
                wanted["members"] = members;
            }
        }

        private static void Member(JObject members, IDictionary<string, string> because, string member, string annotation, string args)
        {
            string attribute;
            JToken value;
            switch (annotation.ToLowerInvariant())
            {
                case "defaultmember":
                    attribute = "VB_UserMemId";
                    value = 0L;
                    break;
                case "enumerator":
                    attribute = "VB_UserMemId";
                    value = -4L;
                    break;
                case "description":
                    attribute = "VB_Description";
                    value = ValueOf(args);
                    break;
                case "memberattribute":
                    var comma = args.IndexOf(',');
                    if (comma < 0 || !args.Substring(0, comma).Trim().StartsWith("VB_", StringComparison.Ordinal))
                    {
                        return;
                    }

                    attribute = args.Substring(0, comma).Trim();
                    value = ValueOf(args.Substring(comma + 1).Trim());
                    break;
                default:
                    return;
            }

            var values = members[member] as JObject;
            if (values == null)
            {
                values = new JObject();
                members[member] = values;
            }

            values[attribute] = value;
            because[member + "." + attribute] = "@" + annotation;
        }

        //The value as an attribute line would hold it: "text", True/False or a number, read the way ModuleAttributes reads them.
        private static JToken ValueOf(string raw)
            => ModuleAttributes.Parse("Attribute x.VB_Value = " + raw)["members"]["x"]["VB_Value"];
    }
}
