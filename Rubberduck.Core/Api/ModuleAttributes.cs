using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace Rubberduck.Api
{
    /// <summary>
    /// The hidden <c>Attribute VB_...</c> lines of a module (<c>VB_PredeclaredId</c>, <c>VB_Description</c>, <c>VB_UserMemId</c>...) as data, and the other way around.
    /// The shape is <c>{ "module": { "VB_PredeclaredId": true }, "members": { "Greet": { "VB_UserMemId": 0 } } }</c>, so nobody has to read or rebuild the exported header.
    /// Pure text work; the VBE's own import is still what writes the module.
    /// </summary>
    public static class ModuleAttributes
    {
        private static readonly Regex AttributeLine = new Regex(@"^Attribute\s+(?:(?<member>\w+)\.)?(?<name>VB_\w+)\s*=\s*(?<value>.*?)\s*$", RegexOptions.Compiled);
        private static readonly Regex DeclarationLine = new Regex(@"^\s*(?:(?:Public|Private|Friend|Global|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(?<name>\w+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ValidKey = new Regex(@"^VB_\w+$", RegexOptions.Compiled);

        public static JObject Parse(string exportedText)
        {
            var module = new JObject();
            var members = new JObject();
            foreach (var line in Lines(exportedText))
            {
                var match = AttributeLine.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                var target = module;
                if (match.Groups["member"].Success)
                {
                    var member = match.Groups["member"].Value;
                    target = members[member] as JObject;
                    if (target == null)
                    {
                        target = new JObject();
                        members[member] = target;
                    }
                }

                target[match.Groups["name"].Value] = ParseValue(match.Groups["value"].Value);
            }

            return new JObject { ["module"] = module, ["members"] = members };
        }

        /// <summary>Sets the attributes named in <paramref name="changes"/> (the shape <see cref="Parse"/> gives), replacing a line that is there and adding one that is not.</summary>
        public static string Apply(string exportedText, JObject changes)
        {
            var lines = Lines(exportedText).ToList();

            if (changes["module"] is JObject module)
            {
                foreach (var property in module.Properties())
                {
                    SetModuleAttribute(lines, Key(property.Name), property.Value);
                }
            }

            if (changes["members"] is JObject members)
            {
                foreach (var member in members.Properties())
                {
                    if (!(member.Value is JObject values))
                    {
                        throw ApiException.BadRequest($"The attributes of member '{member.Name}' must be an object.");
                    }

                    foreach (var property in values.Properties())
                    {
                        SetMemberAttribute(lines, member.Name, Key(property.Name), property.Value);
                    }
                }
            }

            return string.Join("\r\n", lines);
        }

        private static void SetModuleAttribute(List<string> lines, string key, JToken value)
        {
            var text = $"Attribute {key} = {Format(key, value)}";
            var existing = lines.FindIndex(l => { var m = AttributeLine.Match(l); return m.Success && !m.Groups["member"].Success && m.Groups["name"].Value == key; });
            if (existing >= 0)
            {
                lines[existing] = text;
                return;
            }

            var last = lines.FindLastIndex(l => { var m = AttributeLine.Match(l); return m.Success && !m.Groups["member"].Success; });
            if (last < 0)
            {
                throw ApiException.BadRequest("The module has no attribute section to add to.");
            }

            lines.Insert(last + 1, text);
        }

        private static void SetMemberAttribute(List<string> lines, string member, string key, JToken value)
        {
            var text = $"Attribute {member}.{key} = {Format(key, value)}";
            var existing = lines.FindIndex(l => { var m = AttributeLine.Match(l); return m.Success && m.Groups["member"].Success
                                                  && string.Equals(m.Groups["member"].Value, member, StringComparison.OrdinalIgnoreCase) && m.Groups["name"].Value == key; });
            if (existing >= 0)
            {
                lines[existing] = text;
                return;
            }

            var declaration = lines.FindIndex(l => { var m = DeclarationLine.Match(l); return m.Success && string.Equals(m.Groups["name"].Value, member, StringComparison.OrdinalIgnoreCase); });
            if (declaration < 0)
            {
                throw ApiException.BadRequest($"The module has no member '{member}' to put an attribute on.");
            }

            //The attributes go right after the declaration, which may run over several lines.
            var at = declaration;
            while (lines[at].TrimEnd().EndsWith(" _", StringComparison.Ordinal) && at + 1 < lines.Count)
            {
                at++;
            }

            at++;
            while (at < lines.Count && AttributeLine.Match(lines[at]).Success)
            {
                at++;
            }

            lines.Insert(at, text);
        }

        private static string Key(string name)
            => ValidKey.IsMatch(name) ? name : throw ApiException.BadRequest($"'{name}' is not an attribute name (they start with VB_).");

        private static IEnumerable<string> Lines(string text)
            => (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');

        private static JToken ParseValue(string raw)
        {
            if (raw.Length >= 2 && raw[0] == '"' && raw[raw.Length - 1] == '"')
            {
                return raw.Substring(1, raw.Length - 2).Replace("\"\"", "\"");
            }

            if (raw.Equals("True", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (raw.Equals("False", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number) ? (JToken)number : raw;
        }

        private static string Format(string key, JToken value)
        {
            switch (value.Type)
            {
                case JTokenType.Boolean:
                    return value.Value<bool>() ? "True" : "False";
                case JTokenType.Integer:
                    return value.Value<long>().ToString(CultureInfo.InvariantCulture);
                case JTokenType.String:
                    return "\"" + value.Value<string>().Replace("\"", "\"\"") + "\"";
                default:
                    throw ApiException.BadRequest($"The value of {key} must be a string, a number or true/false.");
            }
        }
    }
}
