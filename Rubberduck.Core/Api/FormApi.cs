using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace Rubberduck.Api
{
    /// <summary>
    /// The controls of a UserForm: list, add, change, delete. A form's layout is not code, so it has no hash; every change is backed up like an edit and needs the user's approval.
    /// Only the standard MSForms controls can be added, by short name (TextBox) or ProgID (Forms.TextBox.1); other ActiveX controls are not loaded on an agent's say-so.
    /// </summary>
    public class FormApi
    {
        private static readonly Regex ValidName = new Regex(@"^[A-Za-z][A-Za-z0-9_]{0,30}$", RegexOptions.Compiled);
        private static readonly Regex ValidProperty = new Regex(@"^[A-Za-z][A-Za-z0-9]{0,40}$", RegexOptions.Compiled);

        //A control is a member of the form, so a name the language reserves makes the form's code fail to compile; the designer does not object.
        private static readonly HashSet<string> ReservedWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "AddressOf", "Alias", "And", "As", "Boolean", "ByRef", "Byte", "ByVal", "Call", "Case", "Close", "Const", "Currency", "Date", "Declare",
            "DefBool", "DefByte", "DefCur", "DefDate", "DefDbl", "DefInt", "DefLng", "DefObj", "DefSng", "DefStr", "DefVar", "Dim", "Do", "Double",
            "Each", "Else", "ElseIf", "Empty", "End", "Enum", "Eqv", "Erase", "Error", "Event", "Exit", "False", "For", "Friend", "Function", "Get",
            "Global", "GoSub", "GoTo", "If", "Imp", "Implements", "In", "Input", "Integer", "Is", "Let", "Lib", "Like", "Line", "Lock", "Long", "Loop",
            "LSet", "Me", "Mod", "New", "Next", "Not", "Nothing", "Null", "Object", "On", "Open", "Option", "Optional", "Or", "ParamArray", "Preserve",
            "Print", "Private", "Property", "Public", "Put", "RaiseEvent", "ReDim", "Rem", "Resume", "Return", "RSet", "Seek", "Select", "Set", "Single",
            "Static", "Stop", "String", "Sub", "Then", "To", "True", "Type", "TypeOf", "Unlock", "Until", "Variant", "Wend", "While", "Width", "With",
            "WithEvents", "Write", "Xor"
        };

        /// <summary>Short name to the ProgID the designer adds.</summary>
        public static readonly IReadOnlyDictionary<string, string> ControlTypes = new[]
        {
            "CheckBox", "ComboBox", "CommandButton", "Frame", "Image", "Label", "ListBox", "MultiPage",
            "OptionButton", "ScrollBar", "SpinButton", "TabStrip", "TextBox", "ToggleButton"
        }.ToDictionary(type => type, type => "Forms." + type + ".1", StringComparer.OrdinalIgnoreCase);

        //Properties the routes and tools take as fields of their own, next to the free-form 'properties'.
        private static readonly string[] Direct = { "left", "top", "width", "height", "caption", "text", "tabIndex" };

        public const string Summary = "Controls of a UserForm. Left, Top, Width and Height are in points. 'properties' takes any property the control has, by its VBA name (Caption, Text, Value, RowSource, Style, TabIndex, Visible...). Set a form's code with write_module; the controls are named as in the code ('cmbStatus_Change').";

        private readonly IFormControlStore _store;
        private readonly ApiWriteApproval _approval;

        public FormApi(IFormControlStore store, ApiWriteApproval approval)
        {
            _store = store;
            _approval = approval;
        }

        public JObject List(string project, string module) => _store.ListControls(project, module);

        public JObject Add(string client, string project, string module, string type, string name, JObject properties)
        {
            if (string.IsNullOrWhiteSpace(type) || !ControlTypes.TryGetValue(ShortType(type), out var progId))
            {
                throw ApiException.BadRequest($"type must be one of: {string.Join(", ", ControlTypes.Keys)}.");
            }

            RequireName(name);
            RejectReserved(name);
            properties = Checked(properties);
            if (properties.Properties().Any(p => p.Name.Equals("Name", StringComparison.OrdinalIgnoreCase)))
            {
                throw ApiException.BadRequest("Send the name as 'name', not in 'properties'.");
            }

            _approval.Require(client);
            return _store.AddControl(project, module, progId, name, properties);
        }

        public JObject Set(string client, string project, string module, string name, JObject properties)
        {
            RequireName(name);
            properties = Checked(properties);
            if (!properties.HasValues)
            {
                throw ApiException.BadRequest("'properties' has nothing to set.");
            }

            var rename = properties.Properties().FirstOrDefault(p => p.Name.Equals("Name", StringComparison.OrdinalIgnoreCase));
            if (rename != null && (rename.Value.Type != JTokenType.String || !ValidName.IsMatch(rename.Value.Value<string>())))
            {
                throw ApiException.BadRequest("The new name must start with a letter and use only letters, digits and underscores (31 characters at most).");
            }

            if (rename != null)
            {
                RejectReserved(rename.Value.Value<string>());
            }

            _approval.Require(client);
            return _store.SetControl(project, module, name, properties);
        }

        public object Delete(string client, string project, string module, string name)
        {
            RequireName(name);
            _approval.Require(client);
            _store.DeleteControl(project, module, name);
            return new { deleted = name };
        }

        //Forms.TextBox.1 and textbox both mean TextBox.
        private static string ShortType(string type)
        {
            var name = type.Trim();
            if (name.StartsWith("Forms.", StringComparison.OrdinalIgnoreCase)) name = name.Substring("Forms.".Length);
            if (name.EndsWith(".1", StringComparison.Ordinal)) name = name.Substring(0, name.Length - 2);
            return name;
        }

        private static void RequireName(string name)
        {
            if (name == null || !ValidName.IsMatch(name))
            {
                throw ApiException.BadRequest("The control name must start with a letter and use only letters, digits and underscores (31 characters at most).");
            }
        }

        private static void RejectReserved(string name)
        {
            if (ReservedWords.Contains(name))
            {
                throw ApiException.BadRequest($"'{name}' is a reserved word in VBA, so code on the form could not refer to the control. Use a prefix, for example 'txt{name}'.");
            }
        }

        private static JObject Checked(JObject properties)
        {
            properties = properties ?? new JObject();
            foreach (var property in properties.Properties())
            {
                if (!ValidProperty.IsMatch(property.Name))
                {
                    throw ApiException.BadRequest($"'{property.Name}' is not a property name.");
                }
            }

            return properties;
        }

        /// <summary>The free-form 'properties' of a request, and the fields (left, top, width, height, caption, text, tabIndex) that sit next to it.</summary>
        public static JObject PropertiesOf(JObject body)
        {
            var token = body["properties"];
            if (token != null && token.Type != JTokenType.Null && !(token is JObject))
            {
                throw ApiException.BadRequest("'properties' must be an object.");
            }

            var result = new JObject();
            foreach (var field in Direct)
            {
                var value = body[field];
                if (value != null && value.Type != JTokenType.Null)
                {
                    result[char.ToUpperInvariant(field[0]) + field.Substring(1)] = value;
                }
            }

            //Written last: the free-form object wins over the fields next to it.
            if (token is JObject free)
            {
                foreach (var property in free.Properties())
                {
                    result[property.Name] = property.Value;
                }
            }

            return result;
        }

        #region Routes

        /// <summary>segments: projects, {p}, modules, {m}, controls[, {c}]</summary>
        public ApiResponse Route(ApiRequest request, string method, string[] segments)
        {
            var project = segments[1];
            var module = segments[3];
            if (segments.Length == 5)
            {
                switch (method)
                {
                    case "GET":
                        return ApiResponse.Json(List(project, module));
                    case "POST":
                        var body = RefactorArgs.Body(request.Body);
                        return new ApiResponse { Status = 201, Body = Add(request.Client, project, module, RefactorArgs.Text(body, "type"), RefactorArgs.Text(body, "name"), PropertiesOf(body)) };
                    default:
                        throw new ApiException(405, "method_not_allowed", "Allowed: GET, POST.");
                }
            }

            var name = segments[5];
            switch (method)
            {
                case "PUT":
                    return ApiResponse.Json(Set(request.Client, project, module, name, PropertiesOf(RefactorArgs.Body(request.Body))));
                case "DELETE":
                    return ApiResponse.Json(Delete(request.Client, project, module, name));
                default:
                    throw new ApiException(405, "method_not_allowed", "Allowed: PUT, DELETE.");
            }
        }

        #endregion

        #region Tools

        private static JObject Prop(string type, string description) => new JObject { ["type"] = type, ["description"] = description };

        private IList<RefactorToolSpec> _tools;

        public IEnumerable<RefactorToolSpec> Tools() => _tools ?? (_tools = new List<RefactorToolSpec>
        {
            new RefactorToolSpec
            {
                Name = "list_controls",
                Description = "Lists the controls of a form: name, type and the properties each has (position, size, caption, text, tab index...). " + Summary,
                Properties = new JObject { ["project"] = Prop("string", "Project id or name."), ["module"] = Prop("string", "Name of the form.") },
                Required = new[] { "project", "module" },
                Run = (args, client) => List(Arg(args, "project"), Arg(args, "module"))
            },
            new RefactorToolSpec
            {
                Name = "add_control",
                Description = "Adds a control to a form and sets its properties, e.g. a ComboBox 'cmbStatus' at left 12, top 12, width 200, height 20. All or nothing. Needs the user's approval; the form's previous version is backed up. Types: " + string.Join(", ", ControlTypes.Keys) + ". " + Summary,
                Properties = ControlProperties(withType: true),
                Required = new[] { "project", "module", "type", "name" },
                Run = (args, client) => Add(client, Arg(args, "project"), Arg(args, "module"), Arg(args, "type"), Arg(args, "name"), PropertiesOf(args))
            },
            new RefactorToolSpec
            {
                Name = "set_control_properties",
                Description = "Changes properties of a control: move or resize it, set its caption or text, TabIndex, RowSource... 'properties' {Name: \"newName\"} renames it (code that uses the old name is not updated). Needs the user's approval. " + Summary,
                Properties = ControlProperties(withType: false),
                Required = new[] { "project", "module", "name" },
                Run = (args, client) => Set(client, Arg(args, "project"), Arg(args, "module"), Arg(args, "name"), PropertiesOf(args))
            },
            new RefactorToolSpec
            {
                Name = "delete_control",
                Description = "Removes a control from a form (code that uses it stops compiling). Needs the user's approval; the form's previous version is backed up.",
                Properties = new JObject
                {
                    ["project"] = Prop("string", "Project id or name."),
                    ["module"] = Prop("string", "Name of the form."),
                    ["name"] = Prop("string", "Name of the control.")
                },
                Required = new[] { "project", "module", "name" },
                Run = (args, client) => Delete(client, Arg(args, "project"), Arg(args, "module"), Arg(args, "name"))
            }
        });

        private static JObject ControlProperties(bool withType)
        {
            var properties = new JObject
            {
                ["project"] = Prop("string", "Project id or name."),
                ["module"] = Prop("string", "Name of the form."),
                ["name"] = Prop("string", "Name of the control.")
            };

            if (withType)
            {
                properties["type"] = Prop("string", "One of: " + string.Join(", ", ControlTypes.Keys) + ".");
            }

            properties["left"] = Prop("number", "Distance from the left edge, in points.");
            properties["top"] = Prop("number", "Distance from the top edge, in points.");
            properties["width"] = Prop("number", "Width in points.");
            properties["height"] = Prop("number", "Height in points.");
            properties["caption"] = Prop("string", "Caption (Label, CommandButton, CheckBox, OptionButton, Frame...).");
            properties["text"] = Prop("string", "Text (TextBox, ComboBox).");
            properties["tabIndex"] = Prop("integer", "Position in the tab order.");
            properties["properties"] = Prop("object", "Any other property by its VBA name, e.g. { \"Style\": 2, \"RowSource\": \"Sheet1!A1:A5\", \"MultiLine\": true }.");
            return properties;
        }

        private static string Arg(JObject args, string name)
        {
            var value = args.Value<string>(name);
            return string.IsNullOrWhiteSpace(value) ? throw ApiException.BadRequest($"'{name}' is required.") : value;
        }

        #endregion
    }
}
