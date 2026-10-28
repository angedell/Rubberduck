using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace Rubberduck.Api
{
    /// <summary>Which module-level field to turn into a property, and how.</summary>
    public class EncapsulateFieldRequest
    {
        public const string BackingField = "backingField";
        public const string UdtMember = "udtMember";

        public string Module { get; set; }
        public string Name { get; set; }
        public int? Line { get; set; }
        public int? Column { get; set; }

        /// <summary>Name of the property; Rubberduck proposes one (the field's name, capitalised) if none is given.</summary>
        public string PropertyName { get; set; }

        /// <summary>True: only a Property Get. False: Get and Let/Set. Null: Rubberduck's choice (read/write, unless the field is an array).</summary>
        public bool? ReadOnly { get; set; }

        /// <summary><see cref="BackingField"/> (the field stays, renamed, behind the property) or <see cref="UdtMember"/> (the field becomes a member of the module's private Type). Null: Rubberduck's choice.</summary>
        public string Strategy { get; set; }

        public static EncapsulateFieldRequest From(JObject body)
        {
            var token = body["readOnly"];
            if (token != null && token.Type != JTokenType.Null && token.Type != JTokenType.Boolean)
            {
                throw ApiException.BadRequest("'readOnly' must be true or false.");
            }

            return new EncapsulateFieldRequest
            {
                Module = RefactorArgs.Text(body, "module"),
                Name = RefactorArgs.Text(body, "name"),
                Line = RefactorArgs.Number(body, "line"),
                Column = RefactorArgs.Number(body, "column"),
                PropertyName = RefactorArgs.Text(body, "propertyName"),
                ReadOnly = token == null || token.Type == JTokenType.Null ? (bool?)null : token.Value<bool>(),
                Strategy = RefactorArgs.Text(body, "strategy")
            };
        }
    }

    public class EncapsulatedProperty
    {
        /// <summary>The field being encapsulated (for a field of a private Type, one entry per member of the Type).</summary>
        public string Field { get; set; }

        public string PropertyName { get; set; }

        /// <summary>The field that holds the value behind the property.</summary>
        public string BackingField { get; set; }

        public bool ReadOnly { get; set; }

        /// <summary>False for arrays, which only get a Property Get.</summary>
        public bool CanBeReadWrite { get; set; }

        /// <summary>Other modules assign the field, so the property needs a Let/Set.</summary>
        public bool AssignedExternally { get; set; }
    }

    /// <summary>What encapsulating a field would do, as Rubberduck's refactoring works it out. <see cref="Code"/> is the preview the refactoring's dialog shows.</summary>
    public class EncapsulateFieldPreview
    {
        public string Project { get; set; }
        public string Module { get; set; }
        public string Field { get; set; }
        public int Line { get; set; }
        public string Strategy { get; set; }
        public List<EncapsulatedProperty> Properties { get; set; } = new List<EncapsulatedProperty>();
        public bool Valid { get; set; }
        public List<string> Problems { get; set; } = new List<string>();

        /// <summary>The code that would be added to the module, as text.</summary>
        public string Code { get; set; }
    }

    public class EncapsulateFieldResult
    {
        public string Project { get; set; }
        public string Module { get; set; }
        public string Field { get; set; }
        public string Strategy { get; set; }
        public List<EncapsulatedProperty> Properties { get; set; } = new List<EncapsulatedProperty>();
        public string Note { get; set; }
    }

    /// <summary>Where encapsulations come from: Rubberduck's Encapsulate Field refactoring, run without its dialog.</summary>
    public interface IEncapsulateFieldSource
    {
        /// <summary>Works out what the refactoring would do. Changes nothing.</summary>
        EncapsulateFieldPreview Preview(string project, EncapsulateFieldRequest request);

        /// <summary>Encapsulates the field through the refactoring, so that uses of it are rewritten to go through the property. The caller has checked the preview.</summary>
        EncapsulateFieldResult Apply(string project, EncapsulateFieldRequest request);
    }

    /// <summary>
    /// Turn a module-level field into a property. The refactoring is Rubberduck's own; this checks the request, shows what it would do first,
    /// and asks the user's approval for the change.
    /// </summary>
    public class RefactorEncapsulateFieldApi
    {
        private readonly IEncapsulateFieldSource _source;
        private readonly ApiWriteApproval _approval;

        public RefactorEncapsulateFieldApi(IEncapsulateFieldSource source, ApiWriteApproval approval)
        {
            _source = source;
            _approval = approval;
        }

        public EncapsulateFieldPreview Preview(string project, EncapsulateFieldRequest request)
        {
            Check(request);
            return _source.Preview(project, request);
        }

        public EncapsulateFieldResult Encapsulate(string client, string project, EncapsulateFieldRequest request)
        {
            Check(request);
            //Looks first, so that a request that cannot work is refused before the user is asked anything.
            Verify(_source.Preview(project, request));
            _approval.Require(client);
            return _source.Apply(project, request);
        }

        private static void Check(EncapsulateFieldRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Module))
            {
                throw ApiException.BadRequest("'module' is required.");
            }

            if (request.Strategy != null
                && !request.Strategy.Equals(EncapsulateFieldRequest.BackingField, StringComparison.OrdinalIgnoreCase)
                && !request.Strategy.Equals(EncapsulateFieldRequest.UdtMember, StringComparison.OrdinalIgnoreCase))
            {
                throw ApiException.BadRequest($"'strategy' must be '{EncapsulateFieldRequest.BackingField}' or '{EncapsulateFieldRequest.UdtMember}'.");
            }
        }

        /// <summary>What stops an encapsulation from going ahead. Public because the real source checks again, on the editor's thread, in the instant before it acts.</summary>
        public static void Verify(EncapsulateFieldPreview preview)
        {
            if (!preview.Valid)
            {
                throw RefactorArgs.Refused($"'{preview.Field}' cannot be encapsulated like that: {string.Join(" ", preview.Problems)}");
            }
        }

        /// <summary>The routes under /projects/{p}/refactor/encapsulate-field, which the router hands over. Null if the path is not one of them.</summary>
        public ApiResponse Route(ApiRequest request, string method, string[] segments)
        {
            //segments: projects, {p}, refactor, encapsulate-field[, preview]
            if (segments.Length == 4)
            {
                RefactorArgs.AllowPost(method);
                return ApiResponse.Json(Encapsulate(request.Client, segments[1], EncapsulateFieldRequest.From(RefactorArgs.Body(request.Body))));
            }

            if (segments.Length == 5 && segments[4] == "preview")
            {
                RefactorArgs.AllowPost(method);
                return ApiResponse.Json(Preview(segments[1], EncapsulateFieldRequest.From(RefactorArgs.Body(request.Body))));
            }

            return null;
        }

        private IList<RefactorToolSpec> _tools;

        public IEnumerable<RefactorToolSpec> Tools() => _tools ?? (_tools = new List<RefactorToolSpec>
        {
            new RefactorToolSpec
            {
                Name = "preview_encapsulate_field",
                Description = "Shows what encapsulating a module-level field would do, and changes nothing: the property (name, read-only or not) and backing field Rubberduck would create, whether that is valid (name clashes, arrays, ...), and the code it would add. Always call this before encapsulate_field. Name the field by 'module' and 'name'; add 'line' (and 'column') if the name is not unique in the module.",
                Properties = EncapsulateProperties(),
                Required = new[] { "project", "module" },
                Run = (args, client) => Preview(RefactorArgs.Required(args, "project"), EncapsulateFieldRequest.From(args))
            },
            new RefactorToolSpec
            {
                Name = "encapsulate_field",
                Description = "Turns a module-level field into a property through Rubberduck's Encapsulate Field refactoring: it adds the Property Get (and Let/Set unless read-only), keeps the value in a backing field, and rewrites the uses of the field to go through the property. Call preview_encapsulate_field first. Needs the user's approval like a write. Re-read the module afterwards (its hash changes). Fails with 409 if the editor is not in design mode.",
                Properties = EncapsulateProperties(),
                Required = new[] { "project", "module" },
                Run = (args, client) => Encapsulate(client, RefactorArgs.Required(args, "project"), EncapsulateFieldRequest.From(args))
            }
        });

        private static JObject EncapsulateProperties() => new JObject
        {
            ["project"] = RefactorArgs.String("Project id or name."),
            ["module"] = RefactorArgs.String("The module that declares the field."),
            ["name"] = RefactorArgs.String("The field's name. Omit it when giving 'line' and 'column'."),
            ["line"] = RefactorArgs.Integer("Line where the field is declared (1-based), if 'name' is not unique in the module."),
            ["column"] = RefactorArgs.Integer("A column on that line within the name."),
            ["propertyName"] = RefactorArgs.String("Name of the property. Default: the field's name, capitalised (Rubberduck avoids clashes)."),
            ["readOnly"] = RefactorArgs.Boolean("True for a Property Get only. Default: read/write, except for arrays."),
            ["strategy"] = RefactorArgs.String("'backingField' (the field is renamed and stays behind the property) or 'udtMember' (it becomes a member of the module's private Type). Default: Rubberduck's choice.")
        };
    }
}
