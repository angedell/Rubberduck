using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace Rubberduck.Api
{
    /// <summary>Which Type to build a class for, and how.</summary>
    public class ExtractUdtClassRequest
    {
        public const string Plain = "plain";
        public const string Fluent = "fluent";

        public string Module { get; set; }
        /// <summary>The Type's name.</summary>
        public string Name { get; set; }
        public int? Line { get; set; }
        public int? Column { get; set; }

        /// <summary>Name of the new class. Rubberduck proposes one (the Type's name without its T or UDT prefix) if none is given.</summary>
        public string ClassName { get; set; }

        /// <summary><see cref="Plain"/> or <see cref="Fluent"/>. Null: plain.</summary>
        public string Style { get; set; }

        /// <summary>Names of the members that have to be given when the object is made: they become the arguments of Create. Fluent style only.</summary>
        public List<string> Required { get; set; } = new List<string>();

        public static ExtractUdtClassRequest From(JObject body) => new ExtractUdtClassRequest
        {
            Module = RefactorArgs.Text(body, "module"),
            Name = RefactorArgs.Text(body, "name"),
            Line = RefactorArgs.Number(body, "line"),
            Column = RefactorArgs.Number(body, "column"),
            ClassName = RefactorArgs.Text(body, "className"),
            Style = RefactorArgs.Text(body, "style"),
            Required = Names(body, "required")
        };

        private static List<string> Names(JObject body, string name)
        {
            var token = body[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return new List<string>();
            }

            if (!(token is JArray list) || list.Any(item => item.Type != JTokenType.String))
            {
                throw ApiException.BadRequest($"'{name}' must be a list of member names.");
            }

            return list.Select(item => item.Value<string>()).ToList();
        }
    }

    public class ExtractedMember
    {
        public string Name { get; set; }
        public string Type { get; set; }

        /// <summary>readWrite, readOnly (arrays have a getter only) or left out (a Type inside the Type is not wrapped).</summary>
        public string Access { get; set; }

        /// <summary>An argument of Create.</summary>
        public bool Required { get; set; }
    }

    public class ExtractUdtClassPreview
    {
        public string Project { get; set; }
        public string Module { get; set; }
        public string Type { get; set; }
        public string ClassName { get; set; }
        public string Style { get; set; }
        public List<ExtractedMember> Members { get; set; } = new List<ExtractedMember>();

        /// <summary>The code of the class as it would be added.</summary>
        public string Code { get; set; }

        public List<string> Problems { get; set; } = new List<string>();
        public bool Valid => Problems.Count == 0;
    }

    public class ExtractUdtClassResult
    {
        public string Project { get; set; }
        public string Type { get; set; }
        public string ClassName { get; set; }
        public string Style { get; set; }
        public string Note { get; set; }
    }

    /// <summary>Where classes come from: Rubberduck's Extract Class from Type refactoring, answered from the request.</summary>
    public interface IExtractUdtClassSource
    {
        /// <summary>Works out what the class would be. Changes nothing.</summary>
        ExtractUdtClassPreview Preview(string project, ExtractUdtClassRequest request);

        /// <summary>Adds the class. The caller has checked the preview.</summary>
        ExtractUdtClassResult Apply(string project, ExtractUdtClassRequest request);
    }

    /// <summary>
    /// Add a class module for a Type, with properties for its members and, on request, a predeclared instance and chainable WithX methods.
    /// The Type is left as it is. This checks the request, shows the class first, and asks the user's approval for the new module.
    /// </summary>
    public class RefactorExtractUdtClassApi
    {
        private readonly IExtractUdtClassSource _source;
        private readonly ApiWriteApproval _approval;

        public RefactorExtractUdtClassApi(IExtractUdtClassSource source, ApiWriteApproval approval)
        {
            _source = source;
            _approval = approval;
        }

        public ExtractUdtClassPreview Preview(string project, ExtractUdtClassRequest request)
        {
            Check(request);
            return _source.Preview(project, request);
        }

        public ExtractUdtClassResult Extract(string client, string project, ExtractUdtClassRequest request)
        {
            Check(request);
            //Looks first, so that a request that cannot work is refused before the user is asked anything.
            Verify(_source.Preview(project, request));
            _approval.Require(client);
            return _source.Apply(project, request);
        }

        private static void Check(ExtractUdtClassRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Module))
            {
                throw ApiException.BadRequest("'module' is required.");
            }

            if (request.Style != null
                && !request.Style.Equals(ExtractUdtClassRequest.Plain, StringComparison.OrdinalIgnoreCase)
                && !request.Style.Equals(ExtractUdtClassRequest.Fluent, StringComparison.OrdinalIgnoreCase))
            {
                throw ApiException.BadRequest($"'style' must be '{ExtractUdtClassRequest.Plain}' or '{ExtractUdtClassRequest.Fluent}'.");
            }
        }

        /// <summary>What stops a class from being added. Public because the real source checks again, on the editor's thread, in the instant before it acts.</summary>
        public static void Verify(ExtractUdtClassPreview preview)
        {
            if (!preview.Valid)
            {
                throw RefactorArgs.Refused($"No class can be made for '{preview.Type}' like that: {string.Join(" ", preview.Problems)}");
            }
        }

        /// <summary>The routes under /projects/{p}/refactor/extract-udt-class, which the router hands over. Null if the path is not one of them.</summary>
        public ApiResponse Route(ApiRequest request, string method, string[] segments)
        {
            //segments: projects, {p}, refactor, extract-udt-class[, preview]
            if (segments.Length == 4)
            {
                RefactorArgs.AllowPost(method);
                return ApiResponse.Json(Extract(request.Client, segments[1], ExtractUdtClassRequest.From(RefactorArgs.Body(request.Body))));
            }

            if (segments.Length == 5 && segments[4] == "preview")
            {
                RefactorArgs.AllowPost(method);
                return ApiResponse.Json(Preview(segments[1], ExtractUdtClassRequest.From(RefactorArgs.Body(request.Body))));
            }

            return null;
        }

        private IList<RefactorToolSpec> _tools;

        public IEnumerable<RefactorToolSpec> Tools() => _tools ?? (_tools = new List<RefactorToolSpec>
        {
            new RefactorToolSpec
            {
                Name = "preview_extract_udt_class",
                Description = "Shows the class that extract_udt_class would add for a Type (a user defined type, public or private) and changes nothing: its name, which members get which properties, whether the request is valid (class name taken, member names that clash with Self, Create or a WithX method, ...) and the code of the class. Always call this before extract_udt_class. Name the Type by 'module' and 'name'; add 'line' (and 'column') if the name is not unique in the module.",
                Properties = ExtractProperties(),
                Required = new[] { "project", "module" },
                Run = (args, client) => Preview(RefactorArgs.Required(args, "project"), ExtractUdtClassRequest.From(args))
            },
            new RefactorToolSpec
            {
                Name = "extract_udt_class",
                Description = "Adds a class module for a Type, through Rubberduck's Extract Class from Type refactoring: a Property Get and Let/Set per member and a Self property, and with style 'fluent' also a predeclared instance with a Create function and a chainable WithName-style method per member; members named in 'required' become the arguments of Create. The Type itself is not changed, nor are its uses. Call preview_extract_udt_class first. Needs the user's approval like a write. Fails with 409 if the editor is not in design mode.",
                Properties = ExtractProperties(),
                Required = new[] { "project", "module" },
                Run = (args, client) => Extract(client, RefactorArgs.Required(args, "project"), ExtractUdtClassRequest.From(args))
            }
        });

        private static JObject ExtractProperties() => new JObject
        {
            ["project"] = RefactorArgs.String("Project id or name."),
            ["module"] = RefactorArgs.String("The module that declares the Type."),
            ["name"] = RefactorArgs.String("The Type's name. Omit it when giving 'line' and 'column'."),
            ["line"] = RefactorArgs.Integer("Line where the Type is declared (1-based), if 'name' is not unique in the module."),
            ["column"] = RefactorArgs.Integer("A column on that line within the name."),
            ["className"] = RefactorArgs.String("Name of the new class. Default: the Type's name without a T or UDT prefix (Rubberduck avoids clashes)."),
            ["style"] = RefactorArgs.String("'plain' (properties and Self) or 'fluent' (also a predeclared instance, Create and WithX methods). Default: plain."),
            ["required"] = new JObject
            {
                ["type"] = "array",
                ["items"] = new JObject { ["type"] = "string" },
                ["description"] = "Names of the members that have to be given when the object is made; they become the arguments of Create, in the order of the Type. Fluent style only; arrays and members that are a Type cannot be required."
            }
        };
    }
}
