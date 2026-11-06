using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace Rubberduck.Api
{
    /// <summary>What to split: the whole project, or just one module of it.</summary>
    public class SplitParametersRequest
    {
        /// <summary>Limits the split to this module. Null for every module of the project.</summary>
        public string Module { get; set; }

        public static SplitParametersRequest From(JObject body) => new SplitParametersRequest { Module = RefactorArgs.Text(body, "module") };
    }

    public class SplitMember
    {
        public string Module { get; set; }
        public string Member { get; set; }
        public int Parameters { get; set; }
    }

    /// <summary>The members whose parameter lists would change. A member that is laid out one parameter to a line already is not in it.</summary>
    public class SplitParametersPreview
    {
        public string Project { get; set; }
        public List<SplitMember> Members { get; set; } = new List<SplitMember>();
    }

    public class SplitParametersResult
    {
        public string Project { get; set; }
        public List<SplitMember> Split { get; set; } = new List<SplitMember>();
        public string Note { get; set; }
    }

    /// <summary>Where splits come from: Rubberduck's Split Parameters refactoring, run over a project in one go.</summary>
    public interface ISplitParametersSource
    {
        /// <summary>Works out what would change. Changes nothing.</summary>
        SplitParametersPreview Preview(string project, SplitParametersRequest request);

        /// <summary>Splits the parameter lists, as one rewrite.</summary>
        SplitParametersResult Apply(string project, SplitParametersRequest request);
    }

    /// <summary>
    /// Put every parameter of every member of a project on a line of its own, so that diffs of the signatures stay small. The refactoring is Rubberduck's own;
    /// this shows what it would do first and asks the user's approval for the change.
    /// </summary>
    public class RefactorSplitParametersApi
    {
        private readonly ISplitParametersSource _source;
        private readonly ApiWriteApproval _approval;

        public RefactorSplitParametersApi(ISplitParametersSource source, ApiWriteApproval approval)
        {
            _source = source;
            _approval = approval;
        }

        public SplitParametersPreview Preview(string project, SplitParametersRequest request) => _source.Preview(project, request);

        public SplitParametersResult Split(string client, string project, SplitParametersRequest request)
        {
            //Looks first, so that nothing to do is not something the user is asked to approve.
            var preview = _source.Preview(project, request);
            if (preview.Members.Count == 0)
            {
                return new SplitParametersResult
                {
                    Project = preview.Project,
                    Note = "Every parameter list is laid out one parameter to a line already; nothing was changed."
                };
            }

            _approval.Require(client);
            return _source.Apply(project, request);
        }

        /// <summary>The routes under /projects/{p}/refactor/split-parameters, which the router hands over. Null if the path is not one of them.</summary>
        public ApiResponse Route(ApiRequest request, string method, string[] segments)
        {
            //segments: projects, {p}, refactor, split-parameters[, preview]
            if (segments.Length == 4)
            {
                RefactorArgs.AllowPost(method);
                return ApiResponse.Json(Split(request.Client, segments[1], BodyOf(request)));
            }

            if (segments.Length == 5 && segments[4] == "preview")
            {
                RefactorArgs.AllowPost(method);
                return ApiResponse.Json(Preview(segments[1], BodyOf(request)));
            }

            return null;
        }

        //The body is optional: without one it means the whole project.
        private static SplitParametersRequest BodyOf(ApiRequest request)
            => string.IsNullOrWhiteSpace(request.Body) ? new SplitParametersRequest() : SplitParametersRequest.From(RefactorArgs.Body(request.Body));

        private IList<RefactorToolSpec> _tools;

        public IEnumerable<RefactorToolSpec> Tools() => _tools ?? (_tools = new List<RefactorToolSpec>
        {
            new RefactorToolSpec
            {
                Name = "preview_split_parameters",
                Description = "Shows which procedures, functions, properties and events of a project would have their parameter lists split to one parameter per line (leading commas, closing parenthesis on its own line), and changes nothing. Members that are laid out that way already are left out. Call this before split_parameters.",
                Properties = SplitProperties(),
                Required = new[] { "project" },
                Run = (args, client) => Preview(RefactorArgs.Required(args, "project"), SplitParametersRequest.From(args))
            },
            new RefactorToolSpec
            {
                Name = "split_parameters",
                Description = "Puts each parameter of every member of the project (or of one module) on a line of its own, in one operation, through Rubberduck's Split Parameters refactoring. This keeps diffs of signatures small: adding, removing or reordering a parameter changes one line. Call preview_split_parameters first. Needs the user's approval like a write. Re-read the modules afterwards (their hashes change). Fails with 409 if the editor is not in design mode.",
                Properties = SplitProperties(),
                Required = new[] { "project" },
                Run = (args, client) => Split(client, RefactorArgs.Required(args, "project"), SplitParametersRequest.From(args))
            }
        });

        private static JObject SplitProperties() => new JObject
        {
            ["project"] = RefactorArgs.String("Project id or name."),
            ["module"] = RefactorArgs.String("Only split this module's members. Without it, every module of the project.")
        };
    }
}
