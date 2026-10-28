using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json.Linq;
using Rubberduck.CodeAnalysis.Inspections;

namespace Rubberduck.Api
{
    /// <summary>The id a diagnostic is known by: the same inspection result gets the same id for as long as the code it points at is unchanged.</summary>
    public static class QuickFixIds
    {
        public static string For(IInspectionResult result)
        {
            var where = result.QualifiedSelection.QualifiedName;
            var selection = result.QualifiedSelection.Selection;
            var text = string.Join("|", result.Inspection.Name, where.ProjectId, where.ComponentName,
                selection.StartLine, selection.StartColumn, selection.EndLine, selection.EndColumn, result.Description);
            using (var sha = SHA1.Create())
            {
                return string.Concat(sha.ComputeHash(Encoding.UTF8.GetBytes(text)).Take(6).Select(b => b.ToString("x2")));
            }
        }
    }

    public class QuickFixOption
    {
        /// <summary>The fix's name: what to pass as 'fix' to apply it.</summary>
        public string Name { get; set; }

        /// <summary>What it does to this result, as Rubberduck words it.</summary>
        public string Description { get; set; }

        /// <summary>It can also be applied to every result of the same inspection in the module (scope "module").</summary>
        public bool CanFixInModule { get; set; }
    }

    public class QuickFixOptions
    {
        public string Project { get; set; }
        public string Result { get; set; }
        public string Module { get; set; }
        public string Inspection { get; set; }
        public string Description { get; set; }
        public int StartLine { get; set; }
        public List<QuickFixOption> Fixes { get; set; } = new List<QuickFixOption>();
    }

    public class QuickFixOutcome
    {
        public string Project { get; set; }
        public string Result { get; set; }
        public string Fix { get; set; }

        /// <summary>"result" or "module".</summary>
        public string Scope { get; set; }

        /// <summary>How many results the fix was applied to.</summary>
        public int Fixed { get; set; }

        public string Note { get; set; }
    }

    /// <summary>Where quick fixes come from: Rubberduck's own fix classes, applied to the results of the latest diagnostics run.</summary>
    public interface IQuickFixSource
    {
        /// <summary>The fixes on offer for one diagnostic. 404 if the id is not one of the latest results; 409 if those are out of date.</summary>
        QuickFixOptions List(string project, string result);

        /// <summary>Applies a fix, on the editor's thread. <paramref name="scope"/> is "result" or "module".</summary>
        QuickFixOutcome Apply(string project, string result, string fix, string scope);
    }

    /// <summary>
    /// Applies Rubberduck's quick fixes to diagnostics. A diagnostic's <c>id</c> comes from the diagnostics routes. Listing the fixes is a read; applying one changes code and needs the user's approval like a write.
    /// </summary>
    public class QuickFixApi
    {
        public const string ScopeResult = "result";
        public const string ScopeModule = "module";

        private readonly IQuickFixSource _source;
        private readonly ApiWriteApproval _approval;

        public QuickFixApi(IQuickFixSource source, ApiWriteApproval approval)
        {
            _source = source;
            _approval = approval;
        }

        public QuickFixOptions List(string project, string result)
        {
            if (string.IsNullOrWhiteSpace(result))
            {
                throw ApiException.BadRequest("'result' is required: the id of a diagnostic.");
            }

            return _source.List(project, result);
        }

        public QuickFixOutcome Apply(string client, string project, string result, string fix, string scope)
        {
            if (string.IsNullOrWhiteSpace(result) || string.IsNullOrWhiteSpace(fix))
            {
                throw ApiException.BadRequest("'result' (the id of a diagnostic) and 'fix' (one of its fixes' names) are required.");
            }

            scope = string.IsNullOrWhiteSpace(scope) ? ScopeResult : scope.ToLowerInvariant();
            if (scope != ScopeResult && scope != ScopeModule)
            {
                throw ApiException.BadRequest($"'scope' must be \"{ScopeResult}\" or \"{ScopeModule}\".");
            }

            //Checked before the user is asked: a request that cannot work should not cost them a dialog.
            var options = _source.List(project, result);
            var chosen = options.Fixes.FirstOrDefault(f => string.Equals(f.Name, fix, StringComparison.OrdinalIgnoreCase))
                ?? throw RefactorArgs.Refused($"'{fix}' is not a fix for this diagnostic. Available: {string.Join(", ", options.Fixes.Select(f => f.Name))}.");
            if (scope == ScopeModule && !chosen.CanFixInModule)
            {
                throw RefactorArgs.Refused($"'{chosen.Name}' cannot be applied to a whole module; use scope \"{ScopeResult}\".");
            }

            _approval.Require(client);
            return _source.Apply(project, result, chosen.Name, scope);
        }

        public ApiResponse Route(ApiRequest request, string method, string[] segments)
        {
            //segments: projects, {p}, quickfix
            var project = segments[1];
            switch (method)
            {
                case "GET":
                    request.Query.TryGetValue("result", out var id);
                    return ApiResponse.Json(List(project, id));
                case "POST":
                    var body = RefactorArgs.Body(request.Body);
                    return ApiResponse.Json(Apply(request.Client, project, RefactorArgs.Text(body, "result"), RefactorArgs.Text(body, "fix"), RefactorArgs.Text(body, "scope")));
                default:
                    throw new ApiException(405, "method_not_allowed", "Allowed: GET, POST.");
            }
        }

        private IList<RefactorToolSpec> _tools;

        public IEnumerable<RefactorToolSpec> Tools() => _tools ?? (_tools = new List<RefactorToolSpec>
        {
            new RefactorToolSpec
            {
                Name = "list_quick_fixes",
                Description = "Lists the quick fixes Rubberduck offers for one diagnostic (its 'id' is in the get_diagnostics and check_work results). Changes nothing. Run get_diagnostics first: the id only exists for the latest inspection run and is refused (409) once the code has changed since.",
                Properties = new JObject { ["project"] = RefactorArgs.String("Project id or name."), ["result"] = RefactorArgs.String("The diagnostic's id.") },
                Required = new[] { "project", "result" },
                Run = (args, client) => List(RefactorArgs.Required(args, "project"), RefactorArgs.Required(args, "result"))
            },
            new RefactorToolSpec
            {
                Name = "apply_quick_fix",
                Description = "Applies one of Rubberduck's quick fixes to a diagnostic, through Rubberduck's own fix classes. Needs the user's approval like a write, and the editor in design mode (409 otherwise). With scope \"module\" it fixes every result of the same inspection in the same module, for fixes that allow it (list_quick_fixes says which). It can change several places or modules, so re-read any module you hold afterwards. Run get_diagnostics again afterwards to confirm the diagnostic is gone.",
                Properties = new JObject
                {
                    ["project"] = RefactorArgs.String("Project id or name."),
                    ["result"] = RefactorArgs.String("The diagnostic's id."),
                    ["fix"] = RefactorArgs.String("The fix's name, from list_quick_fixes."),
                    ["scope"] = RefactorArgs.String("\"result\" (default) or \"module\".")
                },
                Required = new[] { "project", "result", "fix" },
                Run = (args, client) => Apply(client, RefactorArgs.Required(args, "project"), RefactorArgs.Required(args, "result"), RefactorArgs.Required(args, "fix"), RefactorArgs.Text(args, "scope"))
            }
        });
    }
}
