using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace Rubberduck.Api
{
    /// <summary>What to rename: a declaration of a module (see <see cref="RefactorTargets.Choose"/>), and the new name.</summary>
    public class RenameRequest
    {
        public string Module { get; set; }
        public string Name { get; set; }
        public string Kind { get; set; }
        public int? Line { get; set; }
        public int? Column { get; set; }
        public string NewName { get; set; }

        /// <summary>Rubberduck renames the interface member, the control or the event instead when the target is its implementation or handler. The caller has to say yes to that.</summary>
        public bool AcceptResolvedTarget { get; set; }

        /// <summary>The new name clashes with another declaration. The dialog lets the user go on anyway; so can the caller.</summary>
        public bool AllowConflicts { get; set; }

        public static RenameRequest From(JObject body) => new RenameRequest
        {
            Module = RefactorArgs.Text(body, "module"),
            Name = RefactorArgs.Text(body, "name"),
            Kind = RefactorArgs.Text(body, "kind"),
            Line = RefactorArgs.Number(body, "line"),
            Column = RefactorArgs.Number(body, "column"),
            NewName = RefactorArgs.Text(body, "newName"),
            AcceptResolvedTarget = RefactorArgs.Flag(body, "acceptResolvedTarget"),
            AllowConflicts = RefactorArgs.Flag(body, "allowConflicts")
        };
    }

    public class RenameTarget
    {
        public string Module { get; set; }
        public string Name { get; set; }
        public string Kind { get; set; }
        public int Line { get; set; }
    }

    /// <summary>What a rename would do, as Rubberduck's own refactoring works it out.</summary>
    public class RenamePreview
    {
        public string Project { get; set; }

        /// <summary>The declaration the request named.</summary>
        public RenameTarget Target { get; set; }

        /// <summary>What is renamed instead, if that is not the target (the interface member an implementation belongs to, the control behind its event handler, a user event); otherwise null.</summary>
        public RenameTarget ResolvedTarget { get; set; }

        /// <summary>Why: "interface member", "control event handler" or "user event".</summary>
        public string ResolvedBecause { get; set; }

        public string NewName { get; set; }

        /// <summary>False if the new name cannot be used (invalid, or only the case differs). Conflicts do not make it invalid: see <see cref="Conflicts"/>.</summary>
        public bool Valid { get; set; }

        public List<string> Problems { get; set; } = new List<string>();

        /// <summary>Declarations the new name would clash with.</summary>
        public List<string> Conflicts { get; set; } = new List<string>();

        /// <summary>How many places use the renamed declaration, and in which modules. These are rewritten with it.</summary>
        public int ReferenceCount { get; set; }

        public List<string> ModulesAffected { get; set; } = new List<string>();
    }

    public class RenameResult
    {
        public string Project { get; set; }
        public string OldName { get; set; }
        public string NewName { get; set; }
        public RenameTarget Renamed { get; set; }
        public int ReferenceCount { get; set; }
        public List<string> ModulesAffected { get; set; } = new List<string>();
        public string Note { get; set; }
    }

    /// <summary>Where renames come from: Rubberduck's rename refactoring, run without its dialog.</summary>
    public interface IRenameSource
    {
        /// <summary>Works out what the rename would do. Changes nothing.</summary>
        RenamePreview Preview(string project, RenameRequest request);

        /// <summary>Renames, through the refactoring, so that every reference is rewritten. The caller has checked the preview.</summary>
        RenameResult Apply(string project, RenameRequest request);
    }

    /// <summary>
    /// Rename a procedure, variable, type, module... and everything that uses it. The refactoring is Rubberduck's own; this checks the request, shows what it would do first,
    /// and asks the user's approval for the change.
    /// </summary>
    public class RefactorRenameApi
    {
        private readonly IRenameSource _source;
        private readonly ApiWriteApproval _approval;

        public RefactorRenameApi(IRenameSource source, ApiWriteApproval approval)
        {
            _source = source;
            _approval = approval;
        }

        public RenamePreview Preview(string project, RenameRequest request)
        {
            Check(request);
            return _source.Preview(project, request);
        }

        public RenameResult Rename(string client, string project, RenameRequest request)
        {
            Check(request);
            //Looks first, so that a name that cannot work is refused before the user is asked anything.
            Verify(_source.Preview(project, request), request);
            _approval.Require(client);
            return _source.Apply(project, request);
        }

        private static void Check(RenameRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Module))
            {
                throw ApiException.BadRequest("'module' is required.");
            }

            if (string.IsNullOrWhiteSpace(request.NewName))
            {
                throw ApiException.BadRequest("'newName' is required.");
            }
        }

        /// <summary>What stops a rename from going ahead. Public because the real source checks again, on the editor's thread, in the instant before it renames.</summary>
        public static void Verify(RenamePreview preview, RenameRequest request)
        {
            if (!preview.Valid)
            {
                throw RefactorArgs.Refused($"'{request.NewName}' cannot be used: {string.Join(" ", preview.Problems)}");
            }

            if (preview.ResolvedTarget != null && !request.AcceptResolvedTarget)
            {
                throw ApiException.Conflict($"Rubberduck would rename {preview.ResolvedTarget.Kind} '{preview.ResolvedTarget.Name}' in '{preview.ResolvedTarget.Module}' instead "
                                            + $"({preview.ResolvedBecause}), not '{preview.Target.Name}'. Send 'acceptResolvedTarget': true to go ahead.");
            }

            if (preview.Conflicts.Count > 0 && !request.AllowConflicts)
            {
                throw ApiException.Conflict($"'{request.NewName}' clashes with {string.Join(", ", preview.Conflicts.Take(5))}. Send 'allowConflicts': true to rename anyway.");
            }
        }

        /// <summary>The routes under /projects/{p}/refactor/rename, which the router hands over. Null if the path is not one of them.</summary>
        public ApiResponse Route(ApiRequest request, string method, string[] segments)
        {
            //segments: projects, {p}, refactor, rename[, preview]
            if (segments.Length == 4)
            {
                Allow(method);
                var rename = RenameRequest.From(RefactorArgs.Body(request.Body));
                return ApiResponse.Json(Rename(request.Client, segments[1], rename));
            }

            if (segments.Length == 5 && segments[4] == "preview")
            {
                Allow(method);
                return ApiResponse.Json(Preview(segments[1], RenameRequest.From(RefactorArgs.Body(request.Body))));
            }

            return null;
        }

        private static void Allow(string method)
        {
            if (method != "POST")
            {
                throw new ApiException(405, "method_not_allowed", "Allowed: POST.");
            }
        }

        private IList<RefactorToolSpec> _tools;

        public IEnumerable<RefactorToolSpec> Tools() => _tools ?? (_tools = new List<RefactorToolSpec>
        {
            new RefactorToolSpec
            {
                Name = "preview_rename",
                Description = "Shows what renaming a declaration would do, and changes nothing: the declaration Rubberduck would really rename (renaming an interface implementation, a control's event handler or a user event renames the original instead), whether the new name is valid, what it would clash with, and how many references in which modules would be rewritten. Always call this before rename_symbol. Name the declaration by 'module' and 'name', adding 'line' (and 'column') where it is declared if the name is not unique in the module; a module is renamed by giving its own name.",
                Properties = RenameProperties(),
                Required = new[] { "project", "module", "newName" },
                Run = (args, client) => Preview(Arg(args, "project"), RenameRequest.From(args))
            },
            new RefactorToolSpec
            {
                Name = "rename_symbol",
                Description = "Renames a procedure, function, variable, constant, parameter, type, enum, event, module or form through Rubberduck's rename refactoring, so every reference to it, in every module, is rewritten. Call preview_rename first. Needs the user's approval like a write. It edits several modules at once, so re-read any module you hold afterwards (their hashes change). Fails with 409 if the editor is not in design mode.",
                Properties = RenameProperties(),
                Required = new[] { "project", "module", "newName" },
                Run = (args, client) => Rename(client, Arg(args, "project"), RenameRequest.From(args))
            }
        });

        private static JObject RenameProperties()
        {
            var properties = new JObject
            {
                ["project"] = RefactorArgs.String("Project id or name."),
                ["module"] = RefactorArgs.String("The module that declares it."),
                ["name"] = RefactorArgs.String("The declaration's current name. Omit it when giving 'line' and 'column'."),
                ["kind"] = RefactorArgs.String("Narrows 'name' by kind, e.g. Function, Variable, Parameter, PropertyGet."),
                ["line"] = RefactorArgs.Integer("Line where the name is declared (1-based), if 'name' is not unique in the module. Parameters and local variables always need it."),
                ["column"] = RefactorArgs.Integer("A column on that line within the name."),
                ["newName"] = RefactorArgs.String("The new name."),
                ["acceptResolvedTarget"] = RefactorArgs.Boolean("Go ahead when Rubberduck would rename another declaration instead (the preview says which)."),
                ["allowConflicts"] = RefactorArgs.Boolean("Go ahead although the new name clashes with another declaration.")
            };
            return properties;
        }

        private static string Arg(JObject args, string name)
        {
            var token = args[name];
            return token != null && token.Type == JTokenType.String
                ? token.Value<string>()
                : throw ApiException.BadRequest($"'{name}' is required and must be a string.");
        }
    }
}
