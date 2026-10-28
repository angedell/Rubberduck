using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace Rubberduck.Api
{
    /// <summary>Which modules to move, and where: a folder in Rubberduck's sense, the '@Folder annotation, with dots between levels ("App.Core.Data").</summary>
    public class MoveToFolderRequest
    {
        public const int MaxModules = 50;

        /// <summary>The modules to move: <c>module</c> and/or <c>modules</c> of the request, without repeats.</summary>
        public List<string> Modules { get; set; } = new List<string>();

        public string Folder { get; set; }

        public static MoveToFolderRequest From(JObject body)
        {
            var request = new MoveToFolderRequest { Folder = RefactorArgs.Text(body, "folder") };
            var one = RefactorArgs.Text(body, "module");
            if (!string.IsNullOrWhiteSpace(one))
            {
                request.Modules.Add(one);
            }

            if (body["modules"] != null && body["modules"].Type != JTokenType.Null)
            {
                if (!(body["modules"] is JArray many) || many.Any(m => m.Type != JTokenType.String))
                {
                    throw ApiException.BadRequest("'modules' must be a list of module names.");
                }

                request.Modules.AddRange(many.Select(m => m.Value<string>()).Where(m => !string.IsNullOrWhiteSpace(m)));
            }

            request.Modules = request.Modules.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            return request;
        }
    }

    public class ModuleMove
    {
        public string Module { get; set; }

        /// <summary>The folder it is in now (the project's name if it has no '@Folder).</summary>
        public string FromFolder { get; set; }

        /// <summary>False if it is there already, so nothing would change.</summary>
        public bool Changes { get; set; }
    }

    /// <summary>What a move would do, as Rubberduck's refactoring works it out.</summary>
    public class MoveToFolderPreview
    {
        public string Project { get; set; }
        public string Folder { get; set; }
        public List<ModuleMove> Moves { get; set; } = new List<ModuleMove>();
        public bool Valid { get; set; }
        public List<string> Problems { get; set; } = new List<string>();
    }

    public class MoveToFolderResult
    {
        public string Project { get; set; }
        public string Folder { get; set; }
        public List<string> Moved { get; set; } = new List<string>();

        /// <summary>The modules that were in that folder already.</summary>
        public List<string> Unchanged { get; set; } = new List<string>();

        public string Note { get; set; }
    }

    /// <summary>Where moves come from: Rubberduck's Move to Folder refactoring, run without its dialog.</summary>
    public interface IMoveToFolderSource
    {
        /// <summary>Works out what the move would do. Changes nothing.</summary>
        MoveToFolderPreview Preview(string project, MoveToFolderRequest request);

        /// <summary>Moves the modules, by writing their '@Folder annotations through the refactoring. The caller has checked the preview.</summary>
        MoveToFolderResult Apply(string project, MoveToFolderRequest request);
    }

    /// <summary>
    /// Move modules to a folder (set their '@Folder annotation). The refactoring is Rubberduck's own; this checks the request, shows what it would do first,
    /// and asks the user's approval for the change.
    /// </summary>
    public class RefactorMoveToFolderApi
    {
        private readonly IMoveToFolderSource _source;
        private readonly ApiWriteApproval _approval;

        public RefactorMoveToFolderApi(IMoveToFolderSource source, ApiWriteApproval approval)
        {
            _source = source;
            _approval = approval;
        }

        public MoveToFolderPreview Preview(string project, MoveToFolderRequest request)
        {
            Check(request);
            return _source.Preview(project, request);
        }

        public MoveToFolderResult Move(string client, string project, MoveToFolderRequest request)
        {
            Check(request);
            //Looks first, so that a request that cannot work is refused before the user is asked anything.
            var preview = _source.Preview(project, request);
            Verify(preview);
            if (preview.Moves.All(m => !m.Changes))
            {
                //Nothing to write, so nothing to approve.
                return new MoveToFolderResult
                {
                    Project = preview.Project,
                    Folder = request.Folder,
                    Unchanged = preview.Moves.Select(m => m.Module).ToList(),
                    Note = "Every module is in that folder already; nothing was changed."
                };
            }

            _approval.Require(client);
            return _source.Apply(project, request);
        }

        private static void Check(MoveToFolderRequest request)
        {
            if (request.Modules.Count == 0)
            {
                throw ApiException.BadRequest("'module' (or a list in 'modules') is required.");
            }

            if (request.Modules.Count > MoveToFolderRequest.MaxModules)
            {
                throw ApiException.BadRequest($"At most {MoveToFolderRequest.MaxModules} modules in one move.");
            }

            if (string.IsNullOrEmpty(request.Folder))
            {
                throw ApiException.BadRequest("'folder' is required.");
            }
        }

        /// <summary>What the dialog checks as the user types a folder: not empty, no control characters, no empty level.</summary>
        public static List<string> FolderProblems(string folder)
        {
            var problems = new List<string>();
            if (string.IsNullOrEmpty(folder))
            {
                problems.Add(Rubberduck.Refactorings.RefactoringsUI.MoveFolders_EmptyFolderName);
                return problems;
            }

            if (folder.Any(char.IsControl))
            {
                problems.Add(Rubberduck.Refactorings.RefactoringsUI.MoveFolders_ControlCharacter);
            }

            if (folder.Split(Rubberduck.InternalApi.Extensions.FolderExtensions.FolderDelimiter).Any(string.IsNullOrEmpty))
            {
                problems.Add(Rubberduck.Refactorings.RefactoringsUI.MoveFolders_EmptySubfolderName);
            }

            return problems;
        }

        /// <summary>What stops a move from going ahead. Public because the real source checks again, on the editor's thread, in the instant before it acts.</summary>
        public static void Verify(MoveToFolderPreview preview)
        {
            if (!preview.Valid)
            {
                throw RefactorArgs.Refused($"'{preview.Folder}' cannot be used: {string.Join(" ", preview.Problems)}");
            }
        }

        /// <summary>The routes under /projects/{p}/refactor/move-to-folder, which the router hands over. Null if the path is not one of them.</summary>
        public ApiResponse Route(ApiRequest request, string method, string[] segments)
        {
            //segments: projects, {p}, refactor, move-to-folder[, preview]
            if (segments.Length == 4)
            {
                RefactorArgs.AllowPost(method);
                return ApiResponse.Json(Move(request.Client, segments[1], MoveToFolderRequest.From(RefactorArgs.Body(request.Body))));
            }

            if (segments.Length == 5 && segments[4] == "preview")
            {
                RefactorArgs.AllowPost(method);
                return ApiResponse.Json(Preview(segments[1], MoveToFolderRequest.From(RefactorArgs.Body(request.Body))));
            }

            return null;
        }

        private IList<RefactorToolSpec> _tools;

        public IEnumerable<RefactorToolSpec> Tools() => _tools ?? (_tools = new List<RefactorToolSpec>
        {
            new RefactorToolSpec
            {
                Name = "preview_move_to_folder",
                Description = "Shows what moving modules to a folder would do, and changes nothing: for each module the folder it is in now and whether the move changes anything, and whether the folder name is valid. Always call this before move_to_folder. A folder is the '@Folder annotation of Rubberduck's Code Explorer, with dots between levels (\"App.Core\").",
                Properties = MoveProperties(),
                Required = new[] { "project", "folder" },
                Run = (args, client) => Preview(RefactorArgs.Required(args, "project"), MoveToFolderRequest.From(args))
            },
            new RefactorToolSpec
            {
                Name = "move_to_folder",
                Description = "Moves one or more modules to a folder through Rubberduck's Move to Folder refactoring, which writes (or updates) the '@Folder annotation at the top of each module. Call preview_move_to_folder first. Needs the user's approval like a write. Re-read the modules afterwards (their hashes change). Fails with 409 if the editor is not in design mode.",
                Properties = MoveProperties(),
                Required = new[] { "project", "folder" },
                Run = (args, client) => Move(client, RefactorArgs.Required(args, "project"), MoveToFolderRequest.From(args))
            }
        });

        private static JObject MoveProperties() => new JObject
        {
            ["project"] = RefactorArgs.String("Project id or name."),
            ["module"] = RefactorArgs.String("The module to move."),
            ["modules"] = new JObject
            {
                ["type"] = "array",
                ["items"] = new JObject { ["type"] = "string" },
                ["description"] = "Several modules to move together, at most " + MoveToFolderRequest.MaxModules + ". Instead of, or as well as, 'module'."
            },
            ["folder"] = RefactorArgs.String("The target folder, with dots between levels, e.g. \"App.Core\". Created by being named; no empty levels.")
        };
    }
}
