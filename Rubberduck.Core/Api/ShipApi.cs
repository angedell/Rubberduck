using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;
using Newtonsoft.Json.Linq;
using NLog;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Ship;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Api
{
    /// <summary>Finds an open project by the API's key and runs work on it on the editor's thread, the only one the VBE can be talked to from.</summary>
    public interface IShipProjects
    {
        /// <exception cref="ApiException">No such open project, or more than one; the editor does not answer.</exception>
        T With<T>(string project, Func<IVBProject, T> work, TimeSpan? timeout = null);

        /// <summary>Finds the project on the editor's thread and hands it back, so that long work, which must not hold that thread, can use it from the caller's own.</summary>
        IVBProject Resolve(string project);
    }

    /// <summary><see cref="IShipProjects"/> on top of the VBE.</summary>
    public class VbeShipProjects : IShipProjects
    {
        private static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(60);

        private readonly IProjectsProvider _projects;
        private readonly IUiDispatcher _dispatcher;

        public VbeShipProjects(IProjectsProvider projects, IUiDispatcher dispatcher)
        {
            _projects = projects;
            _dispatcher = dispatcher;
        }

        public IVBProject Resolve(string project) => With(project, p => p, TimeSpan.FromSeconds(30));

        public T With<T>(string project, Func<IVBProject, T> work, TimeSpan? timeout = null)
        {
            var task = _dispatcher.StartTask(() => work(Find(project)), CancellationToken.None);
            try
            {
                if (!task.Wait(timeout ?? DefaultTimeout))
                {
                    throw ApiException.Unavailable("The editor did not respond in time. It may be showing a dialog or running code, or Ship is still working: look at the result in the workbook's folder.");
                }

                return task.Result;
            }
            catch (AggregateException exception)
            {
                var inner = exception.InnerException ?? exception;
                if (inner is COMException)
                {
                    throw ApiException.Unavailable($"The editor refused the request ({inner.Message}). It may be in break mode or running code.");
                }

                ExceptionDispatchInfo.Capture(inner).Throw();
                throw;
            }
        }

        private IVBProject Find(string key)
        {
            var matches = ProjectIds.Match(key, _projects.Projects().ToList(), p => (p.ProjectId, p.Project.Name, ProjectIds.PathOf(p.Project)));
            switch (matches.Count)
            {
                case 0:
                    throw ApiException.NotFound($"No open project '{key}'.");
                case 1:
                    return matches[0].Project;
                default:
                    throw ApiException.Conflict($"More than one open project is named '{key}'; use its id (see /projects).");
            }
        }
    }

    /// <summary>Ship cannot start (see <see cref="ShipRefusedException"/>). Carries what the answer shows beyond code and message.</summary>
    public class ShipRefusal : ApiException
    {
        public ShipRefusal(int status, string reason, string message) : base(status, "ship_" + reason, message) { Reason = reason; }

        public string Reason { get; }

        /// <summary>For <c>exists</c>: the file in the way.</summary>
        public string Path { get; set; }

        /// <summary>For <c>exists</c>: the values of 'onExisting' that decide it.</summary>
        public string[] Choices { get; set; }
    }

    /// <summary>
    /// Ship: a clean .xlam/.xlsm copy of the saved workbook, without what Rubberduck added (notes, test modules, annotation comments, descriptive attributes, config).
    /// The open project is never changed. The routes read and write the <c>ship</c> settings, preview what would be removed, and ship.
    /// Shipping writes a file, so it needs the user's approval and an explicit <c>confirm</c>, unless the user switched on unattended ship for the document:
    /// that switch is theirs alone and cannot be read as writable here.
    /// </summary>
    public class ShipApi
    {
        private static readonly string[] ExistingChoices = { "ask", "replace", "keepBoth" };
        private static readonly string[] SettingKeys = { "format", "outputFolder", "fileName", "onExisting", "steps" };
        private static readonly TimeSpan ShipTimeout = TimeSpan.FromMinutes(10);
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        /// <summary>What agents are told about Ship, in the tool descriptions and with the settings.</summary>
        public const string Summary = "Ship produces a clean copy of the workbook as .xlam or .xlsm, without what Rubberduck added: notes, test modules and the test library reference, annotation comments ('@Folder, '@Description...), descriptive attributes and rubberduck.json. The comments, notes and tests are stripped from the artifact, never from the open project, which is not changed. It works on the saved file, so the workbook must be an Excel document that has been saved, and what is unsaved in the editor is not in the artifact. Preview first (preview_ship), then ship_project with confirm: true.";

        private readonly IShipService _service;
        private readonly IShipProjects _projects;
        private readonly ApiWriteApproval _approval;

        public ShipApi(IShipService service, IShipProjects projects, ApiWriteApproval approval)
        {
            _service = service;
            _projects = projects;
            _approval = approval;
        }

        #region Operations

        public JObject Settings(string project)
        {
            return Translate(() => _projects.With(project, p => View(p, _service.ReadSettings(p))));
        }

        /// <summary>Changes the settings that are in <paramref name="body"/>; the others stay. A null file name or folder clears it.</summary>
        public JObject PutSettings(string client, string project, JObject body)
        {
            if (body.Property("unattended") != null)
            {
                throw new ApiException(400, "unattended_read_only", "Unattended ship is the user's own switch (in the Ship dialog) and cannot be set through the API. Leave 'unattended' out.");
            }

            var unknown = body.Properties().Select(p => p.Name).Where(n => !SettingKeys.Contains(n)).ToList();
            if (unknown.Count > 0)
            {
                throw ApiException.BadRequest($"Unknown setting(s): {string.Join(", ", unknown)}. Settings are: {string.Join(", ", SettingKeys)}.");
            }

            return Translate(() =>
            {
                var settings = _projects.With(project, p => _service.ReadSettings(p));
                Apply(body, settings);
                _approval.Require(client);
                Logger.Info("Ship settings written for project '{0}' by client '{1}'.", project, client);
                return _projects.With(project, p =>
                {
                    _service.WriteSettings(p, settings);
                    return View(p, _service.ReadSettings(p));
                });
            });
        }

        public JObject Preview(string project)
        {
            return Translate(() => _projects.With(project, p =>
            {
                var report = _service.Preview(p, _service.ReadSettings(p));
                return Report(report, preview: true);
            }, ShipTimeout));
        }

        /// <param name="confirm">Without it, and unless the user switched on unattended ship, this only previews.</param>
        /// <param name="onExisting">"ask", "replace" or "keepBoth": what to do if the file exists. Null: the setting decides.</param>
        public JObject Ship(string client, string project, bool confirm, string onExisting)
        {
            var choice = ParseOnExisting(onExisting);
            return Translate(() =>
            {
                var state = _projects.With(project, p =>
                {
                    var settings = _service.ReadSettings(p);
                    var unattended = _service.IsUnattended(p);
                    return new { settings, unattended, preview = unattended || confirm ? null : _service.Preview(p, settings) };
                }, ShipTimeout);

                if (state.preview != null)
                {
                    var answer = Report(state.preview, preview: true);
                    answer["shipped"] = false;
                    answer["confirmRequired"] = true;
                    answer["note"] = "Nothing was shipped: this is the preview. To ship, send { \"confirm\": true } (and \"onExisting\": \"replace\" or \"keepBoth\" if the file exists). The first write also has to be approved by the user in Rubberduck.";
                    return answer;
                }

                //An existing file is replaced only by the user's own hand: from the dialog, or with their switch on. An agent can edit the folder and the name in the settings,
                //so a replace that it asks for would let it overwrite any file of that name.
                var effective = choice != ShipOnExisting.Ask ? choice : state.settings.OnExisting;
                if (effective == ShipOnExisting.Replace && !state.unattended)
                {
                    throw new ShipRefusal(403, "replace_needs_user", "Replacing an existing file is left to the user: they can do it from the Ship dialog, or switch on \"Ship without asking\". Use \"keepBoth\" or another file name.");
                }

                //The user's own switch lifts both the confirm and the approval; without it, the approval is asked here, off the editor's thread.
                if (!state.unattended)
                {
                    _approval.Require(client);
                }

                Logger.Info("Ship started for project '{0}' by client '{1}' (unattended {2}, onExisting {3}).", project, client, state.unattended, choice);
                //Not on the editor's thread: Ship waits for a second Excel, and its watchdog could not fire from a thread that is itself blocked.
                var report = _service.Ship(_projects.Resolve(project), state.settings, choice, new Quiet());
                var shipped = Report(report, preview: false);
                shipped["shipped"] = true;
                shipped["unattended"] = state.unattended;
                return shipped;
            });
        }

        private class Quiet : IProgress<string>
        {
            public void Report(string value) { }
        }

        #endregion

        #region Translation

        private static T Translate<T>(Func<T> work)
        {
            try
            {
                return work();
            }
            catch (ShipRefusedException refused)
            {
                throw Refusal(refused);
            }
        }

        private static ShipRefusal Refusal(ShipRefusedException refused)
        {
            var reason = refused.Reason ?? "failed";
            switch (reason)
            {
                case "exists":
                    var where = string.IsNullOrEmpty(refused.Path) ? "The artifact's file" : $"'{refused.Path}'";
                    return new ShipRefusal(409, reason, $"{where} exists. Send \"onExisting\": \"replace\" to overwrite it, \"keepBoth\" to ship under a new name next to it, or \"ask\" to leave the choice to the user. {refused.Message}")
                    {
                        Path = refused.Path,
                        Choices = ExistingChoices
                    };
                case "unsaved":
                case "targetIsSource":
                    return new ShipRefusal(409, reason, refused.Message);
                case "restrictedFolder":
                    return new ShipRefusal(422, reason, refused.Message);
                case "unsupportedHost":
                case "notExcel":
                case "folderMissing":
                case "folderNotWritable":
                case "vbomNotTrusted":
                    return new ShipRefusal(422, reason, refused.Message);
                case "excelUnavailable":
                    return new ShipRefusal(503, reason, refused.Message);
                default:
                    return new ShipRefusal(500, reason, refused.Message);
            }
        }

        private static ShipOnExisting ParseOnExisting(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return ShipOnExisting.Ask;
            }

            switch (text.Trim().ToLowerInvariant())
            {
                case "ask":
                    return ShipOnExisting.Ask;
                case "replace":
                    return ShipOnExisting.Replace;
                case "keepboth":
                    return ShipOnExisting.KeepBoth;
                default:
                    throw ApiException.BadRequest("'onExisting' must be \"ask\", \"replace\" or \"keepBoth\".");
            }
        }

        private static string Name(ShipOnExisting value)
            => value == ShipOnExisting.KeepBoth ? "keepBoth" : value.ToString().ToLowerInvariant();

        private static void Apply(JObject body, ShipSettings settings)
        {
            if (body.TryGetValue("format", out var format))
            {
                var text = Str(format, "format");
                if (string.Equals(text, "xlsm", StringComparison.OrdinalIgnoreCase)) settings.Format = ShipFormat.Xlsm;
                else if (string.Equals(text, "xlam", StringComparison.OrdinalIgnoreCase)) settings.Format = ShipFormat.Xlam;
                else throw ApiException.BadRequest("'format' must be \"xlsm\" or \"xlam\".");
            }

            if (body.TryGetValue("outputFolder", out var folder))
            {
                settings.OutputFolder = string.IsNullOrWhiteSpace(Str(folder, "outputFolder")) ? null : Str(folder, "outputFolder").Trim();
            }

            if (body.TryGetValue("fileName", out var fileName))
            {
                settings.FileName = string.IsNullOrWhiteSpace(Str(fileName, "fileName")) ? null : Str(fileName, "fileName").Trim();
            }

            if (body.TryGetValue("onExisting", out var existing))
            {
                settings.OnExisting = ParseOnExisting(Str(existing, "onExisting"));
            }

            if (body.TryGetValue("steps", out var steps))
            {
                var map = steps as JObject ?? throw ApiException.BadRequest($"'steps' must be an object of step name to true/false. Steps: {string.Join(", ", ShipSteps.All)}.");
                foreach (var step in map.Properties())
                {
                    var name = ShipSteps.All.FirstOrDefault(s => string.Equals(s, step.Name, StringComparison.OrdinalIgnoreCase))
                        ?? throw ApiException.BadRequest($"Unknown step '{step.Name}'. Steps: {string.Join(", ", ShipSteps.All)}.");
                    if (step.Value.Type != JTokenType.Boolean)
                    {
                        throw ApiException.BadRequest($"Step '{name}' must be true or false.");
                    }

                    settings.Steps[name] = step.Value.Value<bool>();
                }
            }
        }

        private static string Str(JToken token, string name)
        {
            if (token.Type == JTokenType.Null)
            {
                return null;
            }

            return token.Type == JTokenType.String ? token.Value<string>() : throw ApiException.BadRequest($"'{name}' must be a string.");
        }

        private JObject View(IVBProject project, ShipSettings settings)
        {
            var steps = new JObject();
            foreach (var step in ShipSteps.All)
            {
                steps[step] = settings.IsOn(step);
            }

            return new JObject
            {
                ["format"] = settings.Format.ToString().ToLowerInvariant(),
                ["outputFolder"] = settings.OutputFolder,
                ["resolvedOutputFolder"] = _service.ResolveOutputFolder(project, settings),
                ["fileName"] = settings.FileName,
                ["onExisting"] = Name(settings.OnExisting),
                ["steps"] = steps,
                ["unattended"] = _service.IsUnattended(project),
                ["unattendedReadOnly"] = true,
                ["canShip"] = _service.CanShip(project),
                ["note"] = "'unattended' is the user's own switch: it lifts the confirm and the approval for shipping, and cannot be changed here (a PUT that sends it is refused). 'resolvedOutputFolder' is where the settings lead on this machine; null if no folder is set."
            };
        }

        private static JObject Report(ShipReport report, bool preview)
        {
            var counts = new JObject();
            var steps = new JArray();
            foreach (var step in report.Steps)
            {
                counts[step.Step] = step.Removed;
                steps.Add(new JObject
                {
                    ["step"] = step.Step,
                    ["removed"] = step.Removed,
                    ["details"] = new JArray(step.Details)
                });
            }

            return new JObject
            {
                ["dryRun"] = report.DryRun || preview,
                ["outputPath"] = report.OutputPath,
                ["counts"] = counts,
                ["steps"] = steps,
                ["unparsedModules"] = new JArray(report.Unparsed),
                ["warnings"] = new JArray(report.Warnings)
            };
        }

        #endregion

        #region Routes

        public ApiResponse Route(ApiRequest request, string method, string[] segments)
        {
            //segments: projects, {p}, ship[, settings|preview]
            var project = segments[1];
            var rest = segments.Length == 3 ? null : segments[3];
            try
            {
                if (segments.Length == 3)
                {
                    if (method != "POST")
                    {
                        throw MethodNotAllowed("POST");
                    }

                    var body = string.IsNullOrWhiteSpace(request.Body) ? new JObject() : RefactorArgs.Body(request.Body);
                    return ApiResponse.Json(Ship(request.Client, project, Flag(body, "confirm"), RefactorArgs.Text(body, "onExisting")));
                }

                if (segments.Length == 4 && rest == "preview")
                {
                    if (method != "GET")
                    {
                        throw MethodNotAllowed("GET");
                    }

                    return ApiResponse.Json(Preview(project));
                }

                if (segments.Length == 4 && rest == "settings")
                {
                    switch (method)
                    {
                        case "GET":
                            return ApiResponse.Json(Settings(project));
                        case "PUT":
                            return ApiResponse.Json(PutSettings(request.Client, project, RefactorArgs.Body(request.Body)));
                        default:
                            throw MethodNotAllowed("GET, PUT");
                    }
                }

                throw ApiException.NotFound("No such route.");
            }
            catch (ShipRefusal refusal)
            {
                var error = new JObject { ["code"] = refusal.Code, ["message"] = refusal.Message, ["reason"] = refusal.Reason };
                if (refusal.Path != null)
                {
                    error["path"] = refusal.Path;
                }

                if (refusal.Choices != null)
                {
                    error["choices"] = new JArray(refusal.Choices);
                }

                return new ApiResponse { Status = refusal.Status, Body = new JObject { ["error"] = error } };
            }
        }

        private static ApiException MethodNotAllowed(string allowed) => new ApiException(405, "method_not_allowed", $"Allowed: {allowed}.");

        private static bool Flag(JObject body, string name)
        {
            var token = body[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return false;
            }

            return token.Type == JTokenType.Boolean ? token.Value<bool>() : throw ApiException.BadRequest($"'{name}' must be true or false.");
        }

        #endregion

        #region Tools

        private static JObject Prop(string type, string description) => new JObject { ["type"] = type, ["description"] = description };

        private IList<RefactorToolSpec> _tools;

        public IEnumerable<RefactorToolSpec> Tools() => _tools ?? (_tools = new List<RefactorToolSpec>
        {
            new RefactorToolSpec
            {
                Name = "get_ship_settings",
                Description = Summary + " Reads the 'ship' settings of the project's rubberduck.json: format (xlsm or xlam), outputFolder as written and resolvedOutputFolder as it leads on this machine, fileName, onExisting (ask, replace, keepBoth) and the strip steps (notes, testModules, comments, descriptiveAttributes, config). 'unattended' is shown but is the user's own switch, which only they can change.",
                Properties = new JObject { ["project"] = Prop("string", "Project id or name.") }, Required = new[] { "project" },
                Run = (args, client) => Settings(Arg(args, "project"))
            },
            new RefactorToolSpec
            {
                Name = "set_ship_settings",
                Description = "Changes the 'ship' settings in the project's rubberduck.json; what you leave out stays. Needs the user's approval like a write. Sending 'unattended' is refused: that switch is the user's alone. " + Summary,
                Properties = new JObject
                {
                    ["project"] = Prop("string", "Project id or name."),
                    ["format"] = Prop("string", "\"xlsm\" (a workbook with macros) or \"xlam\" (an add-in)."),
                    ["outputFolder"] = Prop("string", "Where the artifact goes: absolute, relative to the workbook, ../, %ENV%, or a known/OneDrive/Dropbox folder. An empty string clears it."),
                    ["fileName"] = Prop("string", "File name without extension. An empty string means the project's name."),
                    ["onExisting"] = Prop("string", "What to do when the file exists: \"ask\", \"replace\" or \"keepBoth\"."),
                    ["steps"] = Prop("object", "Strip steps to switch on/off, e.g. { \"comments\": false }. Steps: " + string.Join(", ", ShipSteps.All) + ".")
                },
                Required = new[] { "project" },
                Run = (args, client) =>
                {
                    var settings = (JObject)args.DeepClone();
                    settings.Remove("project");
                    return PutSettings(client, Arg(args, "project"), settings);
                }
            },
            new RefactorToolSpec
            {
                Name = "preview_ship",
                Description = "Shows what Ship would remove, step by step, and which modules did not parse and would be shipped unchanged. Writes nothing and needs no approval. Look at it before ship_project. " + Summary,
                Properties = new JObject { ["project"] = Prop("string", "Project id or name.") }, Required = new[] { "project" },
                Run = (args, client) => Preview(Arg(args, "project"))
            },
            new RefactorToolSpec
            {
                Name = "ship_project",
                Description = "Writes the clean artifact (.xlam or .xlsm) to the output folder of the settings and returns the report: counts per step, unparsed modules, the output path. It never changes the open project, and the workbook must have been saved. Needs the user's approval and confirm: true; without confirm it only previews and tells you what to send (both are waived when the user switched on unattended ship). If the file exists the answer is the error ship_exists with the path and the choices: send onExisting \"replace\" or \"keepBoth\", or let the user decide. " + Summary,
                Properties = new JObject
                {
                    ["project"] = Prop("string", "Project id or name."),
                    ["confirm"] = Prop("boolean", "true to ship. Without it you get the preview."),
                    ["onExisting"] = Prop("string", "What to do if the file exists: \"replace\", \"keepBoth\" or \"ask\" (the setting decides when left out).")
                },
                Required = new[] { "project" },
                Run = (args, client) => Ship(client, Arg(args, "project"), args.Value<bool?>("confirm") ?? false, args.Value<string>("onExisting"))
            }
        });

        private static string Arg(JObject args, string name)
        {
            var value = args.Value<string>(name);
            return string.IsNullOrWhiteSpace(value) ? throw ApiException.BadRequest($"'{name}' is required.") : value;
        }

        #endregion
    }
}
