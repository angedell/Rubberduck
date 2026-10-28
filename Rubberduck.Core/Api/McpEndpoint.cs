using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Rubberduck.Api
{
    /// <summary>
    /// The same operations as the REST routes, as an MCP server (Streamable HTTP, JSON responses only): one JSON-RPC message per POST.
    /// Tool failures are reported inside the result with <c>isError</c>, as the protocol wants, so the agent can read what went wrong and carry on.
    /// </summary>
    public class McpEndpoint
    {
        private const string ProtocolVersion = "2025-03-26";

        private readonly ModuleApi _api;
        private readonly DiagnosticsApi _diagnostics;
        private readonly ImmediateApi _immediate;
        private readonly ResourceApi _resources;
        private readonly InsightApi _insight;
        private readonly TestsApi _tests;
        private readonly CheckApi _check;
        private readonly LogApi _log;
        private readonly ParserSync _parser;
        private readonly DumpApi _dump;
        private readonly RefactorRenameApi _renameRefactor;

        public McpEndpoint(ModuleApi api, DiagnosticsApi diagnostics, ImmediateApi immediate, ResourceApi resources = null,
            InsightApi insight = null, TestsApi tests = null, CheckApi check = null, LogApi log = null, ParserSync parser = null, DumpApi dump = null,
            RefactorRenameApi renameRefactor = null)
        {
            _log = log;
            _dump = dump;
            _renameRefactor = renameRefactor;
            _parser = parser;
            _api = api;
            _diagnostics = diagnostics;
            _immediate = immediate;
            _resources = resources;
            _insight = insight;
            _tests = tests;
            _check = check;
        }

        public ApiResponse Handle(ApiRequest request)
        {
            if (!request.Method.Equals("POST", StringComparison.OrdinalIgnoreCase))
            {
                throw new ApiException(405, "method_not_allowed", "The MCP endpoint takes POST only.");
            }

            JObject message;
            try
            {
                message = JToken.Parse(request.Body ?? string.Empty) as JObject;
            }
            catch (JsonException)
            {
                message = null;
            }

            if (message == null)
            {
                return RpcError(null, -32700, "Parse error: send one JSON-RPC message per request.");
            }

            var id = message["id"];
            var method = message.Value<string>("method");
            if (id == null)
            {
                //A notification (or a response we never asked for): nothing to answer.
                return new ApiResponse { Status = 202 };
            }

            try
            {
                switch (method)
                {
                    case "initialize":
                        return Result(id, new JObject
                        {
                            ["protocolVersion"] = ProtocolVersion,
                            ["capabilities"] = new JObject { ["tools"] = new JObject(), ["resources"] = new JObject() },
                            ["serverInfo"] = new JObject { ["name"] = "rubberduck", ["version"] = BuildInfo.Current.Version, ["build"] = BuildInfo.Current.Display },
                            ["instructions"] = "Rubberduck is an open-source add-in for the VBA editor (VBE) of Microsoft Office (Excel, Word, Access, PowerPoint, Outlook...). It parses the VBA of every open project, and this server gives you the user's live editor session: what you read and write is the code they have open, not files on disk.\n\nWHAT YOU CAN DO\n- Navigate: list_projects, list_modules, get_structure (folders from '@Folder annotations), list_libraries (references), list_symbols (what a module declares), find_definition, find_references. Prefer these to reading whole modules.\n- Read and write code: read_module, write_module, edit_lines, batch_edit (several modules atomically), create_module, delete_module (a backup is kept). Every write needs the hash of your last read of that module; a stale hash is refused, so re-read and retry. The first write has to be approved by the user in Rubberduck.\n- Check your work: check_work reports parse errors, inspection findings and optionally test results in one call; get_parser_errors, get_diagnostics (code inspections) and get_metrics (complexity) are the parts. Always run check_work after editing.\n- Unit tests: list_tests and run_tests for Rubberduck's '@TestMethod tests. run_tests EXECUTES the user's VBA and needs their approval.\n- Document resources: list_resources, read_resource, write_resource, delete_resource work on what lives inside the project's document file besides VBA: the ribbon XML (customUI.xml / customUI14.xml), the pictures it uses (customUI/images) and rubberduck.json, Rubberduck's per-project config (indenter and git sync settings). Only saved Open XML documents (xlsm, docm, pptm...) have them. Office keeps an open document locked, so writes are staged and go into the file when the document is closed; reads already include staged changes. Reference a picture in the ribbon XML by its file name without extension.\n- Immediate window: read_immediate and run_immediate type into the editor's Immediate window. run_immediate EXECUTES CODE with the user's rights and is only available if the user switched it on.\n\nHOW TO WORK\n1. list_projects, then list_modules/get_structure to find your way; read a module before changing it.\n2. Change the smallest thing that does the job (edit_lines, or batch_edit when several modules must change together).\n3. check_work. Fix what it reports until 'clean' is true.\n4. Tell the user what you changed. Anything you wrote to a document's resources is staged until they close the document.\n\nTools that run code (run_tests, run_immediate) and the first write all need the user's consent; if one is refused, say so and carry on without it. The same data is also offered as MCP resources (module code, diagnostics, parser errors)."
                        });
                    case "ping":
                        return Result(id, new JObject());
                    case "resources/list":
                        return Result(id, new JObject { ["resources"] = new JArray(ListResources()) });
                    case "resources/templates/list":
                        return Result(id, new JObject { ["resourceTemplates"] = new JArray(ResourceTemplates()) });
                    case "resources/read":
                        return Result(id, ReadResource(message["params"] as JObject));
                    case "tools/list":
                        return Result(id, new JObject { ["tools"] = new JArray(Tools().Select(tool => tool.Describe())) });
                    case "tools/call":
                        return Result(id, CallTool(message["params"] as JObject, request.Client));
                    default:
                        return RpcError(id, -32601, $"Method not found: {method}");
                }
            }
            catch (ApiException exception)
            {
                return RpcError(id, -32602, exception.Message);
            }
        }

        private JObject CallTool(JObject parameters, string client)
        {
            var name = parameters?.Value<string>("name");
            var tool = Tools().FirstOrDefault(t => t.Name == name) ?? throw ApiException.BadRequest($"Unknown tool '{name}'.");
            var arguments = parameters["arguments"] as JObject ?? new JObject();
            try
            {
                var result = tool.Run(arguments, client);
                return Content(JsonConvert.SerializeObject(result, Formatting.Indented), false);
            }
            catch (ApiException exception)
            {
                return Content($"{exception.Code}: {exception.Message}", true);
            }
        }

        private static JObject Content(string text, bool isError) => new JObject
        {
            ["content"] = new JArray(new JObject { ["type"] = "text", ["text"] = text }),
            ["isError"] = isError
        };

        private static ApiResponse Result(JToken id, JObject result)
            => ApiResponse.Json(new JObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = result });

        private static ApiResponse RpcError(JToken id, int code, string message)
            => ApiResponse.Json(new JObject
            {
                ["jsonrpc"] = "2.0",
                ["id"] = id,
                ["error"] = new JObject { ["code"] = code, ["message"] = message }
            });

        private class Tool
        {
            public string Name;
            public string Description;
            public JObject Properties;
            public string[] Required;
            public Func<JObject, string, object> Run;

            public JObject Describe() => new JObject
            {
                ["name"] = Name,
                ["description"] = Description,
                ["inputSchema"] = new JObject
                {
                    ["type"] = "object",
                    ["properties"] = Properties,
                    ["required"] = new JArray(Required)
                }
            };
        }

        private static JObject String(string description) => new JObject { ["type"] = "string", ["description"] = description };
        private static JObject Integer(string description) => new JObject { ["type"] = "integer", ["description"] = description };

        private IList<Tool> _moduleTools;

        //The Immediate window tools are only offered while the user allows them.
        private IEnumerable<Tool> Tools() => ModuleTools()
            .Concat(_resources != null ? ResourceTools() : Enumerable.Empty<Tool>())
            .Concat(_insight != null ? InsightTools() : Enumerable.Empty<Tool>())
            .Concat(_tests != null ? TestTools().Where(t => t.Name != "run_tests" || _tests.IsEnabled) : Enumerable.Empty<Tool>())
            .Concat(_check != null ? CheckTools() : Enumerable.Empty<Tool>())
            .Concat(_log != null ? LogTools() : Enumerable.Empty<Tool>())
            .Concat(_dump != null ? DumpTools() : Enumerable.Empty<Tool>())
            .Concat(_renameRefactor != null ? FromSpecs(_renameRefactor.Tools()) : Enumerable.Empty<Tool>())
            .Concat(_immediate.IsEnabled ? ImmediateTools() : Enumerable.Empty<Tool>());

        //Tools that come with their feature, described by the feature (see RefactorToolSpec).
        private static IEnumerable<Tool> FromSpecs(IEnumerable<RefactorToolSpec> specs)
            => specs.Select(s => new Tool { Name = s.Name, Description = s.Description, Properties = s.Properties, Required = s.Required, Run = s.Run });

        private static JObject Boolean(string description) => new JObject { ["type"] = "boolean", ["description"] = description };

        private IList<Tool> _insightTools;

        private IEnumerable<Tool> InsightTools() => _insightTools ?? (_insightTools = new[]
        {
            new Tool
            {
                Name = "get_parser_errors",
                Description = "Lists the code Rubberduck's parser could not understand (syntax errors), with module, line, column and the text it stumbled on. Inspections of such code are incomplete, so check this first. Waits for the latest edits to be parsed.",
                Properties = new JObject { ["project"] = String("Project id or name.") }, Required = new[] { "project" },
                Run = (args, client) => _insight.ParserErrors(Arg(args, "project"))
            },
            new Tool
            {
                Name = "list_symbols",
                Description = "Lists what a module declares (procedures, functions, properties, module-level variables, constants, types, enums...) with kind, type, accessibility and line range. Far cheaper than reading the module to find a procedure.",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["module"] = String("Module name."),
                    ["includeLocals"] = Boolean("Also list parameters and variables local to procedures."),
                    ["kind"] = String("Only this kind, e.g. Procedure, Function, PropertyGet, Variable, Constant.")
                },
                Required = new[] { "project", "module" },
                Run = (args, client) => _insight.Symbols(Arg(args, "project"), Arg(args, "module"), args.Value<bool?>("includeLocals") ?? false, args.Value<string>("kind"))
            },
            new Tool
            {
                Name = "find_references",
                Description = "Finds every use of a declared name in a project, per declaration (the same name can be declared in several places). Use it before renaming or changing a signature.",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["name"] = String("The identifier."),
                    ["module"] = String("Only declarations in this module."),
                    ["limit"] = Integer("At most this many references per declaration. Default 200.")
                },
                Required = new[] { "project", "name" },
                Run = (args, client) => _insight.References(Arg(args, "project"), args.Value<string>("module"), Arg(args, "name"), args["limit"] != null ? Int(args, "limit") : (int?)null)
            },
            new Tool
            {
                Name = "find_definition",
                Description = "Finds where the identifier at a position (1-based line and column) is declared. Only declarations in the project's own code are returned.",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["module"] = String("Module name."),
                    ["line"] = Integer("1-based line."),
                    ["column"] = Integer("1-based column.")
                },
                Required = new[] { "project", "module", "line", "column" },
                Run = (args, client) => _insight.Definition(Arg(args, "project"), Arg(args, "module"), Int(args, "line"), Int(args, "column"))
            },
            new Tool
            {
                Name = "list_libraries",
                Description = "Lists the libraries a project references (Excel, Scripting Runtime, ADO...), with version, path and whether the reference is broken.",
                Properties = new JObject { ["project"] = String("Project id or name.") }, Required = new[] { "project" },
                Run = (args, client) => _insight.Libraries(Arg(args, "project"))
            },
            new Tool
            {
                Name = "get_structure",
                Description = "Shows how the modules of a project are organised in folders, from their '@Folder annotations.",
                Properties = new JObject { ["project"] = String("Project id or name.") }, Required = new[] { "project" },
                Run = (args, client) => _insight.Structure(Arg(args, "project"))
            },
            new Tool
            {
                Name = "get_metrics",
                Description = "Rubberduck's code metrics (lines of code, cyclomatic complexity, nesting, ...) per module and member. Shows where the complicated code is.",
                Properties = new JObject { ["project"] = String("Project id or name."), ["module"] = String("Only this module.") }, Required = new[] { "project" },
                Run = (args, client) => _insight.Metrics(Arg(args, "project"), args.Value<string>("module"))
            }
        }.ToList());

        private IList<Tool> _testTools;

        private IEnumerable<Tool> TestTools() => _testTools ?? (_testTools = new[]
        {
            new Tool
            {
                Name = "list_tests",
                Description = "Lists the unit tests (Rubberduck's '@TestMethod) of a project with the outcome of their last run since Rubberduck started: Unknown if not run. Does not run anything.",
                Properties = new JObject { ["project"] = String("Project id or name."), ["module"] = String("Only this test module.") }, Required = new[] { "project" },
                Run = (args, client) => _tests.List(Arg(args, "project"), args.Value<string>("module"))
            },
            new Tool
            {
                Name = "run_tests",
                Description = "Runs unit tests and returns their outcomes and output. THIS EXECUTES THE USER'S VBA. The user has to approve the first run. Without module and test, every test of the project runs. 'completed' is false if the wait ran out first; list_tests shows the rest later.",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["module"] = String("Only this test module."),
                    ["test"] = String("Only this test (needs module)."),
                    ["waitSeconds"] = Integer("How long to wait for the run to finish. Default 60, maximum 300.")
                },
                Required = new[] { "project" },
                Run = (args, client) => _tests.Run(client, Arg(args, "project"), args.Value<string>("module"), args.Value<string>("test"), args["waitSeconds"] != null ? Int(args, "waitSeconds") : (int?)null)
            }
        }.ToList());

        private IList<Tool> _logTools;

        private IEnumerable<Tool> LogTools() => _logTools ?? (_logTools = new[]
        {
            new Tool
            {
                Name = "get_log",
                Description = "Reads the end of Rubberduck's own log (what the add-in itself did: parsing, the Code Explorer, errors) and says which level it logs at. Use it to find out why something in Rubberduck does not work. Filters: minLevel (Trace, Debug, Info, Warn, Error, Fatal), contains (text in the message or the class), logger (class name part), tail (how many of the latest entries, default 200). If 'note' says logging is off or the setting did not take effect, call set_log_level.",
                Properties = new JObject
                {
                    ["tail"] = Integer("How many of the latest matching entries. Default 200, maximum 2000."),
                    ["minLevel"] = String("Leave out entries less severe than this: Trace, Debug, Info, Warn, Error or Fatal."),
                    ["contains"] = String("Only entries whose message or class contains this text."),
                    ["logger"] = String("Only entries from classes whose name contains this.")
                },
                Required = new string[0],
                Run = (args, client) => _log.Get(args["tail"] != null ? Int(args, "tail") : (int?)null, args.Value<string>("minLevel"), args.Value<string>("contains"), args.Value<string>("logger"))
            },
            new Tool
            {
                Name = "set_log_level",
                Description = "Changes how much Rubberduck logs: Trace (everything), Debug, Info, Warn, Error, Fatal or Off. It is saved in the user's settings as their own choice and applies at once. The user has to approve it like a write. Verbose levels slow the editor a little; set it back when done. Then reproduce the problem and read the log with get_log.",
                Properties = new JObject { ["level"] = String("Trace, Debug, Info, Warn, Error, Fatal or Off.") },
                Required = new[] { "level" },
                Run = (args, client) => _log.SetLevel(client, Arg(args, "level"))
            }
        }.ToList());

        private IList<Tool> _dumpTools;

        private IEnumerable<Tool> DumpTools() => _dumpTools ?? (_dumpTools = new[]
        {
            new Tool
            {
                Name = "capture_dump",
                Description = "Takes a memory dump of the editor's process with Sysinternals ProcDump, for when the editor or the API hangs (calls time out with 503, the parser stays Busy, Excel shows 'waiting for an OLE action'). It works while the editor's UI thread is stuck. The dump contains everything in memory, the open documents included, so the user has to approve it like a write. Answers with the file; ask the user to share it. If ProcDump is not installed the error says where to put it.",
                Properties = new JObject
                {
                    ["kind"] = String("'full' (default, all memory) or 'mini' (small: threads and stacks only)."),
                    ["waitSeconds"] = Integer("How long to wait for ProcDump. Default 120, at most 300.")
                },
                Required = new string[0],
                Run = (args, client) => _dump.Capture(client, args.Value<string>("kind"), args["waitSeconds"] != null ? Int(args, "waitSeconds") : (int?)null)
            }
        }.ToList());

        private IList<Tool> _checkTools;

        private IEnumerable<Tool> CheckTools() => _checkTools ?? (_checkTools = new[]
        {
            new Tool
            {
                Name = "check_work",
                Description = "Everything to know after an edit, in one call: parser errors, then (if the code parsed) inspection findings of warning or error severity, then optionally the unit tests. 'clean' is true when there is nothing to fix. Running tests executes code and needs the user's approval.",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["module"] = String("Limit the inspections (and tests) to this module."),
                    ["runTests"] = Boolean("Also run the unit tests. Only works if the user allowed running code."),
                    ["waitSeconds"] = Integer("How long to wait for the inspections (default 60, at most 300). If they are not done, 'inspectionsStatus' is 'running' and 'clean' is false: call again, and they are there.")
                },
                Required = new[] { "project" },
                Run = (args, client) => _check.Check(client, Arg(args, "project"), args.Value<string>("module"), args.Value<bool?>("runTests") ?? false,
                    args["waitSeconds"] != null ? Int(args, "waitSeconds") : (int?)null)
            }
        }.ToList());

        private IEnumerable<JObject> ListResources()
        {
            foreach (var project in _api.Projects())
            {
                var root = "rubberduck://" + Uri.EscapeDataString(project.Name);
                yield return Described(root + "/parser-errors", project.Name + " parser errors", "Code the parser could not understand.", "application/json");
                yield return Described(root + "/diagnostics", project.Name + " diagnostics", "Inspection findings for the whole project.", "application/json");
                foreach (var module in _api.Modules(project.Name).Take(200))
                {
                    yield return Described(root + "/modules/" + Uri.EscapeDataString(module.Name), project.Name + "/" + module.Name,
                        "The code of the " + module.Kind + " " + module.Name + ".", "text/plain");
                }
            }
        }

        private static JObject Described(string uri, string name, string description, string mime) => new JObject
        {
            ["uri"] = uri,
            ["name"] = name,
            ["description"] = description,
            ["mimeType"] = mime
        };

        private static IEnumerable<JObject> ResourceTemplates() => new[]
        {
            new JObject
            {
                ["uriTemplate"] = "rubberduck://{project}/modules/{module}",
                ["name"] = "Module code",
                ["description"] = "The code of a module, class or form.",
                ["mimeType"] = "text/plain"
            },
            new JObject
            {
                ["uriTemplate"] = "rubberduck://{project}/diagnostics",
                ["name"] = "Diagnostics",
                ["description"] = "Inspection findings for a project.",
                ["mimeType"] = "application/json"
            },
            new JObject
            {
                ["uriTemplate"] = "rubberduck://{project}/parser-errors",
                ["name"] = "Parser errors",
                ["description"] = "Code the parser could not understand.",
                ["mimeType"] = "application/json"
            }
        };

        private JObject ReadResource(JObject parameters)
        {
            var uri = parameters?.Value<string>("uri") ?? throw ApiException.BadRequest("'uri' is required.");
            const string scheme = "rubberduck://";
            if (!uri.StartsWith(scheme, StringComparison.OrdinalIgnoreCase))
            {
                throw ApiException.BadRequest($"Unknown resource '{uri}'.");
            }

            var parts = uri.Substring(scheme.Length).Split('/').Select(Uri.UnescapeDataString).ToArray();
            string text;
            string mime = "application/json";
            if (parts.Length == 3 && parts[1] == "modules")
            {
                text = _api.Read(parts[0], parts[2], false).Code;
                mime = "text/plain";
            }
            else if (parts.Length == 2 && parts[1] == "diagnostics")
            {
                text = JsonConvert.SerializeObject(_diagnostics.Get(parts[0], null, null, null), Formatting.Indented);
            }
            else if (parts.Length == 2 && parts[1] == "parser-errors" && _insight != null)
            {
                text = JsonConvert.SerializeObject(_insight.ParserErrors(parts[0]), Formatting.Indented);
            }
            else
            {
                throw ApiException.BadRequest($"Unknown resource '{uri}'.");
            }

            return new JObject
            {
                ["contents"] = new JArray(new JObject { ["uri"] = uri, ["mimeType"] = mime, ["text"] = text })
            };
        }

        private IList<Tool> _resourceTools;

        private IEnumerable<Tool> ResourceTools() => _resourceTools ?? (_resourceTools = new[]
        {
            new Tool
            {
                Name = "list_resources",
                Description = "Lists the ribbon XML (customUI.xml, customUI14.xml), the config (rubberduck.json: indenter and sync settings) and the pictures (customUI/images) stored inside a project's document file. Only saved Open XML documents (xlsm, docm, pptm...) have them; 'supported' is false otherwise. Changes you make are staged and go into the file when the document is closed: 'pending' marks them.",
                Properties = new JObject { ["project"] = String("Project id or name.") }, Required = new[] { "project" },
                Run = (args, client) => _resources.List(Arg(args, "project"))
            },
            new Tool
            {
                Name = "read_resource",
                Description = "Reads a ribbon XML part ('xml'), the config ('text') or a picture ('contentBase64') with its hash, staged changes included. Writes need the hash.",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["name"] = String("Resource name from list_resources, e.g. customUI14.xml or logo.png.")
                },
                Required = new[] { "project", "name" },
                Run = (args, client) => _resources.Read(Arg(args, "project"), Arg(args, "name"))
            },
            new Tool
            {
                Name = "write_resource",
                Description = "Adds or replaces the config (name rubberduck.json, give 'text': JSON with optional 'indenter' and 'sync' sections) or ribbon XML (give 'xml'; the name is customUI.xml for Office 2007 or customUI14.xml for 2010 and later, and the root must be <customUI> with the matching namespace) or a picture (give 'contentBase64'; png, jpg, gif, bmp or ico, up to 5 MB). A picture is referred to in the XML by its file name without extension, e.g. image=\"logo\". Replacing needs the hash from your last read. The change is staged until the document is closed. The first write has to be approved by the user.",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["name"] = String("customUI.xml, customUI14.xml, or a picture file name such as logo.png."),
                    ["xml"] = String("The ribbon XML."),
                    ["text"] = String("The JSON of rubberduck.json."),
                    ["contentBase64"] = String("The picture, base64 encoded."),
                    ["expectedHash"] = String("The hash from the read this write is based on. Not needed for a new resource.")
                },
                Required = new[] { "project", "name" },
                Run = (args, client) => _resources.Write(client, Arg(args, "project"), Arg(args, "name"),
                    args.Value<string>("xml"), args.Value<string>("text"), args.Value<string>("contentBase64"), args.Value<string>("expectedHash"))
            },
            new Tool
            {
                Name = "delete_resource",
                Description = "Removes a picture from the document (staged until it is closed). Ribbon XML and rubberduck.json cannot be removed, only replaced.",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["name"] = String("Picture file name."),
                    ["expectedHash"] = String("The hash from your last read of the picture.")
                },
                Required = new[] { "project", "name", "expectedHash" },
                Run = (args, client) => _resources.Delete(client, Arg(args, "project"), Arg(args, "name"), Arg(args, "expectedHash"))
            }
        }.ToList());

        private IEnumerable<Tool> ImmediateTools() => new[]
        {
            new Tool
            {
                Name = "read_immediate",
                Description = "Reads the text of the editor's Immediate window.",
                Properties = new JObject(), Required = new string[0],
                Run = (args, client) => _immediate.Read()
            },
            new Tool
            {
                Name = "run_immediate",
                Description = "Types one or more lines into the editor's Immediate window and runs them, as if the user had pressed Enter: '?expression' prints a value, a statement runs, a procedure name calls it. THIS EXECUTES CODE with the user's rights. The user has to approve the first command. Returns what appeared in the window; 'completed' is false if it was still changing when the wait ran out.",
                Properties = new JObject
                {
                    ["command"] = String("The line(s) to run, separated by newlines."),
                    ["waitMs"] = Integer("How long to wait for the output to settle, at most. Default 5000, maximum 60000.")
                },
                Required = new[] { "command" },
                Run = (args, client) => _immediate.Run(client, Arg(args, "command"), args["waitMs"] != null ? Int(args, "waitMs") : (int?)null)
            }
        };

        private IEnumerable<Tool> ModuleTools() => _moduleTools ?? (_moduleTools = new[]
        {
            new Tool
            {
                Name = "get_diagnostics",
                Description = "Runs Rubberduck's code inspections over a project or one module and returns what they found (worst first), with line and column. Use it after writing code to check it. Waits for Rubberduck to parse the latest edits.",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["module"] = String("Only this module."),
                    ["minSeverity"] = new JObject { ["type"] = "string", ["enum"] = new JArray(DiagnosticsApi.Severities), ["description"] = "Leave out anything less severe. Default: everything." },
                    ["limit"] = Integer("At most this many results. Default 200, maximum 1000."),
                    ["waitSeconds"] = Integer("How long to wait for the inspections (default 60, at most 300; they can take minutes on a large project). If they are not done, the answer has status 'running' and no results: call again, and they are there.")
                },
                Required = new[] { "project" },
                Run = (args, client) => _diagnostics.Get(Arg(args, "project"), args.Value<string>("module"), args.Value<string>("minSeverity"),
                    args["limit"] != null ? Int(args, "limit") : (int?)null, args["waitSeconds"] != null ? Int(args, "waitSeconds") : (int?)null)
            },
            new Tool
            {
                Name = "get_info",
                Description = "Says which build of Rubberduck is answering: the full version (the last number counts up with every build) and the git commit. Mention it when you report a problem.",
                Properties = new JObject(), Required = new string[0],
                Run = (args, client) => ApiRouter.InfoReport()
            },
            new Tool
            {
                Name = "reparse",
                Description = "Does what Refresh in the Code Explorer does: parses everything again and waits until the parser is done. Use it when a workbook was opened after Rubberduck started and list_projects does not show it. Answers with the parser state and the projects.",
                Properties = new JObject(), Required = new string[0],
                Run = (args, client) => ApiRouter.ReparseReport(_parser, _api)
            },
            new Tool
            {
                Name = "list_projects",
                Description = "Lists the VBA projects open in the editor, and 'vbeMode': 'design' (all is well), 'running_or_break' (a macro runs or is stopped: writes, tests and the parser are refused with 409 until the user resets it in the editor) or 'unresponsive' (the editor does not answer, probably a dialog nobody can see; no projects are listed).",
                Properties = new JObject(), Required = new string[0],
                Run = (args, client) => ApiRouter.ProjectsReport(_parser, _api)
            },
            new Tool
            {
                Name = "list_modules",
                Description = "Lists the modules, classes, forms and document modules (sheets, workbook) of a project. Read a module to get the hash its writes need.",
                Properties = new JObject { ["project"] = String("Project id or name.") }, Required = new[] { "project" },
                Run = (args, client) => _api.Modules(Arg(args, "project"))
            },
            new Tool
            {
                Name = "read_module",
                Description = "Reads a module's code and its hash. Pass includeText to also get the exported module with its header (attributes, form layout), or includeAttributes for the hidden attributes as data ({module:{VB_PredeclaredId:true}, members:{Name:{VB_UserMemId:0}}}).",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["module"] = String("Module name."),
                    ["includeText"] = new JObject { ["type"] = "boolean", ["description"] = "Also return the full exported text." },
                    ["includeAttributes"] = new JObject { ["type"] = "boolean", ["description"] = "Also return the hidden attributes as data." }
                },
                Required = new[] { "project", "module" },
                Run = (args, client) => _api.Read(Arg(args, "project"), Arg(args, "module"), args.Value<bool?>("includeText") ?? false, args.Value<bool?>("includeAttributes") ?? false)
            },
            new Tool
            {
                Name = "write_module",
                Description = "Replaces a module's whole code. Send the hash from your last read. Give 'code' for the code alone (the header and a form's layout stay), or 'text' to replace the module as exported, header included. Give 'attributes' ({module:{VB_PredeclaredId:true}, members:{Name:{VB_UserMemId:0}}}) to set hidden attributes without rebuilding the header, alone or together with 'code'.",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["module"] = String("Module name."),
                    ["expectedHash"] = String("The hash from the read this write is based on."),
                    ["code"] = String("The new code, without header."),
                    ["text"] = String("The new module as exported, with header. Use instead of code."),
                    ["attributes"] = new JObject { ["type"] = "object", ["description"] = "Hidden attributes to set: { module: { VB_Name: value }, members: { MemberName: { VB_UserMemId: 0 } } }." },
                    ["expectedAttributesHash"] = String("With 'attributes' alone: the attributesHash from your last read with include=attributes. Refused if the attributes changed since.")
                },
                Required = new[] { "project", "module", "expectedHash" },
                Run = (args, client) =>
                {
                    var attributes = args["attributes"] as JObject;
                    if (args["text"] != null)
                    {
                        return _api.ReplaceText(client, Arg(args, "project"), Arg(args, "module"), Arg(args, "text"), Arg(args, "expectedHash"));
                    }

                    if (args["code"] == null && attributes != null)
                    {
                        return _api.WriteAttributes(client, Arg(args, "project"), Arg(args, "module"), attributes, Arg(args, "expectedHash"), args.Value<string>("expectedAttributesHash"));
                    }

                    var replaced = _api.ReplaceCode(client, Arg(args, "project"), Arg(args, "module"), Arg(args, "code"), Arg(args, "expectedHash"));
                    return attributes == null ? replaced : _api.WriteAttributes(client, Arg(args, "project"), Arg(args, "module"), attributes, replaced.Hash);
                }
            },
            new Tool
            {
                Name = "edit_lines",
                Description = "Replaces lines startLine to endLine (1-based, inclusive) of a module's code. An endLine of startLine - 1 inserts before startLine.",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["module"] = String("Module name."),
                    ["expectedHash"] = String("The hash from the read this edit is based on."),
                    ["startLine"] = Integer("First line to replace."),
                    ["endLine"] = Integer("Last line to replace."),
                    ["replacement"] = String("The new lines. Empty deletes the range.")
                },
                Required = new[] { "project", "module", "expectedHash", "startLine", "endLine", "replacement" },
                Run = (args, client) => _api.EditLines(client, Arg(args, "project"), Arg(args, "module"),
                    Int(args, "startLine"), Int(args, "endLine"), Arg(args, "replacement"), Arg(args, "expectedHash"))
            },
            new Tool
            {
                Name = "batch_edit",
                Description = "Changes several modules in one call, one edit per module. Every hash is checked before anything is changed, so a stale one fails the whole batch untouched. Each edit has an 'op' (replaceCode, editLines, replaceText, create, delete), the 'module', and the fields of that operation: expectedHash for all but create; code (replaceCode, create); text (replaceText); startLine, endLine, replacement (editLines); kind (create). One approval covers the batch.",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["edits"] = new JObject
                    {
                        ["type"] = "array",
                        ["description"] = "The edits, at most " + ModuleApi.MaxBatchEdits + ".",
                        ["items"] = new JObject
                        {
                            ["type"] = "object",
                            ["properties"] = new JObject
                            {
                                ["op"] = String("replaceCode, editLines, replaceText, create or delete."),
                                ["module"] = String("Module name."),
                                ["expectedHash"] = String("Hash from the read this edit is based on."),
                                ["code"] = String("New code (replaceCode, create)."),
                                ["text"] = String("New module as exported (replaceText)."),
                                ["kind"] = String("standard, class or form (create)."),
                                ["startLine"] = Integer("First line (editLines)."),
                                ["endLine"] = Integer("Last line (editLines)."),
                                ["replacement"] = String("New lines (editLines).")
                            },
                            ["required"] = new JArray("op", "module")
                        }
                    }
                },
                Required = new[] { "project", "edits" },
                Run = (args, client) =>
                {
                    var edits = args["edits"] as JArray ?? throw ApiException.BadRequest("'edits' is required and must be a list.");
                    try
                    {
                        return _api.Batch(client, Arg(args, "project"), edits.Select(e => e.ToObject<ModuleEdit>()).ToList());
                    }
                    catch (JsonException)
                    {
                        throw ApiException.BadRequest("Every edit must be an object with string fields and whole-number lines.");
                    }
                }
            },
            new Tool
            {
                Name = "create_module",
                Description = "Adds a standard module, class or (empty) form to a project.",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["module"] = String("Name of the new module."),
                    ["kind"] = new JObject { ["type"] = "string", ["enum"] = new JArray(ModuleKinds.Creatable), ["description"] = "What to create." },
                    ["code"] = String("Its code.")
                },
                Required = new[] { "project", "module", "kind" },
                Run = (args, client) => _api.Create(client, Arg(args, "project"), Arg(args, "module"), Arg(args, "kind"), args.Value<string>("code"))
            },
            new Tool
            {
                Name = "delete_module",
                Description = "Removes a module or form from a project (document modules only lose their code). A copy is kept in Rubberduck's backups.",
                Properties = new JObject
                {
                    ["project"] = String("Project id or name."),
                    ["module"] = String("Module name."),
                    ["expectedHash"] = String("The hash from your last read.")
                },
                Required = new[] { "project", "module", "expectedHash" },
                Run = (args, client) =>
                {
                    _api.Delete(client, Arg(args, "project"), Arg(args, "module"), Arg(args, "expectedHash"));
                    return new { deleted = Arg(args, "module") };
                }
            }
        }.ToList());

        private static string Arg(JObject args, string name)
        {
            var token = args[name];
            return token != null && token.Type == JTokenType.String
                ? token.Value<string>()
                : throw ApiException.BadRequest($"'{name}' is required and must be a string.");
        }

        private static int Int(JObject args, string name)
        {
            var token = args[name];
            return token != null && token.Type == JTokenType.Integer
                ? token.Value<int>()
                : throw ApiException.BadRequest($"'{name}' is required and must be a whole number.");
        }
    }
}
