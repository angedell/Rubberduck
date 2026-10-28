using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Rubberduck.Api
{
    public class ApiRequest
    {
        public string Method { get; set; }
        public string Path { get; set; }
        public IDictionary<string, string> Query { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public string IfMatch { get; set; }
        public string Body { get; set; }

        /// <summary>Who is calling, as far as the user needs to know to approve them.</summary>
        public string Client { get; set; }
    }

    public class ApiResponse
    {
        public int Status { get; set; } = 200;
        public object Body { get; set; }
        public string ETag { get; set; }

        public static ApiResponse Json(object body, string etag = null) => new ApiResponse { Body = body, ETag = etag };

        public static ApiResponse Error(int status, string code, string message)
            => new ApiResponse { Status = status, Body = new { error = new { code, message } } };
    }

    /// <summary>Turns a request into a call on <see cref="ModuleApi"/> and the result into a response. Knows nothing of HTTP itself.</summary>
    public class ApiRouter
    {
        public const string Base = "/api/v1";
        public const string McpPath = "/mcp";

        private readonly ModuleApi _api;
        private readonly DiagnosticsApi _diagnostics;
        private readonly ImmediateApi _immediate;
        private readonly ResourceApi _resources;
        private readonly InsightApi _insight;
        private readonly TestsApi _tests;
        private readonly CheckApi _check;
        private readonly LogApi _log;
        private readonly McpEndpoint _mcp;

        public ApiRouter(ModuleApi api, DiagnosticsApi diagnostics, ImmediateApi immediate, ResourceApi resources = null,
            InsightApi insight = null, TestsApi tests = null, CheckApi check = null, LogApi log = null)
        {
            _log = log;
            _api = api;
            _diagnostics = diagnostics;
            _immediate = immediate;
            _resources = resources;
            _insight = insight;
            _tests = tests;
            _check = check;
            _mcp = new McpEndpoint(api, diagnostics, immediate, resources, insight, tests, check, log);
        }

        public ApiResponse Handle(ApiRequest request)
        {
            try
            {
                var path = request.Path.TrimEnd('/');
                if (path.Equals(McpPath, StringComparison.OrdinalIgnoreCase))
                {
                    return _mcp.Handle(request);
                }

                if (!path.StartsWith(Base, StringComparison.OrdinalIgnoreCase))
                {
                    throw ApiException.NotFound("No such route.");
                }

                var segments = path.Substring(Base.Length).Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(Uri.UnescapeDataString).ToArray();
                return Route(request, segments);
            }
            catch (ApiException exception)
            {
                return ApiResponse.Error(exception.Status, exception.Code, exception.Message);
            }
            catch (JsonException)
            {
                return ApiResponse.Error(400, "bad_json", "The body is not valid JSON.");
            }
        }

        /// <summary>Which build is running, the same as the splash screen says. Public so the MCP tool can give the same answer.</summary>
        public static InfoReport InfoReport()
        {
            var build = BuildInfo.Current;
            return new InfoReport { Name = "Rubberduck", Version = build.Version, Commit = build.Commit, Display = build.Display, Api = "v1" };
        }

        private ApiResponse Route(ApiRequest request, string[] segments)
        {
            var method = request.Method.ToUpperInvariant();

            // /projects
            if (segments.Length == 1 && segments[0] == "projects")
            {
                Allow(method, "GET");
                return ApiResponse.Json(new { projects = _api.Projects() });
            }

            // /info: which build is answering
            if (segments.Length == 1 && segments[0] == "info")
            {
                Allow(method, "GET");
                return ApiResponse.Json(InfoReport());
            }

            // /log: Rubberduck's own log and its level
            if (segments.Length == 1 && segments[0] == "log" && _log != null)
            {
                switch (method)
                {
                    case "GET":
                        var tail = Query(request, "tail");
                        if (tail != null && !int.TryParse(tail, out _))
                        {
                            throw ApiException.BadRequest("tail must be a whole number.");
                        }

                        return ApiResponse.Json(_log.Get(tail == null ? (int?)null : int.Parse(tail), Query(request, "minLevel"), Query(request, "contains"), Query(request, "logger")));
                    case "PUT":
                        return ApiResponse.Json(_log.SetLevel(request.Client, Text(Body(request), "level")));
                    default:
                        throw MethodNotAllowed("GET, PUT");
                }
            }

            // /immediate
            if (segments.Length == 1 && segments[0] == "immediate")
            {
                switch (method)
                {
                    case "GET":
                        return ApiResponse.Json(_immediate.Read());
                    case "POST":
                        var body = Body(request);
                        return ApiResponse.Json(_immediate.Run(request.Client, Text(body, "command"), OptionalInt(body, "waitMs")));
                    default:
                        throw MethodNotAllowed("GET, POST");
                }
            }

            // /projects/{p}/diagnostics and /projects/{p}/modules/{m}/diagnostics
            if (segments.Length >= 3 && segments[0] == "projects"
                && ((segments.Length == 3 && segments[2] == "diagnostics")
                    || (segments.Length == 5 && segments[2] == "modules" && segments[4] == "diagnostics")))
            {
                Allow(method, "GET");
                var module = segments.Length == 5 ? segments[3] : Query(request, "module");
                var limit = Query(request, "limit");
                if (limit != null && !int.TryParse(limit, out _))
                {
                    throw ApiException.BadRequest("limit must be a whole number.");
                }

                return ApiResponse.Json(_diagnostics.Get(segments[1], module, Query(request, "minSeverity"), limit == null ? (int?)null : int.Parse(limit)));
            }

            // /projects/{p}/resources[/{name}]
            if (segments.Length >= 3 && segments.Length <= 4 && segments[0] == "projects" && segments[2] == "resources")
            {
                if (_resources == null)
                {
                    throw ApiException.NotFound("No such route.");
                }

                if (segments.Length == 3)
                {
                    switch (method)
                    {
                        case "GET":
                            return ApiResponse.Json(_resources.List(segments[1]));
                        case "DELETE":
                            //Discards what is staged and not yet in the file; nothing already in the file is touched.
                            return ApiResponse.Json(_resources.DiscardPending(request.Client, segments[1]));
                        default:
                            throw MethodNotAllowed("GET, DELETE");
                    }
                }

                return Resource(request, method, segments[1], segments[3]);
            }

            if (segments.Length >= 3 && segments[0] == "projects")
            {
                var routed = RouteProject(request, method, segments);
                if (routed != null)
                {
                    return routed;
                }
            }

            // /projects/{p}/modules[/{m}]
            if (segments.Length >= 3 && segments.Length <= 4 && segments[0] == "projects" && segments[2] == "modules")
            {
                var project = segments[1];
                if (segments.Length == 3)
                {
                    switch (method)
                    {
                        case "GET":
                            return ApiResponse.Json(new { modules = _api.Modules(project) });
                        case "POST":
                            var created = Body(request);
                            var info = _api.Create(request.Client, project, Text(created, "name"), Text(created, "kind"), Text(created, "code"));
                            if (created["attributes"] is JObject newAttributes)
                            {
                                info = _api.WriteAttributes(request.Client, project, info.Name, newAttributes, info.Hash);
                            }

                            return new ApiResponse { Status = 201, Body = info, ETag = info.Hash };
                        default:
                            throw MethodNotAllowed("GET, POST");
                    }
                }

                return Module(request, method, project, segments[3]);
            }

            throw ApiException.NotFound("No such route.");
        }

        //What Rubberduck knows about a project beyond its code: /projects/{p}/<what>. Null if the route is not one of these.
        private ApiResponse RouteProject(ApiRequest request, string method, string[] segments)
        {
            var project = segments[1];
            var what = segments[2];
            var rest = segments.Length == 3 ? null : segments[3];

            switch (what)
            {
                case "parser-errors" when segments.Length == 3 && _insight != null:
                    Allow(method, "GET");
                    return ApiResponse.Json(_insight.ParserErrors(project));

                case "symbols" when segments.Length == 3 && _insight != null:
                    Allow(method, "GET");
                    return ApiResponse.Json(_insight.Symbols(project, Query(request, "module"), Flag(request, "includeLocals"), Query(request, "kind")));

                case "symbols" when segments.Length == 4 && rest == "references" && _insight != null:
                    Allow(method, "GET");
                    return ApiResponse.Json(_insight.References(project, Query(request, "module"), Query(request, "name"), QueryInt(request, "limit")));

                case "symbols" when segments.Length == 4 && rest == "definition" && _insight != null:
                    Allow(method, "GET");
                    return ApiResponse.Json(_insight.Definition(project, Query(request, "module"), QueryInt(request, "line") ?? 0, QueryInt(request, "column") ?? 0));

                case "libraries" when segments.Length == 3 && _insight != null:
                    Allow(method, "GET");
                    return ApiResponse.Json(new { libraries = _insight.Libraries(project) });

                case "structure" when segments.Length == 3 && _insight != null:
                    Allow(method, "GET");
                    return ApiResponse.Json(new { folders = _insight.Structure(project) });

                case "metrics" when segments.Length == 3 && _insight != null:
                    Allow(method, "GET");
                    return ApiResponse.Json(new { metrics = _insight.Metrics(project, Query(request, "module")) });

                case "tests" when segments.Length == 3 && _tests != null:
                    Allow(method, "GET");
                    return ApiResponse.Json(new { tests = _tests.List(project, Query(request, "module")) });

                case "tests" when segments.Length == 4 && rest == "run" && _tests != null:
                    Allow(method, "POST");
                    var run = Body(request);
                    return ApiResponse.Json(_tests.Run(request.Client, project, Text(run, "module"), Text(run, "test"), OptionalInt(run, "waitSeconds")));

                case "check" when segments.Length == 3 && _check != null:
                    Allow(method, "POST");
                    var check = string.IsNullOrWhiteSpace(request.Body) ? new JObject() : Body(request);
                    return ApiResponse.Json(_check.Check(request.Client, project, Text(check, "module"), Bool(check, "runTests")));

                case "edits" when segments.Length == 3:
                    Allow(method, "POST");
                    var batch = Body(request);
                    var edits = batch["edits"] as JArray ?? throw ApiException.BadRequest("'edits' is required and must be a list.");
                    List<ModuleEdit> list;
                    try
                    {
                        list = edits.Select(e => e.ToObject<ModuleEdit>()).ToList();
                    }
                    catch (JsonException)
                    {
                        throw ApiException.BadRequest("Every edit must be an object with string fields and whole-number lines.");
                    }

                    return ApiResponse.Json(_api.Batch(request.Client, project, list));

                default:
                    return null;
            }
        }

        private static bool Flag(ApiRequest request, string name)
            => bool.TryParse(Query(request, name), out var value) && value;

        private static bool Bool(JObject body, string name)
        {
            var token = body[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return false;
            }

            return token.Type == JTokenType.Boolean ? token.Value<bool>() : throw ApiException.BadRequest($"'{name}' must be true or false.");
        }

        private static int? QueryInt(ApiRequest request, string name)
        {
            var text = Query(request, name);
            if (text == null)
            {
                return null;
            }

            return int.TryParse(text, out var value) ? value : throw ApiException.BadRequest($"{name} must be a whole number.");
        }

        private ApiResponse Resource(ApiRequest request, string method, string project, string name)
        {
            switch (method)
            {
                case "GET":
                    var content = _resources.Read(project, name);
                    return ApiResponse.Json(content, content.Hash);

                case "PUT":
                    var put = Body(request);
                    return ApiResponse.Json(_resources.Write(request.Client, project, name, Text(put, "xml"), Text(put, "text"), Text(put, "contentBase64"), ExpectedHashOrNull(request, put)));

                case "DELETE":
                    return ApiResponse.Json(_resources.Delete(request.Client, project, name, request.IfMatch?.Trim('"')));

                default:
                    throw MethodNotAllowed("GET, PUT, DELETE");
            }
        }

        //A new resource needs no hash, so unlike for modules its absence is not an error here.
        private static string ExpectedHashOrNull(ApiRequest request, JObject body)
            => request.IfMatch?.Trim('"') ?? Text(body, "expectedHash");

        private ApiResponse Module(ApiRequest request, string method, string project, string module)
        {
            switch (method)
            {
                case "GET":
                    var include = request.Query.TryGetValue("include", out var what) ? what : string.Empty;
                    var wanted = include.Split(',').Select(w => w.Trim()).ToList();
                    var content = _api.Read(project, module, wanted.Contains("text", StringComparer.OrdinalIgnoreCase), wanted.Contains("attributes", StringComparer.OrdinalIgnoreCase));
                    return ApiResponse.Json(content, content.Hash);

                case "PUT":
                    var put = Body(request);
                    var hash = ExpectedHash(request, put);
                    ModuleInfo replaced;
                    var attributes = put["attributes"] as JObject;
                    if (put["attributes"] != null && attributes == null)
                    {
                        throw ApiException.BadRequest("'attributes' must be an object.");
                    }

                    if (put["text"] != null && attributes != null)
                    {
                        throw ApiException.BadRequest("'text' already holds the attributes; send either 'text' or 'code' with 'attributes'.");
                    }

                    if (put["text"] != null)
                    {
                        replaced = _api.ReplaceText(request.Client, project, module, Text(put, "text"), hash);
                    }
                    else if (put["code"] != null || attributes == null)
                    {
                        replaced = _api.ReplaceCode(request.Client, project, module, Text(put, "code"), hash);
                        if (attributes != null)
                        {
                            replaced = _api.WriteAttributes(request.Client, project, module, attributes, replaced.Hash);
                        }
                    }
                    else
                    {
                        replaced = _api.WriteAttributes(request.Client, project, module, attributes, hash);
                    }

                    return ApiResponse.Json(replaced, replaced.Hash);

                case "PATCH":
                    var patch = Body(request);
                    var edited = _api.EditLines(request.Client, project, module, Int(patch, "startLine"), Int(patch, "endLine"),
                        Text(patch, "replacement"), ExpectedHash(request, patch));
                    return ApiResponse.Json(edited, edited.Hash);

                case "DELETE":
                    _api.Delete(request.Client, project, module, request.IfMatch?.Trim('"'));
                    return new ApiResponse { Status = 204 };

                default:
                    throw MethodNotAllowed("GET, PUT, PATCH, DELETE");
            }
        }

        private static JObject Body(ApiRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Body))
            {
                throw ApiException.BadRequest("A JSON body is required.");
            }

            return JToken.Parse(request.Body) as JObject ?? throw ApiException.BadRequest("The body must be a JSON object.");
        }

        private static string ExpectedHash(ApiRequest request, JObject body)
            => request.IfMatch?.Trim('"') ?? Text(body, "expectedHash");

        private static string Text(JObject body, string name)
        {
            var token = body[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            return token.Type == JTokenType.String ? token.Value<string>() : throw ApiException.BadRequest($"'{name}' must be a string.");
        }

        private static string Query(ApiRequest request, string name) => request.Query.TryGetValue(name, out var value) ? value : null;

        private static int? OptionalInt(JObject body, string name)
        {
            var token = body[name];
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            return token.Type == JTokenType.Integer ? token.Value<int>() : throw ApiException.BadRequest($"'{name}' must be a whole number.");
        }

        private static int Int(JObject body, string name)
        {
            var token = body[name];
            if (token == null || token.Type != JTokenType.Integer)
            {
                throw ApiException.BadRequest($"'{name}' is required and must be a whole number.");
            }

            return token.Value<int>();
        }

        private static void Allow(string method, string allowed)
        {
            if (method != allowed)
            {
                throw MethodNotAllowed(allowed);
            }
        }

        private static ApiException MethodNotAllowed(string allowed)
            => new ApiException(405, "method_not_allowed", $"Allowed: {allowed}.");
    }
}
