using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Api
{
    /// <summary>What a read answers with (202) when the parser was not ready within the wait. Nothing is read yet; the catch-up goes on, so asking again is the way on.</summary>
    public class ReadyPending
    {
        public string Project { get; set; }

        /// <summary>Always "running", as for diagnostics.</summary>
        public string Status { get; set; } = "running";

        public int RetryAfterSeconds { get; set; }
    }

    /// <summary>
    /// <c>?wait=ready</c> for the reads that depend on the parser: wait, off the UI thread, for the parser to catch up with the editor, and if it is not done within
    /// <c>waitSeconds</c> (default 60, 0 to 300, as for diagnostics) say "running" instead of holding the caller. The catch-up goes on in the background;
    /// the next call joins it and, once it is done, reads at once. Without <c>wait</c> nothing changes: the read waits as it always did.
    /// </summary>
    public class ReadyWait
    {
        public const string Ready = "ready";
        public const int DefaultWaitSeconds = 60;
        public const int RetryAfterSeconds = 10;

        private static readonly ConditionalWeakTable<ParserSync, ReadyWait> Shared = new ConditionalWeakTable<ParserSync, ReadyWait>();

        //The reads that look at what the parser holds. Libraries read the project's references from the editor, and modules, resources, the log and the like do not use the parser.
        private static readonly string[] ParserReads = { "parser-errors", "symbols", "structure", "metrics" };

        private readonly Func<string, (string Id, string Name)> _resolve;
        private readonly Func<ParserState> _catchUp;
        private readonly object _gate = new object();
        private Task<ParserState> _job;

        /// <param name="resolve">Finds the project, and throws if there is none or the key is ambiguous.</param>
        /// <param name="catchUp">Asks for a parse and returns when the parser is done. Runs on a worker thread, never on the UI thread.</param>
        public ReadyWait(Func<string, (string Id, string Name)> resolve, Func<ParserState> catchUp)
        {
            _resolve = resolve;
            _catchUp = catchUp;
        }

        /// <summary>One for each parser, so the REST routes and the MCP tools share the catch-up in progress.</summary>
        public static ReadyWait For(ParserSync parser) => Shared.GetValue(parser, p => new ReadyWait(p.ResolveProject, () => p.CatchUp()));

        /// <summary>Whether <c>wait=ready</c> means anything for this request: a GET of /projects/{p}/&lt;what&gt;[/...] that reads parser results.</summary>
        public static bool Applies(string method, string[] segments)
        {
            if (method != "GET" || segments.Length < 3 || segments[0] != "projects")
            {
                return false;
            }

            return ParserReads.Contains(segments[2]) || (segments[2] == "tests" && segments.Length == 3);
        }

        /// <summary>
        /// Null if the read can go ahead: no wait was asked for, or the parser is ready. Otherwise the answer to give instead (202).
        /// A project that does not exist, a macro that is running and the like throw as they would for the read itself.
        /// </summary>
        public ReadyPending Pending(string project, string wait, int? waitSeconds)
        {
            if (wait == null)
            {
                return null;
            }

            if (!Ready.Equals(wait, StringComparison.OrdinalIgnoreCase))
            {
                throw ApiException.BadRequest($"wait must be '{Ready}'.");
            }

            if (waitSeconds < 0 || waitSeconds > DiagnosticsApi.MaxWaitSeconds)
            {
                throw ApiException.BadRequest($"waitSeconds must be between 0 and {DiagnosticsApi.MaxWaitSeconds}.");
            }

            var (_, name) = _resolve(project);

            Task<ParserState> job;
            lock (_gate)
            {
                //A catch-up that an earlier "running" answer left behind is joined, not started over.
                job = _job ?? (_job = Task.Run(_catchUp));
            }

            bool done;
            try
            {
                done = job.Wait(TimeSpan.FromSeconds(waitSeconds ?? DefaultWaitSeconds));
            }
            catch (AggregateException exception)
            {
                Forget(job);
                //Not in design mode, the parser took too long: the same errors as the read gives.
                throw exception.InnerException ?? exception;
            }

            if (!done)
            {
                return new ReadyPending { Project = name, RetryAfterSeconds = RetryAfterSeconds };
            }

            //The read that follows asks the parser to catch up once more, so edits made in between are seen as well.
            Forget(job);
            return null;
        }

        private void Forget(Task<ParserState> job)
        {
            lock (_gate)
            {
                if (_job == job)
                {
                    _job = null;
                }
            }
        }

        /// <summary>Adds the tool arguments <c>wait</c> and <c>waitSeconds</c> to a tool's properties.</summary>
        public static JObject Describe(JObject properties)
        {
            properties["wait"] = new JObject
            {
                ["type"] = "string",
                ["enum"] = new JArray(Ready),
                ["description"] = "'ready': wait for Rubberduck's parser to catch up with the latest edits, but only waitSeconds. If it is not done by then the answer has status 'running' and no results: call again."
            };
            properties["waitSeconds"] = new JObject
            {
                ["type"] = "integer",
                ["description"] = $"With wait: how long to wait for the parser (default {DefaultWaitSeconds}, at most {DiagnosticsApi.MaxWaitSeconds})."
            };
            return properties;
        }
    }
}
