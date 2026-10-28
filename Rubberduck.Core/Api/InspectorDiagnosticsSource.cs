using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using NLog;
using Rubberduck.CodeAnalysis.Inspections;

namespace Rubberduck.Api
{
    /// <summary>
    /// <see cref="IDiagnosticsSource"/> on Rubberduck's own inspector: waits for the parser to catch up with the editor, then runs the inspections over the parser state.
    /// The inspector looks at everything the parser holds, every open project, so on a large one it takes minutes. The run is therefore a background job whose results
    /// are kept for as long as the parser state stays the same: a caller waits for a while, is told "running" if that is not enough, and gets the results at once next time.
    /// </summary>
    public class InspectorDiagnosticsSource : IDiagnosticsSource
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly TimeSpan DefaultWait = TimeSpan.FromSeconds(60);

        //One run at a time: the inspections keep per-run state in the parse tree listeners they share.
        private static readonly object RunLock = new object();

        private readonly IInspector _inspector;
        private readonly ParserSync _parser;
        private readonly object _gate = new object();

        private int _cachedGeneration = -1;
        private IInspectionResult[] _cached;
        private int _jobGeneration = -1;
        private Task<IInspectionResult[]> _job;

        public InspectorDiagnosticsSource(IInspector inspector, ParserSync parser)
        {
            _inspector = inspector;
            _parser = parser;
        }

        public DiagnosticsReport Run(string project, string module, TimeSpan? wait = null)
        {
            var (id, name) = _parser.ResolveProject(project);
            var finalState = _parser.CatchUp();
            var generation = _parser.Generation;

            Task<IInspectionResult[]> job;
            lock (_gate)
            {
                if (_cachedGeneration == generation)
                {
                    return Report(_cached, id, name, module, finalState);
                }

                if (_job == null || _jobGeneration != generation || _job.IsFaulted || _job.IsCanceled)
                {
                    _jobGeneration = generation;
                    _job = Task.Run(() => Inspect(generation));
                }

                job = _job;
            }

            IInspectionResult[] results;
            try
            {
                if (!job.Wait(wait ?? DefaultWait))
                {
                    return new DiagnosticsReport { Project = name, Module = module, ParserState = finalState.ToString(), Status = "running", RetryAfterSeconds = 10 };
                }

                results = job.Result;
            }
            catch (AggregateException exception) when (exception.InnerException is OperationCanceledException)
            {
                throw ApiException.Unavailable("The inspections were cancelled. Try again.");
            }

            return Report(results, id, name, module, finalState);
        }

        private IInspectionResult[] Inspect(int generation)
        {
            IInspectionResult[] results;
            lock (RunLock)
            {
                var timer = System.Diagnostics.Stopwatch.StartNew();
                Logger.Info($"API inspections started (parser state {generation}).");
                results = _inspector.FindIssuesAsync(_parser.State, CancellationToken.None).GetAwaiter().GetResult().ToArray();
                Logger.Info($"API inspections finished after {(int)timer.Elapsed.TotalSeconds} s: {results.Length} results.");
            }

            lock (_gate)
            {
                if (generation > _cachedGeneration)
                {
                    _cachedGeneration = generation;
                    _cached = results;
                }
            }

            return results;
        }

        private DiagnosticsReport Report(IInspectionResult[] results, string id, string name, string module, Rubberduck.Parsing.VBA.ParserState finalState)
        {
            var report = new DiagnosticsReport { Project = name, Module = module, ParserState = finalState.ToString() };
            foreach (var result in results)
            {
                var where = result.QualifiedSelection.QualifiedName;
                if (!string.Equals(where.ProjectId, id, StringComparison.OrdinalIgnoreCase)
                    || (module != null && !string.Equals(where.ComponentName, module, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                var selection = result.QualifiedSelection.Selection;
                report.Diagnostics.Add(new Diagnostic
                {
                    Module = string.IsNullOrEmpty(where.ComponentName) ? null : where.ComponentName,
                    Inspection = result.Inspection.Name,
                    AnnotationName = result.Inspection.AnnotationName,
                    Severity = result.Inspection.Severity.ToString().ToLowerInvariant(),
                    Type = result.Inspection.InspectionType.ToString(),
                    Description = result.Description,
                    StartLine = selection.StartLine,
                    StartColumn = selection.StartColumn,
                    EndLine = selection.EndLine,
                    EndColumn = selection.EndColumn
                });
            }

            if (module != null && report.Diagnostics.Count == 0 && !_parser.ModuleExists(id, module))
            {
                throw ApiException.NotFound($"The project has no module '{module}'.");
            }

            return report;
        }
    }
}
