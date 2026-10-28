using System;
using System.Linq;
using System.Threading;
using Rubberduck.CodeAnalysis.Inspections;

namespace Rubberduck.Api
{
    /// <summary>
    /// <see cref="IDiagnosticsSource"/> on Rubberduck's own inspector: waits for the parser to catch up with the editor, then runs the inspections over the parser state.
    /// </summary>
    public class InspectorDiagnosticsSource : IDiagnosticsSource
    {
        //One run at a time: the inspections keep per-run state in the parse tree listeners they share.
        private static readonly object RunLock = new object();

        private readonly IInspector _inspector;
        private readonly ParserSync _parser;

        public InspectorDiagnosticsSource(IInspector inspector, ParserSync parser)
        {
            _inspector = inspector;
            _parser = parser;
        }

        public DiagnosticsReport Run(string project, string module)
        {
            var (id, name) = _parser.ResolveProject(project);

            lock (RunLock)
            {
                var finalState = _parser.CatchUp();
                IInspectionResult[] results;
                try
                {
                    results = _inspector.FindIssuesAsync(_parser.State, CancellationToken.None).GetAwaiter().GetResult().ToArray();
                }
                catch (OperationCanceledException)
                {
                    throw ApiException.Unavailable("The inspections were cancelled. Try again.");
                }

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
}
