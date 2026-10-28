using System;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.QuickFixes;

namespace Rubberduck.Api
{
    /// <summary><see cref="IQuickFixSource"/> on Rubberduck's <see cref="IQuickFixProvider"/> and the results of the latest run of <see cref="InspectorDiagnosticsSource"/>.</summary>
    public class RubberduckQuickFixSource : IQuickFixSource
    {
        private readonly RefactorHost _host;
        private readonly InspectorDiagnosticsSource _diagnostics;
        private readonly IQuickFixProvider _fixes;

        public RubberduckQuickFixSource(RefactorHost host, InspectorDiagnosticsSource diagnostics, IQuickFixProvider fixes)
        {
            _host = host;
            _diagnostics = diagnostics;
            _fixes = fixes;
        }

        public QuickFixOptions List(string project, string result)
        {
            var name = _host.CatchUp(project, out var id);
            var found = Find(id, result, out _);
            var options = new QuickFixOptions
            {
                Project = name,
                Result = result,
                Module = found.QualifiedSelection.QualifiedName.ComponentName,
                Inspection = found.Inspection.Name,
                Description = found.Description,
                StartLine = found.QualifiedSelection.Selection.StartLine
            };

            foreach (var fix in _fixes.QuickFixes(found))
            {
                options.Fixes.Add(new QuickFixOption { Name = fix.GetType().Name, Description = fix.Description(found), CanFixInModule = fix.CanFixInModule });
            }

            return options;
        }

        public QuickFixOutcome Apply(string project, string result, string fix, string scope)
        {
            var name = _host.CatchUp(project, out var id);
            var found = Find(id, result, out var all);
            var chosen = _fixes.QuickFixes(found).FirstOrDefault(f => f.GetType().Name == fix)
                ?? throw RefactorArgs.Refused($"'{fix}' is not a fix for this diagnostic.");

            var count = 1;
            _host.Run(() =>
            {
                if (scope == QuickFixApi.ScopeModule)
                {
                    var sameKind = all.Where(r => r.Inspection.GetType() == found.Inspection.GetType()
                        && r.QualifiedSelection.QualifiedName == found.QualifiedSelection.QualifiedName
                        && _fixes.CanFix(chosen, r)).ToList();
                    count = sameKind.Count;
                    _fixes.FixInModule(chosen, found.QualifiedSelection, found.Inspection.GetType(), all);
                }
                else
                {
                    _fixes.Fix(chosen, found);
                }

                return true;
            });

            return new QuickFixOutcome
            {
                Project = name,
                Result = result,
                Fix = fix,
                Scope = scope,
                Fixed = count,
                Note = "The fix was handed to Rubberduck; this does not prove it worked. Run get_diagnostics again: the diagnostic should be gone, and the module's hash has changed."
            };
        }

        private IInspectionResult Find(string projectId, string result, out IInspectionResult[] all)
        {
            all = _diagnostics.CurrentResults()
                ?? throw ApiException.Conflict("The diagnostics are out of date (the code changed or they were never run). Run get_diagnostics again, then use the new ids.");
            return all.FirstOrDefault(r => string.Equals(r.QualifiedSelection.QualifiedName.ProjectId, projectId, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(QuickFixIds.For(r), result, StringComparison.OrdinalIgnoreCase))
                ?? throw ApiException.NotFound("No diagnostic has that id in this project. Ids come from the latest get_diagnostics.");
        }
    }
}
