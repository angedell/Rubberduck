using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Api
{
    /// <summary>Where diagnostics come from. Waits for the parser to catch up with the code, so the answer reflects what is in the editor now.</summary>
    public interface IDiagnosticsSource
    {
        /// <param name="project">Project id or name.</param>
        /// <param name="module">A module's name, or null for the whole project.</param>
        /// <param name="wait">How long to wait for the inspections; if they are not done by then the report says "running" and they carry on in the background. Null: a default.</param>
        DiagnosticsReport Run(string project, string module, TimeSpan? wait = null);
    }

    /// <summary>Rubberduck's code inspections, for agents to check what they wrote. Read-only, so it needs no approval.</summary>
    public class DiagnosticsApi
    {
        public const int DefaultLimit = 200;
        public const int MaxLimit = 1000;

        //Least to most severe; a minimum severity keeps everything from its position on.
        public static readonly IReadOnlyList<string> Severities = new[] { "hint", "suggestion", "warning", "error" };

        private readonly IDiagnosticsSource _source;

        public DiagnosticsApi(IDiagnosticsSource source)
        {
            _source = source;
        }

        public const int MaxWaitSeconds = 300;

        public DiagnosticsReport Get(string project, string module, string minSeverity, int? limit, int? waitSeconds = null)
        {
            if (waitSeconds < 0 || waitSeconds > MaxWaitSeconds)
            {
                throw ApiException.BadRequest($"waitSeconds must be between 0 and {MaxWaitSeconds}.");
            }

            var minimum = 0;
            if (!string.IsNullOrWhiteSpace(minSeverity))
            {
                minimum = Severities.ToList().FindIndex(s => s.Equals(minSeverity, StringComparison.OrdinalIgnoreCase));
                if (minimum < 0)
                {
                    throw ApiException.BadRequest($"minSeverity must be one of: {string.Join(", ", Severities)}.");
                }
            }

            if (limit < 1 || limit > MaxLimit)
            {
                throw ApiException.BadRequest($"limit must be between 1 and {MaxLimit}.");
            }

            var report = _source.Run(project, string.IsNullOrWhiteSpace(module) ? null : module, waitSeconds.HasValue ? TimeSpan.FromSeconds(waitSeconds.Value) : (TimeSpan?)null);
            report.CountsBySeverity = Severities
                .Select(s => (Severity: s, Count: report.Diagnostics.Count(d => s.Equals(d.Severity, StringComparison.OrdinalIgnoreCase))))
                .Where(x => x.Count > 0)
                .ToDictionary(x => x.Severity, x => x.Count);
            var kept = report.Diagnostics
                .Where(d => Severities.ToList().FindIndex(s => s.Equals(d.Severity, StringComparison.OrdinalIgnoreCase)) >= minimum)
                //Worst first, then in reading order, so a cut list loses the least important.
                .OrderByDescending(d => Severities.ToList().FindIndex(s => s.Equals(d.Severity, StringComparison.OrdinalIgnoreCase)))
                .ThenBy(d => d.Module, StringComparer.OrdinalIgnoreCase)
                .ThenBy(d => d.StartLine)
                .ThenBy(d => d.StartColumn)
                .ToList();

            var max = limit ?? DefaultLimit;
            report.Total = kept.Count;
            report.Truncated = kept.Count > max;
            report.Diagnostics = kept.Take(max).ToList();
            return report;
        }
    }
}
