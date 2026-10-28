using System.Linq;

namespace Rubberduck.Api
{
    /// <summary>"Check my work": what an agent wants after an edit, in one call. Parse errors first; then the inspections; then, if asked, the tests.</summary>
    public class CheckApi
    {
        private readonly InsightApi _insight;
        private readonly DiagnosticsApi _diagnostics;
        private readonly TestsApi _tests;

        public CheckApi(InsightApi insight, DiagnosticsApi diagnostics, TestsApi tests)
        {
            _insight = insight;
            _diagnostics = diagnostics;
            _tests = tests;
        }

        public bool CanRunTests => _tests.IsEnabled;

        public CheckReport Check(string client, string project, string module, bool runTests, int? waitSeconds = null)
        {
            var report = new CheckReport { Project = project, Module = module };
            report.ParserErrors = _insight.ParserErrors(project);
            report.Project = report.ParserErrors.Project;

            //Inspections of code that did not parse are noise; leave them out and say why through the parser errors.
            var parsed = report.ParserErrors.Errors.Count == 0;
            if (parsed)
            {
                report.Diagnostics = _diagnostics.Get(project, module, "warning", null, waitSeconds);
                report.InspectionsStatus = report.Diagnostics.Status;
                report.Suppressed = report.Diagnostics.CountsBySeverity
                    .Where(c => DiagnosticsApi.Severities.ToList().IndexOf(c.Key) < DiagnosticsApi.Severities.ToList().IndexOf("warning"))
                    .ToDictionary(c => c.Key, c => c.Value);
            }

            if (runTests && parsed)
            {
                report.Tests = _tests.Run(client, project, module, null, null);
            }
            else if (parsed)
            {
                var count = _tests.List(project, module).Count;
                if (count > 0)
                {
                    report.TestsNotRun = $"{count} unit test(s) exist and were not run. Call check_work again with runTests: true; that also shows the Test Explorer to the user.";
                }
            }

            report.Clean = parsed
                           && report.InspectionsStatus == null
                           && report.Diagnostics.Diagnostics.Count == 0
                           && (report.Tests == null || (report.Tests.Failed == 0 && report.Tests.Completed));
            return report;
        }
    }
}
