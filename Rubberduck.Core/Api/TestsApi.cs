using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Api
{
    public interface ITestSource
    {
        IReadOnlyList<TestInfo> List(string project, string module);

        /// <summary>Runs the matching tests and waits up to <paramref name="wait"/> for them to finish.</summary>
        TestRunReport Run(string project, string module, string test, TimeSpan wait);
    }

    /// <summary>
    /// Rubberduck's unit tests for an agent. Listing only reads. Running executes the user's VBA, so it needs the same opt-in
    /// as the Immediate window and the same one-time approval.
    /// </summary>
    public class TestsApi
    {
        public static readonly TimeSpan DefaultWait = TimeSpan.FromSeconds(60);
        public static readonly TimeSpan MaxWait = TimeSpan.FromMinutes(5);

        private readonly ITestSource _source;
        private readonly ApiExecuteApproval _approval;
        private readonly Func<bool> _enabled;

        public TestsApi(ITestSource source, ApiExecuteApproval approval, Func<bool> enabled)
        {
            _source = source;
            _approval = approval;
            _enabled = enabled;
        }

        /// <summary>Whether running code is switched on in the settings. Listing tests does not depend on it.</summary>
        public bool IsEnabled => _enabled();

        public IReadOnlyList<TestInfo> List(string project, string module)
            => _source.List(project, string.IsNullOrWhiteSpace(module) ? null : module);

        public TestRunReport Run(string client, string project, string module, string test, int? waitSeconds)
        {
            if (!_enabled())
            {
                throw ApiException.Forbidden("execute_disabled", "Running code is switched off. The user can allow it in Rubberduck's settings.");
            }

            if (waitSeconds < 1 || waitSeconds > MaxWait.TotalSeconds)
            {
                throw ApiException.BadRequest($"wait must be between 1 and {(int)MaxWait.TotalSeconds} seconds.");
            }

            if (!string.IsNullOrWhiteSpace(test) && string.IsNullOrWhiteSpace(module))
            {
                throw ApiException.BadRequest("'module' is required when 'test' is given.");
            }

            _approval.Require(client);
            return _source.Run(project,
                string.IsNullOrWhiteSpace(module) ? null : module,
                string.IsNullOrWhiteSpace(test) ? null : test,
                waitSeconds.HasValue ? TimeSpan.FromSeconds(waitSeconds.Value) : DefaultWait);
        }

        public static TestRunReport Summarize(string project, bool completed, IEnumerable<TestInfo> tests)
        {
            var list = tests.ToList();
            return new TestRunReport
            {
                Project = project,
                Completed = completed,
                Succeeded = list.Count(t => t.Outcome == "Succeeded"),
                Failed = list.Count(t => t.Outcome == "Failed"),
                Inconclusive = list.Count(t => t.Outcome == "Inconclusive"),
                Ignored = list.Count(t => t.Outcome == "Ignored"),
                Tests = list
            };
        }
    }
}
