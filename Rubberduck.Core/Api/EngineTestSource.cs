using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.UnitTesting;

namespace Rubberduck.Api
{
    /// <summary>
    /// <see cref="ITestSource"/> on Rubberduck's test engine. The engine does not keep results where anyone can ask for them,
    /// so this listens to every run, whoever started it, and remembers the last result of each test.
    /// </summary>
    public class EngineTestSource : ITestSource
    {
        private static readonly object RunLock = new object();

        private readonly ITestEngine _engine;
        private readonly ParserSync _parser;
        private readonly IDialogWatcher _dialogs;
        private readonly ConcurrentDictionary<string, TestResult> _results = new ConcurrentDictionary<string, TestResult>(StringComparer.OrdinalIgnoreCase);

        private ManualResetEventSlim _finished;

        /// <param name="dialogs">Looks for the editor's error dialogs while tests run; without it, an error dialog is left for the user and the run waits.</param>
        public EngineTestSource(ITestEngine engine, ParserSync parser, IDialogWatcher dialogs = null)
        {
            _engine = engine;
            _parser = parser;
            _dialogs = dialogs;
            _engine.TestCompleted += (sender, args) => _results[KeyOf(args.Test)] = args.Result;
            _engine.TestRunCompleted += (sender, args) => _finished?.Set();
        }

        public IReadOnlyList<TestInfo> List(string project, string module)
        {
            var (id, _) = _parser.ResolveProject(project);
            _parser.CatchUp();
            return Select(id, module, null).Select(ToInfo).ToList();
        }

        public TestRunReport Run(string project, string module, string test, TimeSpan wait)
        {
            var (id, name) = _parser.ResolveProject(project);

            //One run at a time: the engine reports to everybody listening, and a second run would be mixed into the first.
            if (!System.Threading.Monitor.TryEnter(RunLock))
            {
                throw ApiException.Conflict("Tests are already running.");
            }

            try
            {
                _parser.CatchUp();
                var selected = Select(id, module, test).ToList();
                if (selected.Count == 0)
                {
                    if (module != null && !_parser.ModuleExists(id, module))
                    {
                        throw ApiException.NotFound($"The project has no module '{module}'.");
                    }

                    return TestsApi.Summarize(name, true, Enumerable.Empty<TestInfo>());
                }

                if (!_engine.CanRun)
                {
                    throw ApiException.Unavailable("Tests cannot run now: the parser is not ready, or the editor is not in design mode (a macro is running or stopped).");
                }

                foreach (var item in selected)
                {
                    _results.TryRemove(KeyOf(item), out _);
                }

                using (var finished = new ManualResetEventSlim(false))
                {
                    _finished = finished;
                    try
                    {
                        _engine.Run(selected);
                        var waited = DialogAwareWait.Wait(slice => finished.Wait(slice), wait, _dialogs);
                        var report = TestsApi.Summarize(name, waited.Completed, selected.Select(ToInfo));
                        report.Errors = waited.Errors;
                        report.BlockedByDialog = waited.Blocked;
                        return report;
                    }
                    finally
                    {
                        _finished = null;
                    }
                }
            }
            finally
            {
                System.Threading.Monitor.Exit(RunLock);
            }
        }

        private IEnumerable<TestMethod> Select(string projectId, string module, string test)
            => (_engine.Tests ?? Enumerable.Empty<TestMethod>())
                .Where(t => string.Equals(t.Declaration.ProjectId, projectId, StringComparison.OrdinalIgnoreCase)
                            && (module == null || string.Equals(t.Declaration.ComponentName, module, StringComparison.OrdinalIgnoreCase))
                            && (test == null || string.Equals(t.Declaration.IdentifierName, test, StringComparison.OrdinalIgnoreCase)));

        private TestInfo ToInfo(TestMethod method)
        {
            var info = new TestInfo
            {
                Module = method.Declaration.ComponentName,
                Name = method.Declaration.IdentifierName,
                Category = method.Category.Name,
                Ignored = method.IsIgnored,
                Outcome = method.IsIgnored ? TestOutcome.Ignored.ToString() : TestOutcome.Unknown.ToString()
            };

            if (_results.TryGetValue(KeyOf(method), out var result))
            {
                info.Outcome = result.Outcome.ToString();
                info.Output = result.Output;
                info.DurationMs = result.Duration;
            }

            return info;
        }

        private static string KeyOf(TestMethod method)
            => $"{method.Declaration.ProjectId}/{method.Declaration.ComponentName}/{method.Declaration.IdentifierName}";
    }
}
