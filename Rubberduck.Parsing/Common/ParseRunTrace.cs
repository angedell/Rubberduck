using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using NLog;

namespace Rubberduck.Parsing.Common
{
    /// <summary>
    /// What one parsing run did and why: who asked for it, which modules it parsed or resolved and for what reason, and how long each step took.
    /// It is written as one entry at the end of the run, so that a slow run can be explained from the log alone.
    /// </summary>
    public sealed class ParseRunTrace
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private const int MaxNames = 8;

        private readonly Stopwatch _total = Stopwatch.StartNew();
        private readonly string _requestor;
        private readonly object _lock = new object();
        private readonly List<(string Name, long Milliseconds)> _steps = new List<(string, long)>();
        private readonly List<string> _facts = new List<string>();

        public ParseRunTrace(object requestor)
        {
            _requestor = requestor?.GetType().Name ?? "unknown";
        }

        /// <summary>Times everything up to the disposal of the returned scope. A step that runs several times is added up, with its count.</summary>
        public IDisposable Step(string name) => new StepScope(this, name);

        /// <summary>Something that explains the run, as a sentence ("Referenced libraries: none loaded, 1 unloaded").</summary>
        public void Fact(string text)
        {
            lock (_lock)
            {
                _facts.Add(text);
            }
        }

        /// <summary>A fact about a set of things: how many, and the first few by name.</summary>
        public void Fact(string what, IReadOnlyCollection<string> names)
        {
            Fact(Describe(what, names));
        }

        /// <summary>Counts per reason, with the first few modules of each: "Modules to parse: 3 (2 content changed: A, B; 1 new: C)".</summary>
        public void Fact(string what, IEnumerable<KeyValuePair<string, string>> nameAndReason)
        {
            var list = nameAndReason.ToList();
            if (list.Count == 0)
            {
                Fact($"{what}: none");
                return;
            }

            var groups = list.GroupBy(pair => pair.Value)
                .OrderByDescending(group => group.Count())
                .Select(group => $"{group.Count()} {group.Key}: {Names(group.Select(pair => pair.Key).ToList())}");
            Fact($"{what}: {list.Count} ({string.Join("; ", groups)})");
        }

        public void Finish(string outcome)
        {
            _total.Stop();
            var text = new StringBuilder();
            text.Append($"Parsing run requested by {_requestor} {outcome} after {_total.ElapsedMilliseconds} ms.");

            lock (_lock)
            {
                foreach (var fact in _facts)
                {
                    text.Append("\n  ").Append(fact);
                }

                var steps = _steps.GroupBy(step => step.Name)
                    .Select(group => (Name: group.Key, Count: group.Count(), Total: group.Sum(step => step.Milliseconds)))
                    .OrderByDescending(step => step.Total)
                    .ToList();
                var timed = steps.Sum(step => step.Total);
                text.Append("\n  Steps (longest first): ")
                    .Append(string.Join("; ", steps.Select(step => step.Count > 1 ? $"{step.Name} {step.Total} ms (x{step.Count})" : $"{step.Name} {step.Total} ms")))
                    .Append($"; not in any step {Math.Max(0, _total.ElapsedMilliseconds - timed)} ms.");
            }

            Logger.Info(text.ToString());
        }

        private static string Describe(string what, IReadOnlyCollection<string> names)
            => names.Count == 0 ? $"{what}: none" : $"{what}: {names.Count} ({Names(names)})";

        private static string Names(IReadOnlyCollection<string> names)
            => string.Join(", ", names.Take(MaxNames)) + (names.Count > MaxNames ? $", and {names.Count - MaxNames} more" : string.Empty);

        private void Add(string name, long milliseconds)
        {
            lock (_lock)
            {
                _steps.Add((name, milliseconds));
            }
        }

        private sealed class StepScope : IDisposable
        {
            private readonly ParseRunTrace _trace;
            private readonly string _name;
            private readonly Stopwatch _watch = Stopwatch.StartNew();

            public StepScope(ParseRunTrace trace, string name)
            {
                _trace = trace;
                _name = name;
            }

            public void Dispose()
            {
                _watch.Stop();
                _trace.Add(_name, _watch.ElapsedMilliseconds);
            }
        }
    }
}
