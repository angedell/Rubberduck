using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Rubberduck.Api
{
    /// <summary>One managed thread of a dump, as the analyzer read it. Frames are "Namespace.Type.Method", the top of the stack first.</summary>
    public class DumpThread
    {
        public int ManagedThreadId { get; set; }

        public uint OsThreadId { get; set; }

        /// <summary>The thread the host's windows run on: the one that has to stay free for the editor to answer.</summary>
        public bool IsMain { get; set; }

        /// <summary>The exception the thread is handling, as "Type: message", or null.</summary>
        public string Exception { get; set; }

        /// <summary>What the thread is waiting for, when its top frames say so.</summary>
        public string Waiting { get; set; }

        /// <summary>The topmost Rubberduck frame, which says which feature the thread is in.</summary>
        public string RubberduckFrame { get; set; }

        public List<string> Frames { get; set; } = new List<string>();
    }

    public class DumpFinding
    {
        /// <summary>"problem" (probably why the editor hangs), "warning" or "info".</summary>
        public string Severity { get; set; }

        public string Title { get; set; }

        public string Detail { get; set; }

        /// <summary>The managed thread ids the finding is about.</summary>
        public List<int> Threads { get; set; } = new List<int>();
    }

    public class DumpAnalysis
    {
        public string File { get; set; }

        public long SizeBytes { get; set; }

        public string Runtime { get; set; }

        public string Architecture { get; set; }

        public int ThreadCount { get; set; }

        /// <summary>One sentence: what is wrong, or that nothing recognizable is.</summary>
        public string Summary { get; set; }

        public List<DumpFinding> Findings { get; set; } = new List<DumpFinding>();

        /// <summary>The threads worth a look: the main one, those handling an exception and those inside Rubberduck.</summary>
        public List<DumpThread> Threads { get; set; } = new List<DumpThread>();

        /// <summary>How many other managed threads there are (framework, finalizer and the like).</summary>
        public int OtherThreads { get; set; }

        public string Note { get; set; }

        /// <summary>Whether the dump was deleted after it was read.</summary>
        public bool Deleted { get; set; }
    }

    /// <summary>Reads a memory dump of the editor's process.</summary>
    public interface IDumpAnalyzer
    {
        /// <exception cref="ApiException">When the dump cannot be read.</exception>
        DumpAnalysis Analyze(string file, CancellationToken token);
    }

    /// <summary>
    /// Turns the threads of a dump into findings: what is blocked, on what, and who else is in the way. Knows nothing of how the threads were read,
    /// so it is tested with made-up stacks.
    /// </summary>
    public static class DumpDiagnosis
    {
        private const string RubberduckPrefix = "Rubberduck.";
        private const int WaitDepth = 8;

        // The first of these found among a thread's top frames says what it waits for. The more specific ones come first.
        private static readonly (string Frame, string Label)[] Waits =
        {
            ("System.Windows.Threading.Dispatcher.Invoke", "a Dispatcher.Invoke call, which waits for the UI thread"),
            ("System.Windows.Threading.Dispatcher.InvokeImpl", "a Dispatcher.Invoke call, which waits for the UI thread"),
            ("System.Threading.Tasks.Task.Wait", "a Task to finish"),
            ("System.Threading.Tasks.Task.InternalWait", "a Task to finish"),
            ("System.Threading.Tasks.Parallel.ForEach", "a Parallel.ForEach to finish"),
            ("System.Threading.Tasks.Parallel.For", "a Parallel.For to finish"),
            ("System.Threading.Monitor.Enter", "a lock (Monitor.Enter) held by another thread"),
            ("System.Threading.Monitor.ReliableEnter", "a lock (Monitor.Enter) held by another thread"),
            ("System.Threading.Monitor.Wait", "a Monitor.Wait signal"),
            ("System.Threading.SemaphoreSlim.Wait", "a semaphore"),
            ("System.Threading.ManualResetEventSlim.Wait", "an event"),
            ("System.Threading.WaitHandle.WaitOne", "a wait handle"),
            ("System.Threading.WaitHandle.WaitAny", "a wait handle"),
            ("System.Threading.WaitHandle.WaitAll", "wait handles"),
            ("System.Threading.Thread.Join", "another thread to end (Thread.Join)"),
            ("System.Threading.Thread.Sleep", "a sleep")
        };

        // A thread in one of these is talking to the host (Excel, the editor) through COM.
        private static readonly string[] HostCalls =
        {
            "Microsoft.Vbe.Interop.",
            "Microsoft.Office.Interop.",
            "System.Runtime.InteropServices.Marshal.",
            "System.RuntimeType.InvokeDispMethod",
            "System.Runtime.InteropServices.ComTypes.",
            "Rubberduck.VBEditor.SafeComWrappers."
        };

        private static readonly string[] BackgroundWaits = { "Task", "Parallel", "Monitor", "semaphore", "event", "wait handle", "Thread.Join" };

        public static void Fill(DumpAnalysis analysis, IEnumerable<DumpThread> allThreads)
        {
            var threads = allThreads.ToList();
            foreach (var thread in threads)
            {
                thread.Waiting = thread.Waiting ?? WaitingFor(thread);
                thread.RubberduckFrame = thread.RubberduckFrame ?? thread.Frames.FirstOrDefault(f => f.StartsWith(RubberduckPrefix));
            }

            analysis.ThreadCount = threads.Count;
            analysis.Findings = Findings(threads);
            var interesting = threads.Where(t => t.IsMain || t.Exception != null || t.RubberduckFrame != null).ToList();
            analysis.Threads = interesting;
            analysis.OtherThreads = threads.Count - interesting.Count;
            var worst = analysis.Findings.FirstOrDefault(f => f.Severity == "problem") ?? analysis.Findings.FirstOrDefault(f => f.Severity == "warning");
            analysis.Summary = worst != null
                ? worst.Title
                : "No hang pattern found in the managed threads. If the editor hung, the cause may be in native code (an OLE call or a dialog in Excel) which a managed reading cannot see.";
        }

        private static string WaitingFor(DumpThread thread)
        {
            foreach (var frame in thread.Frames.Take(WaitDepth))
            {
                foreach (var wait in Waits)
                {
                    if (frame.StartsWith(wait.Frame))
                    {
                        return wait.Label;
                    }
                }
            }

            return null;
        }

        private static bool CallsTheHost(DumpThread thread) => thread.Frames.Any(f => HostCalls.Any(f.StartsWith));

        private static string Where(DumpThread thread) => thread.RubberduckFrame ?? thread.Frames.FirstOrDefault() ?? "unknown";

        private static List<DumpFinding> Findings(IList<DumpThread> threads)
        {
            var findings = new List<DumpFinding>();
            var main = threads.FirstOrDefault(t => t.IsMain);

            if (main != null && main.Waiting != null)
            {
                findings.Add(new DumpFinding
                {
                    Severity = "problem",
                    Title = $"The main thread is blocked, waiting for {main.Waiting}.",
                    Detail = $"Managed thread {main.ManagedThreadId} runs the editor's windows, so while it waits nothing is drawn and API calls that need it time out. It was in {Where(main)}.",
                    Threads = new List<int> { main.ManagedThreadId }
                });

                var callers = threads.Where(t => !t.IsMain && CallsTheHost(t)).ToList();
                if (callers.Any())
                {
                    findings.Add(new DumpFinding
                    {
                        Severity = "problem",
                        Title = "Probable deadlock: the main thread waits while another thread calls into the host.",
                        Detail = "Calls into Excel or the editor are carried out on the main thread. A thread that makes one while the main thread waits for it (or for something it holds) cannot be answered. "
                                 + "Look at " + string.Join(", ", callers.Select(t => $"thread {t.ManagedThreadId} in {Where(t)}")) + ".",
                        Threads = new[] { main.ManagedThreadId }.Concat(callers.Select(t => t.ManagedThreadId)).ToList()
                    });
                }
                else if (BackgroundWaits.Any(w => main.Waiting.Contains(w)))
                {
                    var busy = threads.Where(t => !t.IsMain && t.RubberduckFrame != null).ToList();
                    if (busy.Any())
                    {
                        findings.Add(new DumpFinding
                        {
                            Severity = "warning",
                            Title = "The main thread is waiting on work that other threads are still doing.",
                            Detail = "Still running: " + string.Join(", ", busy.Select(t => $"thread {t.ManagedThreadId} in {Where(t)}")) + ".",
                            Threads = busy.Select(t => t.ManagedThreadId).ToList()
                        });
                    }
                }
            }

            foreach (var thread in threads.Where(t => t.Exception != null))
            {
                findings.Add(new DumpFinding
                {
                    Severity = "warning",
                    Title = $"Thread {thread.ManagedThreadId} is handling an exception: {thread.Exception}",
                    Detail = $"It was in {Where(thread)}.",
                    Threads = new List<int> { thread.ManagedThreadId }
                });
            }

            var inspections = threads
                .Select(t => (Thread: t, Frame: t.Frames.FirstOrDefault(f => f.StartsWith(RubberduckPrefix) && f.Contains(".Inspections.Concrete."))))
                .Where(x => x.Frame != null)
                .ToList();
            if (inspections.Any())
            {
                var names = inspections.Select(x => InspectionName(x.Frame)).Distinct().OrderBy(n => n).ToList();
                findings.Add(new DumpFinding
                {
                    Severity = "warning",
                    Title = $"{names.Count} inspection{(names.Count == 1 ? " is" : "s are")} still running: {string.Join(", ", names)}.",
                    Detail = "An inspection that never returns keeps the whole inspection run, and the Code Inspections window, empty. Take a second dump a minute later: one that is in the same place both times is the stuck one.",
                    Threads = inspections.Select(x => x.Thread.ManagedThreadId).ToList()
                });
            }

            var others = threads.Where(t => !t.IsMain && t.Exception == null && t.RubberduckFrame != null
                                            && !inspections.Any(x => x.Thread == t)).ToList();
            if (others.Any())
            {
                findings.Add(new DumpFinding
                {
                    Severity = "info",
                    Title = $"{others.Count} other thread{(others.Count == 1 ? " is" : "s are")} inside Rubberduck.",
                    Detail = string.Join("; ", others.Select(t => $"thread {t.ManagedThreadId}: {t.RubberduckFrame}" + (t.Waiting != null ? $" (waiting for {t.Waiting})" : string.Empty))),
                    Threads = others.Select(t => t.ManagedThreadId).ToList()
                });
            }

            return findings;
        }

        /// <summary>"Rubberduck.CodeAnalysis.Inspections.Concrete.ShadowedDeclarationInspection.DoGetInspectionResults" is "ShadowedDeclarationInspection".</summary>
        private static string InspectionName(string frame)
        {
            var parts = frame.Split('.');
            return parts.FirstOrDefault(p => p.EndsWith("Inspection")) ?? parts.Last();
        }
    }
}
