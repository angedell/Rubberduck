using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.Diagnostics.Runtime;

namespace Rubberduck.Api
{
    /// <summary>
    /// Reads a full dump of the editor's process with ClrMD. The dump is read where it was taken, so the CLR's own debugging library (mscordacwks / mscordaccore)
    /// is the one on this machine; on another machine ClrMD would look for it on the Microsoft symbol server instead.
    /// </summary>
    public class ClrMdDumpAnalyzer : IDumpAnalyzer
    {
        private const int MaxFrames = 40;

        public DumpAnalysis Analyze(string file, CancellationToken token)
        {
            try
            {
                using (var target = DataTarget.LoadDump(file))
                {
                    var analysis = new DumpAnalysis
                    {
                        File = file,
                        Architecture = target.DataReader.Architecture.ToString()
                    };

                    var version = target.ClrVersions.FirstOrDefault();
                    if (version == null)
                    {
                        analysis.Note = "The dump holds no .NET runtime, so there are no managed threads to read. It may be a mini dump of a process that had not loaded Rubberduck, or the wrong process.";
                        analysis.Summary = analysis.Note;
                        return analysis;
                    }

                    analysis.Runtime = version.Version.ToString();
                    using (var runtime = version.CreateRuntime())
                    {
                        var threads = new List<DumpThread>();
                        foreach (var thread in runtime.Threads.Where(t => t.IsAlive))
                        {
                            token.ThrowIfCancellationRequested();
                            threads.Add(Read(thread));
                        }

                        MarkMainThread(threads);
                        DumpDiagnosis.Fill(analysis, threads);
                    }

                    return analysis;
                }
            }
            catch (OperationCanceledException)
            {
                throw ApiException.Unavailable("Reading the dump took longer than the limit and was stopped.");
            }
            catch (ApiException)
            {
                throw;
            }
            catch (Exception exception)
            {
                var cause = exception.GetBaseException();
                throw ApiException.Unavailable("The dump could not be read: " + exception.Message + (cause == exception ? string.Empty : " (" + cause.Message + ")")
                    + " Reading a managed stack needs a full dump ('kind': 'full') taken by the same .NET runtime that is installed here, and a debugging library for it (mscordacwks.dll or mscordaccore.dll).");
            }
        }

        private static DumpThread Read(ClrThread thread)
        {
            var frames = new List<string>();
            foreach (var frame in thread.EnumerateStackTrace().Take(MaxFrames))
            {
                var method = frame.Method;
                frames.Add(method == null ? frame.FrameName ?? "[native]" : $"{method.Type?.Name}.{method.Name}");
            }

            var exception = thread.CurrentException;
            return new DumpThread
            {
                ManagedThreadId = thread.ManagedThreadId,
                OsThreadId = thread.OSThreadId,
                Exception = exception == null ? null : $"{exception.Type?.Name}: {exception.Message}",
                Frames = frames
            };
        }

        /// <summary>
        /// The editor's main thread is managed thread 1, the first the runtime saw (the host's own pump is native, so it loads the add-in there). Failing that,
        /// a thread pumping a WPF dispatcher, which is what a dialog on the main thread looks like.
        /// </summary>
        private static void MarkMainThread(IList<DumpThread> threads)
        {
            var main = threads.FirstOrDefault(t => t.ManagedThreadId == 1)
                       ?? threads.FirstOrDefault(t => t.Frames.Any(f => f.StartsWith("System.Windows.Threading.Dispatcher.Run") || f.StartsWith("System.Windows.Threading.Dispatcher.PushFrame")));
            if (main != null)
            {
                main.IsMain = true;
            }
        }
    }
}
