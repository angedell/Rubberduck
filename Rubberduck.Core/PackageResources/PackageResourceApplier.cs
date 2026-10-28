using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using NLog;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.PackageResources
{
    /// <summary>
    /// Writes staged resource changes into the document files once the applications let go of them: when a project is closed, when Rubberduck starts
    /// (for changes that were still waiting when it last stopped), and, as a last try, when it shuts down.
    /// A document that is still open just stays staged; nothing is ever forced.
    /// </summary>
    public sealed class PackageResourceApplier : IDisposable
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly TimeSpan RetryEvery = TimeSpan.FromSeconds(2);
        private static readonly TimeSpan ClosingPatience = TimeSpan.FromMinutes(2);
        private static readonly TimeSpan ShutdownPatience = TimeSpan.FromSeconds(5);

        private readonly IPackageResourceService _resources;
        private readonly IVbeEvents _vbeEvents;
        private readonly CancellationTokenSource _stopped = new CancellationTokenSource();
        private int _running;
        private bool _started;

        public PackageResourceApplier(IPackageResourceService resources, IVbeEvents vbeEvents)
        {
            _resources = resources;
            _vbeEvents = vbeEvents;
        }

        public void Start()
        {
            if (_started)
            {
                return;
            }

            _started = true;
            _vbeEvents.ProjectRemoved += ProjectRemoved;
            Task.Run(() => ApplyPending(TimeSpan.Zero, _stopped.Token));
        }

        /// <summary>One last try, for documents the application has already closed. Blocks for a few seconds at most.</summary>
        public void Shutdown()
        {
            _stopped.Cancel();
            Detach();
            ApplyPending(ShutdownPatience, CancellationToken.None);
        }

        private void ProjectRemoved(object sender, ProjectEventArgs e)
            => Task.Run(() => ApplyPending(ClosingPatience, _stopped.Token));

        /// <summary>Tries every document with staged changes until none is busy or the patience runs out.</summary>
        public void ApplyPending(TimeSpan patience, CancellationToken cancel)
        {
            //One pass at a time; a project closing while another pass waits is picked up by that pass.
            if (Interlocked.CompareExchange(ref _running, 1, 0) != 0)
            {
                return;
            }

            try
            {
                var deadline = DateTime.UtcNow + patience;
                while (!cancel.IsCancellationRequested)
                {
                    var busy = _resources.PendingDocuments().Count(document =>
                    {
                        var outcome = _resources.TryApply(document, out var message);
                        if (outcome == PackageApplyOutcome.Failed)
                        {
                            Logger.Warn("Staged resource changes for {0} could not be applied: {1}", document, message);
                        }

                        return outcome == PackageApplyOutcome.Busy;
                    });

                    if (busy == 0 || DateTime.UtcNow + RetryEvery > deadline)
                    {
                        return;
                    }

                    cancel.WaitHandle.WaitOne(RetryEvery);
                }
            }
            catch (Exception exception)
            {
                Logger.Warn(exception, "Applying staged resource changes failed.");
            }
            finally
            {
                Interlocked.Exchange(ref _running, 0);
            }
        }

        private void Detach()
        {
            if (_started)
            {
                _vbeEvents.ProjectRemoved -= ProjectRemoved;
                _started = false;
            }
        }

        public void Dispose()
        {
            Detach();
            _stopped.Cancel();
            _stopped.Dispose();
        }
    }
}
