using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer.GitSync
{
    public enum SyncStage
    {
        ReadingModules,
        ReadingFiles,
        Comparing,
        Preparing,
        Syncing
    }

    /// <summary>Where the sync reports how far it is, and where it learns that the user gave up.</summary>
    public interface ISyncProgress
    {
        void Report(SyncStage stage, int done, int total);

        /// <summary>Raised when the user cancels. Safe to read from any thread.</summary>
        CancellationToken Token { get; }
    }

    public sealed class NoProgress : ISyncProgress
    {
        public static readonly NoProgress Instance = new NoProgress();

        public void Report(SyncStage stage, int done, int total) { }
        public CancellationToken Token => CancellationToken.None;
    }

    /// <summary>
    /// A small window with a bar and a Cancel button. The sync talks to the VBE on the thread that called it, so the window cannot have a thread
    /// of its own; <see cref="Report"/> lets it handle its messages now and then instead.
    /// </summary>
    public sealed class GitSyncProgressDialog : Form, ISyncProgress
    {
        private static readonly TimeSpan PumpInterval = TimeSpan.FromMilliseconds(80);

        private readonly CancellationTokenSource _cancellation = new CancellationTokenSource();
        private readonly Label _label;
        private readonly ProgressBar _bar;
        private DateTime _lastPump = DateTime.MinValue;

        public GitSyncProgressDialog()
        {
            Text = CodeExplorerUI.GitSync_DialogTitle;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            ControlBox = false;
            ShowInTaskbar = false;
            TopMost = true;
            ClientSize = new Size(380, 100);

            _label = new Label { Location = new Point(12, 12), Size = new Size(356, 20) };
            _bar = new ProgressBar { Location = new Point(12, 36), Size = new Size(356, 18), Style = ProgressBarStyle.Continuous };
            var cancel = new Button { Text = CodeExplorerUI.BulkRemove_Cancel, Location = new Point(288, 64), Size = new Size(80, 26) };
            cancel.Click += (s, e) =>
            {
                cancel.Enabled = false;
                _cancellation.Cancel();
            };

            Controls.AddRange(new Control[] { _label, _bar, cancel });
        }

        public CancellationToken Token => _cancellation.Token;

        public void Report(SyncStage stage, int done, int total)
        {
            if (IsDisposed)
            {
                return;
            }

            //Keep the window responsive without spending the time on repainting for each module.
            var now = DateTime.UtcNow;
            if (now - _lastPump < PumpInterval && done < total)
            {
                return;
            }

            _lastPump = now;
            _label.Text = Describe(stage, done, total);
            _bar.Maximum = Math.Max(total, 1);
            _bar.Value = Math.Min(Math.Max(done, 0), _bar.Maximum);
            Application.DoEvents();
        }

        private static string Describe(SyncStage stage, int done, int total)
        {
            string format;
            switch (stage)
            {
                case SyncStage.ReadingModules: format = CodeExplorerUI.GitSync_Progress_ReadingModules; break;
                case SyncStage.ReadingFiles: format = CodeExplorerUI.GitSync_Progress_ReadingFiles; break;
                case SyncStage.Comparing: format = CodeExplorerUI.GitSync_Progress_Comparing; break;
                case SyncStage.Preparing: format = CodeExplorerUI.GitSync_Progress_Preparing; break;
                default: format = CodeExplorerUI.GitSync_Progress_Syncing; break;
            }

            return string.Format(format, done, total);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _cancellation.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
