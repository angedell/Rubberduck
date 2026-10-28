using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Api
{
    /// <summary>How a wait for running code ended.</summary>
    public class DialogAwareWaitResult
    {
        /// <summary>The thing waited for happened.</summary>
        public bool Completed { get; set; }

        /// <summary>
        /// What the editor's own error dialogs said while waiting, each message once. They were closed (End or OK), so the editor is not left waiting for a click.
        /// A message that came back several times is followed by how many: "Compile error: ... (x7)".
        /// </summary>
        public List<string> Errors => _counts.Select(c => c.Value > 1 ? $"{c.Key} (x{c.Value})" : c.Key).ToList();

        private readonly List<KeyValuePair<string, int>> _counts = new List<KeyValuePair<string, int>>();

        /// <returns>How many times this message has come up, this one included.</returns>
        internal int AddError(string message)
        {
            var at = _counts.FindIndex(c => c.Key == message);
            if (at < 0)
            {
                _counts.Add(new KeyValuePair<string, int>(message, 1));
                return 1;
            }

            _counts[at] = new KeyValuePair<string, int>(message, _counts[at].Value + 1);
            return _counts[at].Value;
        }

        /// <summary>A dialog nobody could answer for the caller, such as a message box of the code that ran. The wait ended because of it.</summary>
        public BlockingDialog Blocked { get; set; }
    }

    /// <summary>
    /// Waits for code that is running, and keeps an eye on the dialogs it can raise. A compile or run-time error shows a dialog that stops everything until a person
    /// clicks it, so a plain wait would sit there until it ran out. Here the editor's error dialogs are closed and their text is kept; any other dialog ends the wait.
    /// </summary>
    public static class DialogAwareWait
    {
        private static readonly TimeSpan Slice = TimeSpan.FromMilliseconds(200);

        /// <param name="waitSlice">Waits up to the given time; true as soon as what is waited for has happened.</param>
        /// <param name="dialogs">Null means nobody looks for dialogs, and this is a plain wait.</param>
        /// <param name="clock">Time since some fixed point; replaceable so tests need not wait.</param>
        /// <param name="onError">Told of each error dialog that was closed: its text and how many times that text has come up.</param>
        public static DialogAwareWaitResult Wait(Func<TimeSpan, bool> waitSlice, TimeSpan total, IDialogWatcher dialogs, Func<TimeSpan> clock = null, Action<string, int> onError = null)
        {
            var result = new DialogAwareWaitResult();
            if (clock == null)
            {
                var timer = Stopwatch.StartNew();
                clock = () => timer.Elapsed;
            }

            var deadline = clock() + total;
            while (true)
            {
                var remaining = deadline - clock();
                if (remaining < TimeSpan.Zero)
                {
                    remaining = TimeSpan.Zero;
                }

                if (waitSlice(dialogs == null || remaining < Slice ? remaining : Slice))
                {
                    result.Completed = true;
                    return result;
                }

                var dialog = dialogs?.Find();
                if (dialog != null)
                {
                    if (dialog.IsEditorError && dialogs.Acknowledge(dialog))
                    {
                        var text = dialog.Message ?? dialog.Title;
                        var times = result.AddError(text);
                        onError?.Invoke(text, times);
                        continue;
                    }

                    result.Blocked = dialog;
                    return result;
                }

                if (clock() >= deadline)
                {
                    return result;
                }
            }
        }
    }
}
