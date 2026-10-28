using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace Rubberduck.Api
{
    /// <summary>The Immediate window, as far as the API needs it.</summary>
    public interface IImmediateWindow
    {
        /// <summary>The window's whole text.</summary>
        /// <exception cref="ApiException">The window could not be found or read.</exception>
        string Read();

        /// <summary>The window's text if it is open right now, otherwise null. Unlike <see cref="Read"/> it never shows the window, and it avoids the editor's thread, so it can be called often.</summary>
        string Peek();

        /// <summary>Types one line into the window and presses Enter, which makes the editor run it. Does not wait.</summary>
        void Send(string line);
    }

    /// <summary>
    /// The user's say over running code. Separate from <see cref="ApiWriteApproval"/>: editing code is one thing, executing it is another.
    /// </summary>
    public class ApiExecuteApproval : ApiWriteApproval
    {
        public ApiExecuteApproval(Func<string, bool> ask, Func<DateTime> now = null) : base(ask, now)
        {
        }
    }

    /// <summary>
    /// Lets an agent look at the Immediate window and enter commands in it. Entering a command <b>runs it</b>: <c>?expression</c>, a statement, or a call to any procedure of the open projects.
    /// So it is off unless the user switched it on in the settings, and the first command needs the user's approval, which is asked separately from writes.
    /// </summary>
    public class ImmediateApi : IDisposable
    {
        public const int MaxLines = 50;
        public const int MaxCharacters = 10000;
        public const int DefaultWaitMs = 5000;
        public const int MaxWaitMs = 60000;
        private const int ReadLimit = 20000;
        private const int MinOverlap = 8;

        //The window has to stay unchanged this long before the output is taken as complete.
        private static readonly TimeSpan Quiet = TimeSpan.FromMilliseconds(500);
        private static readonly TimeSpan Poll = TimeSpan.FromMilliseconds(100);

        private readonly IImmediateWindow _window;
        private readonly ApiExecuteApproval _approval;
        private readonly Func<bool> _enabled;
        private readonly Action<TimeSpan> _sleep;
        private readonly Func<TimeSpan> _clock;
        private readonly IDialogWatcher _dialogs;
        private readonly Func<string> _editorMode;

        /// <param name="editorMode">Where the editor is after a command, so a macro that stopped it is seen at once; left out, the answer has no mode.</param>
        /// <param name="dialogs">Looks for the editor's error dialogs while a command runs; without it, an error dialog is left for the user.</param>
        /// <param name="sleep">Waits between reads; replaceable so tests need not wait.</param>
        /// <param name="clock">Time since some fixed point; must move when <paramref name="sleep"/> is called.</param>
        /// <param name="enabled">Whether agents may run commands. That includes reading, whatever <paramref name="canRead"/> says.</param>
        /// <param name="canRead">Whether agents may read the window; left out, reading is allowed exactly when running is.</param>
        public ImmediateApi(IImmediateWindow window, ApiExecuteApproval approval, Func<bool> enabled, Action<TimeSpan> sleep = null, Func<TimeSpan> clock = null, IDialogWatcher dialogs = null, Func<string> editorMode = null, Func<bool> canRead = null)
        {
            _canRead = () => enabled() || (canRead?.Invoke() ?? false);
            _dialogs = dialogs;
            _editorMode = editorMode;
            _window = window;
            _approval = approval;
            _enabled = enabled;
            if (sleep == null)
            {
                var timer = Stopwatch.StartNew();
                _sleep = Thread.Sleep;
                _clock = () => timer.Elapsed;
            }
            else
            {
                _sleep = sleep;
                _clock = clock ?? throw new ArgumentNullException(nameof(clock));
            }

            _capture = new ImmediateCapture(window, _canRead, _sleep, _clock);
        }

        private readonly ImmediateCapture _capture;
        private readonly Func<bool> _canRead;

        /// <summary>Starts keeping what the window prints, so <see cref="Output"/> has it even when nobody was asking at the time.</summary>
        public void StartCapture() => _capture.Start();

        public void Dispose() => _capture.Dispose();

        /// <summary>Agents may run commands.</summary>
        public bool IsEnabled => _enabled();

        /// <summary>Agents may read the window, which is also true wherever they may run commands.</summary>
        public bool CanRead => _canRead();

        /// <summary>What was printed to the window, <c>Debug.Print</c> included, from the line number <paramref name="since"/> on (the <c>Next</c> of the last answer); without it, the newest lines.</summary>
        public ImmediateOutput Output(long? since, int? limit, string contains, int? waitMs)
        {
            RequireRead();
            return _capture.Get(since, limit, contains, waitMs ?? 0);
        }

        public ImmediateContent Read()
        {
            RequireRead();
            var text = _window.Read();
            return text.Length > ReadLimit
                ? new ImmediateContent { Text = text.Substring(text.Length - ReadLimit), Truncated = true }
                : new ImmediateContent { Text = text };
        }

        /// <param name="waitMs">How long to wait for the output to settle, at most.</param>
        public ImmediateRun Run(string client, string command, int? waitMs)
        {
            RequireEnabled();
            if (string.IsNullOrWhiteSpace(command))
            {
                throw ApiException.BadRequest("'command' is required.");
            }

            var lines = command.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n').Where(line => line.Trim().Length > 0).ToList();
            if (lines.Count > MaxLines || command.Length > MaxCharacters)
            {
                throw ApiException.BadRequest($"At most {MaxLines} lines and {MaxCharacters} characters per call.");
            }

            //The editor's Immediate window takes the ANSI code page; anything else would arrive as question marks, silently changing the command.
            var offending = command.FirstOrDefault(c => c > 0xFF || (char.IsControl(c) && c != '\n' && c != '\r' && c != '\t'));
            if (offending != default(char))
            {
                throw new ApiException(422, "unsupported_character", $"The command contains a character the Immediate window cannot take (U+{(int)offending:X4}).");
            }

            var wait = waitMs ?? DefaultWaitMs;
            if (wait < 0 || wait > MaxWaitMs)
            {
                throw ApiException.BadRequest($"waitMs must be between 0 and {MaxWaitMs}.");
            }

            _approval.Require(client);

            var before = _window.Read();
            var run = new ImmediateRun();
            var settled = true;
            var after = before;
            foreach (var line in lines)
            {
                _window.Send(line);
                //Each line gets its own wait; the next one is not typed into a window that is still busy with the last.
                settled = Settle(TimeSpan.FromMilliseconds(wait), run, out after);
                if (!settled)
                {
                    break;
                }
            }

            run.Output = Delta(before, after);
            run.Completed = settled;
            run.VbeMode = _editorMode?.Invoke();
            return run;
        }

        //Reads until the text has stopped changing, or the wait is over. A command that takes a while shows up as "not completed".
        //An error dialog stops everything until somebody clicks it, so it is looked for on every round, before the window is read.
        private bool Settle(TimeSpan wait, ImmediateRun run, out string text)
        {
            var current = _window.Read();
            var unchangedSince = _clock();
            var deadline = unchangedSince + wait;
            while (true)
            {
                _sleep(Poll);

                var dialog = _dialogs?.Find();
                if (dialog != null)
                {
                    if (dialog.IsEditorError && _dialogs.Acknowledge(dialog))
                    {
                        run.Errors.Add(dialog.Message ?? dialog.Title);
                        unchangedSince = _clock();
                        continue;
                    }

                    run.BlockedByDialog = dialog;
                    text = current;
                    return false;
                }

                var now = _window.Read();
                if (now != current)
                {
                    current = now;
                    unchangedSince = _clock();
                }
                else if (_clock() - unchangedSince >= Quiet)
                {
                    text = current;
                    return true;
                }

                if (_clock() >= deadline)
                {
                    text = current;
                    return false;
                }
            }
        }

        /// <summary>What was added to the window. The window drops its oldest lines when it fills, so the old text may not be a prefix any more.</summary>
        public static string Delta(string before, string after)
        {
            if (after.StartsWith(before, StringComparison.Ordinal))
            {
                return EndWithNewLine(after.Substring(before.Length));
            }

            //The text pattern does not always give back the same line endings, or the last one, so compare without them.
            var old = Normalize(before).TrimEnd('\n');
            var current = Normalize(after);
            if (current.StartsWith(old, StringComparison.Ordinal))
            {
                return EndWithNewLine(ToCrlf(current.Substring(old.Length).TrimStart('\n')));
            }

            //The window dropped lines from the top: the longest end of the old text that the new text begins with is what is left of it.
            //Overlaps of a few characters could be chance, so they do not count.
            for (var start = 1; old.Length - start >= MinOverlap; start++)
            {
                var length = old.Length - start;
                if (current.Length >= length && string.CompareOrdinal(current, 0, old, start, length) == 0)
                {
                    return EndWithNewLine(ToCrlf(current.Substring(length).TrimStart('\n')));
                }
            }

            return EndWithNewLine(ToCrlf(current));
        }

        private static string ToCrlf(string text) => text.Replace("\n", "\r\n");

        private static string Normalize(string text) => text.Replace("\r\n", "\n").Replace('\r', '\n');

        //Output that is joined to the next command's echo is hard to read; every answer ends its last line.
        private static string EndWithNewLine(string text)
            => text.Length == 0 || text.EndsWith("\n", StringComparison.Ordinal) || text.EndsWith("\r", StringComparison.Ordinal) ? text : text + "\r\n";

        private void RequireEnabled()
        {
            if (!_enabled())
            {
                throw ApiException.Forbidden("immediate_disabled", _canRead()
                    ? "The user has allowed reading the Immediate window, but not running commands in it. 'Also let agents run code' can be switched on in Rubberduck's settings (General)."
                    : "The user has not allowed access to the Immediate window. It can be switched on in Rubberduck's settings (General).");
            }
        }

        private void RequireRead()
        {
            if (!_canRead())
            {
                throw ApiException.Forbidden("immediate_disabled",
                    "The user has not allowed access to the Immediate window. Reading it can be switched on in Rubberduck's settings (General).");
            }
        }
    }
}
