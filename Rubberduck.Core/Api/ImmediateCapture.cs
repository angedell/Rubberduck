using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using NLog;

namespace Rubberduck.Api
{
    /// <summary>
    /// Keeps what the Immediate window prints, <c>Debug.Print</c> included, whether or not an agent was asking at the time.
    /// The editor has no event for it, so the window is looked at twice a second and what was added since the last look is kept as numbered lines.
    /// An agent asks for the lines from the number it saw last. The buffer holds the newest <see cref="Capacity"/> lines.
    /// </summary>
    public class ImmediateCapture : IDisposable
    {
        public const int Capacity = 5000;
        public const int DefaultLimit = 200;
        public const int MaxLimit = 1000;
        public const int MaxWaitMs = 30000;

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly TimeSpan Interval = TimeSpan.FromMilliseconds(500);
        private static readonly TimeSpan Poll = TimeSpan.FromMilliseconds(100);

        //While the window is closed the editor is not asked again for this many rounds; looking for it costs the editor's thread.
        private const int ClosedBackoffRounds = 10;

        private readonly IImmediateWindow _window;
        private readonly Func<bool> _enabled;
        private readonly Action<TimeSpan> _sleep;
        private readonly Func<TimeSpan> _clock;

        private readonly object _pollLock = new object();
        private readonly object _bufferLock = new object();
        private readonly LinkedList<(long Number, string Text)> _lines = new LinkedList<(long, string)>();

        private long _next;
        private string _baseline;
        private string _tailSeen = string.Empty;
        private int _backoff;
        private Timer _timer;

        public ImmediateCapture(IImmediateWindow window, Func<bool> enabled, Action<TimeSpan> sleep, Func<TimeSpan> clock)
        {
            _window = window;
            _enabled = enabled;
            _sleep = sleep;
            _clock = clock;
        }

        /// <summary>Starts looking at the window in the background. Does nothing while the user has not allowed agents the Immediate window.</summary>
        public void Start()
        {
            lock (_bufferLock)
            {
                _timer = _timer ?? new Timer(_ => Tick(), null, Interval, Timeout.InfiniteTimeSpan);
            }
        }

        private void Tick()
        {
            try
            {
                if (_enabled())
                {
                    if (_backoff > 0)
                    {
                        _backoff--;
                    }
                    else if (!Capture(false))
                    {
                        _backoff = ClosedBackoffRounds;
                    }
                }
                else
                {
                    //Nothing is kept while the window is off limits, and what was seen before is not valid any more.
                    Forget();
                }
            }
            catch (Exception exception)
            {
                Logger.Debug(exception, "Looking at the Immediate window failed.");
            }
            finally
            {
                lock (_bufferLock)
                {
                    _timer?.Change(Interval, Timeout.InfiniteTimeSpan);
                }
            }
        }

        /// <summary>Looks at the window now and keeps what is new. With <paramref name="show"/> the window is opened if it is closed, and a window that cannot be read throws.</summary>
        /// <returns>False if the window is not open.</returns>
        public bool Capture(bool show)
        {
            lock (_pollLock)
            {
                var text = show ? _window.Read() : _window.Peek();
                if (text == null)
                {
                    return false;
                }

                var committed = Committed(text);
                if (_baseline == null)
                {
                    //What was there before the first look is history, not output.
                    _baseline = committed;
                    return true;
                }

                //A line that was taken for finished while it was being typed, and is being typed on: leave it, it is not output.
                var lastLineWithdrawn = _baseline.Length > committed.Length && _baseline.StartsWith(committed, StringComparison.Ordinal)
                    && _baseline.Substring(committed.Length).TrimEnd('\r', '\n').IndexOfAny(new[] { '\r', '\n' }) < 0;
                if (committed != _baseline && !lastLineWithdrawn)
                {
                    var added = ImmediateApi.Delta(_baseline, committed);
                    _baseline = committed;
                    Append(added);
                }

                return true;
            }
        }

        //A line still being typed is not output. The text up to the last line break is complete; the line after it counts once it has stayed the same for a look.
        private string Committed(string text)
        {
            //After output the window ends with an empty line for the next command, which that command then overwrites: it is the prompt, not output.
            //A blank line the code printed last waits until something follows it.
            var content = text.TrimEnd('\r', '\n');
            if (content.Length < text.Length)
            {
                text = content.Length == 0 ? string.Empty : content + "\r\n";
            }

            var cut = text.LastIndexOfAny(new[] { '\r', '\n' }) + 1;
            var tail = text.Substring(cut);
            var stable = tail.Length > 0 && tail == _tailSeen;
            _tailSeen = tail;
            return stable ? text + "\r\n" : text.Substring(0, cut);
        }

        private void Append(string added)
        {
            var lines = added.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n').ToList();
            if (lines.Count > 0 && lines[lines.Count - 1].Length == 0)
            {
                lines.RemoveAt(lines.Count - 1);
            }

            lock (_bufferLock)
            {
                foreach (var line in lines)
                {
                    _lines.AddLast((_next++, line));
                    if (_lines.Count > Capacity)
                    {
                        _lines.RemoveFirst();
                    }
                }
            }
        }

        private void Forget()
        {
            lock (_pollLock)
            {
                _baseline = null;
                _tailSeen = string.Empty;
            }
        }

        /// <param name="since">The <c>Next</c> of the last answer: the lines from that number on. Left out, the newest lines.</param>
        /// <param name="contains">Only lines with this text in them (case does not matter).</param>
        /// <param name="waitMs">If there is nothing to report, wait this long for something to be printed.</param>
        public ImmediateOutput Get(long? since, int? limit, string contains, int waitMs)
        {
            var take = limit ?? DefaultLimit;
            if (take < 1 || take > MaxLimit)
            {
                throw ApiException.BadRequest($"limit must be between 1 and {MaxLimit}.");
            }

            if (waitMs < 0 || waitMs > MaxWaitMs)
            {
                throw ApiException.BadRequest($"waitMs must be between 0 and {MaxWaitMs}.");
            }

            if (since < 0)
            {
                throw ApiException.BadRequest("since must not be negative.");
            }

            //An unreadable window is the caller's business: say so rather than answer "nothing printed".
            Capture(true);
            var answer = Select(since, take, contains);
            var deadline = _clock() + TimeSpan.FromMilliseconds(waitMs);
            while (answer.LineCount == 0 && _clock() < deadline)
            {
                _sleep(Poll);
                Capture(false);
                answer = Select(since, take, contains);
            }

            return answer;
        }

        private ImmediateOutput Select(long? since, int take, string contains)
        {
            lock (_bufferLock)
            {
                var first = _lines.Count > 0 ? _lines.First.Value.Number : _next;
                IEnumerable<(long Number, string Text)> candidates = _lines;
                if (!string.IsNullOrEmpty(contains))
                {
                    candidates = candidates.Where(line => line.Text.IndexOf(contains, StringComparison.OrdinalIgnoreCase) >= 0);
                }

                var output = new ImmediateOutput { Next = _next, Missed = since.HasValue && since.Value < first };
                List<string> lines;
                if (since.HasValue)
                {
                    var wanted = candidates.Where(line => line.Number >= since.Value).Take(take + 1).ToList();
                    if (wanted.Count > take)
                    {
                        wanted.RemoveAt(take);
                        output.More = true;
                        output.Next = wanted[wanted.Count - 1].Number + 1;
                    }

                    lines = wanted.Select(line => line.Text).ToList();
                }
                else
                {
                    lines = candidates.Reverse().Take(take).Reverse().Select(line => line.Text).ToList();
                }

                output.LineCount = lines.Count;
                output.Text = string.Join("\n", lines);
                return output;
            }
        }

        public void Dispose()
        {
            lock (_bufferLock)
            {
                _timer?.Dispose();
                _timer = null;
            }
        }
    }
}
