using System;

namespace Rubberduck.Api
{
    /// <summary>
    /// The user's say over writes. The first write after the API starts asks them once; yes lets every later write through until the API is switched off or its token is regenerated.
    /// A no is remembered for a minute, so an agent that keeps trying does not bury the user in dialogs.
    /// </summary>
    public class ApiWriteApproval
    {
        private static readonly TimeSpan DenialMemory = TimeSpan.FromMinutes(1);

        private readonly Func<string, bool> _ask;
        private readonly Func<DateTime> _now;
        private readonly object _lock = new object();
        private bool _approved;
        private DateTime? _deniedAt;

        /// <param name="ask">Shows the question to the user, whatever thread that takes, and returns the answer.</param>
        public ApiWriteApproval(Func<string, bool> ask, Func<DateTime> now = null)
        {
            _ask = ask;
            _now = now ?? (() => DateTime.UtcNow);
        }

        public bool IsApproved
        {
            get
            {
                lock (_lock)
                {
                    return _approved;
                }
            }
        }

        /// <exception cref="ApiException">The user said no.</exception>
        /// <param name="client">Who is asking, for the user to see.</param>
        public void Require(string client)
        {
            //Held while the dialog is open: a second write waits for the answer to the first instead of asking again.
            lock (_lock)
            {
                if (_approved)
                {
                    return;
                }

                if (_deniedAt.HasValue && _now() - _deniedAt.Value < DenialMemory)
                {
                    throw Denied();
                }

                if (_ask(client))
                {
                    _approved = true;
                    _deniedAt = null;
                    return;
                }

                _deniedAt = _now();
                throw Denied();
            }
        }

        /// <summary>Forgets the approval, so the next write asks again.</summary>
        public void Reset()
        {
            lock (_lock)
            {
                _approved = false;
                _deniedAt = null;
            }
        }

        private static ApiException Denied()
            => ApiException.Forbidden("write_denied", "The user did not allow writes. Ask them to approve in Rubberduck, then retry in a minute.");
    }
}
