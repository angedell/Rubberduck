using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Api
{
    /// <summary>
    /// Decides whether a request may be served at all, before anything is looked at: it has to come from an agent that knows the token,
    /// and not from a web page. A page can reach the loopback address too, and an agent never sends an <c>Origin</c> header, so any request with one is refused.
    /// Too many wrong tokens from one address and that address is turned away for a while.
    /// </summary>
    public class ApiRequestGuard
    {
        private const int MaxFailures = 10;
        private static readonly TimeSpan Lockout = TimeSpan.FromMinutes(1);
        private static readonly string[] LoopbackHosts = { "127.0.0.1", "localhost", "[::1]" };

        private readonly ApiToken _token;
        private readonly Func<bool> _allowNetwork;
        private readonly Func<DateTime> _now;
        private readonly Dictionary<string, (int Failures, DateTime Since)> _failures = new Dictionary<string, (int, DateTime)>();

        public ApiRequestGuard(ApiToken token, Func<bool> allowNetwork, Func<DateTime> now = null)
        {
            _token = token;
            _allowNetwork = allowNetwork;
            _now = now ?? (() => DateTime.UtcNow);
        }

        /// <exception cref="ApiException">The request must not be served.</exception>
        /// <param name="host">The Host header, with its port.</param>
        /// <param name="origin">The Origin header, or null.</param>
        /// <param name="authorization">The Authorization header, or null.</param>
        /// <param name="remote">Who is asking, to count wrong tokens.</param>
        public void Check(string host, string origin, string authorization, string remote)
        {
            if (!string.IsNullOrEmpty(origin))
            {
                throw ApiException.Forbidden("origin_refused", "Requests from web pages are refused.");
            }

            if (!_allowNetwork() && !LoopbackHosts.Contains(HostName(host), StringComparer.OrdinalIgnoreCase))
            {
                throw ApiException.Forbidden("host_refused", "Network access is switched off; use 127.0.0.1 or localhost.");
            }

            if (IsLockedOut(remote))
            {
                throw new ApiException(429, "too_many_failures", "Too many wrong tokens. Try again in a minute.");
            }

            if (!_token.Matches(BearerToken(authorization)))
            {
                RecordFailure(remote);
                throw ApiException.Unauthorized("Send the token as 'Authorization: Bearer <token>'.");
            }

            ClearFailures(remote);
        }

        public static string BearerToken(string authorization)
        {
            const string scheme = "Bearer ";
            return authorization != null && authorization.StartsWith(scheme, StringComparison.OrdinalIgnoreCase)
                ? authorization.Substring(scheme.Length).Trim()
                : null;
        }

        //"host:port", "[::1]:port" or "host".
        public static string HostName(string host)
        {
            if (string.IsNullOrEmpty(host))
            {
                return string.Empty;
            }

            if (host.StartsWith("[", StringComparison.Ordinal))
            {
                var end = host.IndexOf(']');
                return end < 0 ? host : host.Substring(0, end + 1);
            }

            var colon = host.LastIndexOf(':');
            return colon < 0 ? host : host.Substring(0, colon);
        }

        private bool IsLockedOut(string remote)
        {
            lock (_failures)
            {
                if (!_failures.TryGetValue(remote ?? string.Empty, out var entry))
                {
                    return false;
                }

                if (_now() - entry.Since > Lockout)
                {
                    _failures.Remove(remote ?? string.Empty);
                    return false;
                }

                return entry.Failures >= MaxFailures;
            }
        }

        private void RecordFailure(string remote)
        {
            lock (_failures)
            {
                var key = remote ?? string.Empty;
                _failures.TryGetValue(key, out var entry);
                _failures[key] = (entry.Failures + 1, entry.Failures == 0 ? _now() : entry.Since);
            }
        }

        private void ClearFailures(string remote)
        {
            lock (_failures)
            {
                _failures.Remove(remote ?? string.Empty);
            }
        }
    }
}
