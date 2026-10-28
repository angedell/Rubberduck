using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using NLog;

namespace Rubberduck.Api
{
    /// <summary>Whether the server is answering on the network right now. Shared by the server, which sets it, and the guard, which reads it.</summary>
    public class ApiState
    {
        public volatile bool Network;

        /// <summary>The user allows agents to use the Immediate window, and the server is running.</summary>
        public volatile bool Immediate;
    }

    public enum ApiStartProblem
    {
        /// <summary>Windows has not been told that this user may listen on the network.</summary>
        NeedsUrlAcl,
        PortInUse,
        Other
    }

    /// <summary>The server could not start. The message is for the user; <see cref="Problem"/> tells what can be offered to fix it.</summary>
    public class ApiStartException : InvalidOperationException
    {
        public ApiStartException(string message, ApiStartProblem problem, int port, Exception inner) : base(message, inner)
        {
            Problem = problem;
            Port = port;
        }

        public ApiStartProblem Problem { get; }
        public int Port { get; }
    }

    /// <summary>
    /// The HTTP side: listens, checks each request with <see cref="ApiRequestGuard"/>, hands it to <see cref="ApiRouter"/> and writes the answer.
    /// Off unless the user switched it on. By default it answers on the loopback address only; with network access on it answers on every address of the machine.
    /// </summary>
    public class ApiServer : IDisposable
    {
        public const int DefaultPort = 7437;
        private const int MaxBodyBytes = 5 * 1024 * 1024;

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private static readonly JsonSerializerSettings JsonSettings = new JsonSerializerSettings
        {
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
            NullValueHandling = NullValueHandling.Ignore,
            Formatting = Formatting.Indented
        };

        private readonly ApiRouter _router;
        private readonly ApiRequestGuard _guard;
        private readonly ApiWriteApproval _approval;
        private readonly ApiExecuteApproval _executeApproval;
        private readonly ApiState _state;
        private readonly object _lock = new object();

        private HttpListener _listener;
        private int _port;
        private bool _network;

        public ApiServer(ApiRouter router, ApiRequestGuard guard, ApiWriteApproval approval, ApiExecuteApproval executeApproval, ApiState state)
        {
            _router = router;
            _guard = guard;
            _approval = approval;
            _executeApproval = executeApproval;
            _state = state;
        }

        public bool IsRunning
        {
            get
            {
                lock (_lock)
                {
                    return _listener != null;
                }
            }
        }

        public int Port => _port;

        /// <summary>Where agents can reach the server, as far as this machine knows.</summary>
        public IEnumerable<string> Urls()
        {
            yield return $"http://127.0.0.1:{_port}";
            if (_network)
            {
                foreach (var address in Dns.GetHostAddresses(Dns.GetHostName()).Where(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork))
                {
                    yield return $"http://{address}:{_port}";
                }
            }
        }

        /// <summary>
        /// Starts, stops or restarts the server so that it matches the settings. Switching it off also takes back the user's approvals;
        /// so does taking away the Immediate window.
        /// </summary>
        /// <exception cref="ApiStartException">The server could not listen; the message says what to do about it.</exception>
        public void Apply(bool enabled, bool network, int port, bool allowImmediate)
        {
            lock (_lock)
            {
                _state.Immediate = enabled && allowImmediate;
                if (!_state.Immediate)
                {
                    _executeApproval.Reset();
                }

                if (_listener != null && enabled && _network == network && _port == port)
                {
                    return;
                }

                StopListening();
                if (!enabled)
                {
                    _approval.Reset();
                    return;
                }

                Listen(network, port <= 0 || port > 65535 ? DefaultPort : port);
            }
        }

        private void Listen(bool network, int port)
        {
            var listener = new HttpListener();
            if (network)
            {
                listener.Prefixes.Add($"http://+:{port}/");
            }
            else
            {
                listener.Prefixes.Add($"http://127.0.0.1:{port}/");
                listener.Prefixes.Add($"http://localhost:{port}/");
            }

            try
            {
                listener.Start();
            }
            catch (HttpListenerException exception)
            {
                listener.Close();
                throw Explain(exception, network, port);
            }

            _listener = listener;
            _network = network;
            _state.Network = network;
            _port = port;
            Logger.Info($"API listening on port {port}{(network ? " (network)" : " (loopback only)")}.");
            Task.Run(() => Accept(listener));
        }

        private static ApiStartException Explain(HttpListenerException exception, bool network, int port)
        {
            switch (exception.ErrorCode)
            {
                case 5 when network:
                    return new ApiStartException(
                        "Windows has to be told once that Rubberduck may listen on the network.", ApiStartProblem.NeedsUrlAcl, port, exception);
                case 32:
                case 183:
                    return new ApiStartException($"Port {port} is already in use. Choose another port.", ApiStartProblem.PortInUse, port, exception);
                default:
                    return new ApiStartException($"The API could not start on port {port}: {exception.Message}", ApiStartProblem.Other, port, exception);
            }
        }

        private async Task Accept(HttpListener listener)
        {
            while (listener.IsListening)
            {
                HttpListenerContext context;
                try
                {
                    context = await listener.GetContextAsync().ConfigureAwait(false);
                }
                catch (Exception exception) when (exception is HttpListenerException || exception is ObjectDisposedException || exception is InvalidOperationException)
                {
                    return;
                }

                //One request must not hold up the next; the VBE work serialises itself on the UI thread.
                var _ = Task.Run(() => Serve(context));
            }
        }

        private void Serve(HttpListenerContext context)
        {
            ApiResponse response;
            try
            {
                response = Process(context.Request);
            }
            catch (ApiException exception)
            {
                response = ApiResponse.Error(exception.Status, exception.Code, exception.Message);
            }
            catch (Exception exception)
            {
                Logger.Error(exception, "API request failed.");
                response = ApiResponse.Error(500, "internal_error", exception.Message);
            }

            try
            {
                Write(context.Response, response);
            }
            catch (Exception exception) when (exception is HttpListenerException || exception is System.IO.IOException || exception is ObjectDisposedException)
            {
                Logger.Debug(exception, "API client went away.");
            }
        }

        private ApiResponse Process(HttpListenerRequest request)
        {
            var remote = request.RemoteEndPoint?.Address.ToString();
            _guard.Check(request.Headers["Host"], request.Headers["Origin"], request.Headers["Authorization"], remote);

            if (request.ContentLength64 > MaxBodyBytes)
            {
                throw new ApiException(413, "too_large", $"The body is larger than {MaxBodyBytes / 1024 / 1024} MB.");
            }

            string body = null;
            if (request.HasEntityBody)
            {
                using (var reader = new System.IO.StreamReader(request.InputStream, request.ContentEncoding ?? Encoding.UTF8))
                {
                    body = reader.ReadToEnd();
                }
            }

            return _router.Handle(new ApiRequest
            {
                Method = request.HttpMethod,
                Path = request.Url.AbsolutePath,
                Query = request.QueryString.AllKeys.Where(k => k != null).ToDictionary(k => k, k => request.QueryString[k], StringComparer.OrdinalIgnoreCase),
                IfMatch = request.Headers["If-Match"],
                Body = body,
                Client = string.IsNullOrWhiteSpace(request.UserAgent) ? remote : $"{request.UserAgent} at {remote}"
            });
        }

        private static void Write(HttpListenerResponse response, ApiResponse result)
        {
            response.StatusCode = result.Status;
            response.Headers["Cache-Control"] = "no-store";
            if (result.ETag != null)
            {
                response.Headers["ETag"] = $"\"{result.ETag}\"";
            }

            //Unavailable means busy parsing, and that passes: say when to ask again.
            if (result.Status == 503)
            {
                response.Headers["Retry-After"] = "2";
            }

            if (result.Body == null)
            {
                response.Close();
                return;
            }

            var bytes = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(result.Body, JsonSettings));
            response.ContentType = "application/json; charset=utf-8";
            response.ContentLength64 = bytes.Length;
            response.OutputStream.Write(bytes, 0, bytes.Length);
            response.Close();
        }

        private void StopListening()
        {
            if (_listener == null)
            {
                return;
            }

            try
            {
                _listener.Stop();
                _listener.Close();
            }
            catch (ObjectDisposedException)
            {
                //Already closed.
            }

            _listener = null;
            _state.Network = false;
            Logger.Info("API stopped.");
        }

        public void Dispose()
        {
            lock (_lock)
            {
                StopListening();
            }
        }
    }
}
