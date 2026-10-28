using System;
using System.Linq;

namespace Rubberduck.Api
{
    /// <summary>What the user needs to hand an agent, and the one thing they can do about a token that got out.</summary>
    public class ApiConnection
    {
        private readonly ApiServer _server;
        private readonly ApiToken _token;
        private readonly ApiWriteApproval _approval;

        public ApiConnection(ApiServer server, ApiToken token, ApiWriteApproval approval)
        {
            _server = server;
            _token = token;
            _approval = approval;
        }

        /// <summary>The addresses, the token and the MCP endpoint, as text to paste into an agent's setup.</summary>
        public string Details()
        {
            var urls = _server.IsRunning ? _server.Urls().ToList() : new System.Collections.Generic.List<string> { $"http://127.0.0.1:{ApiServer.DefaultPort}" };
            var lines = urls.SelectMany(url => new[] { $"REST: {url}{ApiRouter.Base}", $"MCP:  {url}{ApiRouter.McpPath}" }).ToList();
            lines.Add($"Authorization: Bearer {_token.Value}");
            if (!_server.IsRunning)
            {
                lines.Add("(The API is switched off, or has not started yet.)");
            }

            return string.Join(Environment.NewLine, lines);
        }

        /// <summary>Locks out whoever has the old token, and makes the user approve writes again.</summary>
        public void Regenerate()
        {
            _token.Regenerate();
            _approval.Reset();
        }
    }
}
