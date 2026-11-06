using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using Newtonsoft.Json.Linq;

namespace Rubberduck.Git
{
    public class AiException : Exception
    {
        public AiException(string message) : base(message)
        {
        }
    }

    /// <summary>A language model that answers one question at a time.</summary>
    public interface IAiClient
    {
        /// <summary>Where the question goes, so the user can be told before anything is sent.</summary>
        string Endpoint { get; }

        /// <exception cref="AiException">The service refused or could not be reached; the message is for the user.</exception>
        string Complete(string system, string prompt, int maxTokens = 600);
    }

    /// <summary>The Anthropic Messages API.</summary>
    public class AnthropicClient : IAiClient
    {
        public const string DefaultModel = "claude-sonnet-5-5";
        public const string Url = "https://api.anthropic.com/v1/messages";

        private readonly string _key;
        private readonly string _model;

        public AnthropicClient(string apiKey, string model)
        {
            _key = apiKey;
            _model = string.IsNullOrWhiteSpace(model) ? DefaultModel : model.Trim();
        }

        public string Endpoint => Url;

        public string Complete(string system, string prompt, int maxTokens = 600)
        {
            var body = new JObject
            {
                ["model"] = _model,
                ["max_tokens"] = maxTokens,
                ["system"] = system,
                ["messages"] = new JArray(new JObject { ["role"] = "user", ["content"] = prompt })
            };

            var handler = new HttpClientHandler { Proxy = WebRequest.GetSystemWebProxy(), UseProxy = true };
            handler.Proxy.Credentials = CredentialCache.DefaultCredentials;
            using (var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(60) })
            using (var request = new HttpRequestMessage(HttpMethod.Post, Url))
            {
                request.Headers.Add("x-api-key", _key);
                request.Headers.Add("anthropic-version", "2023-06-01");
                request.Content = new StringContent(body.ToString(), Encoding.UTF8, "application/json");

                HttpResponseMessage response;
                string text;
                try
                {
                    response = client.SendAsync(request).GetAwaiter().GetResult();
                    text = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                }
                catch (Exception exception) when (exception is HttpRequestException || exception is System.Threading.Tasks.TaskCanceledException)
                {
                    throw new AiException("The AI service could not be reached: " + exception.Message);
                }

                if (!response.IsSuccessStatusCode)
                {
                    throw new AiException(Describe(response.StatusCode, text));
                }

                return Reply(text);
            }
        }

        internal static string Reply(string json)
        {
            try
            {
                var parts = JObject.Parse(json)["content"] as JArray;
                var answer = string.Concat(parts?.Where(part => (string)part["type"] == "text").Select(part => (string)part["text"]) ?? new string[0]).Trim();
                if (answer.Length == 0)
                {
                    throw new AiException("The AI service answered with nothing.");
                }

                return answer;
            }
            catch (Newtonsoft.Json.JsonException)
            {
                throw new AiException("The AI service's answer could not be read.");
            }
        }

        //The service's own words, without the request: the key is never in a response, and never in a message here.
        private static string Describe(HttpStatusCode status, string json)
        {
            string detail = null;
            try
            {
                detail = (string)JObject.Parse(json)["error"]?["message"];
            }
            catch (Newtonsoft.Json.JsonException)
            {
            }

            var prefix = status == HttpStatusCode.Unauthorized ? "The API key was refused" : $"The AI service refused ({(int)status})";
            return string.IsNullOrEmpty(detail) ? prefix + "." : prefix + ": " + detail;
        }
    }
}
