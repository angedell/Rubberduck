using System;
using System.Collections.Generic;
using System.Text;

namespace Rubberduck.Git
{
    /// <summary>What a remote call needs to sign in: environment variables for the child process and the strings that must never be logged.</summary>
    public class GitRemoteAuth
    {
        public GitRemoteAuth(IReadOnlyDictionary<string, string> environment, IReadOnlyList<string> secrets)
        {
            Environment = environment;
            Secrets = secrets;
        }

        public IReadOnlyDictionary<string, string> Environment { get; }
        public IReadOnlyList<string> Secrets { get; }

        public static readonly GitRemoteAuth None = new GitRemoteAuth(new Dictionary<string, string>(), new string[0]);

        /// <summary>
        /// Signs in to one host with a token. git reads <c>GIT_CONFIG_COUNT</c> and its key and value variables as configuration for the call (2.31 or later),
        /// so the header goes out with the requests to that host only (the key names the host), and appears neither on the command line nor in any file.
        /// </summary>
        public static GitRemoteAuth ForHost(string host, StoredCredential credential)
        {
            //GitHub and Azure DevOps both take a token as the password of basic authentication; the user name is not checked.
            var user = string.IsNullOrEmpty(credential.User) ? "x-access-token" : credential.User;
            var basic = Convert.ToBase64String(Encoding.UTF8.GetBytes(user + ":" + credential.Secret));
            return new GitRemoteAuth(
                new Dictionary<string, string>
                {
                    ["GIT_CONFIG_COUNT"] = "1",
                    ["GIT_CONFIG_KEY_0"] = "http.https://" + host + "/.extraHeader",
                    ["GIT_CONFIG_VALUE_0"] = "Authorization: Basic " + basic
                },
                new[] { credential.Secret, basic });
        }

        /// <summary>The host of an https URL, or null for anything else (no ssh, no local path, no ext::).</summary>
        public static string HostOf(string url)
        {
            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps || string.IsNullOrEmpty(uri.Host) || !string.IsNullOrEmpty(uri.UserInfo))
            {
                return null;
            }

            return uri.Host.ToLowerInvariant();
        }
    }
}
