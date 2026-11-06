using System;
using System.Text.RegularExpressions;

namespace Rubberduck.Git
{
    /// <summary>What <c>git --version</c> says, as a comparable version.</summary>
    public static class GitVersion
    {
        /// <summary>
        /// The oldest git the client works with. Credentials are passed to a call through <c>GIT_CONFIG_COUNT</c> and its keys,
        /// which git understands from 2.31.
        /// </summary>
        public static readonly Version Minimum = new Version(2, 31);

        private static readonly Regex Output = new Regex(@"^git version (\d+)\.(\d+)(?:\.(\d+))?", RegexOptions.CultureInvariant);

        /// <summary>Reads "git version 2.43.0.windows.1" and the like. Returns null for anything else.</summary>
        public static Version Parse(string output)
        {
            var match = Output.Match(output?.Trim() ?? string.Empty);
            if (!match.Success)
            {
                return null;
            }

            var patch = match.Groups[3].Success ? int.Parse(match.Groups[3].Value) : 0;
            return new Version(int.Parse(match.Groups[1].Value), int.Parse(match.Groups[2].Value), patch);
        }

        public static bool IsSupported(Version version) => version != null && version >= Minimum;
    }
}
