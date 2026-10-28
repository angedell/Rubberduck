using System;
using System.Linq;
using System.Reflection;

namespace Rubberduck
{
    /// <summary>
    /// Which build is running: the full version (<c>build.cmd</c> stamps <c>2.5.92.N</c>, N counting up with every build) and the git commit it was built from.
    /// The splash screen, the REST API (<c>GET /api/v1/info</c>) and the MCP server's <c>serverInfo</c> all say the same thing, so a report of "it still does that" can be matched to a build.
    /// </summary>
    public sealed class BuildInfo
    {
        private const int CommitLength = 9;

        /// <summary>All four parts, e.g. <c>2.5.92.28</c>.</summary>
        public string Version { get; }

        /// <summary>The start of the commit hash the build was made from, or null if the build did not record one.</summary>
        public string Commit { get; }

        /// <summary>What to show people: the version, and the commit in brackets when there is one.</summary>
        public string Display => Commit == null ? Version : $"{Version} ({Commit})";

        private BuildInfo(string version, string commit)
        {
            Version = version;
            Commit = commit;
        }

        /// <summary>The build of Rubberduck's own assembly.</summary>
        public static BuildInfo Current { get; } = Describe(typeof(BuildInfo).Assembly);

        public static BuildInfo Describe(Assembly assembly)
        {
            var informational = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
            return Describe(assembly.GetName().Version, informational);
        }

        /// <param name="version">The assembly version. Missing parts count as 0, so the answer always has four.</param>
        /// <param name="informational">The informational version, <c>2.5.92.28+0123abc...</c>: the part after '+' is the commit.</param>
        public static BuildInfo Describe(Version version, string informational)
        {
            var four = version == null
                ? "0.0.0.0"
                : $"{version.Major}.{version.Minor}.{Math.Max(version.Build, 0)}.{Math.Max(version.Revision, 0)}";

            string commit = null;
            var plus = informational?.IndexOf('+') ?? -1;
            if (plus >= 0)
            {
                //What follows may be more metadata after the hash (a "-dirty" mark, say); only the hash itself is kept.
                var hash = new string(informational.Substring(plus + 1).TakeWhile(Uri.IsHexDigit).ToArray());
                if (hash.Length >= 7)
                {
                    commit = hash.Substring(0, Math.Min(hash.Length, CommitLength));
                }
            }

            return new BuildInfo(four, commit);
        }
    }
}
