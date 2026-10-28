using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using NLog;

namespace Rubberduck.Api
{
    /// <summary>What Windows Firewall would do with a connection from another machine to the API port. Every value can be "unknown"; asking never fails.</summary>
    public class FirewallStatus
    {
        /// <summary>The port that was checked.</summary>
        public int Port { get; set; }

        /// <summary>The network profile(s) the machine is on right now: "private", "public", "domain", several joined with a comma, or "unknown".</summary>
        public string Profile { get; set; } = FirewallCheck.Unknown;

        /// <summary>"on" if the firewall is switched on for the active profile, "off", or "unknown".</summary>
        public string Firewall { get; set; } = FirewallCheck.Unknown;

        /// <summary>"yes" if an enabled inbound allow rule covers the port (TCP) for the active profile, "no", or "unknown". Block rules are not looked at.</summary>
        public string InboundRule { get; set; } = FirewallCheck.Unknown;

        /// <summary>
        /// "yes" if a program on another machine should get through (the firewall is off, or a rule allows it), "no" if it will probably be blocked,
        /// "unknown" if that could not be told.
        /// </summary>
        public string NetworkReachable { get; set; } = FirewallCheck.Unknown;

        /// <summary>One sentence for the user, or null if there is nothing to add.</summary>
        public string Hint { get; set; }
    }

    /// <summary>Reads the state of Windows Firewall. A seam, so that the API and the settings page can be tested without a firewall.</summary>
    public interface IFirewallCheck
    {
        /// <summary>For the port the API listens on now (or would listen on by default).</summary>
        FirewallStatus Check();

        FirewallStatus Check(int port);
    }

    /// <summary>One inbound rule, as far as the check needs it. Kept apart from COM so that the matching can be tested.</summary>
    public class FirewallRule
    {
        public bool Enabled { get; set; }
        public bool Inbound { get; set; }
        public bool Allow { get; set; }

        /// <summary>IANA protocol number; 6 is TCP, 256 is "any".</summary>
        public int Protocol { get; set; }

        /// <summary>As Windows writes it: "*", "7437", "80,443", "7000-7500". Null or empty means any.</summary>
        public string LocalPorts { get; set; }

        /// <summary>Bitmask: 1 domain, 2 private, 4 public; 0x7FFFFFFF all.</summary>
        public int Profiles { get; set; }

        /// <summary>Set if the rule is for one program only.</summary>
        public string ApplicationName { get; set; }

        /// <summary>Set if the rule is for one service only.</summary>
        public string ServiceName { get; set; }
    }

    /// <summary>Asks Windows Firewall through HNetCfg.FwPolicy2 (late bound, no reference needed).</summary>
    public class FirewallCheck : IFirewallCheck
    {
        public const string Unknown = "unknown";

        private const int Domain = 1;
        private const int Private = 2;
        private const int Public = 4;

        private static readonly ILogger Logger = LogManager.GetCurrentClassLogger();

        private readonly ApiState _state;
        private readonly Func<string> _programPath;

        public FirewallCheck(ApiState state) : this(state, () => Process.GetCurrentProcess().MainModule?.FileName)
        {
        }

        public FirewallCheck(ApiState state, Func<string> programPath)
        {
            _state = state;
            _programPath = programPath;
        }

        public FirewallStatus Check() => Check(_state != null && _state.Port > 0 ? _state.Port : ApiServer.DefaultPort);

        public FirewallStatus Check(int port)
        {
            try
            {
                return Read(port);
            }
            catch (Exception exception)
            {
                Logger.Debug(exception, "The firewall could not be queried.");
                return Evaluate(port, 0, null, null);
            }
        }

        private FirewallStatus Read(int port)
        {
            var type = Type.GetTypeFromProgID("HNetCfg.FwPolicy2");
            if (type == null)
            {
                return Evaluate(port, 0, null, null);
            }

            dynamic policy = Activator.CreateInstance(type);
            int active = Convert.ToInt32(policy.CurrentProfileTypes);

            var enabled = new Dictionary<int, bool>();
            foreach (var profile in new[] { Domain, Private, Public }.Where(p => (active & p) != 0))
            {
                try
                {
                    enabled[profile] = (bool)policy.FirewallEnabled[profile];
                }
                catch (Exception exception)
                {
                    Logger.Debug(exception, "The firewall state of one profile could not be read.");
                }
            }

            List<FirewallRule> rules = null;
            try
            {
                rules = new List<FirewallRule>();
                foreach (dynamic rule in policy.Rules)
                {
                    try
                    {
                        // Cheap checks first: most rules are disabled or outbound, and each property is a COM call.
                        if (!(bool)rule.Enabled || Convert.ToInt32(rule.Direction) != 1)
                        {
                            continue;
                        }

                        rules.Add(new FirewallRule
                        {
                            Enabled = true,
                            Inbound = true,
                            Allow = Convert.ToInt32(rule.Action) == 1,
                            Protocol = Convert.ToInt32(rule.Protocol),
                            LocalPorts = (string)rule.LocalPorts,
                            Profiles = Convert.ToInt32(rule.Profiles),
                            ApplicationName = (string)rule.ApplicationName,
                            ServiceName = (string)rule.ServiceName
                        });
                    }
                    catch (Exception)
                    {
                        //A rule that cannot be read is skipped; it must not hide the others.
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Debug(exception, "The firewall rules could not be listed.");
                rules = null;
            }

            return Evaluate(port, active, enabled, rules, SafePath());
        }

        private string SafePath()
        {
            try
            {
                return _programPath?.Invoke();
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>Turns what Windows said into an answer. Public so the logic can be tested; null for a part means it could not be read.</summary>
        public static FirewallStatus Evaluate(int port, int activeProfiles, IDictionary<int, bool> enabledByProfile, IEnumerable<FirewallRule> rules, string programPath = null)
        {
            var status = new FirewallStatus { Port = port };
            if (activeProfiles == 0)
            {
                status.Hint = "The firewall could not be queried.";
                return status;
            }

            status.Profile = ProfileName(activeProfiles);

            var on = enabledByProfile == null ? null : enabledByProfile.Where(p => (activeProfiles & p.Key) != 0).Select(p => (bool?)p.Value).ToList();
            if (on != null && on.Count > 0)
            {
                status.Firewall = on.Any(v => v == true) ? "on" : "off";
            }

            if (rules != null)
            {
                var covered = 0;
                foreach (var rule in rules.Where(r => Applies(r, port, programPath)))
                {
                    covered |= rule.Profiles;
                }

                status.InboundRule = (covered & activeProfiles) == activeProfiles ? "yes" : "no";
            }

            if (status.Firewall == "off")
            {
                status.NetworkReachable = "yes";
            }
            else if (status.InboundRule == "yes")
            {
                status.NetworkReachable = "yes";
            }
            else if (status.Firewall == "on" && status.InboundRule == "no")
            {
                status.NetworkReachable = "no";
                status.Hint = $"Windows Firewall is on for the {status.Profile} network and has no rule that lets other machines reach port {port}. " +
                              "Only if a sandbox or VM cannot connect, run the firewall command from the API setup.";
            }

            return status;
        }

        /// <summary>Whether the rule lets inbound TCP traffic to the port through.</summary>
        public static bool Applies(FirewallRule rule, int port, string programPath = null)
        {
            if (rule == null || !rule.Enabled || !rule.Inbound || !rule.Allow || (!string.IsNullOrEmpty(rule.ServiceName) && !rule.ServiceName.Equals("*")))
            {
                return false;
            }

            if (rule.Protocol != 6 && rule.Protocol != 256)
            {
                return false;
            }

            if (!string.IsNullOrEmpty(rule.ApplicationName) && !string.Equals(rule.ApplicationName, programPath, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return PortMatches(rule.LocalPorts, port);
        }

        /// <summary>Windows writes ports as "*", a number, a range "7000-7500", or a comma separated list of those.</summary>
        public static bool PortMatches(string localPorts, int port)
        {
            if (string.IsNullOrWhiteSpace(localPorts) || localPorts.Trim() == "*")
            {
                return true;
            }

            foreach (var part in localPorts.Split(','))
            {
                var bounds = part.Trim().Split('-');
                if (bounds.Length == 1 && int.TryParse(bounds[0], out var single) && single == port)
                {
                    return true;
                }

                if (bounds.Length == 2 && int.TryParse(bounds[0], out var low) && int.TryParse(bounds[1], out var high) && port >= low && port <= high)
                {
                    return true;
                }
            }

            return false;
        }

        public static string ProfileName(int profiles)
        {
            var names = new List<string>();
            if ((profiles & Private) != 0) { names.Add("private"); }
            if ((profiles & Domain) != 0) { names.Add("domain"); }
            if ((profiles & Public) != 0) { names.Add("public"); }
            return names.Count == 0 ? Unknown : string.Join(",", names);
        }
    }
}
