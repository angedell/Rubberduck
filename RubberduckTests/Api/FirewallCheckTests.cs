using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Api;
using Rubberduck.UI.Settings;

namespace RubberduckTests.Api
{
    [TestFixture]
    [Category("Api")]
    public class FirewallCheckTests
    {
        private const int Private = 2;
        private const int Public = 4;

        private class FakeFirewall : IFirewallCheck
        {
            public FirewallStatus Status = new FirewallStatus { Port = 7437, Profile = "private", Firewall = "on", InboundRule = "no", NetworkReachable = "no", Hint = "closed" };
            public int? AskedPort;

            public FirewallStatus Check() => Status;

            public FirewallStatus Check(int port)
            {
                AskedPort = port;
                return Status;
            }
        }

        private static FirewallRule Allow(string ports = "7437", int profiles = Private, int protocol = 6)
            => new FirewallRule { Enabled = true, Inbound = true, Allow = true, Protocol = protocol, LocalPorts = ports, Profiles = profiles };

        private static FirewallStatus Evaluate(int active, bool enabled, params FirewallRule[] rules)
            => FirewallCheck.Evaluate(7437, active, new Dictionary<int, bool> { { active, enabled } }, rules);

        [TestCase("*", true)]
        [TestCase("", true)]
        [TestCase(null, true)]
        [TestCase("7437", true)]
        [TestCase("80,7437", true)]
        [TestCase("7000-7500", true)]
        [TestCase("7438", false)]
        [TestCase("80, 443", false)]
        [TestCase("1-100", false)]
        [TestCase("nonsense", false)]
        public void PortsAreMatchedTheWayWindowsWritesThem(string ports, bool expected)
            => Assert.AreEqual(expected, FirewallCheck.PortMatches(ports, 7437));

        [Test]
        public void AnAllowRuleForThePort_Covers()
        {
            var status = Evaluate(Private, true, Allow());

            Assert.AreEqual("private", status.Profile);
            Assert.AreEqual("on", status.Firewall);
            Assert.AreEqual("yes", status.InboundRule);
            Assert.AreEqual("yes", status.NetworkReachable);
            Assert.IsNull(status.Hint);
        }

        [Test]
        public void ARuleForAnotherPort_DoesNotCover()
        {
            var status = Evaluate(Private, true, Allow("80"));

            Assert.AreEqual("no", status.InboundRule);
            Assert.AreEqual("no", status.NetworkReachable);
            StringAssert.Contains("7437", status.Hint);
        }

        [Test]
        public void ARuleForAnotherProfile_DoesNotCover()
            => Assert.AreEqual("no", Evaluate(Public, true, Allow(profiles: Private)).InboundRule);

        [Test]
        public void ADisabledBlockOrOutboundRule_DoesNotCover()
        {
            var disabled = Allow(); disabled.Enabled = false;
            var block = Allow(); block.Allow = false;
            var outbound = Allow(); outbound.Inbound = false;
            var udp = Allow(protocol: 17);

            Assert.AreEqual("no", Evaluate(Private, true, disabled, block, outbound, udp).InboundRule);
        }

        [Test]
        public void ARuleForOneOtherProgram_DoesNotCover_ButOurOwnDoes()
        {
            var rule = Allow(); rule.ApplicationName = @"C:\Other\app.exe";
            var ours = Allow(); ours.ApplicationName = @"C:\Office\EXCEL.EXE";

            Assert.AreEqual("no", FirewallCheck.Evaluate(7437, Private, new Dictionary<int, bool> { { Private, true } }, new[] { rule }, @"C:\Office\excel.exe").InboundRule);
            Assert.AreEqual("yes", FirewallCheck.Evaluate(7437, Private, new Dictionary<int, bool> { { Private, true } }, new[] { ours }, @"C:\Office\excel.exe").InboundRule);
        }

        [Test]
        public void ARuleForAnyProtocolOrAServiceWildcard_Covers()
        {
            var rule = Allow(protocol: 256); rule.ServiceName = "*";

            Assert.AreEqual("yes", Evaluate(Private, true, rule).InboundRule);
        }

        [Test]
        public void TwoProfilesAreActive_BothMustBeCovered()
        {
            var both = FirewallCheck.Evaluate(7437, Private | Public, new Dictionary<int, bool> { { Private, true }, { Public, true } }, new[] { Allow(profiles: Private) });
            var covered = FirewallCheck.Evaluate(7437, Private | Public, new Dictionary<int, bool> { { Private, true }, { Public, true } }, new[] { Allow(profiles: Private), Allow(profiles: Public) });

            Assert.AreEqual("private,public", both.Profile);
            Assert.AreEqual("no", both.InboundRule);
            Assert.AreEqual("yes", covered.InboundRule);
        }

        [Test]
        public void TheFirewallBeingOff_MeansReachable_WithoutARule()
        {
            var status = Evaluate(Public, false);

            Assert.AreEqual("off", status.Firewall);
            Assert.AreEqual("no", status.InboundRule);
            Assert.AreEqual("yes", status.NetworkReachable);
        }

        [Test]
        public void NothingReadable_IsUnknown_NotAnError()
        {
            var status = FirewallCheck.Evaluate(7437, 0, null, null);

            Assert.AreEqual("unknown", status.Profile);
            Assert.AreEqual("unknown", status.Firewall);
            Assert.AreEqual("unknown", status.InboundRule);
            Assert.AreEqual("unknown", status.NetworkReachable);
        }

        [Test]
        public void RulesThatCouldNotBeListed_LeaveTheRuleAnswerUnknown()
        {
            var status = FirewallCheck.Evaluate(7437, Private, new Dictionary<int, bool> { { Private, true } }, null);

            Assert.AreEqual("on", status.Firewall);
            Assert.AreEqual("unknown", status.InboundRule);
            Assert.AreEqual("unknown", status.NetworkReachable);
        }

        [Test]
        public void TheRealCheck_NeverThrows()
        {
            var check = new FirewallCheck(new ApiState());

            Assert.DoesNotThrow(() => check.Check());
            Assert.AreEqual(1234, check.Check(1234).Port);
        }

        [Test]
        public void WithoutARunningServer_TheDefaultPortIsChecked()
            => Assert.AreEqual(ApiServer.DefaultPort, new FirewallCheck(new ApiState()).Check().Port);

        private static ApiRouter Router(IFirewallCheck firewall)
            => new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api, firewall: firewall);

        [Test]
        public void TheInfoRoute_CarriesTheFirewallAnswer()
        {
            var response = Router(new FakeFirewall()).Handle(new ApiRequest { Method = "GET", Path = ApiRouter.Base + "/info", Client = "test" });

            Assert.AreEqual(200, response.Status);
            var firewall = ((InfoReport)response.Body).Firewall;
            Assert.AreEqual("private", firewall.Profile);
            Assert.AreEqual("on", firewall.Firewall);
            Assert.AreEqual("no", firewall.InboundRule);
            Assert.AreEqual("no", firewall.NetworkReachable);
        }

        [Test]
        public void TheInfoRoute_WithoutACheck_StillAnswers()
        {
            var response = Router(null).Handle(new ApiRequest { Method = "GET", Path = ApiRouter.Base + "/info", Client = "test" });

            Assert.AreEqual(200, response.Status);
            Assert.IsNull(((InfoReport)response.Body).Firewall);
        }

        [Test]
        public void TheMcpInfoTool_CarriesTheFirewallAnswer()
        {
            var router = Router(new FakeFirewall());
            var body = new JObject { ["jsonrpc"] = "2.0", ["id"] = 1, ["method"] = "tools/call", ["params"] = new JObject { ["name"] = "get_info", ["arguments"] = new JObject() } };

            var response = router.Handle(new ApiRequest { Method = "POST", Path = ApiRouter.McpPath, Client = "test", Body = body.ToString() });

            Assert.AreEqual(200, response.Status);
            var text = (string)JObject.FromObject(response.Body)["result"]["content"][0]["text"];
            StringAssert.Contains("networkreachable", text.ToLowerInvariant());
            StringAssert.Contains("inboundrule", text.ToLowerInvariant());
        }

        [Test]
        public void TheSettingsHint_SaysWhyItIsClosed_OrThatItIsOpen()
        {
            Assert.AreEqual("closed", GeneralSettingsViewModel.FirewallHintText(new FakeFirewall().Status));
            StringAssert.Contains("can reach port 7437", GeneralSettingsViewModel.FirewallHintText(new FirewallStatus { Port = 7437, Profile = "private", Firewall = "off", NetworkReachable = "yes" }));
        }

        [Test]
        public void TheSettingsHint_IsEmptyWhenNothingIsKnown()
        {
            Assert.IsNull(GeneralSettingsViewModel.FirewallHintText(new FirewallStatus()));
            Assert.IsNull(GeneralSettingsViewModel.FirewallHintText(null));
        }
    }
}
