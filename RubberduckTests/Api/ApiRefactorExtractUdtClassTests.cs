using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    internal class FakeExtractUdtClass : IExtractUdtClassSource
    {
        public List<string> Calls = new List<string>();
        public List<string> Problems = new List<string>();
        public ExtractUdtClassRequest Last;

        public ExtractUdtClassPreview Preview(string project, ExtractUdtClassRequest request)
        {
            Calls.Add("preview");
            Last = request;
            return new ExtractUdtClassPreview { Project = project, Type = "TPerson", ClassName = "Person", Problems = Problems.ToList() };
        }

        public ExtractUdtClassResult Apply(string project, ExtractUdtClassRequest request)
        {
            Calls.Add("apply");
            return new ExtractUdtClassResult { Project = project, Type = "TPerson", ClassName = "Person" };
        }
    }

    [TestFixture]
    [Category("Api")]
    public class RefactorExtractUdtClassApiTests
    {
        private FakeExtractUdtClass _source;
        private int _asked;
        private bool _allow;
        private RefactorExtractUdtClassApi _api;

        private static ExtractUdtClassRequest Request(string style = null) => new ExtractUdtClassRequest { Module = "Module1", Name = "TPerson", Style = style };

        [SetUp]
        public void SetUp()
        {
            _source = new FakeExtractUdtClass();
            _asked = 0;
            _allow = true;
            _api = new RefactorExtractUdtClassApi(_source, new ApiWriteApproval(_ => { _asked++; return _allow; }));
        }

        [Test]
        public void Preview_ChangesNothing_AsksNobody()
        {
            _api.Preview("p", Request());

            Assert.AreEqual(0, _asked);
            CollectionAssert.AreEqual(new[] { "preview" }, _source.Calls);
        }

        [Test]
        public void Extract_LooksFirst_ThenAsksTheUser_ThenAdds()
        {
            var result = _api.Extract("agent", "p", Request(ExtractUdtClassRequest.Fluent));

            Assert.AreEqual("Person", result.ClassName);
            Assert.AreEqual(1, _asked);
            CollectionAssert.AreEqual(new[] { "preview", "apply" }, _source.Calls);
        }

        [Test]
        public void Extract_WithoutApproval_IsForbidden_AndAddsNothing()
        {
            _allow = false;

            Assert.AreEqual(403, Assert.Throws<ApiException>(() => _api.Extract("agent", "p", Request())).Status);
            CollectionAssert.DoesNotContain(_source.Calls, "apply");
        }

        [Test]
        public void AProblemInThePreview_RefusesTheRequest_BeforeTheUserIsAsked()
        {
            _source.Problems.Add("'Self' clashes with the Self property of the class.");

            var failure = Assert.Throws<ApiException>(() => _api.Extract("agent", "p", Request()));

            Assert.AreEqual(422, failure.Status);
            StringAssert.Contains("clashes", failure.Message);
            Assert.AreEqual(0, _asked);
            CollectionAssert.DoesNotContain(_source.Calls, "apply");
        }

        [Test]
        public void TheModuleIsRequired()
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Preview("p", new ExtractUdtClassRequest { Name = "TPerson" })).Status);
        }

        [Test]
        public void TheStyleIsPlainOrFluent()
        {
            Assert.AreEqual(400, Assert.Throws<ApiException>(() => _api.Preview("p", Request("chatty"))).Status);
            Assert.DoesNotThrow(() => _api.Preview("p", Request("Fluent")));
        }

        [Test]
        public void TheRouteReadsTheBody()
        {
            var body = "{\"module\":\"Module1\",\"name\":\"TPerson\",\"className\":\"Human\",\"style\":\"fluent\"}";

            var response = _api.Route(new ApiRequest { Client = "agent", Body = body }, "POST", new[] { "projects", "p", "refactor", "extract-udt-class", "preview" });

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual("Human", _source.Last.ClassName);
            Assert.AreEqual("fluent", _source.Last.Style);
        }

        [Test]
        public void ThereAreTwoTools()
        {
            CollectionAssert.AreEqual(new[] { "preview_extract_udt_class", "extract_udt_class" }, _api.Tools().Select(t => t.Name));
        }
    }
}
