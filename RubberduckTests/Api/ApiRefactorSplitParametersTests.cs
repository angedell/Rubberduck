using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    internal class FakeSplitParameters : ISplitParametersSource
    {
        public List<string> Calls = new List<string>();
        public bool Anything = true;

        public SplitParametersPreview Preview(string project, SplitParametersRequest request)
        {
            Calls.Add("preview");
            return new SplitParametersPreview
            {
                Project = project,
                Members = Anything ? new List<SplitMember> { new SplitMember { Module = "Module1", Member = "Create", Parameters = 5 } } : new List<SplitMember>()
            };
        }

        public SplitParametersResult Apply(string project, SplitParametersRequest request)
        {
            Calls.Add("apply");
            return new SplitParametersResult { Project = project, Split = Preview(project, request).Members };
        }
    }

    [TestFixture]
    [Category("Api")]
    public class RefactorSplitParametersApiTests
    {
        private FakeSplitParameters _source;
        private int _asked;
        private bool _allow;
        private RefactorSplitParametersApi _api;

        [SetUp]
        public void SetUp()
        {
            _source = new FakeSplitParameters();
            _asked = 0;
            _allow = true;
            _api = new RefactorSplitParametersApi(_source, new ApiWriteApproval(_ => { _asked++; return _allow; }));
        }

        [Test]
        public void Preview_ChangesNothing_AsksNobody()
        {
            Assert.AreEqual(1, _api.Preview("p", new SplitParametersRequest()).Members.Count);
            Assert.AreEqual(0, _asked);
            CollectionAssert.AreEqual(new[] { "preview" }, _source.Calls);
        }

        [Test]
        public void Split_LooksFirst_ThenAsksTheUser_ThenSplits()
        {
            var result = _api.Split("agent", "p", new SplitParametersRequest());

            Assert.AreEqual("Create", result.Split.Single().Member);
            Assert.AreEqual(1, _asked);
            Assert.IsTrue(_source.Calls.IndexOf("preview") < _source.Calls.IndexOf("apply"));
        }

        [Test]
        public void Split_WithoutApproval_IsForbidden_AndSplitsNothing()
        {
            _allow = false;

            Assert.AreEqual(403, Assert.Throws<ApiException>(() => _api.Split("agent", "p", new SplitParametersRequest())).Status);
            CollectionAssert.DoesNotContain(_source.Calls, "apply");
        }

        [Test]
        public void WhenThereIsNothingToSplit_NothingIsWritten_AndNobodyIsAsked()
        {
            _source.Anything = false;

            var result = _api.Split("agent", "p", new SplitParametersRequest());

            CollectionAssert.IsEmpty(result.Split);
            CollectionAssert.DoesNotContain(_source.Calls, "apply");
            Assert.AreEqual(0, _asked);
        }

        [Test]
        public void TheRouteTakesNoBody_AndMeansTheWholeProject()
        {
            var response = _api.Route(new ApiRequest { Client = "agent", Body = null }, "POST", new[] { "projects", "p", "refactor", "split-parameters" });

            Assert.AreEqual(200, response.Status);
        }

        [Test]
        public void ThereAreTwoTools()
        {
            CollectionAssert.AreEqual(new[] { "preview_split_parameters", "split_parameters" }, _api.Tools().Select(t => t.Name));
        }
    }
}
