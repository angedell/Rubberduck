using System;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    public class ApiDiagnosticsWaitTests
    {
        private class Source : IDiagnosticsSource
        {
            public bool Running;
            public TimeSpan? AskedWait;

            public DiagnosticsReport Run(string project, string module, TimeSpan? wait = null)
            {
                AskedWait = wait;
                return Running
                    ? new DiagnosticsReport { Project = project, ParserState = "Ready", Status = "running", RetryAfterSeconds = 10 }
                    : new DiagnosticsReport { Project = project, ParserState = "Ready" };
            }
        }

        private Source _source;
        private ApiRouter _router;

        [SetUp]
        public void SetUp()
        {
            _source = new Source();
            _router = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), new DiagnosticsApi(_source), Fakes.Immediate(() => false).Api);
        }

        private ApiResponse Get(string path, string waitSeconds = null)
        {
            var request = new ApiRequest { Method = "GET", Path = path };
            if (waitSeconds != null)
            {
                request.Query["waitSeconds"] = waitSeconds;
            }

            return _router.Handle(request);
        }

        [Test]
        public void WhenTheInspectionsAreStillRunning_TheAnswerIs202_WithNoResultsYet()
        {
            _source.Running = true;

            var response = Get("/api/v1/projects/p1/diagnostics");

            Assert.AreEqual(202, response.Status);
            Assert.AreEqual("running", ((DiagnosticsReport)response.Body).Status);
        }

        [Test]
        public void WhenTheyAreDone_TheAnswerIs200_AndHasNoStatus()
        {
            var response = Get("/api/v1/projects/p1/diagnostics");

            Assert.AreEqual(200, response.Status);
            Assert.IsNull(((DiagnosticsReport)response.Body).Status);
        }

        [Test]
        public void WaitSeconds_IsPassedOn_AndZeroMeansDoNotWait()
        {
            Get("/api/v1/projects/p1/diagnostics", "0");

            Assert.AreEqual(TimeSpan.Zero, _source.AskedWait);
        }

        [Test]
        public void WithoutWaitSeconds_TheSourceChoosesItsDefault()
        {
            Get("/api/v1/projects/p1/diagnostics");

            Assert.IsNull(_source.AskedWait);
        }

        [TestCase("-1")]
        [TestCase("301")]
        [TestCase("soon")]
        public void AnInvalidWaitSeconds_IsABadRequest(string wait)
        {
            Assert.AreEqual(400, Get("/api/v1/projects/p1/diagnostics", wait).Status);
        }
    }
}
