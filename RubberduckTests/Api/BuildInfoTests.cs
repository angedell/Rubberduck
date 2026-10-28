using System;
using NUnit.Framework;
using Rubberduck;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    public class BuildInfoTests
    {
        [Test]
        public void ShowsAllFourPartsOfTheVersion_AndTheCommit()
        {
            var build = BuildInfo.Describe(new Version(2, 5, 92, 28), "2.5.92.28+00a288206ad07dc6c8d774b0886ff15a8a9bfea3");

            Assert.AreEqual("2.5.92.28", build.Version);
            Assert.AreEqual("00a288206", build.Commit);
            Assert.AreEqual("2.5.92.28 (00a288206)", build.Display);
        }

        [Test]
        public void WithoutACommit_ItIsTheVersionAlone()
        {
            var build = BuildInfo.Describe(new Version(2, 5, 92, 28), "2.5.92.28");

            Assert.IsNull(build.Commit);
            Assert.AreEqual("2.5.92.28", build.Display);
        }

        [Test]
        public void ADirtyMarkAfterTheHash_IsNotPartOfTheCommit()
            => Assert.AreEqual("00a288206", BuildInfo.Describe(new Version(1, 2, 3, 4), "1.2.3.4+00a288206ad07dc6c8d774b0886ff15a8a9bfea3-dirty").Commit);

        [TestCase("1.2.3.4+notahash")]
        [TestCase("1.2.3.4+abc")]
        [TestCase(null)]
        public void SomethingThatIsNotAHash_IsNoCommit(string informational)
            => Assert.IsNull(BuildInfo.Describe(new Version(1, 2, 3, 4), informational).Commit);

        [Test]
        public void AVersionWithMissingParts_StillHasFour()
            => Assert.AreEqual("2.5.0.0", BuildInfo.Describe(new Version(2, 5), null).Version);

        [Test]
        public void TheRunningBuild_CanBeDescribed()
            => Assert.AreEqual(4, BuildInfo.Current.Version.Split('.').Length);

        [Test]
        public void TheApiAndTheSplash_SayTheSame()
        {
            var info = ApiRouter.InfoReport();

            Assert.AreEqual(BuildInfo.Current.Version, info.Version);
            Assert.AreEqual(BuildInfo.Current.Display, info.Display);
            Assert.AreEqual("v1", info.Api);
        }

        [Test]
        public void TheInfoRoute_AnswersWithoutAnyProject()
        {
            var router = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api);

            var response = router.Handle(new ApiRequest { Method = "GET", Path = ApiRouter.Base + "/info", Client = "test" });

            Assert.AreEqual(200, response.Status);
            Assert.AreEqual(BuildInfo.Current.Version, ((InfoReport)response.Body).Version);
        }

        [Test]
        public void TheInfoRoute_IsReadOnly()
        {
            var router = new ApiRouter(new ModuleApi(new FakeStore(), new ApiWriteApproval(_ => true)), Fakes.Diagnostics(), Fakes.Immediate(() => false).Api);

            Assert.AreEqual(405, router.Handle(new ApiRequest { Method = "POST", Path = ApiRouter.Base + "/info", Client = "test" }).Status);
        }
    }
}
