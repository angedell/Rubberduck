using System;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    /// <summary>
    /// Reads a real dump, so it runs only when RUBBERDUCK_TEST_DUMP names one (a full dump of an Excel process with Rubberduck loaded,
    /// taken by the .NET runtime installed on this machine, from POST /api/v1/dump).
    /// </summary>
    [TestFixture]
    public class ClrMdDumpAnalyzerTests
    {
        private static string DumpFile
        {
            get
            {
                var file = Environment.GetEnvironmentVariable("RUBBERDUCK_TEST_DUMP");
                if (string.IsNullOrWhiteSpace(file))
                {
                    Assert.Ignore("RUBBERDUCK_TEST_DUMP is not set.");
                }

                return file;
            }
        }

        [Test]
        public void ReadsTheManagedThreadsOfARealDump()
        {
            var analysis = new ClrMdDumpAnalyzer().Analyze(DumpFile, CancellationToken.None);

            TestContext.WriteLine($"Runtime {analysis.Runtime}, {analysis.Architecture}, {analysis.ThreadCount} threads ({analysis.OtherThreads} not shown)");
            TestContext.WriteLine("Summary: " + analysis.Summary);
            foreach (var finding in analysis.Findings)
            {
                TestContext.WriteLine($"[{finding.Severity}] {finding.Title} {finding.Detail}");
            }

            foreach (var thread in analysis.Threads)
            {
                TestContext.WriteLine($"thread {thread.ManagedThreadId} main={thread.IsMain} waiting={thread.Waiting} rubberduck={thread.RubberduckFrame}");
            }

            Assert.IsNotNull(analysis.Summary);
            Assert.Greater(analysis.ThreadCount, 0, analysis.Note);
        }

        [Test]
        public void ACancelledRead_IsReportedAsATimeout()
        {
            var file = DumpFile;
            using (var cancelled = new CancellationTokenSource())
            {
                cancelled.Cancel();

                var failure = Assert.Throws<ApiException>(() => new ClrMdDumpAnalyzer().Analyze(file, cancelled.Token));

                Assert.AreEqual(503, failure.Status);
            }
        }

        [Test]
        public void AFileThatIsNoDump_IsReportedNotThrownRaw()
        {
            var failure = Assert.Throws<ApiException>(() => new ClrMdDumpAnalyzer().Analyze(System.Reflection.Assembly.GetExecutingAssembly().Location, CancellationToken.None));

            Assert.AreEqual(503, failure.Status);
            StringAssert.Contains("could not be read", failure.Message);
        }
    }
}
