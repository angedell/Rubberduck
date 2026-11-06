using System.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    public class DumpDiagnosisTests
    {
        private static DumpThread Thread(int id, bool main = false, string exception = null, params string[] frames)
            => new DumpThread { ManagedThreadId = id, IsMain = main, Exception = exception, Frames = frames.ToList() };

        private static DumpAnalysis Diagnose(params DumpThread[] threads)
        {
            var analysis = new DumpAnalysis();
            DumpDiagnosis.Fill(analysis, threads);
            return analysis;
        }

        private const string Inspection = "Rubberduck.CodeAnalysis.Inspections.Concrete.ShadowedDeclarationInspection.DoGetInspectionResults";

        [Test]
        public void NothingRecognizable_SaysSoAndAdmitsNativeHangsAreOutOfSight()
        {
            var analysis = Diagnose(Thread(1, main: true, frames: "System.Windows.Forms.Application.Run"), Thread(5, frames: "System.Threading.ThreadPoolWorkQueue.Dispatch"));

            Assert.IsEmpty(analysis.Findings);
            StringAssert.Contains("native code", analysis.Summary);
        }

        [Test]
        public void MainThreadWaitingOnALock_IsAProblem_NamingWhatItWaitsFor()
        {
            var analysis = Diagnose(Thread(1, main: true, frames: new[] { "System.Threading.Monitor.Enter", "Rubberduck.Parsing.VBA.RubberduckParserState.GetDeclarations" }));

            var finding = analysis.Findings.Single();
            Assert.AreEqual("problem", finding.Severity);
            StringAssert.Contains("a lock", finding.Title);
            StringAssert.Contains("RubberduckParserState.GetDeclarations", finding.Detail);
            Assert.AreEqual(finding.Title, analysis.Summary);
        }

        [Test]
        public void MainThreadWaitingWhileAnotherThreadCallsTheHost_IsAProbableDeadlock()
        {
            var analysis = Diagnose(
                Thread(1, main: true, frames: new[] { "System.Threading.Tasks.Task.Wait", "Rubberduck.Parsing.VBA.ParseCoordinator.Parse" }),
                Thread(9, frames: new[] { "System.RuntimeType.InvokeDispMethod", "Rubberduck.VBEditor.SafeComWrappers.VBA.VBComponent.get_Name", "Rubberduck.Parsing.VBA.ComponentParseTask.Parse" }));

            var deadlock = analysis.Findings.Single(f => f.Title.StartsWith("Probable deadlock"));
            Assert.AreEqual("problem", deadlock.Severity);
            CollectionAssert.AreEquivalent(new[] { 1, 9 }, deadlock.Threads);
            StringAssert.Contains("thread 9", deadlock.Detail);
        }

        [Test]
        public void MainThreadWaitingOnWorkOthersAreStillDoing_NamesThatWork()
        {
            var analysis = Diagnose(
                Thread(1, main: true, frames: new[] { "System.Threading.Tasks.Parallel.ForEach", "Rubberduck.CodeAnalysis.Inspections.Inspector.RunInspectionsInParallel" }),
                Thread(7, frames: Inspection));

            Assert.IsTrue(analysis.Findings.Any(f => f.Severity == "problem" && f.Title.Contains("main thread is blocked")));
            Assert.IsTrue(analysis.Findings.Any(f => f.Severity == "warning" && f.Title.Contains("waiting on work that other threads are still doing")));
            Assert.IsFalse(analysis.Findings.Any(f => f.Title.StartsWith("Probable deadlock")));
        }

        [Test]
        public void InspectionsStillRunning_AreNamedOnce_AndBecomeTheSummaryWhenNothingWorseIsThere()
        {
            var other = "Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseInspection.DoGetInspectionResults";
            var analysis = Diagnose(Thread(1, main: true, frames: "Excel.Idle"), Thread(7, frames: Inspection), Thread(8, frames: Inspection), Thread(9, frames: other));

            var finding = analysis.Findings.Single();
            Assert.AreEqual("warning", finding.Severity);
            StringAssert.Contains("2 inspections are still running: ShadowedDeclarationInspection, UnreachableCaseInspection.", finding.Title);
            CollectionAssert.AreEquivalent(new[] { 7, 8, 9 }, finding.Threads);
            Assert.AreEqual(finding.Title, analysis.Summary);
        }

        [Test]
        public void AThreadHandlingAnException_IsReportedWithWhereItWas()
        {
            var analysis = Diagnose(Thread(4, exception: "System.NullReferenceException: Object reference not set", frames: new[] { "Rubberduck.CodeAnalysis.Inspections.Concrete.IIfSideEffectInspection.Extract" }));

            var finding = analysis.Findings.First(f => f.Title.Contains("exception"));
            StringAssert.Contains("NullReferenceException", finding.Title);
            StringAssert.Contains("IIfSideEffectInspection", finding.Detail);
        }

        [Test]
        public void OnlyTheTopFramesSayWhatAThreadWaitsFor()
        {
            var deep = new[] { "A.B", "A.C", "A.D", "A.E", "A.F", "A.G", "A.H", "A.I", "System.Threading.Monitor.Enter" };

            var analysis = Diagnose(Thread(1, main: true, frames: deep));

            Assert.IsEmpty(analysis.Findings);
        }

        [Test]
        public void ThreadsAreSplitIntoThoseWorthALookAndTheRest()
        {
            var analysis = Diagnose(
                Thread(1, main: true, frames: "Excel.Idle"),
                Thread(2, frames: "System.Threading.ThreadPoolWorkQueue.Dispatch"),
                Thread(3, frames: "Rubberduck.Parsing.VBA.ParseCoordinator.Parse"),
                Thread(4, frames: "System.GC.Finalize"));

            Assert.AreEqual(4, analysis.ThreadCount);
            CollectionAssert.AreEqual(new[] { 1, 3 }, analysis.Threads.Select(t => t.ManagedThreadId));
            Assert.AreEqual(2, analysis.OtherThreads);
        }

        [Test]
        public void ThreadsInsideRubberduckThatAreNeitherStuckNorInspecting_AreListedAsInfo()
        {
            var analysis = Diagnose(Thread(3, frames: new[] { "System.Threading.Monitor.Wait", "Rubberduck.Parsing.VBA.ParseCoordinator.Parse" }));

            var finding = analysis.Findings.Single();
            Assert.AreEqual("info", finding.Severity);
            StringAssert.Contains("waiting for a Monitor.Wait signal", finding.Detail);
        }
    }
}
