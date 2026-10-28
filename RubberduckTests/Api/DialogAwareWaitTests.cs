using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Api;

namespace RubberduckTests.Api
{
    [TestFixture]
    public class DialogAwareWaitTests
    {
        private class Dialogs : IDialogWatcher
        {
            public Queue<BlockingDialog> Showing = new Queue<BlockingDialog>();
            public int Acknowledged;
            public bool CanAcknowledge = true;

            public BlockingDialog Find() => Showing.Count > 0 ? Showing.Peek() : null;

            public bool Acknowledge(BlockingDialog dialog)
            {
                if (!CanAcknowledge)
                {
                    return false;
                }

                Acknowledged++;
                Showing.Dequeue();
                return true;
            }
        }

        private TimeSpan _now;

        //Each wait takes the time it was given, as a real one does when what it waits for does not happen.
        private bool Waiting(TimeSpan time)
        {
            _now += time;
            return false;
        }

        private DialogAwareWaitResult Run(Func<TimeSpan, bool> wait, IDialogWatcher dialogs, int seconds = 5)
            => DialogAwareWait.Wait(wait, TimeSpan.FromSeconds(seconds), dialogs, () => _now);

        [Test]
        public void WhatIsWaitedFor_Happening_Completes()
        {
            var result = Run(slice => true, new Dialogs());

            Assert.IsTrue(result.Completed);
            Assert.IsEmpty(result.Errors);
        }

        [Test]
        public void WithoutWatcher_ItIsAPlainWait_ThatEndsWhenTheTimeIsUp()
        {
            var result = Run(Waiting, null);

            Assert.IsFalse(result.Completed);
            Assert.IsNull(result.Blocked);
            Assert.GreaterOrEqual(_now, TimeSpan.FromSeconds(5));
        }

        [Test]
        public void AnEditorErrorDialog_IsClosed_AndItsTextKept_AndTheWaitGoesOn()
        {
            var dialogs = new Dialogs();
            dialogs.Showing.Enqueue(new BlockingDialog { Message = "Compile error: Method or data member not found", IsEditorError = true });
            var calls = 0;

            var result = Run(slice => ++calls > 3, dialogs);

            Assert.IsTrue(result.Completed);
            Assert.AreEqual(1, dialogs.Acknowledged);
            CollectionAssert.AreEqual(new[] { "Compile error: Method or data member not found" }, result.Errors);
        }

        [Test]
        public void TheSameErrorSeenSeveralTimes_IsReportedOnce_WithACount()
        {
            var dialogs = new Dialogs();
            for (var i = 0; i < 7; i++)
            {
                dialogs.Showing.Enqueue(new BlockingDialog { Message = "Compile error: Method or data member not found", IsEditorError = true });
            }

            dialogs.Showing.Enqueue(new BlockingDialog { Message = "Something else", IsEditorError = true });
            var calls = 0;

            var result = Run(slice => ++calls > 10, dialogs);

            CollectionAssert.AreEqual(new[] { "Compile error: Method or data member not found (x7)", "Something else" }, result.Errors);
        }

        [Test]
        public void AMessageBox_EndsTheWait_AndIsReported_BecauseNobodyCanAnswerIt()
        {
            var dialogs = new Dialogs();
            dialogs.Showing.Enqueue(new BlockingDialog { Title = "Excel", Message = "Continue?", IsEditorError = false });

            var result = Run(Waiting, dialogs);

            Assert.IsFalse(result.Completed);
            Assert.AreEqual("Continue?", result.Blocked.Message);
            Assert.AreEqual(0, dialogs.Acknowledged);
        }

        [Test]
        public void AnErrorDialogThatCannotBeClosed_IsReportedAsBlocking()
        {
            var dialogs = new Dialogs { CanAcknowledge = false };
            dialogs.Showing.Enqueue(new BlockingDialog { Message = "Run-time error '11'", IsEditorError = true });

            var result = Run(Waiting, dialogs);

            Assert.IsNotNull(result.Blocked);
            Assert.IsEmpty(result.Errors);
        }

        [Test]
        public void ItLooksOftenEnoughToNotSitOutALongWait()
        {
            var slices = new List<TimeSpan>();

            Run(slice => { slices.Add(slice); _now += slice; return false; }, new Dialogs(), 60);

            Assert.IsTrue(slices.All(s => s <= TimeSpan.FromSeconds(1)));
            Assert.Greater(slices.Count, 10);
        }
    }
}
