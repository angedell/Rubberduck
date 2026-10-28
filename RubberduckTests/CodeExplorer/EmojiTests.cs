using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Documents;
using NUnit.Framework;
using Rubberduck.UI.CodeExplorer;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    public class EmojiMatcherTests
    {
        private static readonly string[] Known = { "1f680", "1fab5", "2764", "1f44d_1f3fd", "1f468_200d_1f469_200d_1f467", "0031_20e3", "0031", "00a9", "1f1fa_1f1f8" };

        private static string[] Pieces(string text)
            => EmojiMatcher.Split(text, Known.Contains).Select(piece => (piece.IsEmoji ? "[" + piece.Key + "]" : piece.Text)).ToArray();

        [Test]
        public void PlainText_IsOnePiece()
            => CollectionAssert.AreEqual(new[] { "no emoji here" }, Pieces("no emoji here"));

        [Test]
        public void AnEmoji_IsSplitOutOfTheText()
            => CollectionAssert.AreEqual(new[] { "a ", "[1fab5]", " b" }, Pieces("a \U0001FAB5 b"));

        [Test]
        public void Sequences_AreOneEmoji_NotTheirParts()
        {
            CollectionAssert.AreEqual(new[] { "[1f44d_1f3fd]" }, Pieces("\U0001F44D\U0001F3FD"));
            CollectionAssert.AreEqual(new[] { "[1f468_200d_1f469_200d_1f467]" }, Pieces("\U0001F468‍\U0001F469‍\U0001F467"));
            CollectionAssert.AreEqual(new[] { "[1f1fa_1f1f8]" }, Pieces("\U0001F1FA\U0001F1F8"));
        }

        [Test]
        public void ARedHeartNeedsItsSelector_ElseItIsTheTextHeart()
        {
            CollectionAssert.AreEqual(new[] { "[2764]" }, Pieces("❤️"));
            CollectionAssert.AreEqual(new[] { "❤" }, Pieces("❤"));
        }

        [Test]
        public void DigitsAndSymbols_StayText_UnlessTheyAreKeycaps()
        {
            CollectionAssert.AreEqual(new[] { "1 © 2" }, Pieces("1 © 2"));
            CollectionAssert.AreEqual(new[] { "[0031_20e3]" }, Pieces("1️⃣"));
        }

        [Test]
        public void AnEmojiWithoutAPicture_StaysText()
            => CollectionAssert.AreEqual(new[] { "\U0001F9FF" }, Pieces("\U0001F9FF"));

        [Test]
        public void TextOf_IsTheInverseOfTheKey()
            => Assert.AreEqual("\U0001F468‍\U0001F469", EmojiMatcher.TextOf("1f468_200d_1f469"));
    }

    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class EmojiImagesTests
    {
        [Test]
        public void ThePackage_HasTheNewestEmoji_AndTheirPictures()
        {
            Assert.IsTrue(EmojiImages.Shared.IsKnown("1fab5"), "wood");
            Assert.IsTrue(EmojiImages.Shared.IsKnown("1f680"), "rocket");
            Assert.IsNotNull(EmojiImages.Shared.Png("1fab5"));
            Assert.IsNull(EmojiImages.Shared.Png("ffffff"));
        }

        [Test]
        public void ToDocument_PutsAPictureWhereTheEmojiWas_AndKeepsTheText()
        {
            var document = MarkdownRenderer.ToDocument("# Logs \U0001FAB5\r\n\r\n- \U0001FAB5 docs/debug-logging.md: where it is", "Segoe UI");

            var pictures = Pictures(document);
            Assert.AreEqual(2, pictures.Count);
            var text = new TextRange(document.ContentStart, document.ContentEnd).Text;
            StringAssert.Contains("Logs", text);
            StringAssert.Contains("docs/debug-logging.md: where it is", text);
            Assert.AreEqual(2, ((EmojiLayer)document.Tag).Count);
        }

        [Test]
        public void ToDocument_LeavesCodeBlocksAlone()
        {
            var document = MarkdownRenderer.ToDocument("```text\r\nlog \U0001FAB5\r\n```");

            Assert.AreEqual(0, Pictures(document).Count);
        }

        [Test]
        public void TheSvg_ComesFromTheFolderWithoutAnyDownload()
        {
            var folder = Path.Combine(Path.GetTempPath(), "emoji-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(folder, "svg"));
            try
            {
                File.WriteAllText(Path.Combine(folder, "svg", "emoji_u1fab5.svg"), "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 128 128\"><rect width=\"128\" height=\"128\" fill=\"#a52a2a\"/></svg>");
                var images = new EmojiImages(() => null, () => null, folder, () => throw new InvalidOperationException("must not ask the network"));

                System.Windows.Media.ImageSource svg = null;
                images.RequestSvg("1fab5", false, source => svg = source);

                Assert.IsNotNull(svg);
            }
            finally
            {
                Directory.Delete(folder, true);
            }
        }

        [Test]
        public void TheSvg_IsNotFetched_WhenTheConnectionIsMeteredOrDownloadsAreOff()
        {
            var folder = Path.Combine(Path.GetTempPath(), "emoji-" + Guid.NewGuid().ToString("N"));
            var downloads = 0;
            Task<byte[]> Download(string address) { downloads++; return Task.FromResult(new byte[] { 1 }); }

            var metered = new EmojiImages(() => null, () => null, folder, () => false, Download);
            metered.RequestSvg("1fab5", true, source => Assert.Fail("metered"));
            var off = new EmojiImages(() => null, () => null, folder, () => true, Download);
            off.RequestSvg("1fab5", false, source => Assert.Fail("off"));

            Assert.AreEqual(0, downloads);
            Assert.IsFalse(Directory.Exists(folder));
        }

        [Test]
        public void TheSvg_IsFetchedKeptAndDelivered_WhenTheConnectionAllows()
        {
            var folder = Path.Combine(Path.GetTempPath(), "emoji-" + Guid.NewGuid().ToString("N"));
            var svg = System.Text.Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 128 128\"><circle cx=\"64\" cy=\"64\" r=\"60\" fill=\"#00f\"/></svg>");
            try
            {
                var images = new EmojiImages(() => null, () => null, folder, () => true, address =>
                {
                    StringAssert.EndsWith("/2D/svg/emoji_u1fab5.svg", address);
                    return Task.FromResult(svg);
                });
                using (var delivered = new ManualResetEventSlim())
                {
                    System.Windows.Media.ImageSource got = null;
                    SynchronizationContext.SetSynchronizationContext(new InlineContext());
                    images.RequestSvg("1fab5", true, source => { got = source; delivered.Set(); });

                    Assert.IsTrue(delivered.Wait(TimeSpan.FromSeconds(10)), "delivered");
                    Assert.IsNotNull(got);
                    Assert.IsTrue(File.Exists(Path.Combine(folder, "svg", "emoji_u1fab5.svg")), "kept");
                }
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(null);
                if (Directory.Exists(folder))
                {
                    Directory.Delete(folder, true);
                }
            }
        }

        [Test]
        public void ThePackage_HoldsTheCore_AndTheIndexKnowsTheRest()
        {
            var images = EmojiImages.Shared;

            Assert.IsNotNull(images.Png("1f680"), "a single character is packaged");
            Assert.IsNotNull(images.Png("0031_20e3"), "a keycap is packaged");
            Assert.IsTrue(images.IsKnown("1f44d_1f3fd"), "a skin tone is known from the index");
            Assert.IsTrue(images.IsKnown("1f468_200d_1f469_200d_1f467"), "a family is known from the index");
            Assert.IsTrue(images.Missing().Contains("1f44d_1f3fd"), "and is to be fetched");
            Assert.IsFalse(images.Missing().Contains("1f680"));
        }

        [Test]
        public void AnEmojiWithoutItsPicture_IsLeftAsText_ThenBecomesAPictureInPlace()
        {
            var folder = Path.Combine(Path.GetTempPath(), "emoji-" + Guid.NewGuid().ToString("N"));
            var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
            using (var arrived = new ManualResetEventSlim())
            {
                var images = new EmojiImages(() => null, () => new MemoryStream(System.Text.Encoding.ASCII.GetBytes("1f44d_1f3fd\n")), folder, () => true,
                    address => Task.FromResult(png), isIdle: () => true, pause: TimeSpan.Zero);
                try
                {
                    SynchronizationContext.SetSynchronizationContext(new InlineContext());
                    images.PictureArrived += key => arrived.Set();
                    Assert.IsNull(images.Png("1f44d_1f3fd"), "not here yet");

                    var document = new FlowDocument(new Paragraph(new Run("ok \U0001F44D\U0001F3FD")));
                    var layer = EmojiLayer.Apply(document, images, true);
                    Assert.IsTrue(layer.IsWaiting);
                    Assert.AreEqual(0, Pictures(document).Count);

                    images.StartFetching();
                    Assert.IsTrue(arrived.Wait(TimeSpan.FromSeconds(10)), "arrived");
                    layer.Refresh();

                    Assert.IsFalse(layer.IsWaiting);
                    Assert.AreEqual(1, Pictures(document).Count);
                    Assert.IsTrue(File.Exists(Path.Combine(folder, "png", "emoji_u1f44d_1f3fd.png")), "kept for the next time");
                    StringAssert.StartsWith("ok ", new TextRange(document.ContentStart, document.ContentEnd).Text);
                }
                finally
                {
                    images.StopFetching();
                    SynchronizationContext.SetSynchronizationContext(null);
                    if (Directory.Exists(folder))
                    {
                        Directory.Delete(folder, true);
                    }
                }
            }
        }

        [Test]
        public void TheBackgroundFetching_WaitsForTheUserToBeIdle_AndForAnUnmeteredConnection()
        {
            var folder = Path.Combine(Path.GetTempPath(), "emoji-" + Guid.NewGuid().ToString("N"));
            var downloads = 0;
            var images = new EmojiImages(() => null, () => new MemoryStream(System.Text.Encoding.ASCII.GetBytes("1f44d_1f3fd\n")), folder, () => false,
                address => { downloads++; return Task.FromResult(new byte[] { 1 }); }, isIdle: () => false, pause: TimeSpan.Zero);
            try
            {
                images.StartFetching();
                Thread.Sleep(500);

                Assert.AreEqual(0, downloads, "metered and busy");
            }
            finally
            {
                images.StopFetching();
            }
        }

        [Test]
        public void Copying_GivesTheEmojiAsCharacters()
        {
            var document = MarkdownRenderer.ToDocument("Wood \U0001FAB5 and a rocket \U0001F680 here");
            var all = new TextRange(document.ContentStart, document.ContentEnd);

            var copied = EmojiText.Of(all);

            StringAssert.Contains("Wood \U0001FAB5 and a rocket \U0001F680 here", copied);
            Assert.AreEqual(2, Pictures(document).Count, "they are pictures in the document");
        }

        private static System.Collections.Generic.List<Image> Pictures(FlowDocument document)
        {
            var found = new System.Collections.Generic.List<Image>();
            void Walk(System.Windows.DependencyObject parent)
            {
                foreach (var child in System.Windows.LogicalTreeHelper.GetChildren(parent).OfType<System.Windows.DependencyObject>())
                {
                    if (child is InlineUIContainer container && container.Child is Image image)
                    {
                        found.Add(image);
                    }
                    else
                    {
                        Walk(child);
                    }
                }
            }

            Walk(document);
            return found;
        }

        //Posts run right away, on the thread that posts: enough for a test that has no message loop.
        private sealed class InlineContext : SynchronizationContext
        {
            public override void Post(SendOrPostCallback d, object state) => d(state);
        }
    }
}
