using System.Linq;
using System.Threading;
using System.Windows.Documents;
using System.Windows.Media;
using NUnit.Framework;
using Rubberduck.UI.CodeExplorer;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class MarkdownRendererTests
    {
        [Test]
        public void ToDocument_RendersHeadingsAndParagraphs()
        {
            var document = MarkdownRenderer.ToDocument("# Title\r\n\r\nSome **bold** text.");

            var text = new TextRange(document.ContentStart, document.ContentEnd).Text;
            StringAssert.Contains("Title", text);
            StringAssert.Contains("Some bold text.", text);
            Assert.IsTrue(document.Blocks.Count >= 2);
        }

        [Test]
        public void ToDocument_ColoursVbaFences()
        {
            var document = MarkdownRenderer.ToDocument("```vba\r\nIf True Then ' note\r\n```");

            var runs = document.Blocks.SelectMany(block => (block as Paragraph)?.Inlines.OfType<Run>() ?? Enumerable.Empty<Run>()).ToList();
            var keyword = runs.FirstOrDefault(run => run.Text == "If");
            var comment = runs.FirstOrDefault(run => run.Text.StartsWith("'"));
            Assert.IsNotNull(keyword, "keyword run; blocks: " + string.Join(" | ", document.Blocks.Select(b => b.GetType().Name + ":" + new TextRange(b.ContentStart, b.ContentEnd).Text)) + " runs: " + string.Join(",", runs.Select(r => "[" + r.Text + "]")));
            Assert.IsNotNull(comment, "comment run");
            Assert.AreEqual(Colors.DarkBlue, ((SolidColorBrush)keyword.Foreground).Color);
            Assert.AreEqual(Colors.Green, ((SolidColorBrush)comment.Foreground).Color);
        }

        [Test]
        public void Highlight_Vba_ColoursKeywords()
        {
            var names = string.Join(",", typeof(ResourceSyntaxColors).Assembly.GetManifestResourceNames().Where(n => n.EndsWith(".xshd")));
            Assert.IsNotNull(ResourceSyntaxColors.Vba, "definition; resources: " + names);
            var runs = MarkdownRenderer.Highlight("If True Then ' note", "vba");
            Assert.IsNotNull(runs);
            Assert.AreEqual("If", runs[0].Text);
        }

        [Test]
        public void Highlight_UnknownLanguage_GivesNone()
        {
            Assert.IsNull(MarkdownRenderer.Highlight("anything", "brainfuck"));
            Assert.IsNull(MarkdownRenderer.Highlight("anything", null));
        }

        [Test]
        public void ToDocument_UnknownFenceStillShowsTheCode()
        {
            var document = MarkdownRenderer.ToDocument("```text\r\nplain code\r\n```");

            StringAssert.Contains("plain code", new TextRange(document.ContentStart, document.ContentEnd).Text);
        }
    }
}
