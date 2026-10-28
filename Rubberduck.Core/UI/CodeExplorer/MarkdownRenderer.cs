using System;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Highlighting;
using Markdig;
using Markdig.Renderers;
using Markdig.Syntax;


namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>Turns markdown into a document to show. Code in fences is coloured by the same definitions the editors use (see <see cref="ResourceSyntaxColors"/>).</summary>
    public static class MarkdownRenderer
    {
        private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder().UseSupportedExtensions().Build();

        //Fonts WPF falls back to, in order, for a character the chosen font lacks. Whichever are installed are used: the newest emoji need a recent Segoe UI Emoji (Windows 11) or one of the others.
        //WPF draws a font's glyphs in the text colour; colour glyphs (COLR/CBDT) are not supported by the framework, so emoji come out as monochrome outlines whatever the font.
        private static readonly string[] EmojiFonts = { "Segoe UI Emoji", "Noto Emoji", "Noto Color Emoji", "Twemoji Mozilla", "Segoe UI Symbol" };

        /// <summary>The font family string for the document: the chosen font, then the emoji fonts.</summary>
        public static string FontChain(string fontName)
        {
            var name = string.IsNullOrWhiteSpace(fontName) ? ResourceEditorPreferences.DefaultPreviewFontName : fontName.Trim();
            return string.Join(", ", new[] { name }.Concat(EmojiFonts).Distinct(StringComparer.OrdinalIgnoreCase));
        }

        /// <summary>Marks the paragraph of a code block, where emoji stay text.</summary>
        public const string CodeBlockTag = "code-block";

        /// <summary>
        /// The document for the markdown. Emoji are pictures in colour (see <see cref="EmojiLayer"/>); the layer is in the document's Tag, and wants to be told the zoom.
        /// <paramref name="allowDownload"/> lets a big zoom fetch the SVG of an emoji from the internet, where the connection is not metered.
        /// </summary>
        public static FlowDocument ToDocument(string markdown, string fontName = null, bool allowDownload = false)
        {
            var renderer = new WpfRenderer();
            //First in the list, so it is asked before the stock code block renderer.
            renderer.ObjectRenderers.Insert(0, new HighlightingCodeBlockRenderer());
            var document = Markdig.Wpf.Markdown.ToFlowDocument(markdown ?? string.Empty, Pipeline, renderer);
            document.FontFamily = new FontFamily(FontChain(fontName));
            document.Tag = EmojiLayer.Apply(document, EmojiImages.Shared, allowDownload);
            if (allowDownload)
            {
                //The emoji that are not packaged are fetched in the background, the ones this document waits for first.
                EmojiImages.Shared.StartFetching();
            }

            return document;
        }

        private static MarkdownPipelineBuilder UseSupportedExtensions(this MarkdownPipelineBuilder builder)
            => builder.UsePipeTables().UseEmphasisExtras().UseTaskLists().UseAutoLinks();

        /// <summary>The runs for the code, one brush per character as the highlighter has it; null when no definition is known for the language.</summary>
        public static Run[] Highlight(string code, string language)
        {
            var definition = ResourceSyntaxColors.ForLanguage(language);
            if (definition == null || string.IsNullOrEmpty(code))
            {
                return null;
            }

            var document = new TextDocument(code);
            var highlighter = new DocumentHighlighter(document, definition);
            var runs = new System.Collections.Generic.List<Run>();

            for (var number = 1; number <= document.LineCount; number++)
            {
                var line = document.GetLineByNumber(number);
                var brushes = new Brush[line.Length];
                foreach (var section in highlighter.HighlightLine(number).Sections)
                {
                    var brush = section.Color?.Foreground?.GetBrush(null);
                    if (brush == null)
                    {
                        continue;
                    }

                    for (var offset = Math.Max(section.Offset, line.Offset); offset < Math.Min(section.Offset + section.Length, line.EndOffset); offset++)
                    {
                        brushes[offset - line.Offset] = brush;
                    }
                }

                var start = 0;
                while (start < brushes.Length)
                {
                    var end = start + 1;
                    while (end < brushes.Length && ReferenceEquals(brushes[end], brushes[start]))
                    {
                        end++;
                    }

                    var run = new Run(document.GetText(line.Offset + start, end - start));
                    if (brushes[start] != null)
                    {
                        run.Foreground = brushes[start];
                    }

                    runs.Add(run);
                    start = end;
                }

                if (number < document.LineCount)
                {
                    runs.Add(new Run("\n"));
                }
            }

            return runs.ToArray();
        }

        private sealed class HighlightingCodeBlockRenderer : MarkdownObjectRenderer<WpfRenderer, CodeBlock>
        {
            protected override void Write(WpfRenderer renderer, CodeBlock block)
            {
                var paragraph = new Paragraph { Tag = CodeBlockTag };
                paragraph.SetResourceReference(FrameworkContentElement.StyleProperty, Markdig.Wpf.Styles.CodeBlockStyleKey);

                var code = block.Lines.ToString();
                var language = (block as FencedCodeBlock)?.Info;
                var runs = Highlight(code, language);
                if (runs != null)
                {
                    paragraph.Inlines.AddRange(runs);
                }
                else
                {
                    paragraph.Inlines.Add(new Run(code));
                }

                renderer.WriteBlock(paragraph);
            }
        }
    }
}
