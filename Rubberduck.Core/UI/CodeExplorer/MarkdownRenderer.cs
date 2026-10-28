using System;
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

        public static FlowDocument ToDocument(string markdown)
        {
            var renderer = new WpfRenderer();
            //First in the list, so it is asked before the stock code block renderer.
            renderer.ObjectRenderers.Insert(0, new HighlightingCodeBlockRenderer());
            return Markdig.Wpf.Markdown.ToFlowDocument(markdown ?? string.Empty, Pipeline, renderer);
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
                var paragraph = new Paragraph();
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
