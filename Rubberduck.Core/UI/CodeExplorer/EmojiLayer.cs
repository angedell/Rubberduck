using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>
    /// Puts the emoji of a document in colour: each emoji in the text becomes a picture the size of the line's font (<see cref="EmojiImages"/>). The picture is the 128 px PNG;
    /// when the zoom makes it bigger than that on the screen it is swapped for the SVG, which is fetched if it has to be and the connection allows.
    /// </summary>
    public sealed class EmojiLayer
    {
        //The size of the picture next to the font size of its line: emoji are a little taller than the letters.
        private const double SizeFactor = 1.25;
        private const double PngPixels = 128;

        private readonly EmojiImages _images;
        private readonly bool _allowDownload;
        private readonly List<Entry> _entries = new List<Entry>();
        private readonly List<Waiting> _waiting = new List<Waiting>();

        private EmojiLayer(EmojiImages images, bool allowDownload)
        {
            _images = images;
            _allowDownload = allowDownload;
        }

        /// <summary>How many pictures the document got.</summary>
        public int Count => _entries.Count;

        /// <summary>Replaces the emoji in the text of the document with pictures. Code blocks are left as they are.</summary>
        public static EmojiLayer Apply(FlowDocument document, EmojiImages images, bool allowDownload)
        {
            var layer = new EmojiLayer(images, allowDownload);
            var runs = new List<Run>();
            Collect(document, runs);
            foreach (var run in runs)
            {
                layer.Replace(run);
            }

            return layer;
        }

        /// <summary>Tells the layer how far the document is zoomed (100 is normal): pictures that are bigger than their PNG on the screen become SVG, and go back to PNG when they no longer are.</summary>
        public void SetZoom(double percent)
        {
            foreach (var entry in _entries)
            {
                var height = entry.Image.Height;
                if (double.IsNaN(height) || height <= 0)
                {
                    continue;
                }

                var dpi = VisualTreeHelper.GetDpi(entry.Image).DpiScaleY;
                var wantSvg = height * percent / 100d * dpi > PngPixels;
                if (wantSvg == entry.WantsSvg)
                {
                    continue;
                }

                entry.WantsSvg = wantSvg;
                if (!wantSvg)
                {
                    entry.Image.Source = entry.Png;
                    //An SVG that never came (no connection, a metered one) is asked for again the next time the zoom needs it.
                    entry.SvgRequested = entry.Svg != null;
                }
                else if (entry.Svg != null)
                {
                    entry.Image.Source = entry.Svg;
                }
                else if (!entry.SvgRequested)
                {
                    entry.SvgRequested = true;
                    var captured = entry;
                    _images.RequestSvg(entry.Key, _allowDownload, svg =>
                    {
                        captured.Svg = svg;
                        if (captured.WantsSvg)
                        {
                            captured.Image.Source = svg;
                        }
                    });
                }
            }
        }

        private static void Collect(DependencyObject parent, List<Run> runs)
        {
            foreach (var child in LogicalTreeHelper.GetChildren(parent).OfType<DependencyObject>())
            {
                if (child is Paragraph paragraph && Equals(paragraph.Tag, MarkdownRenderer.CodeBlockTag))
                {
                    continue;
                }

                if (child is Run run)
                {
                    if (run.Text.Any(character => character >= 0xA9 || character == '#' || character == '*' || char.IsDigit(character)))
                    {
                        runs.Add(run);
                    }

                    continue;
                }

                if (!(child is InlineUIContainer) && !(child is BlockUIContainer))
                {
                    Collect(child, runs);
                }
            }
        }

        private void Replace(Run run)
        {
            var pieces = EmojiMatcher.Split(run.Text, _images.IsKnown);
            if (!pieces.Any(piece => piece.IsEmoji) || !(run.Parent is TextElement parent))
            {
                return;
            }

            var inlines = (parent as Paragraph)?.Inlines ?? (parent as Span)?.Inlines;
            if (inlines == null)
            {
                return;
            }

            var replacement = new List<Inline>();
            foreach (var piece in pieces)
            {
                var png = piece.IsEmoji ? _images.Png(piece.Key) : null;
                if (png != null)
                {
                    replacement.Add(Picture(parent, piece.Key, piece.Text, png));
                }
                else
                {
                    var text = CopyOf(run, piece.Text);
                    replacement.Add(text);
                    if (piece.IsEmoji)
                    {
                        //Its picture is on the way; until it arrives the emoji is left to the font (see Refresh).
                        _waiting.Add(new Waiting { Key = piece.Key, Text = text });
                    }
                }
            }

            foreach (var inline in replacement)
            {
                inlines.InsertBefore(run, inline);
            }

            inlines.Remove(run);
        }

        /// <summary>
        /// Turns the emoji that were waiting for their pictures into pictures, for those that have arrived (see <see cref="EmojiImages.PictureArrived"/>). Done in place:
        /// the document is not built again, so the reader keeps their place. Call <see cref="SetZoom"/> after it.
        /// </summary>
        public void Refresh()
        {
            foreach (var waiting in _waiting.ToList())
            {
                var png = _images.Png(waiting.Key);
                var parent = waiting.Text.Parent as TextElement;
                var inlines = (parent as Paragraph)?.Inlines ?? (parent as Span)?.Inlines;
                if (png == null || inlines == null)
                {
                    continue;
                }

                inlines.InsertBefore(waiting.Text, Picture(parent, waiting.Key, waiting.Text.Text, png));
                inlines.Remove(waiting.Text);
                _waiting.Remove(waiting);
            }
        }

        /// <summary>Whether some emoji are still waiting for their pictures.</summary>
        public bool IsWaiting => _waiting.Count > 0;

        private Inline Picture(TextElement parent, string key, string text, ImageSource png)
        {
            var image = new Image { Source = png, Stretch = Stretch.Uniform, ToolTip = text };
            image.SetBinding(FrameworkElement.HeightProperty, new Binding(nameof(TextElement.FontSize)) { Source = parent, Converter = SizeConverter.Instance });
            image.SetBinding(FrameworkElement.WidthProperty, new Binding(nameof(TextElement.FontSize)) { Source = parent, Converter = SizeConverter.Instance });
            System.Windows.Automation.AutomationProperties.SetName(image, text);
            //What it stands for, for copying (see EmojiText).
            image.Tag = text;
            _entries.Add(new Entry { Key = key, Image = image, Png = png });
            return new InlineUIContainer(image) { BaselineAlignment = BaselineAlignment.Center };
        }

        //A piece of the text keeps what the run had: the style (inline code), the colours and the font changes locally set.
        private static Run CopyOf(Run run, string text)
        {
            var copy = new Run(text) { Style = run.Style };
            foreach (var property in new[] { TextElement.ForegroundProperty, TextElement.BackgroundProperty, TextElement.FontFamilyProperty, TextElement.FontSizeProperty, TextElement.FontStyleProperty, TextElement.FontWeightProperty, Inline.TextDecorationsProperty })
            {
                var local = run.ReadLocalValue(property);
                if (local != DependencyProperty.UnsetValue && !(local is Expression))
                {
                    copy.SetValue(property, local);
                }
            }

            return copy;
        }

        private sealed class Waiting
        {
            public string Key;
            public Run Text;
        }

        private sealed class Entry
        {
            public string Key;
            public Image Image;
            public ImageSource Png;
            public ImageSource Svg;
            public bool WantsSvg;
            public bool SvgRequested;
        }

        private sealed class SizeConverter : IValueConverter
        {
            public static readonly SizeConverter Instance = new SizeConverter();

            public object Convert(object value, Type targetType, object parameter, CultureInfo culture) => value is double size ? size * SizeFactor : DependencyProperty.UnsetValue;

            public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotSupportedException();
        }
    }
}
