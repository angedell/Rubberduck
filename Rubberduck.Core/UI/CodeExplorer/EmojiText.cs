using System;
using System.Text;
using System.Windows;
using System.Windows.Documents;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>The text of a selection in a preview, with every emoji picture (see <see cref="EmojiLayer"/>) back as the emoji character it stands for.</summary>
    public static class EmojiText
    {
        public static string Of(TextRange range)
        {
            //The text of a range has a blank where a picture is, so the range is walked instead; the plain text is good enough when there is no picture in it.
            var result = new StringBuilder();
            var pictures = 0;
            for (var position = range.Start; position != null && position.CompareTo(range.End) < 0; position = position.GetNextContextPosition(LogicalDirection.Forward))
            {
                switch (position.GetPointerContext(LogicalDirection.Forward))
                {
                    case TextPointerContext.Text:
                        var text = position.GetTextInRun(LogicalDirection.Forward);
                        var room = position.GetOffsetToPosition(range.End);
                        result.Append(room < text.Length ? text.Substring(0, Math.Max(0, room)) : text);
                        break;
                    case TextPointerContext.ElementStart:
                        var element = position.GetAdjacentElement(LogicalDirection.Forward);
                        if (element is InlineUIContainer container && container.Child is FrameworkElement picture && picture.Tag is string character)
                        {
                            result.Append(character);
                            pictures++;
                        }
                        else if (element is LineBreak)
                        {
                            result.Append("\r\n");
                        }

                        break;
                    case TextPointerContext.ElementEnd:
                        if (position.GetAdjacentElement(LogicalDirection.Forward) is Paragraph)
                        {
                            result.Append("\r\n");
                        }

                        break;
                }
            }

            return pictures == 0 ? range.Text : result.ToString();
        }
    }
}
