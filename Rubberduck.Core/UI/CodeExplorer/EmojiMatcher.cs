using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>One stretch of a text: plain, or one emoji (a character or a sequence such as a flag, a skin tone or a family) with the key of its picture.</summary>
    public struct EmojiPiece
    {
        public EmojiPiece(string text, string key)
        {
            Text = text;
            Key = key;
        }

        public string Text { get; }

        /// <summary>The code points in lower-case hex, joined by underscores, without the variation selector: the name of the picture. Null for plain text.</summary>
        public string Key { get; }

        public bool IsEmoji => Key != null;
    }

    /// <summary>Finds the emoji in a text. Which sequences exist is decided by the pictures there are (<see cref="EmojiImages"/>), so the matching never claims one that cannot be drawn.</summary>
    public static class EmojiMatcher
    {
        private const int VariationSelector16 = 0xFE0F;
        private const int MaxCodePoints = 16;
        //Below this, a single code point is emoji only when asked for with the variation selector: © ™ ↔ and the digits are text otherwise.
        private const int FirstPictograph = 0x1F000;

        public static IReadOnlyList<EmojiPiece> Split(string text, Func<string, bool> isKnown)
        {
            var pieces = new List<EmojiPiece>();
            if (string.IsNullOrEmpty(text))
            {
                return pieces;
            }

            var codePoints = new List<int>();
            var offsets = new List<int>();
            for (var i = 0; i < text.Length; i++)
            {
                offsets.Add(i);
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    codePoints.Add(char.ConvertToUtf32(text[i], text[i + 1]));
                    i++;
                }
                else
                {
                    codePoints.Add(text[i]);
                }
            }

            offsets.Add(text.Length);
            var plainFrom = 0;
            var index = 0;
            while (index < codePoints.Count)
            {
                string key = null;
                var length = CanStart(codePoints[index]) ? LongestAt(codePoints, index, isKnown, out key) : 0;
                if (length == 0)
                {
                    index++;
                    continue;
                }

                if (offsets[index] > plainFrom)
                {
                    pieces.Add(new EmojiPiece(text.Substring(plainFrom, offsets[index] - plainFrom), null));
                }

                pieces.Add(new EmojiPiece(text.Substring(offsets[index], offsets[index + length] - offsets[index]), key));
                index += length;
                plainFrom = offsets[index];
            }

            if (plainFrom < text.Length)
            {
                pieces.Add(new EmojiPiece(text.Substring(plainFrom), null));
            }

            return pieces;
        }

        private static bool CanStart(int codePoint) => codePoint >= 0xA9 || codePoint == '#' || codePoint == '*' || (codePoint >= '0' && codePoint <= '9');

        private static int LongestAt(IList<int> codePoints, int start, Func<string, bool> isKnown, out string key)
        {
            for (var length = Math.Min(MaxCodePoints, codePoints.Count - start); length > 0; length--)
            {
                var kept = new List<int>(length);
                for (var i = start; i < start + length; i++)
                {
                    if (codePoints[i] != VariationSelector16)
                    {
                        kept.Add(codePoints[i]);
                    }
                }

                //A sequence cannot end in a bare joiner, and cannot begin with a selector.
                if (kept.Count == 0)
                {
                    continue;
                }

                if (kept.Count == 1 && kept[0] < FirstPictograph && !(length > 1 && codePoints[start + 1] == VariationSelector16))
                {
                    continue;
                }

                var candidate = string.Join("_", kept.Select(code => code.ToString("x4")));
                if (isKnown(candidate))
                {
                    key = candidate;
                    return length;
                }
            }

            key = null;
            return 0;
        }

        /// <summary>The characters of a key, e.g. "1f468_200d_1f469" back to the text.</summary>
        public static string TextOf(string key)
        {
            var builder = new StringBuilder();
            foreach (var part in key.Split('_'))
            {
                builder.Append(char.ConvertFromUtf32(Convert.ToInt32(part, 16)));
            }

            return builder.ToString();
        }
    }
}
