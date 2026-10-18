using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeTranslate.Core
{
    // Only completed Chat answers use this helper. Preserve Markdown source and
    // line endings verbatim; never log the input or run it on user messages.
    internal static class ChatCjkSpacing
    {
        internal static string Clean(string text)
        {
            var result = new StringBuilder(text.Length);
            var proseStart = 0; var fenceLength = 0; var fenceChar = '\0';
            for (var start = 0; start < text.Length;)
            {
                var end = start;
                while (end < text.Length && text[end] != '\r' && text[end] != '\n') end++;
                var next = end;
                if (next < text.Length && text[next++] == '\r' && next < text.Length && text[next] == '\n') next++;
                var marker = start;
                while (marker < end && text[marker] == ' ') marker++;
                var count = 0;
                if (marker < end && (text[marker] == '`' || text[marker] == '~'))
                    while (marker + count < end && text[marker + count] == text[marker]) count++;
                if (fenceLength == 0 && marker - start <= 3 && count >= 3)
                {
                    AppendProse(result, text.Substring(proseStart, start - proseStart));
                    fenceLength = count; fenceChar = text[marker]; proseStart = start;
                }
                else if (fenceLength > 0 && count >= fenceLength && text[marker] == fenceChar && OnlyFenceWhitespace(text, marker + count, end))
                {
                    result.Append(text, proseStart, next - proseStart);
                    proseStart = next; fenceLength = 0;
                }
                start = next;
            }
            if (fenceLength > 0) result.Append(text, proseStart, text.Length - proseStart);
            else AppendProse(result, text.Substring(proseStart));
            return result.ToString();
        }

        private static bool OnlyFenceWhitespace(string text, int start, int end)
        {
            for (var i = start; i < end; i++) if (text[i] != ' ' && text[i] != '\t') return false;
            return true;
        }

        private static void AppendProse(StringBuilder result, string text)
        {
            // Index delimiter runs once so unmatched/repeated backticks stay
            // linear rather than rescanning the remaining answer each time.
            var runs = new Dictionary<int, Queue<int>>();
            for (var i = 0; i < text.Length; i++)
            {
                if (text[i] != '`') continue;
                var start = i;
                while (i + 1 < text.Length && text[i + 1] == '`') i++;
                var length = i - start + 1;
                if (!runs.TryGetValue(length, out var positions)) runs[length] = positions = new Queue<int>();
                positions.Enqueue(start);
            }
            for (var i = 0; i < text.Length;)
            {
                if (text[i] == '`')
                {
                    var end = i + 1;
                    while (end < text.Length && text[end] == '`') end++;
                    var length = end - i; var positions = runs[length];
                    while (positions.Count > 0 && positions.Peek() <= i) positions.Dequeue();
                    if (!IsEscaped(text, i) && positions.Count > 0)
                        end = positions.Dequeue() + length;
                    result.Append(text, i, end - i); i = end;
                }
                else if (IsGap(text[i]))
                {
                    var end = i + 1;
                    while (end < text.Length && IsGap(text[end])) end++;
                    if (i == 0 || end == text.Length || !IsCjkBefore(text, i) || !IsCjkAt(text, end))
                        result.Append(text, i, end - i);
                    i = end;
                }
                else result.Append(text[i++]);
            }
        }

        private static bool IsEscaped(string text, int position)
        {
            var slashes = 0;
            while (position > 0 && text[--position] == '\\') slashes++;
            return slashes % 2 != 0;
        }
        private static bool IsGap(char c) => c == ' ' || c == '\u3000' || c == '\t';
        private static bool IsCjkBefore(string text, int end) => IsCjkAt(text,
            end >= 2 && char.IsSurrogatePair(text, end - 2) ? end - 2 : end - 1);
        private static bool IsCjkAt(string text, int position)
        {
            var c = char.IsSurrogatePair(text, position) ? char.ConvertToUtf32(text, position) : text[position];
            // Han (including supplementary extensions and compatibility forms),
            // kana and the requested punctuation/fullwidth blocks. No Hangul.
            return (c >= 0x3400 && c <= 0x4DBF) || (c >= 0x4E00 && c <= 0x9FFF) ||
                (c >= 0xF900 && c <= 0xFAFF) || (c >= 0x20000 && c <= 0x2A6DF) ||
                (c >= 0x2A700 && c <= 0x2B73F) || (c >= 0x2B740 && c <= 0x2B81F) ||
                (c >= 0x2B820 && c <= 0x2CEAF) || (c >= 0x2CEB0 && c <= 0x2EBEF) ||
                (c >= 0x2EBF0 && c <= 0x2EE5F) || (c >= 0x2F800 && c <= 0x2FA1F) ||
                (c >= 0x30000 && c <= 0x3134F) || (c >= 0x31350 && c <= 0x323AF) ||
                (c >= 0x323B0 && c <= 0x3347F) || (c >= 0x3040 && c <= 0x309F) ||
                (c >= 0x30A0 && c <= 0x30FF) || (c >= 0x3000 && c <= 0x303F) ||
                (c >= 0xFF00 && c <= 0xFFEF);
        }
    }
}
