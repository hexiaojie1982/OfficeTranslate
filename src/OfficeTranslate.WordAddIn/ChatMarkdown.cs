using System;
using System.Collections.Generic;
using System.Globalization;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeTranslate.WordAddIn
{
    // Deliberately a small, offline Markdown subset. Input is always escaped as
    // RTF text: HTML, URLs and RTF control words can never execute or load assets.
    internal static class ChatMarkdown
    {
        private static readonly Regex Heading = new Regex(@"^ {0,3}(#{1,6})[ \t]+(.*)$");
        private static readonly Regex List = new Regex(@"^( *)(?:([-+*])|([0-9]{1,9})[.)])[ \t]+(.*)$");
        internal static string Normalize(string text) => text.Replace("\r\n", "\n").Replace('\r', '\n');
        internal static string Escape(string text)
        {
            var result = new StringBuilder();
            foreach (var c in text)
            {
                if (c == '\\' || c == '{' || c == '}') result.Append('\\').Append(c);
                else if (c == '\n') result.Append(@"\par ");
                else if (c == '\t') result.Append(@"\tab ");
                else if (c >= 32 && c < 127) result.Append(c);
                else result.Append(@"\u").Append(unchecked((short)c).ToString(CultureInfo.InvariantCulture)).Append('?');
            }
            return result.ToString();
        }
        internal static string Document(IEnumerable<ChatParagraph> paragraphs, Color? foreground = null)
        {
            var color = foreground ?? Color.Black;
            var rtf = new StringBuilder(@"{\rtf1\ansi\deff0\uc1{\fonttbl{\f0 Microsoft YaHei UI;}}{\colortbl ;\red242\green244\blue247;\red90\green104\blue122;");
            // Keep index 0 automatic for unshaded backgrounds. Paragraph() uses
            // index 3 for the foreground without rewriting escaped literal text.
            rtf.Append(@"\red").Append(color.R).Append(@"\green").Append(color.G).Append(@"\blue").Append(color.B).Append(";}");
            foreach (var paragraph in paragraphs) rtf.Append(paragraph.Rtf);
            return rtf.Append('}').ToString();
        }
        internal static List<ChatParagraph> Plain(string text, int fontSize)
        {
            var result = new List<ChatParagraph>();
            AddLines(result, Normalize(text), fontSize, false);
            return result;
        }
        private static void AddLines(List<ChatParagraph> result, string text, int size, bool code)
        {
            var lines = text.Split('\n');
            for (var i = 0; i < lines.Length; i++)
            {
                if (i == lines.Length - 1 && lines[i].Length == 0) break;
                var newline = i < lines.Length - 1;
                var value = lines[i] + (newline ? "\n" : "");
                result.Add(new ChatParagraph(value, Paragraph(size) + (code ? @"\cbpat1\highlight1 " : "") + Escape(value)));
            }
        }
        private static string Paragraph(int size) => @"\pard\plain\f0\fs" + size.ToString(CultureInfo.InvariantCulture) + @"\cf3\highlight0\cbpat0\sa0\sb0 ";
        internal static List<ChatParagraph> Parse(string source, int fontSize)
        {
            var result = new List<ChatParagraph>();
            var lines = Normalize(source).Split('\n');
            var fence = 0; var fenceChar = '\0';
            for (var i = 0; i < lines.Length; i++)
            {
                var line = lines[i]; var trimmed = line.TrimStart(' ');
                var fenceLength = 0;
                if (trimmed.StartsWith("```", StringComparison.Ordinal) || trimmed.StartsWith("~~~", StringComparison.Ordinal))
                    while (fenceLength < trimmed.Length && trimmed[fenceLength] == trimmed[0]) fenceLength++;
                if (fence == 0 && fenceLength >= 3 && line.Length - trimmed.Length <= 3)
                { fence = fenceLength; fenceChar = trimmed[0]; continue; }
                if (fence > 0 && fenceLength >= fence && trimmed[0] == fenceChar && trimmed.Substring(fenceLength).Trim().Length == 0)
                { fence = 0; continue; }
                if (i == lines.Length - 1 && line.Length == 0) break;
                var newline = i < lines.Length - 1;
                if (fence > 0)
                {
                    AddLines(result, line + (newline ? "\n" : ""), fontSize, true);
                    continue;
                }
                var rtf = new StringBuilder(Paragraph(fontSize)); var text = new StringBuilder();
                var quoteDepth = 0;
                while (trimmed.StartsWith(">", StringComparison.Ordinal))
                { quoteDepth++; trimmed = trimmed.Substring(1); if (trimmed.StartsWith(" ", StringComparison.Ordinal)) trimmed = trimmed.Substring(1); }
                if (quoteDepth > 0)
                {
                    line = trimmed;
                    rtf.Append(@"\li").Append(Math.Min(quoteDepth, 8) * 240).Append(@"\cf2\brdrl\brdrs\brdrw15\brdrcf2 ");
                }
                var heading = Heading.Match(line); var list = List.Match(line);
                if (heading.Success)
                {
                    line = Regex.Replace(heading.Groups[2].Value, @"[ \t]+#+[ \t]*$", "");
                    rtf.Append(@"\b\fs").Append(fontSize + (7 - heading.Groups[1].Length) * 2).Append(@"\sb100\sa80 ");
                }
                else if (list.Success)
                {
                    var indent = Math.Min(list.Groups[1].Length / 2, 8) * 240 + quoteDepth * 240 + 300;
                    var marker = list.Groups[3].Success ? list.Groups[3].Value + "." : "•";
                    rtf.Append(@"\li").Append(indent).Append(@"\fi-240\tx").Append(indent).Append(' ').Append(Escape(marker)).Append(@"\tab ");
                    text.Append(marker).Append('\t'); line = list.Groups[4].Value;
                }
                Inline(line, 0, line.Length, rtf, text, 0);
                if (newline) { rtf.Append(@"\par "); text.Append('\n'); }
                result.Add(new ChatParagraph(text.ToString(), rtf.ToString()));
            }
            return result;
        }
        private static int Run(string text, int at, int end, char marker)
        { var next = at; while (next < end && text[next] == marker) next++; return next - at; }
        private static int Closing(string text, int at, int end, char marker, int width)
        {
            for (var i = at; i < end; i++)
            {
                if (marker != '`' && text[i] == '\\') { i++; continue; }
                if (marker != '`' && text[i] == '`')
                {
                    var ticks = Run(text, i, end, '`'); var close = Closing(text, i + ticks, end, '`', ticks);
                    if (close >= 0) { i = close + ticks - 1; continue; }
                }
                if (text[i] != marker) continue;
                var count = Run(text, i, end, marker);
                if (marker == '`' ? count == width : count >= width && (i > at && !char.IsWhiteSpace(text[i - 1])))
                    return marker == '*' ? i + count - width : i;
                i += count - 1;
            }
            return -1;
        }
        private static void Inline(string source, int start, int end, StringBuilder rtf, StringBuilder text, int depth)
        {
            for (var i = start; i < end;)
            {
                var c = source[i];
                if (c == '\\' && i + 1 < end && "\\`*{}[]()#+-.!>_".IndexOf(source[i + 1]) >= 0)
                { i++; rtf.Append(Escape(source[i].ToString())); text.Append(source[i++]); continue; }
                if (c == '`' || (c == '*' && depth < 16))
                {
                    var count = Run(source, i, end, c); var width = c == '`' ? count : Math.Min(count, 3);
                    var close = (i + width < end && !char.IsWhiteSpace(source[i + width])) || c == '`'
                        ? Closing(source, i + width, end, c, width) : -1;
                    if (close >= 0)
                    {
                        if (c == '`')
                        { var value = source.Substring(i + width, close - i - width); rtf.Append(@"{\highlight1 ").Append(Escape(value)).Append('}'); text.Append(value); }
                        else
                        {
                            rtf.Append(width == 3 ? @"{\b\i " : width == 2 ? @"{\b " : @"{\i ");
                            Inline(source, i + width, close, rtf, text, depth + 1); rtf.Append('}');
                        }
                        i = close + width; continue;
                    }
                    // Keep incomplete streaming delimiters literal until closed.
                    rtf.Append(Escape(source.Substring(i, count))); text.Append(source, i, count); i += count; continue;
                }
                rtf.Append(Escape(c.ToString())); text.Append(c); i++;
            }
        }
    }
    internal sealed class ChatParagraph
    {
        internal string Text { get; }
        internal string Rtf { get; }
        internal ChatParagraph(string text, string rtf) { Text = text; Rtf = rtf; }
    }
}
