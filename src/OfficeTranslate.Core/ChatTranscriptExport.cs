using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeTranslate.Core
{
    // Formats only a committed in-memory history snapshot. No context, transport, logging or host access.
    public static class ChatTranscriptExport
    {
        public static string Format(IReadOnlyList<ChatMessage> history, bool markdown, Func<string, string> text)
        {
            if (history.Count == 0) return "";
            var result = new StringBuilder();
            result.Append(markdown ? "# " : "").Append(text("ChatExportHeader")).Append("\r\n\r\n");
            for (var i = 0; i + 1 < history.Count; i += 2)
            {
                var question = history[i]; var answer = history[i + 1];
                var timestamp = question.Timestamp.ToString("yyyy-MM-dd HH:mm:ss zzz", CultureInfo.InvariantCulture);
                result.Append(markdown ? "## " : "----------------------------------------\r\n")
                    .Append(text("ChatExportTimestamp")).Append(": ").Append(timestamp).Append("\r\n\r\n");
                AppendContent(result, text("ChatExportQuestion"), question.Content, markdown);
                AppendContent(result, text("ChatExportAnswer"), answer.Content, markdown);
            }
            return result.ToString();
        }

        private static void AppendContent(StringBuilder result, string label, string content, bool markdown)
        {
            result.Append(markdown ? "### " : "").Append(label).Append("\r\n");
            if (markdown)
                foreach (var line in content.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
                    result.Append("> ").Append(line).Append("\r\n");
            else result.Append(content).Append("\r\n");
            result.Append("\r\n");
        }
    }
}
