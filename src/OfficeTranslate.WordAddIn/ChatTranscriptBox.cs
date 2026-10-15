using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OfficeTranslate.WordAddIn
{
    internal sealed class ChatTranscriptBox : RichTextBox
    {
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hwnd, int message, IntPtr wParam, IntPtr lParam);
        private List<ChatParagraph> _displayed = new List<ChatParagraph>();
        private List<ChatParagraph> _prefix = new List<ChatParagraph>();
        private int FontSize => (int)Math.Round(Font.SizeInPoints * 2);
        internal ChatTranscriptBox()
        {
            DetectUrls = false;
            // Reserve scrollbar width up front. Auto-showing the RichEdit bar
            // changes client bounds and would relayout the surrounding task pane.
            ScrollBars = RichTextBoxScrollBars.ForcedVertical;
        }

        internal void ShowConversation(IReadOnlyList<ChatMessage> history, string userLabel, string assistantLabel,
            string empty, string? question, string output, string error)
        {
            var paragraphs = new List<ChatParagraph>();
            foreach (var message in history)
            {
                paragraphs.AddRange(ChatMarkdown.Plain((message.Role == "user" ? userLabel : assistantLabel) + "\n", FontSize));
                paragraphs.AddRange(message.Role == "user" ? ChatMarkdown.Plain(message.Content, FontSize) : ChatMarkdown.Parse(message.Content, FontSize));
                // Terminate the content paragraph before resetting its paragraph
                // format; otherwise a final quote/code line loses its indent/fill.
                paragraphs.Add(new ChatParagraph("\n", @"\par "));
                paragraphs.AddRange(ChatMarkdown.Plain("\n", FontSize));
            }
            if (question != null)
                paragraphs.AddRange(ChatMarkdown.Plain(userLabel + "\n" + question + "\n\n" + assistantLabel + "\n", FontSize));
            else if (history.Count == 0) paragraphs.AddRange(ChatMarkdown.Plain(empty, FontSize));
            _prefix = paragraphs;
            ShowOutput(output, error);
        }
        internal void ShowOutput(string output, string error = "")
        {
            var next = new List<ChatParagraph>(_prefix);
            next.AddRange(ChatMarkdown.Parse(output, FontSize));
            if (error.StartsWith("\r\n", StringComparison.Ordinal))
            { next.Add(new ChatParagraph("\n", @"\par ")); error = error.Substring(2); }
            next.AddRange(ChatMarkdown.Plain(error, FontSize));
            ReplaceChangedParagraphs(next);
        }
        private void ReplaceChangedParagraphs(List<ChatParagraph> next)
        {
            var common = 0; var start = 0;
            while (common < _displayed.Count && common < next.Count && _displayed[common].Rtf == next[common].Rtf)
            { start += _displayed[common].Text.Length; common++; }
            if (common == _displayed.Count && common == next.Count) return;
            // The 50ms timer calls us once per batch. Keep committed paragraphs
            // intact, freeze only this HWND, and never invalidate sibling controls.
            var redraw = IsHandleCreated && Visible;
            if (redraw) SendMessage(Handle, 0x000B /* WM_SETREDRAW */, IntPtr.Zero, IntPtr.Zero);
            try
            {
                Select(start, TextLength - start);
                SelectedRtf = ChatMarkdown.Document(next.GetRange(common, next.Count - common));
                _displayed = next;
                Select(TextLength, 0); ScrollToCaret();
            }
            finally
            {
                if (redraw) { SendMessage(Handle, 0x000B, new IntPtr(1), IntPtr.Zero); Invalidate(); }
            }
        }
    }
}
