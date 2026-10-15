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
        private int FontSize => (int)Math.Round(Font.SizeInPoints * 2);
        internal ChatTranscriptBox()
        {
            DetectUrls = false;
            ScrollBars = RichTextBoxScrollBars.None;
        }

        // A bubble owns one message; completed messages are separate controls.
        // Never interpret user input as Markdown.
        internal void ShowMessage(string text, bool markdown, string error = "")
        {
            var next = markdown ? ChatMarkdown.Parse(text, FontSize) : ChatMarkdown.Plain(text, FontSize);
            next.AddRange(ChatMarkdown.Plain(error, FontSize));
            RenderMessage(next);
        }
        private void RenderMessage(List<ChatParagraph> next)
        {
            var common = 0;
            while (common < _displayed.Count && common < next.Count && _displayed[common].Rtf == next[common].Rtf)
                common++;
            if (common == _displayed.Count && common == next.Count) return;
            // The 50ms timer calls once per batch. Stream the active message's
            // complete RTF: replacing a selection can inherit the destination's
            // final paragraph format (e.g. drop quote indentation). Freeze only
            // this HWND; completed messages and sibling controls stay untouched.
            var redraw = IsHandleCreated && Visible;
            if (redraw) SendMessage(Handle, 0x000B /* WM_SETREDRAW */, IntPtr.Zero, IntPtr.Zero);
            try
            {
                Rtf = ChatMarkdown.Document(next, ForeColor);
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
