using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OfficeTranslate.WordAddIn
{
    // Only the conversation viewport lays out during a stream batch. The Office
    // task pane, command rows and completed RichEdit documents stay untouched.
    internal sealed class ChatConversationView : UserControl
    {
        private readonly Panel _viewport = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        private readonly Panel _canvas = new Panel();
        private readonly Label _empty = new Label { AutoSize = false, ForeColor = Color.FromArgb(90, 104, 122), Padding = new Padding(16) };
        private readonly List<ChatBubbleRow> _messages = new List<ChatBubbleRow>();
        private ChatBubbleRow? _active;
        private bool _arranging;
        internal IReadOnlyList<ChatBubbleRow> Messages => _messages;

        internal ChatConversationView()
        {
            AutoSize = false; DoubleBuffered = true;
            BackColor = Color.FromArgb(240, 244, 250);
            _viewport.BackColor = _canvas.BackColor = BackColor;
            Controls.Add(_viewport); _viewport.Controls.Add(_canvas); _canvas.Controls.Add(_empty);
            _viewport.SizeChanged += (s, e) => ArrangeMessages(false);
        }

        // A readable representation for accessibility/automation; the actual
        // selectable text and timestamp belong to each message's native control.
        public override string Text
        {
            get => _messages.Count == 0 ? _empty.Text : string.Join("\n\n", _messages.Select(m => m.RoleLabel + "\n" + m.Content.Text));
            set => base.Text = value;
        }

        internal void ShowConversation(IReadOnlyList<ChatMessage> history, string userLabel, string assistantLabel,
            string empty, string? question, string output, string error, string userAvatar, string aiAvatar, DateTimeOffset turnTime)
        {
            var wasAtBottom = AtBottom();
            var count = history.Count + (question == null ? 0 : 2);
            SuspendLayout(); _canvas.SuspendLayout();
            try
            {
                while (_messages.Count > count)
                {
                    var row = _messages[_messages.Count - 1]; _messages.RemoveAt(_messages.Count - 1);
                    _canvas.Controls.Remove(row); row.Dispose();
                }
                while (_messages.Count < count)
                {
                    var row = new ChatBubbleRow { Font = Font, BackColor = BackColor };
                    _messages.Add(row); _canvas.Controls.Add(row);
                }
                for (var i = 0; i < count; i++)
                {
                    var message = i < history.Count ? history[i] : new ChatMessage(i == history.Count ? "user" : "assistant", i == history.Count ? question! : output, turnTime);
                    var user = message.Role == "user";
                    _messages[i].ShowMessage(message.Content, user, user ? userLabel : assistantLabel,
                        user ? userAvatar : aiAvatar, message.Timestamp, i == count - 1 && question != null ? error : "");
                }
                _active = question == null ? null : _messages.Last();
                _empty.Text = empty; _empty.Visible = count == 0;
            }
            finally { _canvas.ResumeLayout(false); ResumeLayout(false); }
            ArrangeMessages(wasAtBottom); OnTextChanged(EventArgs.Empty);
        }

        internal void ShowOutput(string output, string error = "")
        {
            if (_active == null || !_active.UpdateOutput(output, error)) return;
            var follow = AtBottom();
            ArrangeMessages(follow); OnTextChanged(EventArgs.Empty);
        }

        private bool AtBottom() => -_viewport.AutoScrollPosition.Y + _viewport.ClientSize.Height >= _canvas.Height - 24;
        private void ArrangeMessages(bool follow)
        {
            if (_arranging || IsDisposed) return;
            _arranging = true;
            try
            {
                // Always reserve the vertical bar width; its appearance cannot
                // narrow the bubbles and trigger a second wrap pass.
                var width = Math.Max(1, _viewport.Width - SystemInformation.VerticalScrollBarWidth - 2);
                var y = 12;
                foreach (var row in _messages)
                {
                    row.SetBounds(0, y, width, row.Height);
                    row.ArrangeBubble(); y += row.Height + 14;
                }
                _empty.SetBounds(0, 0, width, Math.Max(80, _viewport.Height));
                if (_messages.Count == 0) y = Math.Max(y, _viewport.ClientSize.Height);
                _canvas.Size = new Size(width, Math.Max(1, y));
                if (follow) _viewport.AutoScrollPosition = new Point(0, Math.Max(0, y - _viewport.ClientSize.Height));
            }
            finally { _arranging = false; }
        }
    }

    internal sealed class ChatBubbleRow : Control
    {
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);
        internal ChatTranscriptBox Content { get; } = new ChatTranscriptBox
        {
            ReadOnly = true, Multiline = true, AutoSize = false, BorderStyle = BorderStyle.None,
            ScrollBars = RichTextBoxScrollBars.None, WordWrap = true, MaxLength = 0, TabStop = true
        };
        internal bool IsUser { get; private set; }
        internal string RoleLabel { get; private set; } = "";
        internal DateTimeOffset Timestamp { get; private set; }
        internal Rectangle BubbleBounds { get; private set; }
        private readonly Label _time = new Label { AutoSize = false, ForeColor = Color.FromArgb(90, 104, 122) };
        private string _avatar = "", _source = "", _error = "";
        private int _contentHeight;
        private int _arrangedWidth = -1;
        private bool _measuring, _shown, _dirty = true;
        private Font? _timeFont;
        internal ChatBubbleRow()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            Controls.Add(Content); Controls.Add(_time);
            Content.ContentsResized += (s, e) => _contentHeight = e.NewRectangle.Height;
            Content.MouseWheel += (s, e) =>
            {
                // RichEdit has no inner scrollbar; route the wheel to the viewport.
                if (Parent?.Parent is ScrollableControl viewport)
                    viewport.AutoScrollPosition = new Point(0, Math.Max(0, -viewport.AutoScrollPosition.Y - e.Delta));
            };
        }
        internal void ShowMessage(string source, bool user, string role, string avatar, DateTimeOffset timestamp, string error)
        {
            // Assigning RichTextBox.Font/ForeColor resets existing character
            // formats. Set them only when initializing the message, otherwise a
            // language refresh would flatten already-rendered Markdown styles.
            if (!_shown || IsUser != user)
            {
                Content.Font = Font;
                Content.BackColor = user ? Color.FromArgb(45, 108, 223) : Color.White;
                Content.ForeColor = user ? Color.White : Color.FromArgb(32, 55, 88);
                _shown = false;
            }
            IsUser = user; RoleLabel = role; _avatar = avatar; Timestamp = timestamp;
            Content.AccessibleName = role;
            _time.Text = timestamp.ToLocalTime().ToString("HH:mm", CultureInfo.InvariantCulture);
            _time.AccessibleName = role + " " + timestamp.ToLocalTime().ToString("g", CultureInfo.CurrentCulture);
            var timeSize = Math.Max(7F, Font.SizeInPoints - 1F);
            if (_time.Font.SizeInPoints != timeSize || _time.Font.FontFamily.Name != Font.FontFamily.Name)
            {
                var old = _timeFont;
                _timeFont = new Font(Font.FontFamily, timeSize); _time.Font = _timeFont;
                old?.Dispose();
            }
            UpdateOutput(source, error); Invalidate();
        }
        internal bool UpdateOutput(string source, string error)
        {
            if (_shown && _source == source && _error == error) return false;
            _shown = _dirty = true; _source = source; _error = error;
            Content.ShowMessage(source, !IsUser, error); return true;
        }

        internal void ArrangeBubble()
        {
            if (_measuring || (!_dirty && _arrangedWidth == Width)) return;
            _measuring = true;
            Content.CreateControl();
            SendMessage(Content.Handle, 0x000B, IntPtr.Zero, IntPtr.Zero);
            try
            {
                var scale = DeviceDpi / 96F;
                int Px(int n) => Math.Max(1, (int)Math.Round(n * scale));
                var inset = Px(8); var padding = Px(12); var avatar = Px(30); var gap = Px(8);
                var max = Math.Max(1, Math.Min((int)(Width * .8), Width - inset * 2 - avatar - gap));
                var minimum = Math.Min(max, Px(64));
                // Measure the native RichEdit after wrapping at the maximum width,
                // including heading sizes/indents rather than counting characters.
                Content.Width = Math.Max(1, max - padding * 2);
                SendMessage(Content.Handle, 0x0441 /* EM_REQUESTRESIZE */, IntPtr.Zero, IntPtr.Zero);
                Content.Select(0, 0); Content.ScrollToCaret();
                var right = 0;
                var text = Content.Text;
                var lines = Content.GetLineFromCharIndex(Content.TextLength);
                for (var i = 0; i <= lines; i++)
                {
                    var start = Content.GetFirstCharIndexFromLine(i);
                    var next = Content.GetFirstCharIndexFromLine(i + 1);
                    var end = next < 0 ? Content.TextLength : next;
                    if (end > start)
                    {
                        var position = Content.GetPositionFromCharIndex(end - 1);
                        var charWidth = TextRenderer.MeasureText(text.Substring(end - 1, 1), Content.Font, Size.Empty, TextFormatFlags.NoPadding).Width;
                        right = Math.Max(right, position.X + charWidth + Px(8));
                    }
                }
                var bubbleWidth = Math.Max(minimum, Math.Min(max, right + padding * 2));
                Content.Width = Math.Max(1, bubbleWidth - padding * 2);
                SendMessage(Content.Handle, 0x0441, IntPtr.Zero, IntPtr.Zero);
                var height = Math.Max(Content.Font.Height + Px(6), _contentHeight + Px(4));
                var x = IsUser ? Width - inset - avatar - gap - bubbleWidth : inset + avatar + gap;
                BubbleBounds = new Rectangle(x, 0, bubbleWidth, height + padding * 2);
                Content.SetBounds(x + padding, padding, Content.Width, height);
                _time.TextAlign = IsUser ? ContentAlignment.MiddleRight : ContentAlignment.MiddleLeft;
                _time.SetBounds(x, BubbleBounds.Bottom + Px(3), bubbleWidth, Px(18));
                Height = _time.Bottom; _dirty = false; _arrangedWidth = Width; Invalidate();
            }
            finally { SendMessage(Content.Handle, 0x000B, new IntPtr(1), IntPtr.Zero); Content.Invalidate(); _measuring = false; }
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e); e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var b = BubbleBounds; if (b.Width <= 0 || b.Height <= 0) return;
            var radius = Math.Min(16 * DeviceDpi / 96, Math.Min(b.Width, b.Height));
            using (var path = new GraphicsPath())
            using (var brush = new SolidBrush(Content.BackColor))
            {
                path.AddArc(b.Left, b.Top, radius, radius, 180, 90); path.AddArc(b.Right - radius, b.Top, radius, radius, 270, 90);
                path.AddArc(b.Right - radius, b.Bottom - radius, radius, radius, 0, 90); path.AddArc(b.Left, b.Bottom - radius, radius, radius, 90, 90);
                path.CloseFigure(); e.Graphics.FillPath(brush, path);
            }
            var size = 30 * DeviceDpi / 96;
            var circle = new Rectangle(IsUser ? Width - 8 * DeviceDpi / 96 - size : 8 * DeviceDpi / 96, 0, size, size);
            using (var brush = new SolidBrush(IsUser ? Color.FromArgb(45, 108, 223) : Color.FromArgb(215, 227, 245))) e.Graphics.FillEllipse(brush, circle);
            TextRenderer.DrawText(e.Graphics, _avatar, Font, circle, IsUser ? Color.White : Color.FromArgb(32, 55, 88), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        }
        protected override void Dispose(bool disposing)
        { if (disposing) _timeFont?.Dispose(); base.Dispose(disposing); }
    }
}
