using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Threading;
using System.Windows.Forms;
using Geometry = System.Windows.Media.Geometry;
using PolyLineSegment = System.Windows.Media.PolyLineSegment;
using LineSegment = System.Windows.Media.LineSegment;

namespace OfficeTranslate.WordAddIn
{
    // XAML path mini-language geometries, drawn by the existing WinForms surface.
    // No bitmap resources or WPF visual tree/dispatcher are needed in Office.
    internal sealed class ChatCopyButton : Control
    {
        private static readonly Geometry CopyGeometry = Icon("M 8,5 L 5,5 L 5,19 L 16,19 L 16,16 M 9,2 L 19,2 L 19,15 L 9,15 Z");
        private static readonly Geometry CheckGeometry = Icon("M 4,11 L 9,16 L 20,5");
        private readonly ToolTip _tips = new ToolTip();
        private readonly System.Windows.Forms.Timer _feedbackTimer = new System.Windows.Forms.Timer { Interval = 1500 };
        private readonly int _uiThread = Thread.CurrentThread.ManagedThreadId;
        private string _source = "", _language = "zh-CN", _feedback = "ChatCopy";
        private bool _available, _hover, _user;
        internal Action<string> ClipboardWriter { get; set; } = text => Clipboard.SetText(text, TextDataFormat.UnicodeText);
        internal bool IsCopied => _feedback == "ChatCopied";

        internal ChatCopyButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.Selectable | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent; Cursor = Cursors.Hand; TabStop = true;
            AccessibleRole = AccessibleRole.PushButton;
            _feedbackTimer.Tick += (s, e) => ResetFeedback();
            RefreshTip();
        }
        private static Geometry Icon(string path)
        { var geometry = Geometry.Parse(path); geometry.Freeze(); return geometry; }

        internal void Configure(string source, bool streaming, bool user, string language)
        {
            if (_source != source) ResetFeedback();
            _source = source; _user = user; _language = language;
            _available = !streaming && !string.IsNullOrWhiteSpace(source);
            if (!_available) ResetFeedback();
            Visible = _available; RefreshTip(); Invalidate();
        }
        private void RefreshTip()
        {
            AccessibleName = UiText.Get(_language, "ChatCopy");
            var text = UiText.Get(_language, _feedback);
            AccessibleDescription = text; _tips.SetToolTip(this, text);
        }
        private void ResetFeedback()
        {
            _feedbackTimer.Stop(); _tips.Hide(this); _feedback = "ChatCopy";
            RefreshTip(); Invalidate();
        }
        internal void Copy()
        {
            if (IsDisposed || Disposing) return;
            if (Thread.CurrentThread.ManagedThreadId != _uiThread)
            {
                // InvokeRequired alone is false before handle creation. Never let
                // a worker create a handle or touch the clipboard on that path.
                try { if (IsHandleCreated) BeginInvoke((Action)Copy); }
                catch (InvalidOperationException) { } // Handle destroyed during dispatch.
                return;
            }
            if (!_available) return;
            _feedbackTimer.Stop();
            try
            {
                if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException();
                ClipboardWriter(_source);
                _feedback = "ChatCopied";
            }
            catch
            {
                // Clipboard contention/permission failures are local UI feedback.
                // Never log exception text, source, prompts or clipboard contents.
                _feedback = "ChatCopyFailed";
            }
            if (IsDisposed || Disposing) return;
            RefreshTip(); Invalidate();
            _tips.Show(UiText.Get(_language, _feedback), this, Width / 2, Height, _feedbackTimer.Interval);
            _feedbackTimer.Start();
        }
        protected override void OnClick(EventArgs e) { base.OnClick(e); Copy(); }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space) { e.Handled = e.SuppressKeyPress = true; Copy(); }
        }
        protected override void OnMouseEnter(EventArgs e) { base.OnMouseEnter(e); _hover = true; Invalidate(); }
        protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); _hover = false; Invalidate(); }
        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e); e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var size = Math.Min(Width, Height) * .75F;
            var state = e.Graphics.Save();
            try
            {
                e.Graphics.TranslateTransform((Width - size) / 2, (Height - size) / 2);
                e.Graphics.ScaleTransform(size / 24F, size / 24F);
                var color = _hover ? Color.FromArgb(45, 60, 80) :
                    (_user ? Color.FromArgb(201, 216, 240) : Color.FromArgb(158, 168, 183));
                using (var pen = new System.Drawing.Pen(color, 1.7F) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = System.Drawing.Drawing2D.LineJoin.Round })
                using (var path = new GraphicsPath())
                {
                    foreach (var figure in (IsCopied ? CheckGeometry : CopyGeometry).GetFlattenedPathGeometry().Figures)
                    {
                        path.StartFigure(); var last = new PointF((float)figure.StartPoint.X, (float)figure.StartPoint.Y);
                        foreach (var segment in figure.Segments)
                        {
                            if (segment is PolyLineSegment lines)
                                foreach (var p in lines.Points) { var next = new PointF((float)p.X, (float)p.Y); path.AddLine(last, next); last = next; }
                            else if (segment is LineSegment line)
                            { var next = new PointF((float)line.Point.X, (float)line.Point.Y); path.AddLine(last, next); last = next; }
                        }
                        if (figure.IsClosed) path.CloseFigure();
                    }
                    e.Graphics.DrawPath(pen, path);
                }
            }
            finally { e.Graphics.Restore(state); }
            if (Focused) ControlPaint.DrawFocusRectangle(e.Graphics, ClientRectangle);
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing) { _feedbackTimer.Stop(); _feedbackTimer.Dispose(); _tips.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
