using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OfficeTranslate.Core;

namespace OfficeTranslate.WordAddIn
{
    internal sealed class TranslationResultForm : Form
    {
        private readonly Button _close = new Button();
        private readonly Timer _timer = new Timer();
        private readonly string _uiLanguage;
        private readonly bool _hasUnplacedNotes;
        private int _remainingSeconds = 5;

        private TranslationResultForm(string uiLanguage, TranslationTaskSummary summary)
        {
            _uiLanguage = uiLanguage;
            _hasUnplacedNotes = summary.UnplacedImageNotes.Count > 0;
            Text = UiText.Get(uiLanguage, "SummaryTitle");
            Width = 520; Height = 300; MinimumSize = new Size(480, 280);
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.Manual;
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.FromArgb(247, 249, 252);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(22),
                RowCount = 3,
                ColumnCount = 2
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 52));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var icon = new Label
            {
                Text = _hasUnplacedNotes ? "!" : "✓",
                AutoSize = false,
                Width = 42,
                Height = 42,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 22F, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = _hasUnplacedNotes ? Color.DarkOrange : Color.FromArgb(38, 166, 91),
                Margin = new Padding(0, 0, 10, 12)
            };
            var heading = new Label
            {
                Text = _hasUnplacedNotes
                    ? string.Format(UiText.Get(uiLanguage, "SummaryUnplaced"), summary.UnplacedImageNotes.Count)
                    : string.Format(UiText.Get(uiLanguage, "SummaryCompleted"), summary.Total),
                AutoSize = true,
                Font = new Font(Font.FontFamily, 14F, FontStyle.Bold),
                ForeColor = Color.FromArgb(32, 55, 88),
                Margin = new Padding(0, 7, 0, 12)
            };
            root.Controls.Add(icon, 0, 0);
            root.Controls.Add(heading, 1, 0);

            var details = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                MaxLength = 0,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(42, 58, 78),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Text = string.Format(UiText.Get(uiLanguage, "SummaryDetails"), summary.Translated,
                    summary.CacheHits, summary.SkippedNoSource, summary.NeedsReview, summary.OcrNoText)
            };
            if (summary.TextUnchangedNeedsReview > 0)
                details.Text += "\r\n\r\n" + string.Format(UiText.Get(uiLanguage, "SummaryReviewHint"), summary.TextUnchangedNeedsReview);
            if (summary.TextPartiallyTranslated > 0)
                details.Text += "\r\n\r\n" + string.Format(UiText.Get(uiLanguage, "SummaryPartialHint"), summary.TextPartiallyTranslated);
            if (summary.FallbackTranslatedParagraphs + summary.FallbackReviewParagraphs > 0)
                details.Text += "\r\n" + string.Format(UiText.Get(uiLanguage, "SummaryParagraphFallback"),
                    summary.FallbackTranslatedParagraphs, summary.FallbackReviewParagraphs);
            if (summary.OcrImages > 0)
                details.Text += "\r\n\r\n" + string.Format(UiText.Get(uiLanguage, "OcrRegionDetails"),
                    summary.OcrImages, summary.OcrRawRegions, summary.OcrKeptRegions, summary.OcrDroppedRegions,
                    summary.OcrBboxNotes, summary.OcrEmptyTranslations, summary.ImageNeedsReview);
            if (_hasUnplacedNotes)
                details.Text += "\r\n\r\n" + UiText.Get(uiLanguage, "UnplacedNoteDetails") + "\r\n\r\n"
                    + string.Join("\r\n\r\n", summary.UnplacedImageNotes);
            details.Text = details.Text.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r\n");
            root.SetColumnSpan(details, 2);
            root.Controls.Add(details, 0, 1);

            _close.AutoSize = true;
            _close.Height = 36;
            _close.Padding = new Padding(16, 4, 16, 4);
            _close.Anchor = AnchorStyles.Right;
            _close.FlatStyle = FlatStyle.Flat;
            _close.FlatAppearance.BorderSize = 0;
            _close.BackColor = Color.FromArgb(45, 108, 223);
            _close.ForeColor = Color.White;
            _close.Click += (s, e) => Close();
            UpdateCloseText();
            root.SetColumnSpan(_close, 2);
            root.Controls.Add(_close, 0, 2);
            Controls.Add(root);

            AcceptButton = _close;
            CancelButton = _close;
            FormClosed += (s, e) =>
            {
                _timer.Stop();
                _timer.Dispose();
                Dispose();
            };
            _timer.Interval = 1000;
            _timer.Tick += (s, e) =>
            {
                _remainingSeconds--;
                if (_remainingSeconds <= 0) Close();
                else UpdateCloseText();
            };
        }

        public static void ShowResult(string uiLanguage, TranslationTaskSummary summary, IntPtr ownerHandle)
        {
            var form = new TranslationResultForm(uiLanguage, summary);
            if (GetWindowRect(ownerHandle, out var rect))
                form.Location = new Point(rect.Left + Math.Max(0, (rect.Right - rect.Left - form.Width) / 2),
                    rect.Top + Math.Max(0, (rect.Bottom - rect.Top - form.Height) / 2));
            form.Show(ownerHandle == IntPtr.Zero ? null : new WindowHandle(ownerHandle));
            form.Activate();
            // Failed text must remain copyable until acknowledged, not vanish
            // after five seconds. Successful tasks keep the existing countdown.
            if (!form._hasUnplacedNotes) form._timer.Start();
        }

        private void UpdateCloseText()
        {
            _close.Text = _hasUnplacedNotes ? UiText.Get(_uiLanguage, "Close")
                : string.Format(UiText.Get(_uiLanguage, "ResultCloseCountdown"), _remainingSeconds);
        }

        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out NativeRect rect);
        [StructLayout(LayoutKind.Sequential)] private struct NativeRect { public int Left, Top, Right, Bottom; }
        private sealed class WindowHandle : IWin32Window
        {
            public WindowHandle(IntPtr handle) { Handle = handle; }
            public IntPtr Handle { get; }
        }
    }
}
