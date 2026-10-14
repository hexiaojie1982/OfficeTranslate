using System;
using System.Drawing;
using System.Windows.Forms;
using OfficeTranslate.Core;

namespace OfficeTranslate.WordAddIn
{
    // Source-only wording and a copyable, persistent recovery view for unplaced notes.
    internal sealed class OcrResultForm : Form
    {
        private readonly Button _close = new Button();
        private readonly Timer _timer = new Timer();
        private readonly string _uiLanguage;
        private readonly bool _keepOpen;
        private int _remainingSeconds = 5;

        internal OcrResultForm(string uiLanguage, TranslationTaskSummary? summary, Exception? error)
        {
            _uiLanguage = uiLanguage;
            _keepOpen = error != null || summary?.UnplacedImageNotes.Count > 0;
            Text = UiText.Get(uiLanguage, "OcrResultTitle");
            Width = 520; Height = 300; MinimumSize = new Size(480, 280);
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
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

            var needsAttention = _keepOpen;
            var icon = new Label
            {
                Text = needsAttention ? "!" : "✓",
                AutoSize = false,
                Width = 42,
                Height = 42,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 22F, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = needsAttention ? Color.DarkOrange : Color.FromArgb(38, 166, 91),
                Margin = new Padding(0, 0, 10, 12)
            };
            var heading = new Label
            {
                Text = UiText.Get(uiLanguage, error == null ? "OcrResultCompleted" : error is OperationCanceledException
                    ? "OcrResultCancelled" : "OcrResultFailed"),
                AutoSize = true,
                Font = new Font(Font.FontFamily, 14F, FontStyle.Bold),
                ForeColor = Color.FromArgb(32, 55, 88),
                Margin = new Padding(0, 7, 0, 12)
            };
            root.Controls.Add(icon, 0, 0);
            root.Controls.Add(heading, 1, 0);

            var message = error == null ? UiText.Get(uiLanguage, "OcrResultCompleted") : error is OperationCanceledException
                ? UiText.Get(uiLanguage, "OcrResultCancelledDetails")
                : string.Format(UiText.Get(uiLanguage, "OcrResultFailedDetails"), error.Message);
            if (summary != null)
            {
                message += "\r\n" + string.Format(UiText.Get(uiLanguage, "OcrResultDetails"),
                    summary.OcrImages, summary.OcrKeptRegions, summary.OcrNoText, summary.ImageNeedsReview);
                if (summary.UnplacedImageNotes.Count > 0)
                    message += "\r\n\r\n" + UiText.Get(uiLanguage, "OcrResultUnplacedNotes") + "\r\n\r\n" +
                        string.Join("\r\n\r\n", summary.UnplacedImageNotes);
            }
            message = message.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r\n");
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
                Text = message
            };
            root.SetColumnSpan(details, 2);
            root.Controls.Add(details, 0, 1);

            _close.AutoSize = true;
            _close.Height = 36;
            _close.Padding = new Padding(16, 4, 16, 4);
            _close.Anchor = AnchorStyles.Right;
            _close.FlatStyle = FlatStyle.Flat;
            _close.BackColor = Color.FromArgb(45, 108, 223);
            _close.ForeColor = Color.White;
            _close.DialogResult = DialogResult.OK;
            _close.Click += (s, e) => Close();
            _close.FlatAppearance.BorderSize = 0;
            UpdateCloseText();
            root.SetColumnSpan(_close, 2);
            root.Controls.Add(_close, 0, 2);
            Controls.Add(root);
            AcceptButton = _close; CancelButton = _close;
            FormClosed += (s, e) => _timer.Stop();
            _timer.Interval = 1000;
            _timer.Tick += (s, e) =>
            {
                _remainingSeconds--;
                if (_remainingSeconds <= 0) Close();
                else UpdateCloseText();
            };
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (!_keepOpen && !IsDisposed) _timer.Start();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _timer.Dispose();
            base.Dispose(disposing);
        }

        private void UpdateCloseText()
        {
            _close.Text = _keepOpen ? UiText.Get(_uiLanguage, "Close")
                : string.Format(UiText.Get(_uiLanguage, "OcrResultCloseCountdown"), _remainingSeconds);
        }

        public static void ShowResult(string uiLanguage, TranslationTaskSummary? summary, IntPtr owner, Exception? error = null)
        {
            using (var form = new OcrResultForm(uiLanguage, summary, error))
            {
                if (owner == IntPtr.Zero) form.ShowDialog();
                else form.ShowDialog(new TranslationProgressForm.WindowHandle(owner));
            }
        }
    }
}
