using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace OfficeTranslate.WordAddIn
{
    internal sealed class TranslationProgressForm : Form
    {
        private readonly Label _status = new Label();
        private readonly ProgressBar _progress = new ProgressBar();
        private readonly Button _cancel = new Button();
        private readonly Action _cancelAction;
        private readonly string _cancellingText;
        private bool _cancellationRequested;

        public TranslationProgressForm(string uiLanguage, Action cancelAction)
        {
            _cancelAction = cancelAction;
            _cancellingText = UiText.Get(uiLanguage, "CancellingTranslation");
            Text = UiText.Get(uiLanguage, "TranslationProgress");
            Width = 460; Height = 170; MinimumSize = new Size(420, 170); MaximizeBox = false; MinimizeBox = false;
            FormBorderStyle = FormBorderStyle.FixedDialog; ShowInTaskbar = false; ControlBox = false;
            StartPosition = FormStartPosition.Manual; Font = new Font("Microsoft YaHei UI", 9F); BackColor = Color.FromArgb(247, 249, 252);

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), RowCount = 3, ColumnCount = 1 };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _status.Text = UiText.Get(uiLanguage, "PreparingTranslation"); _status.AutoEllipsis = true; _status.Dock = DockStyle.Fill;
            _status.TextAlign = ContentAlignment.MiddleLeft; _status.ForeColor = Color.FromArgb(32, 55, 88);
            _progress.Dock = DockStyle.Fill; _progress.Height = 18; _progress.Style = ProgressBarStyle.Marquee; _progress.MarqueeAnimationSpeed = 25;
            _cancel.Text = UiText.Get(uiLanguage, "CancelTranslation"); _cancel.AutoSize = true; _cancel.Height = 34;
            _cancel.Padding = new Padding(14, 4, 14, 4); _cancel.Anchor = AnchorStyles.Right; _cancel.FlatStyle = FlatStyle.Flat;
            _cancel.BackColor = Color.White; _cancel.ForeColor = Color.FromArgb(42, 58, 78); _cancel.FlatAppearance.BorderColor = Color.FromArgb(190, 200, 214);
            _cancel.Click += CancelClicked;
            root.Controls.Add(_status, 0, 0); root.Controls.Add(_progress, 0, 1); root.Controls.Add(_cancel, 0, 2); Controls.Add(root);
        }

        public void ShowFor(IntPtr ownerHandle)
        {
            if (GetWindowRect(ownerHandle, out var rect))
                Location = new Point(rect.Left + Math.Max(0, (rect.Right - rect.Left - Width) / 2), rect.Top + Math.Max(0, (rect.Bottom - rect.Top - Height) / 2));
            Show(new WindowHandle(ownerHandle)); Activate();
        }

        public void SetStatus(string text)
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action<string>(SetStatus), text); return; }
            _status.Text = text;
            var match = Regex.Match(text ?? string.Empty, @"(\d+)\s*/\s*(\d+)");
            if (match.Success && int.TryParse(match.Groups[1].Value, out var current) && int.TryParse(match.Groups[2].Value, out var total) && total > 0)
            {
                _progress.Style = ProgressBarStyle.Continuous; _progress.MarqueeAnimationSpeed = 0;
                _progress.Maximum = total; _progress.Value = Math.Max(0, Math.Min(total, current));
            }
        }

        public void CloseSafely()
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(CloseSafely)); return; }
            Close(); Dispose();
        }

        private void CancelClicked(object sender, EventArgs e)
        {
            if (_cancellationRequested) return;
            _cancellationRequested = true; _cancel.Enabled = false; _status.Text = _cancellingText;
            _progress.Style = ProgressBarStyle.Marquee; _progress.MarqueeAnimationSpeed = 25;
            _cancelAction();
        }

        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out NativeRect rect);
        [StructLayout(LayoutKind.Sequential)] private struct NativeRect { public int Left, Top, Right, Bottom; }
        private sealed class WindowHandle : IWin32Window { public WindowHandle(IntPtr handle) { Handle = handle; } public IntPtr Handle { get; } }
    }
}
