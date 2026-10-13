using System;
using System.Drawing;
using System.Windows.Forms;
using OfficeTranslate.Core;

namespace OfficeTranslate.WordAddIn
{
    // Source-only wording and a copyable, persistent recovery view for unplaced notes.
    internal sealed class OcrResultForm : Form
    {
        internal OcrResultForm(TranslationTaskSummary? summary, Exception? error)
        {
            Text = "OCR 识别结果";
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

            var needsAttention = error != null || summary?.UnplacedImageNotes.Count > 0;
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
                Text = error == null ? "OCR 识别完成" : error is OperationCanceledException
                    ? "OCR 识别已取消" : "OCR 识别未全部完成",
                AutoSize = true,
                Font = new Font(Font.FontFamily, 14F, FontStyle.Bold),
                ForeColor = Color.FromArgb(32, 55, 88),
                Margin = new Padding(0, 7, 0, 12)
            };
            root.Controls.Add(icon, 0, 0);
            root.Controls.Add(heading, 1, 0);

            var message = error == null ? "OCR 识别完成" : error is OperationCanceledException
                ? "OCR 识别已取消；已写回的内容保留。" : "OCR 识别未全部完成：" + error.Message;
            if (summary != null)
            {
                message += $"\r\n已处理图片 {summary.OcrImages} 张，保留原文区域 {summary.OcrKeptRegions} 个。" +
                    $"\r\n未识别到文字 {summary.OcrNoText} 张；需要检查排版/旁注 {summary.ImageNeedsReview} 项。";
                if (summary.UnplacedImageNotes.Count > 0)
                    message += "\r\n\r\n以下旁注未放置，请复制保存；本窗口不会自动关闭。\r\n\r\n" +
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

            var close = new Button
            {
                Text = "关闭",
                AutoSize = true,
                Height = 36,
                Padding = new Padding(16, 4, 16, 4),
                Anchor = AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(45, 108, 223),
                ForeColor = Color.White,
                DialogResult = DialogResult.OK
            };
            close.FlatAppearance.BorderSize = 0;
            root.SetColumnSpan(close, 2);
            root.Controls.Add(close, 0, 2);
            Controls.Add(root);
            AcceptButton = close; CancelButton = close;
        }

        public static void ShowResult(TranslationTaskSummary? summary, IntPtr owner, Exception? error = null)
        {
            using (var form = new OcrResultForm(summary, error))
            {
                if (owner == IntPtr.Zero) form.ShowDialog();
                else form.ShowDialog(new TranslationProgressForm.WindowHandle(owner));
            }
        }
    }
}
