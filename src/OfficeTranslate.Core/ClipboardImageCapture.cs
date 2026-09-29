using System;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace OfficeTranslate.Core
{
    // Captures an Office shape as PNG bytes through the clipboard.
    // Must run on an STA thread (the Office UI thread in all three add-ins).
    //
    // Why this exists: the previous code read whatever image happened to be on
    // the clipboard after CopyAsPicture()/Copy(), so a leftover image from the
    // user (or another app) could be OCR-translated as if it were the shape,
    // and the user's own clipboard content was silently destroyed.
    //
    // What it fixes: the user's clipboard is saved and restored (best-effort),
    // and the clipboard is cleared before copying so a stale image can never be
    // mistaken for the shape.
    //
    // Residual risk (documented, not fixable via clipboard APIs): the clipboard is
    // a shared global. If the user or another program writes an image to the
    // clipboard after our Clear() and before we read the shape's image back
    // (the copy call plus up to 20 x 50 ms of polling, i.e. roughly one second
    // worst case), that foreign image would be captured instead. Saving and
    // restoring the clipboard cannot close this window; a fully race-free
    // capture would require shape-to-file export, which Word/Excel do not
    // expose for arbitrary shapes (PowerPoint uses temp-file export).
    public static class ClipboardImageCapture
    {
        public static byte[] CapturePng(Action copyToClipboard, CancellationToken token) =>
            CapturePng(copyToClipboard, null, token);

        public static byte[] CapturePng(Action copyToClipboard, Action? pumpMessages, CancellationToken token)
        {
            if (copyToClipboard == null) throw new ArgumentNullException(nameof(copyToClipboard));
            var saved = SafeGetDataObject();
            try
            {
                // Must actually clear: if a stale image survives here it would be
                // captured as the shape.
                ClearWithRetry();
                copyToClipboard();

                System.Drawing.Image? image = null;
                for (var attempt = 0; attempt < 20 && image == null; attempt++)
                {
                    token.ThrowIfCancellationRequested();
                    pumpMessages?.Invoke();
                    if (Clipboard.ContainsImage()) image = Clipboard.GetImage();
                    if (image == null) Thread.Sleep(50);
                }
                if (image == null)
                    throw new InvalidOperationException("无法从 Office 图片获取可识别图像。");

                using (image)
                using (var stream = new MemoryStream())
                {
                    image.Save(stream, ImageFormat.Png);
                    return stream.ToArray();
                }
            }
            finally
            {
                // Best-effort: a clipboard failure here must never break the task.
                try
                {
                    if (saved != null) Clipboard.SetDataObject(saved, true);
                    else SafeClear();
                }
                catch { }
            }
        }

        private static IDataObject? SafeGetDataObject()
        {
            try { return Clipboard.GetDataObject(); }
            catch { return null; }
        }

        private static void SafeClear()
        {
            try { Clipboard.Clear(); }
            catch { }
        }

        // The clipboard is a shared global; another app may hold it open. Retry
        // briefly, but never silently continue: a stale image left behind could
        // be captured as the shape, so persistent failure aborts the capture.
        private static void ClearWithRetry()
        {
            for (var attempt = 0; attempt < 5; attempt++)
            {
                try { Clipboard.Clear(); return; }
                catch when (attempt < 4) { Thread.Sleep(100); }
            }
        }
    }
}
