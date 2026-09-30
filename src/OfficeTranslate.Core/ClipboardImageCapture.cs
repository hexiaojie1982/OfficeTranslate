using System;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
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
    // S1: the clipboard is a shared global. Opening it can fail transiently
    // (CLIPBRD_E_CANT_OPEN, "Requested Clipboard operation did not succeed")
    // while another thread or process holds it, and Office delayed rendering
    // may need its message loop pumped before the image can be read back.
    // The capture therefore retries the whole clear -> copy -> read cycle a
    // bounded number of times (message pumping between attempts when the
    // caller supplies one). Every attempt re-clears first, so a stale image
    // can never be mistaken for the shape, and cancellation is honored
    // before every attempt and inside the read poll. Only the transient
    // CLIPBRD_E_CANT_OPEN is retried; every other failure aborts at once
    // with the phase, attempt, thread, and HResult attached.
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
        private const int MaxAttempts = 3;
        private const int ClipbrdECantOpen = unchecked((int)0x800401D0);

        public static byte[] CapturePng(Action copyToClipboard, CancellationToken token) =>
            CapturePng(copyToClipboard, null, token);

        public static byte[] CapturePng(Action copyToClipboard, Action? pumpMessages, CancellationToken token)
        {
            if (copyToClipboard == null) throw new ArgumentNullException(nameof(copyToClipboard));
            int uiThread = Thread.CurrentThread.ManagedThreadId;
            var saved = SafeGetDataObject();
            Exception? lastTransient = null;
            try
            {
                for (var attempt = 1; attempt <= MaxAttempts; attempt++)
                {
                    token.ThrowIfCancellationRequested();
                    try
                    {
                        return TryCaptureOnce(copyToClipboard, pumpMessages, token, attempt, uiThread);
                    }
                    catch (Exception ex) when (IsTransientClipboardFailure(ex))
                    {
                        lastTransient = ex;
                        if (attempt < MaxAttempts)
                            BackoffWithPump(pumpMessages, token);
                    }
                }
                throw new InvalidOperationException(
                    "Office 图片捕获失败：剪贴板被占用，重试 " + MaxAttempts + " 次后仍未成功" +
                    "（线程 " + uiThread + "）。最后一次：" + Describe(lastTransient) + "。",
                    lastTransient);
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

        private static byte[] TryCaptureOnce(Action copyToClipboard, Action? pumpMessages, CancellationToken token, int attempt, int uiThread)
        {
            // Must actually clear: if a stale image survives here it would be
            // captured as the shape. ClearWithRetry aborts (never silently
            // continues) when the clipboard stays unavailable.
            try { ClearWithRetry(); }
            catch (Exception ex) when (!IsCancellation(ex))
            { throw PhaseException("清空剪贴板", attempt, uiThread, ex); }

            try { copyToClipboard(); }
            catch (Exception ex) when (!IsCancellation(ex))
            { throw PhaseException("复制图片到剪贴板", attempt, uiThread, ex); }

            System.Drawing.Image? image = null;
            Exception? readError = null;
            try
            {
                for (var i = 0; i < 20 && image == null; i++)
                {
                    token.ThrowIfCancellationRequested();
                    try
                    {
                        pumpMessages?.Invoke();
                        if (Clipboard.ContainsImage()) image = Clipboard.GetImage();
                    }
                    catch (Exception ex) when (IsTransientClipboardFailure(ex))
                    {
                        // Transient: keep polling inside this attempt's
                        // window; the delayed renderer may need the pumped
                        // messages to answer. A persistent failure falls
                        // through to the outer retry, which re-copies.
                        readError = ex;
                    }
                    if (image == null) Thread.Sleep(50);
                }
            }
            catch (Exception ex) when (!IsCancellation(ex))
            { throw PhaseException("读取剪贴板图片", attempt, uiThread, ex); }

            if (image == null)
            {
                if (readError != null)
                    throw PhaseException("读取剪贴板图片", attempt, uiThread, readError);
                throw new InvalidOperationException(
                    "无法从 Office 图片获取可识别图像（尝试 " + attempt + "/" + MaxAttempts +
                    "，线程 " + uiThread + "）。");
            }

            using (image)
            using (var stream = new MemoryStream())
            {
                image.Save(stream, ImageFormat.Png);
                return stream.ToArray();
            }
        }

        // Brief backoff between attempts with message pumping: the clipboard
        // owner may need its message loop to release the clipboard or to
        // answer a delayed-render request.
        private static void BackoffWithPump(Action? pumpMessages, CancellationToken token)
        {
            for (var i = 0; i < 4; i++)
            {
                token.ThrowIfCancellationRequested();
                try { pumpMessages?.Invoke(); }
                catch { }
                Thread.Sleep(100);
            }
        }

        // S1 diagnostics: every capture failure names the phase, the
        // attempt, the thread, and the underlying exception type + HResult,
        // so a menu failure can be attributed without guessing. No image
        // bytes, clipboard text, prompts, or keys are ever recorded.
        private static InvalidOperationException PhaseException(string phase, int attempt, int uiThread, Exception inner)
        {
            return new InvalidOperationException(
                "Office 图片捕获在“" + phase + "”失败（尝试 " + attempt + "/" + MaxAttempts +
                "，线程 " + uiThread + "，" + Describe(inner) + "）。",
                inner);
        }

        private static string Describe(Exception? ex)
        {
            if (ex == null) return "无详细异常";
            return ex.GetType().Name + "，HResult=0x" +
                Marshal.GetHRForException(ex).ToString("X8") + "：" + ex.Message;
        }

        private static bool IsCancellation(Exception ex) => ex is OperationCanceledException;

        // Only CLIPBRD_E_CANT_OPEN ("Requested Clipboard operation did not
        // succeed") is treated as transient: the clipboard is a shared
        // global and another holder usually releases it quickly. The check
        // walks the inner-exception chain because phase wrappers nest the
        // original error. Every other failure propagates immediately.
        internal static bool IsTransientClipboardFailure(Exception? ex)
        {
            while (ex != null)
            {
                if (ex is ExternalException &&
                    Marshal.GetHRForException(ex) == ClipbrdECantOpen)
                    return true;
                ex = ex.InnerException;
            }
            return false;
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
