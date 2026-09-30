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
    // before every attempt and inside the read poll.
    //
    // S2: a copy that completes WITHOUT throwing but leaves no recognizable
    // image inside the read window is its OWN failure category
    // (NoImageCaptureException), never lumped into transient-busy. It gets
    // the same bounded re-capture (each attempt re-clears then re-copies)
    // because delayed rendering can legitimately need a second chance, but
    // a persistent no-image still aborts with explicit diagnostics instead
    // of being misreported as a clipboard-busy retry. The exception carries
    // no HResult and no underlying exception: there is none to report, and
    // inventing one would mislead the next diagnosis.
    //
    // S2/D1 diagnostics: every capture failure names the phase (or NoImage),
    // the attempt, the thread, and -- for the no-image path -- the clipboard
    // sequence numbers before/after the copy (did the copy write anything at
    // all?), the poll count, whether ContainsImage was ever true, and the
    // format names present after the copy. No image bytes, clipboard text,
    // prompts, or keys are ever recorded.
    //
    // D1: ClearWithRetry below predates the outer cycle and retries Clear()
    // itself up to 5 times (catch-all, short sleeps) before aborting; the
    // "only transient-busy / no-image are retried, everything else aborts at
    // once" rule applies to the outer clear -> copy -> read cycle and to the
    // copy/read phases, not to Clear's own internal retries.
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

        public static byte[] CapturePng(Action<int> copyToClipboard, CancellationToken token) =>
            CapturePng(copyToClipboard, null, token);

        // D1: the copy delegate receives the real 1-based capture attempt.
        // The attempt loop clears the clipboard BEFORE invoking the
        // delegate, and a clear-phase failure retries without ever calling
        // it -- so a counter inside the delegate would undercount. Passing
        // the attempt explicitly keeps per-attempt diagnostics (and the
        // "attempt N/3" message) on the loop's numbering by construction.
        public static byte[] CapturePng(Action<int> copyToClipboard, Action? pumpMessages, CancellationToken token)
        {
            if (copyToClipboard == null) throw new ArgumentNullException(nameof(copyToClipboard));
            int uiThread = Thread.CurrentThread.ManagedThreadId;
            var saved = SafeGetDataObject();
            Exception? lastTransient = null;
            Exception? lastNoImage = null;
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
                        lastNoImage = null;
                        if (attempt < MaxAttempts)
                            BackoffWithPump(pumpMessages, token);
                    }
                    catch (Exception ex) when (ex is NoImageCaptureException)
                    {
                        // S2: bounded re-capture for the no-image category.
                        // Each attempt re-clears then re-copies, so a delayed
                        // render gets another chance but a stale image can
                        // never be read. Kept separate from transient-busy:
                        // the two have different causes and different fixes.
                        lastNoImage = ex;
                        lastTransient = null;
                        if (attempt < MaxAttempts)
                            BackoffWithPump(pumpMessages, token);
                    }
                }
                if (lastNoImage != null) throw lastNoImage;
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

        private static byte[] TryCaptureOnce(Action<int> copyToClipboard, Action? pumpMessages, CancellationToken token, int attempt, int uiThread)
        {
            // Must actually clear: if a stale image survives here it would be
            // captured as the shape. ClearWithRetry aborts (never silently
            // continues) when the clipboard stays unavailable.
            try { ClearWithRetry(); }
            catch (Exception ex) when (!IsCancellation(ex))
            { throw PhaseException("清空剪贴板", attempt, uiThread, ex); }

            uint seqBeforeCopy, seqAfterCopy;
            string[] formatsAfterCopy;
            try
            {
                seqBeforeCopy = ClipboardSequenceNumber();
                copyToClipboard(attempt);
                seqAfterCopy = ClipboardSequenceNumber();
                formatsAfterCopy = SafeGetFormats();
            }
            catch (Exception ex) when (!IsCancellation(ex))
            { throw PhaseException("复制图片到剪贴板", attempt, uiThread, ex); }

            System.Drawing.Image? image = null;
            Exception? readError = null;
            int polls = 0;
            bool containsImageSeen = false;
            try
            {
                for (var i = 0; i < 20 && image == null; i++)
                {
                    token.ThrowIfCancellationRequested();
                    polls++;
                    try
                    {
                        pumpMessages?.Invoke();
                        if (Clipboard.ContainsImage())
                        {
                            containsImageSeen = true;
                            image = Clipboard.GetImage();
                        }
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
                throw NoImageException(attempt, uiThread, seqBeforeCopy, seqAfterCopy,
                    polls, containsImageSeen, formatsAfterCopy);
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

        // S2: the no-image failure is its own category. It carries the
        // attempt count that ACTUALLY ran (no "1/3 implies three ran"
        // ambiguity), an explicit NoImage result, and the copy/read
        // observations -- never a fabricated HResult or a fake
        // CLIPBRD_E_CANT_OPEN.
        internal sealed class NoImageCaptureException : InvalidOperationException
        {
            public NoImageCaptureException(string message) : base(message) { }
        }

        private static NoImageCaptureException NoImageException(
            int attempt, int uiThread,
            uint seqBeforeCopy, uint seqAfterCopy,
            int polls, bool containsImageSeen, string[] formatsAfterCopy)
        {
            // The clipboard sequence number is process-global: it can also
            // change because of another process, and the observation window
            // sits between Copy's return and the read poll. Record the
            // observation only; do not present it as proof of the root cause.
            string seqPart;
            if (seqBeforeCopy == 0 && seqAfterCopy == 0)
                seqPart = "剪贴板序号不可用";
            else if (seqBeforeCopy == seqAfterCopy)
                seqPart = "复制返回后剪贴板序号未变化（未观察到剪贴板写入；全局序号也可能被其他进程改变）";
            else
                seqPart = "复制返回后剪贴板序号变化（观察到剪贴板写入，但不是可识别图像；全局序号也可能被其他进程改变）";
            string formatPart = formatsAfterCopy == null || formatsAfterCopy.Length == 0
                ? "无"
                : string.Join(",", formatsAfterCopy);
            string getImagePart = containsImageSeen
                ? "ContainsImage 曾为真但 GetImage 返回 null"
                : "ContainsImage 从未为真";
            return new NoImageCaptureException(
                "Office 图片捕获在“读取剪贴板图片”失败：复制完成但 " + polls +
                " 次轮询内没有可识别图像（实际尝试 " + attempt + "/" + MaxAttempts +
                "，线程 " + uiThread + "）。结果=NoImage，无底层异常。" +
                seqPart + "（序号 " + seqBeforeCopy + "→" + seqAfterCopy + "）。" +
                getImagePart + "。复制后剪贴板格式：" + formatPart + "。");
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
        // succeed") is treated as transient-busy: the clipboard is a shared
        // global and another holder usually releases it quickly. The check
        // walks the inner-exception chain because phase wrappers nest the
        // original error. NoImageCaptureException never matches (it has no
        // inner exception and no HResult); it is retried by its own
        // category in the outer loop, never here. Every other failure
        // propagates immediately.
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

        private static string[] SafeGetFormats()
        {
            try
            {
                var data = Clipboard.GetDataObject();
                return data != null ? data.GetFormats() : Array.Empty<string>();
            }
            catch { return Array.Empty<string>(); }
        }

        // S2: the OS clipboard sequence number. It changes every time the
        // clipboard content is replaced, so comparing before/after the copy
        // tells whether the copy wrote anything at all. Windows-only; 0
        // means unavailable (never fabricated into a diagnosis).
        private static class NativeClipboard
        {
            [DllImport("user32.dll")]
            public static extern uint GetClipboardSequenceNumber();
        }

        private static uint ClipboardSequenceNumber()
        {
            try
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    return NativeClipboard.GetClipboardSequenceNumber();
            }
            catch { }
            return 0;
        }

        // The clipboard is a shared global; another app may hold it open. Retry
        // briefly, but never silently continue: a stale image left behind could
        // be captured as the shape, so persistent failure aborts the capture.
        // D1: this internal 5-attempt catch-all predates the outer cycle and
        // covers Clear() only; the outer "only transient-busy / no-image are
        // retried" rule is unchanged.
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
