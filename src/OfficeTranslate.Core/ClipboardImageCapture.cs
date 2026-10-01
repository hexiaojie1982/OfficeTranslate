using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
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
    // 4ca3b46 review P1/P2: when the bitmap poll finds nothing, a native
    // CF_ENHMETAFILE offer is tried via Win32 handle duplication (the
    // managed DataObject does not reliably recognize clipboard metafiles).
    // The NoImage diagnostic now carries the native format enumeration
    // (IDs/names only, never content) and the per-stage EMF diagnostic, so
    // "clipboard truly empty" can be told apart from "EMF offered but
    // unreadable". Cancellation is never swallowed by the EMF path;
    // temporarily unavailable data gets bounded retries; rasterization has
    // hard pixel caps and releases the Bitmap on every failure path.
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
            return CapturePng(copyToClipboard, pumpMessages, token, new Win32ClipboardReader());
        }

        // 4ca3b46 review P2: internal test seam -- the clipboard is injected
        // so unit tests can drive the capture loop without a real clipboard.
        internal static byte[] CapturePng(Action<int> copyToClipboard, Action? pumpMessages, CancellationToken token, IClipboardReader clipboard)
        {
            if (copyToClipboard == null) throw new ArgumentNullException(nameof(copyToClipboard));
            if (clipboard == null) throw new ArgumentNullException(nameof(clipboard));
            int uiThread = Thread.CurrentThread.ManagedThreadId;
            IDataObject? saved = null;
            try { saved = clipboard.GetDataObject(); }
            catch { saved = null; }
            Exception? lastTransient = null;
            // 4ca3b46 review P2: keep EVERY attempt's NoImage evidence, not
            // just the most recent -- a hard failure after several NoImages
            // used to drop the early attempts.
            var noImageAttempts = new List<NoImageCaptureException>();
            try
            {
                for (var attempt = 1; attempt <= MaxAttempts; attempt++)
                {
                    token.ThrowIfCancellationRequested();
                    try
                    {
                        return TryCaptureOnce(copyToClipboard, pumpMessages, token, attempt, uiThread, clipboard);
                    }
                    catch (Exception ex) when (IsTransientClipboardFailure(ex))
                    {
                        lastTransient = ex;
                        noImageAttempts.Clear();
                        if (attempt < MaxAttempts)
                            BackoffWithPump(pumpMessages, token);
                    }
                    catch (Exception ex) when (ex is NoImageCaptureException noImage)
                    {
                        // S2: bounded re-capture for the no-image category.
                        // Each attempt re-clears then re-copies, so a delayed
                        // render gets another chance but a stale image can
                        // never be read. Kept separate from transient-busy:
                        // the two have different causes and different fixes.
                        noImageAttempts.Add(noImage);
                        lastTransient = null;
                        if (attempt < MaxAttempts)
                            BackoffWithPump(pumpMessages, token);
                    }
                    catch (Exception ex) when (!IsCancellation(ex))
                    {
                        // f2cb7de review item 2 + 4ca3b46 review P2: a hard
                        // copy-phase failure (e.g. Word 0x800A11FD "command
                        // invalid") must not silently discard the earlier
                        // attempts' NoImage evidence (clipboard sequence
                        // numbers, formats, poll counts). Chain ALL attempts
                        // so the final error -- and the host's capture-failure
                        // log line -- keeps the per-attempt evidence instead
                        // of only the last exception. Cancellation still
                        // propagates unwrapped.
                        if (noImageAttempts.Count > 0)
                            throw new InvalidOperationException(
                                "Office 图片捕获失败：复制阶段出错（" + Describe(ex) + "）。"
                                + CombineAttemptEvidence(noImageAttempts),
                                ex);
                        throw;
                    }
                }
                if (noImageAttempts.Count > 0)
                    throw new InvalidOperationException(
                        "Office 图片捕获失败：各轮尝试均未取得图片。"
                        + CombineAttemptEvidence(noImageAttempts),
                        noImageAttempts[noImageAttempts.Count - 1]);
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
                    if (saved != null) clipboard.SetDataObject(saved);
                    else clipboard.Clear();
                }
                catch { }
            }
        }

        // 4ca3b46 review P2: aggregates per-attempt NoImage evidence with a
        // length cap, so a hard failure after several NoImages keeps the
        // early attempts instead of only the most recent one.
        // 3b2a252 review: the cap is applied AFTER building the full text --
        // checking before each append still lets one long round push the
        // total over the limit.
        private static string CombineAttemptEvidence(List<NoImageCaptureException> attempts)
        {
            var sb = new StringBuilder("此前各轮取图证据（共 " + attempts.Count + " 轮）：");
            for (var i = 0; i < attempts.Count; i++)
            {
                sb.Append("[第").Append(i + 1).Append("轮 ").Append(attempts[i].Message).Append("] ");
            }
            // afa3812 review P2: the old code truncated to maxLength and THEN
            // appended the marker, so the result could strictly exceed the
            // cap. Reserve the marker's length up front, and never leave a
            // dangling high surrogate at the cut point.
            const int maxLength = 2000;
            const string truncMark = "…（已截断）";
            if (sb.Length > maxLength)
            {
                sb.Length = maxLength - truncMark.Length;
                if (sb.Length > 0 && char.IsHighSurrogate(sb[sb.Length - 1])) sb.Length--;
                sb.Append(truncMark);
            }
            return sb.ToString();
        }

        private static byte[] TryCaptureOnce(Action<int> copyToClipboard, Action? pumpMessages, CancellationToken token, int attempt, int uiThread, IClipboardReader clipboard)
        {
            // Must actually clear: if a stale image survives here it would be
            // captured as the shape. ClearWithRetry aborts (never silently
            // continues) when the clipboard stays unavailable.
            try { ClearWithRetry(clipboard); }
            catch (Exception ex) when (!IsCancellation(ex))
            { throw PhaseException("清空剪贴板", attempt, uiThread, ex); }

            uint seqBeforeCopy, seqAfterCopy;
            string[]? formatsAfterCopy;
            try
            {
                seqBeforeCopy = clipboard.GetSequenceNumber();
                copyToClipboard(attempt);
                seqAfterCopy = clipboard.GetSequenceNumber();
                formatsAfterCopy = clipboard.GetFormats();
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
                        if (clipboard.ContainsImage())
                        {
                            containsImageSeen = true;
                            image = clipboard.GetImage();
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

            // 4ca3b46 review P1: when the bitmap path finds nothing, try the
            // NATIVE CF_ENHMETAFILE offer. The managed DataObject does not
            // reliably recognize clipboard metafiles (documented
            // limitation), so the authoritative check duplicates the Win32
            // handle directly. Purely additive: the bitmap path above is
            // untouched, and when no metafile is present (or it cannot be
            // rendered) the original NoImage failure is reported unchanged
            // -- now with the native format enumeration and the per-stage
            // EMF diagnostic, so the next round can tell "clipboard truly
            // empty" apart from "EMF offered but unreadable".
            string emfDiag = "absent";
            if (image == null)
                image = TryPollMetafileImage(clipboard, pumpMessages, token, out emfDiag);

            List<string>? nativeFormats = null;
            if (image == null)
                nativeFormats = clipboard.GetNativeFormats();

            if (image == null)
            {
                if (readError != null)
                    throw PhaseException("读取剪贴板图片", attempt, uiThread, readError);
                throw NoImageException(attempt, uiThread, seqBeforeCopy, seqAfterCopy,
                    polls, containsImageSeen, formatsAfterCopy, nativeFormats, emfDiag);
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

        // 4ca3b46 review P1: polls for a native CF_ENHMETAFILE offer and
        // rasterizes it. Cancellation propagates (never swallowed, even from
        // the pump callback); temporarily unavailable data is retried
        // boundedly with the final stage recorded; unexpected errors are
        // recorded as "error:<type>", never disguised as "no EMF".
        // emfDiag is the final stage: "ok", "absent", "open-failed",
        // "dup-failed", "unreadable", "unavailable", "over-limit:WxH",
        // "bad-size", "error:<type>", "pump-error:<type>", "transient-busy".
        private static System.Drawing.Image? TryPollMetafileImage(
            IClipboardReader clipboard, Action? pumpMessages, CancellationToken token, out string emfDiag)
        {
            emfDiag = "absent";
            string lastStage = "absent";
            for (var i = 0; i < 5; i++)
            {
                // Cancellation must propagate, never be swallowed.
                token.ThrowIfCancellationRequested();
                try { pumpMessages?.Invoke(); }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex) when (IsTransientClipboardFailure(ex))
                { lastStage = "transient-busy"; Thread.Sleep(50); continue; }
                catch (Exception ex)
                { lastStage = "pump-error:" + ex.GetType().Name; Thread.Sleep(50); continue; }

                Metafile? mf;
                string stage;
                try { mf = clipboard.GetNativeEnhMetafile(out stage); }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex) when (IsTransientClipboardFailure(ex))
                { lastStage = "transient-busy"; Thread.Sleep(50); continue; }
                catch (Exception ex)
                { lastStage = "error:" + ex.GetType().Name; Thread.Sleep(50); continue; }

                if (mf == null)
                {
                    // Offered-but-not-ready (delayed rendering) and absent
                    // are told apart by the stage; both get bounded retries.
                    lastStage = stage;
                    Thread.Sleep(50);
                    continue;
                }

                using (mf)
                {
                    var img = RasterizeMetafile(mf, out var renderDiag);
                    if (img != null) { emfDiag = "ok"; return img; }
                    lastStage = renderDiag;
                    // A definitive render verdict (oversize, bad size,
                    // corrupt) cannot become readable by polling more;
                    // stop early with the reason kept.
                    if (renderDiag.StartsWith("over-limit", StringComparison.Ordinal)
                        || renderDiag.StartsWith("error", StringComparison.Ordinal)
                        || renderDiag.StartsWith("bad-size", StringComparison.Ordinal))
                    {
                        emfDiag = lastStage;
                        return null;
                    }
                    Thread.Sleep(50);
                }
            }
            emfDiag = lastStage;
            return null;
        }

        // 4ca3b46 review P2: hard caps for EMF rasterization. 8000px per
        // side and 32M pixels total (~128MB at 32bpp) -- far above any Word
        // inline image at print resolution, far below the ~576MB a
        // 12000x12000 bitmap would allocate.
        private const int MaxMetafileDimensionPx = 8000;
        private const long MaxMetafilePixels = 32_000_000;

        // 4ca3b46 review P2: rasterizes with guaranteed resource release.
        // The Bitmap is disposed on every failure path; ownership transfers
        // to the caller only on success. diag: "ok", "bad-size",
        // "over-limit:WxH", "error:<type>".
        internal static System.Drawing.Image? RasterizeMetafile(Metafile? mf, out string diag)
        {
            diag = "error:Unknown";
            Bitmap? bmp = null;
            try
            {
                if (mf == null) { diag = "bad-size"; return null; }
                int w = mf.Width, h = mf.Height;
                if (w <= 0 || h <= 0) { diag = "bad-size"; return null; }
                if (w > MaxMetafileDimensionPx || h > MaxMetafileDimensionPx
                    || (long)w * h > MaxMetafilePixels)
                {
                    diag = "over-limit:" + w + "x" + h;
                    return null;
                }
                bmp = new Bitmap(w, h);
                bmp.SetResolution(mf.HorizontalResolution, mf.VerticalResolution);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.White);
                    g.DrawImage(mf, 0, 0, w, h);
                }
                diag = "ok";
                var result = bmp;
                bmp = null; // ownership transfers to the caller
                return result;
            }
            catch (Exception ex)
            {
                diag = "error:" + ex.GetType().Name;
                return null;
            }
            finally
            {
                // Non-null only when we did NOT hand the bitmap to the
                // caller: a SetResolution/DrawImage failure used to leak it.
                bmp?.Dispose();
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
            int polls, bool containsImageSeen, string[]? formatsAfterCopy,
            List<string>? nativeFormats, string emfDiag)
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
            // 4ca3b46 review P2: a failed format READ is reported as
            // read-failed, never as "none" -- "none" now strictly means the
            // enumeration succeeded and the clipboard offered no formats.
            // The managed enumeration can miss native metafile offers
            // (documented DataObject limitation), so the native list is the
            // authoritative one for the EMF question.
            string formatPart = formatsAfterCopy == null
                ? "读取失败"
                : formatsAfterCopy.Length == 0 ? "无" : string.Join(",", formatsAfterCopy);
            string nativePart = nativeFormats == null
                ? "读取失败"
                : nativeFormats.Count == 0 ? "无" : string.Join(",", nativeFormats);
            string getImagePart = containsImageSeen
                ? "ContainsImage 曾为真但 GetImage 返回 null"
                : "ContainsImage 从未为真";
            return new NoImageCaptureException(
                "Office 图片捕获在“读取剪贴板图片”失败：复制完成但 " + polls +
                " 次轮询内没有可识别图像（实际尝试 " + attempt + "/" + MaxAttempts +
                "，线程 " + uiThread + "）。结果=NoImage，无底层异常。" +
                seqPart + "（序号 " + seqBeforeCopy + "→" + seqAfterCopy + "）。" +
                getImagePart + "。复制后剪贴板格式（托管）：" + formatPart +
                "；原生格式：" + nativePart + "；EMF：" + emfDiag + "。");
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

        // 4ca3b46 review P2: null = the read itself failed (distinct from
        // "no formats"). A failed read must never be reported as "none".
        private static string[]? SafeGetFormats()
        {
            try
            {
                var data = Clipboard.GetDataObject();
                return data != null ? data.GetFormats() : Array.Empty<string>();
            }
            catch { return null; }
        }

        // The real clipboard behind IClipboardReader. Runs on the Office UI
        // (STA) thread like all clipboard access in this class.
        private sealed class Win32ClipboardReader : IClipboardReader
        {
            public void Clear() => Clipboard.Clear();
            public IDataObject? GetDataObject()
            {
                try { return Clipboard.GetDataObject(); }
                catch { return null; }
            }
            public void SetDataObject(IDataObject data) => Clipboard.SetDataObject(data, true);
            public uint GetSequenceNumber() => ClipboardSequenceNumber();
            public bool ContainsImage() => Clipboard.ContainsImage();
            public System.Drawing.Image? GetImage() => Clipboard.GetImage();
            public string[]? GetFormats() => SafeGetFormats();
            public List<string>? GetNativeFormats() => EnumNativeClipboardFormats();
            public Metafile? GetNativeEnhMetafile(out string diag) => ReadNativeEnhMetafile(out diag);
        }

        // S2: the OS clipboard sequence number. It changes every time the
        // clipboard content is replaced, so comparing before/after the copy
        // tells whether the copy wrote anything at all. Windows-only; 0
        // means unavailable (never fabricated into a diagnosis).
        private const uint CF_ENHMETAFILE = 14;

        private static class NativeClipboard
        {
            [DllImport("user32.dll")]
            public static extern uint GetClipboardSequenceNumber();

            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool OpenClipboard(IntPtr hWndNewOwner);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool CloseClipboard();

            // 3b2a252 review: SetLastError is required -- a 0 return means
            // "no more formats" OR failure, and only GetLastError() ==
            // ERROR_SUCCESS proves a clean end of enumeration.
            [DllImport("user32.dll", SetLastError = true)]
            public static extern uint EnumClipboardFormats(uint format);

            [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            public static extern int GetClipboardFormatName(uint format, StringBuilder lpszFormatName, int cchMaxCount);

            [DllImport("user32.dll")]
            public static extern IntPtr GetClipboardData(uint uFormat);

            [DllImport("gdi32.dll", SetLastError = true)]
            public static extern IntPtr CopyEnhMetaFile(IntPtr hemfSrc, string? lpszFile);

            [DllImport("gdi32.dll", SetLastError = true)]
            public static extern bool DeleteEnhMetaFile(IntPtr hemf);
        }

        // 4ca3b46 review P1: native clipboard format enumeration for
        // diagnostics ONLY -- format IDs/names, never content. Returns null
        // when the clipboard cannot be opened (open-failed is distinct from
        // empty). This is the authoritative answer to "is EMF actually on
        // the clipboard?" -- the managed GetFormats() can miss native
        // metafile offers (documented DataObject limitation), which is why
        // the previous round's "formats: none" proved nothing.
        private static List<string>? EnumNativeClipboardFormats()
        {
            var formats = new List<string>();
            try
            {
                if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return null;
                if (!NativeClipboard.OpenClipboard(IntPtr.Zero)) return null;
                try
                {
                    // 3b2a252 review: per MS docs a 0 return means "no more
                    // formats" OR failure. Only GetLastError() ==
                    // ERROR_SUCCESS (0) proves a clean end of enumeration;
                    // any other code means the list may be incomplete and
                    // must be reported as read-failed, never as "none".
                    // https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumclipboardformats
                    bool reachedEnd = false;
                    bool hitCap = false;
                    uint f = 0;
                    while (true)
                    {
                        // afa3812 review P2: the old silent 64-item cap could
                        // be misread as a complete enumeration; mark the
                        // truncation so a partial list is never presented as
                        // the full native format set.
                        if (formats.Count >= 64) { hitCap = true; break; }
                        f = NativeClipboard.EnumClipboardFormats(f);
                        if (f == 0) { reachedEnd = true; break; }
                        formats.Add(NativeFormatLabel(f));
                    }
                    if (hitCap) formats.Add("…(仅列出前64项)");
                    if (reachedEnd && Marshal.GetLastWin32Error() != 0) return null;
                }
                finally { NativeClipboard.CloseClipboard(); }
            }
            catch { return null; }
            return formats;
        }

        private static string NativeFormatLabel(uint f)
        {
            switch (f)
            {
                case 1: return "1=CF_TEXT";
                case 2: return "2=CF_BITMAP";
                case 3: return "3=CF_METAFILEPICT";
                case 8: return "8=CF_DIB";
                case 13: return "13=CF_UNICODETEXT";
                case 14: return "14=CF_ENHMETAFILE";
                case 16: return "16=CF_LOCALE";
                case 17: return "17=CF_DIBV5";
                default: break;
            }
            try
            {
                var sb = new StringBuilder(128);
                if (NativeClipboard.GetClipboardFormatName(f, sb, sb.Capacity) > 0)
                    return f + "=" + sb.ToString();
            }
            catch { }
            return f.ToString();
        }

        // 4ca3b46 review P1: Win32 CF_ENHMETAFILE retrieval. Duplicates the
        // native handle -- the clipboard owns the original and it must never
        // be deleted. The returned Metafile takes ownership of the duplicate
        // and deletes it on dispose. diag stages: "absent" (no EMF offered),
        // "open-failed", "dup-failed", "unreadable", "error:<type>".
        private static Metafile? ReadNativeEnhMetafile(out string diag)
        {
            diag = "absent";
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return null;
            IntPtr copy = IntPtr.Zero;
            try
            {
                if (!NativeClipboard.OpenClipboard(IntPtr.Zero)) { diag = "open-failed"; return null; }
                try
                {
                    IntPtr hEmf = NativeClipboard.GetClipboardData(CF_ENHMETAFILE);
                    if (hEmf == IntPtr.Zero) { diag = "absent"; return null; }
                    copy = NativeClipboard.CopyEnhMetaFile(hEmf, null);
                    if (copy == IntPtr.Zero) { diag = "dup-failed"; return null; }
                }
                finally { NativeClipboard.CloseClipboard(); }
            }
            catch (Exception ex)
            {
                diag = "error:" + ex.GetType().Name;
                return null;
            }
            try
            {
                return new Metafile(copy, true);
            }
            catch
            {
                NativeClipboard.DeleteEnhMetaFile(copy);
                diag = "unreadable";
                return null;
            }
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
        private static void ClearWithRetry(IClipboardReader clipboard)
        {
            for (var attempt = 0; attempt < 5; attempt++)
            {
                try { clipboard.Clear(); return; }
                catch when (attempt < 4) { Thread.Sleep(100); }
            }
        }
    }

    // 4ca3b46 review P1/P2: test seam. All clipboard access in the capture
    // loop goes through this abstraction so unit tests can drive the loop
    // (real EMF, temporarily unavailable, no data, exceptions, cancellation,
    // oversize) without a real clipboard or Office.
    internal interface IClipboardReader
    {
        // Clipboard lifecycle (save/clear/restore). May throw; the caller
        // keeps the existing best-effort semantics around them.
        void Clear();
        IDataObject? GetDataObject();
        void SetDataObject(IDataObject data);
        uint GetSequenceNumber();
        bool ContainsImage();
        System.Drawing.Image? GetImage();
        // Managed format names; null = the read itself failed (distinct from
        // "no formats"). A failed read must never be reported as "none".
        string[]? GetFormats();
        // Native format IDs/names via Win32 enumeration; null = open/enum
        // failed. Diagnostic only -- no clipboard content is ever read.
        List<string>? GetNativeFormats();
        // Win32 CF_ENHMETAFILE retrieval with handle duplication (the
        // clipboard owns the original handle). Returns a caller-owned
        // Metafile, or null; diag is the final stage: "absent",
        // "open-failed", "dup-failed", "unreadable", "error:<type>".
        Metafile? GetNativeEnhMetafile(out string diag);
    }
}
