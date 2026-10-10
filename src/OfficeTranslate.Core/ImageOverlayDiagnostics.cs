using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeTranslate.Core
{
    // Per-image OCR overlay diagnostics. Records ONLY geometry numbers:
    // captured pixel size, the model's raw 0-1000 bbox, the Office image rect
    // (with its reference frame noted), and the final textbox rect. Never
    // image bytes, prompt text, API keys, or the OCR text itself, so the log
    // is safe to attach to a bug report. On a real machine, comparing the
    // logged bbox against the PNG tells model error apart from
    // coordinate-conversion error.
    public sealed class ImageOverlayDiagnosticEntry
    {
        public string Host = string.Empty;
        public string ImageKind = string.Empty;
        public int PixelWidth;
        public int PixelHeight;
        public float BboxX1;
        public float BboxY1;
        public float BboxX2;
        public float BboxY2;
        public float ShapeLeft;
        public float ShapeTop;
        public float ShapeWidth;
        public float ShapeHeight;
        public string FrameNote = string.Empty;
        public float RotationDegrees;
        public bool FlipHorizontal;
        public bool FlipVertical;
        public float OutLeft;
        public float OutTop;
        public float OutWidth;
        public float OutHeight;
        public float FontSize;
        public float WidthGrowthFactor;
        public float HeightGrowthFactor;
        public string Verdict = string.Empty;
        public string Reason = string.Empty;
        public int TextLength;
        public string OwnerId = string.Empty;
    }

    public static class ImageOverlayDiagnostics
    {
        private const string LogFileName = "image-overlay.log";
        private static readonly object Gate = new object();

        // Records which add-in build actually served a translation request.
        // The menu-acceptance test reads the newest line of this log after
        // clicking the ribbon button to prove the expected DLL was loaded,
        // instead of inferring it from registry keys. The build id (module
        // MVID, unique per compilation) distinguishes candidate builds that
        // share the same assembly version, so a stale cached DLL cannot be
        // mistaken for the candidate under test.
        public static void LogVersion(string host, string version, string buildId)
        {
            if (string.IsNullOrEmpty(host)) return;
            var line = string.Format(CultureInfo.InvariantCulture,
                "{0:O} host={1} addin_version={2} build={3} event=translate_start",
                DateTime.UtcNow, host,
                string.IsNullOrEmpty(version) ? "unknown" : version,
                string.IsNullOrEmpty(buildId) ? "unknown" : buildId);
            Trace.WriteLine("OfficeTranslate: " + line);
            try
            {
                var directory = Path.Combine(Path.GetTempPath(), "OfficeTranslate");
                lock (Gate)
                {
                    Directory.CreateDirectory(directory);
                    File.AppendAllText(Path.Combine(directory, LogFileName), line + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch (Exception)
            {
                // Diagnostics must never break translation.
            }
        }

        // M2: thread/apartment diagnostics. Records which thread served each
        // stage (menu entry, around awaits, around each capture): managed
        // thread id, COM apartment state, the ambient sync-context type and
        // whether it is the captured Office UI thread. No document content.
        public static void LogThreadProbe(string host, string stage, int threadId, string apartment, string syncContext, bool isUiThread)
        {
            if (string.IsNullOrEmpty(host)) return;
            var line = string.Format(CultureInfo.InvariantCulture,
                "{0:O} host={1} event=thread_probe stage={2} thread={3} apartment={4} syncctx={5} is_ui={6}",
                DateTime.UtcNow, host, stage, threadId,
                string.IsNullOrEmpty(apartment) ? "unknown" : apartment,
                string.IsNullOrEmpty(syncContext) ? "null" : syncContext,
                isUiThread);
            Trace.WriteLine("OfficeTranslate: " + line);
            try
            {
                var directory = Path.Combine(Path.GetTempPath(), "OfficeTranslate");
                lock (Gate)
                {
                    Directory.CreateDirectory(directory);
                    File.AppendAllText(Path.Combine(directory, LogFileName), line + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch (Exception)
            {
                // Diagnostics must never break translation.
            }
        }

        public static void Log(ImageOverlayDiagnosticEntry entry)
        {
            if (entry == null) return;
            var line = Format(entry);
            AppendLine(line);
        }

        // S2/D1: capture failures must reach the diagnostics log file, not
        // just the popup. detail carries the phase/NoImage result, attempt,
        // thread, HResult (when there is one), clipboard sequence numbers,
        // poll observations and format names -- never image bytes, clipboard
        // text, prompts, or keys. Newlines are flattened so the log stays
        // one line per event.
        public static void LogCaptureFailure(string host, string detail)
        {
            if (string.IsNullOrEmpty(host)) return;
            var safe = (detail ?? "unknown").Replace("\r", " ").Replace("\n", " ");
            var line = string.Format(CultureInfo.InvariantCulture,
                "{0:O} host={1} event=capture_failed detail={2}",
                DateTime.UtcNow, host, safe);
            AppendLine(line);
        }

        // Only allow known finish-reason tokens into the metadata log; arbitrary
        // service strings or exceptions could contain echoed prompts/content.
        public static void LogOcrSummary(string host, int width, int height, ImageRecognitionStatistics? stats,
            string finish, bool truncated, string status, long milliseconds, string captureSource = "unknown", string captureReason = "unknown") =>
            AppendLine(FormatOcrSummary(host, width, height, stats, finish, truncated, status, milliseconds, captureSource, captureReason));

        internal static string FormatOcrSummary(string host, int width, int height, ImageRecognitionStatistics? s,
            string finish, bool truncated, string status, long milliseconds, string captureSource = "unknown", string captureReason = "unknown")
        {
            var safeFinish = new[] { "stop", "length", "end_turn", "eos_token", "tool_calls", "content_filter", "unknown" };
            if (Array.IndexOf(safeFinish, finish) < 0) finish = "other";
            if (host != "Word" && host != "Excel" && host != "PowerPoint") host = "Image";
            return string.Format(CultureInfo.InvariantCulture,
                "{0:O} host={1} event=ocr_summary png={2}x{3} parsed={4} raw={5} kept={6} dropped={7} bbox_invalid={8} bbox_note={9} empty_translation={10} empty_translation_dropped={11} non_dictionary={12} review={13} finish_reason={14} truncated={15} status={16} elapsed_ms={17} capture={18} capture_reason={19}",
                DateTime.UtcNow, host, width, height, s != null, s?.RawCount ?? -1, s?.RetainedCount ?? -1,
                s?.DroppedCount ?? -1, s?.InvalidBboxCount ?? -1, s?.InvalidBboxDowngradedCount ?? -1,
                s?.EmptyTranslationCount ?? -1, s?.EmptyTranslationDroppedCount ?? -1, s?.NonDictionaryCount ?? -1,
                s?.ReviewCount ?? -1, finish, truncated, status, milliseconds,
                captureSource == "original" || captureSource == "fallback" ? captureSource : "unknown", SafeCaptureReason(captureReason));
        }

        private static string SafeCaptureReason(string reason) =>
            !string.IsNullOrEmpty(reason) && reason.Length <= 180 && System.Text.RegularExpressions.Regex.IsMatch(reason, "^[a-zA-Z0-9_-]+$") ? reason : "unknown";

        public static void LogOcrCapture(string host, string path, string reason, int sourceWidth, int sourceHeight, byte[] png)
        {
            PngDimensions.TryRead(png, out var width, out var height);
            AppendLine(string.Format(CultureInfo.InvariantCulture,
                "{0:O} host={1} event=ocr_capture path={2} reason={3} source={4}x{5} png={6}x{7} bytes={8}",
                DateTime.UtcNow, host, path, reason, sourceWidth, sourceHeight, width, height, png.Length));
        }

        private static void AppendLine(string line)
        {
            Trace.WriteLine("OfficeTranslate: " + line);
            try
            {
                var directory = Path.Combine(Path.GetTempPath(), "OfficeTranslate");
                lock (Gate)
                {
                    Directory.CreateDirectory(directory);
                    File.AppendAllText(Path.Combine(directory, LogFileName), line + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch (Exception)
            {
                // Diagnostics must never break translation.
            }
        }

        internal static string Format(ImageOverlayDiagnosticEntry e)
        {
            var culture = CultureInfo.InvariantCulture;
            return string.Format(culture,
                "{0:O} host={1} kind={2} png={3}x{4} bbox=[{5},{6},{7},{8}] shape=[{9},{10},{11},{12}]pt frame={13} rot={14} flipH={15} flipV={16} out=[{17},{18},{19},{20}] font={21} textLen={22} verdict={23} reason={24} owner={25} grow={26}",
                DateTime.UtcNow, e.Host, e.ImageKind, e.PixelWidth, e.PixelHeight,
                F(e.BboxX1), F(e.BboxY1), F(e.BboxX2), F(e.BboxY2),
                F(e.ShapeLeft), F(e.ShapeTop), F(e.ShapeWidth), F(e.ShapeHeight), e.FrameNote,
                F(e.RotationDegrees), e.FlipHorizontal, e.FlipVertical,
                F(e.OutLeft), F(e.OutTop), F(e.OutWidth), F(e.OutHeight),
                F(e.FontSize), e.TextLength, e.Verdict, e.Reason ?? string.Empty, e.OwnerId ?? string.Empty,
                e.Verdict == ImageOverlayVerdict.SideNote.ToString() ? "0.0"
                    : e.WidthGrowthFactor.ToString("0.0", culture) + "x" + e.HeightGrowthFactor.ToString("0.0", culture));
        }

        private static string F(float value)
        {
            return value.ToString("0.##", CultureInfo.InvariantCulture);
        }
    }
}
