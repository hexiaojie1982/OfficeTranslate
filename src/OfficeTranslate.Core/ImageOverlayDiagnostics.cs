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
        // instead of inferring it from registry keys.
        public static void LogVersion(string host, string version)
        {
            if (string.IsNullOrEmpty(host)) return;
            var line = string.Format(CultureInfo.InvariantCulture,
                "{0:O} host={1} addin_version={2} event=translate_start",
                DateTime.UtcNow, host, string.IsNullOrEmpty(version) ? "unknown" : version);
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

        private static string Format(ImageOverlayDiagnosticEntry e)
        {
            var culture = CultureInfo.InvariantCulture;
            return string.Format(culture,
                "{0:O} host={1} kind={2} png={3}x{4} bbox=[{5},{6},{7},{8}] shape=[{9},{10},{11},{12}]pt frame={13} rot={14} flipH={15} flipV={16} out=[{17},{18},{19},{20}] font={21} textLen={22} verdict={23} reason={24} owner={25}",
                DateTime.UtcNow, e.Host, e.ImageKind, e.PixelWidth, e.PixelHeight,
                F(e.BboxX1), F(e.BboxY1), F(e.BboxX2), F(e.BboxY2),
                F(e.ShapeLeft), F(e.ShapeTop), F(e.ShapeWidth), F(e.ShapeHeight), e.FrameNote,
                F(e.RotationDegrees), e.FlipHorizontal, e.FlipVertical,
                F(e.OutLeft), F(e.OutTop), F(e.OutWidth), F(e.OutHeight),
                F(e.FontSize), e.TextLength, e.Verdict, e.Reason ?? string.Empty, e.OwnerId ?? string.Empty);
        }

        private static string F(float value)
        {
            return value.ToString("0.##", CultureInfo.InvariantCulture);
        }
    }
}
