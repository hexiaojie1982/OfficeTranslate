using Excel = Microsoft.Office.Interop.Excel;
using Target = OfficeTranslate.Core.ExcelTextTarget;
using Office = Microsoft.Office.Core;
using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeTranslate.ExcelAddIn
{
    // Separate orchestration preserves the translation path byte-for-byte.
    // Capture primitives, planner, real-layout checks and note splitter are shared.
    internal sealed partial class ExcelTranslationService
    {
        public bool HasSelectedOcrImage()
        {
            try { return ReadSelectedOcrImages().Count > 0; }
            catch (COMException) { return false; }
        }

        private List<Excel.Shape> ReadSelectedOcrImages() => ReadSelectionImages();

        public async Task<TranslationTaskSummary> RecognizeSelectionAsync(TranslationSettings settings,
            CancellationToken token, Action<string> progress, OfficeUiDispatcher ui)
        {
            Summary = null;

            var images = ui.Invoke(ReadSelectedOcrImages);
            if (images.Count == 0) throw new InvalidOperationException("请先选中一张图片");
            ImageOverlayDiagnostics.LogVersion("Excel-OCR",
                typeof(ExcelTranslationService).Assembly.GetName().Version?.ToString() ?? "unknown",
                typeof(ExcelTranslationService).Module.ModuleVersionId.ToString());
            using (var client = new TranslationClient())
            {
                var summary = Summary = new TranslationTaskSummary(images.Count);
                for (var i = 0; i < images.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    progress($"OfficeTranslate：正在 OCR 识别 {i + 1}/{images.Count}");
                    await RecognizeSelectedImageAsync(images[i], i + 1, i + 1, client, settings, summary, token, ui);
                }
                return summary;
            }
        }

        private async Task RecognizeSelectedImageAsync(Excel.Shape image, int imageOrdinal, int taskIndex, TranslationClient client, TranslationSettings settings, TranslationTaskSummary summary, CancellationToken token, OfficeUiDispatcher ui)
        {
            var preNetwork = await ui.InvokeAsync(() =>
            {
                ui.LogProbe("capture_before");
                byte[] bytes;
                var captureSource = "fallback";
                var captureReason = "sheet-identity-not-proven";
                var sourceWidth = 0; var sourceHeight = 0;
                var imageSheet = image.Parent as Excel.Worksheet;
                var capturedSheet = imageSheet ?? _excel.ActiveSheet as Excel.Worksheet
                    ?? throw new InvalidOperationException("无法确定图片所在的工作表。");
                try
                {
                    token.ThrowIfCancellationRequested();
                    bytes = Array.Empty<byte>();
                    var original = false;
                    if (imageSheet != null)
                    {
                        try
                        {
                            var workbook = (Excel.Workbook)imageSheet.Parent;
                            original = OriginalImageCapture.TryCaptureFile(workbook.Path.Length == 0 ? "" : workbook.FullName,
                                workbook.Saved, "Excel", imageSheet.Name, image.Name, image.Width, image.Height,
                                ImageCaptureBudget.MaxSinglePngBytes, out bytes, out captureReason, out sourceWidth, out sourceHeight);
                        }
                        catch (ImageCaptureLimitException) { throw; }
                        catch (Exception ex) { captureReason = "excel-package-" + ex.GetType().Name; }
                    }
                    if (original) captureSource = "original";
                    else bytes = ClipboardImageCapture.CapturePng(_ =>
                    {
                        string copyError;
                        try { image.Copy(); return; }
                        catch (COMException ex) { copyError = DescribeComError("Copy", ex); }
                        try { image.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap); }
                        catch (COMException ex)
                        {
                            throw new InvalidOperationException(
                                "Excel 无法复制图片" + ShapeId(image, imageOrdinal) + "（" + ShapeState(image) + "）：" +
                                "Copy 失败[" + copyError + "]；CopyPicture 失败[" + DescribeComError("CopyPicture", ex) + "]。" +
                                "请选择该图片后重试。", ex);
                        }
                    }, () => Application.DoEvents(), token);
                    token.ThrowIfCancellationRequested();
                    ImageOverlayDiagnostics.LogOcrCapture("Excel", captureSource, captureReason, sourceWidth, sourceHeight, bytes);
                }
                catch (Exception ex) when (!(ex is OperationCanceledException))
                {
                    try
                    {
                        ImageOverlayDiagnostics.LogCaptureFailure("Excel",
                            "image=" + imageOrdinal + " task=" + taskIndex + " shape=" + ShapeId(image, imageOrdinal) +
                            " " + ShapeState(image) +
                            " error=" + ex.GetType().Name +
                            " msg=" + Flatten(ex.Message));
                    }
                    catch { }
                    throw;
                }
                ui.LogProbe("capture_after");
                var owner = ImageOverlayIdentity.ForNamedShape(
                    "xl", SafeString(() => image.Name), bytes);
                return new { Bytes = bytes, Owner = owner, Sheet = capturedSheet, CaptureSource = captureSource, CaptureReason = captureReason };
            });
            var imageBytes = preNetwork.Bytes;
            var ownerId = preNetwork.Owner;

            var hasPixels = PngDimensions.TryRead(imageBytes, out var pixelWidth, out var pixelHeight);
            var recognition = await client.RecognizeImageAsync(imageBytes, settings, token, "Excel", preNetwork.CaptureSource, preNetwork.CaptureReason);
            var regions = recognition.Regions;
            if (regions.Count == 0) { summary.RecordImageRecognition(recognition); return; }
            token.ThrowIfCancellationRequested();
            await ui.InvokeAsync(() =>
            {
                token.ThrowIfCancellationRequested();
                var sheet = preNetwork.Sheet;
                var imageLeft = image.Left; var imageTop = image.Top;
                var imageWidth = image.Width; var imageHeight = image.Height;
                var rotation = SafeFloat(() => image.Rotation);
                var plans = new List<Tuple<ImageTranslationRegion, PlannedOverlay, string>>();
                foreach (var region in regions)
                {
                    var overlayText = region.Source;
                    var plan = region.RequiresReview
                        ? new PlannedOverlay(ImageOverlayVerdict.SideNote, region.ReviewReason, 0F, 0F, 0F, 0F, 0F)
                        : hasPixels
                        ? ImageOverlayPlanner.Plan(
                            region.X1, region.Y1, region.X2, region.Y2,
                            pixelWidth, pixelHeight, imageWidth, imageHeight, overlayText,
                            rotation)
                        : new PlannedOverlay(ImageOverlayVerdict.SideNote,
                            "无法读取捕获图像的像素尺寸，坐标无法可靠换算。", 0F, 0F, 0F, 0F, 0F);
                    LogOverlay(region, pixelWidth, pixelHeight, imageLeft, imageTop, imageWidth, imageHeight,
                        rotation, plan, overlayText, ownerId);
                    plans.Add(Tuple.Create(region, plan, overlayText));
                }
                OcrRemovePreviousResults(sheet, ownerId);

                var noteEntries = new List<string>();
                foreach (var item in plans)
                {
                    var plan = item.Item2;
                    var overlayText = item.Item3;
                    if (plan.Verdict == ImageOverlayVerdict.SideNote)
                    {
                        noteEntries.Add(plan.Reason + "\r" + overlayText);
                        continue;
                    }
                    var created = ImageOverlayWriteback.TryCreate(
                        () => sheet.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                            imageLeft + plan.Left, imageTop + plan.Top, plan.Width, plan.Height),
                        (Excel.Shape newOverlay) =>
                    {
                        newOverlay.AlternativeText = ImageRecognitionTags.OverlayFor(ownerId);
                        newOverlay.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;
                        newOverlay.Fill.Visible = Office.MsoTriState.msoTrue; newOverlay.Fill.ForeColor.RGB = 0xFFFFFF; newOverlay.Fill.Transparency = 0F;
                        newOverlay.Line.Visible = Office.MsoTriState.msoFalse;
                        newOverlay.TextFrame2.MarginLeft = 2; newOverlay.TextFrame2.MarginRight = 2; newOverlay.TextFrame2.MarginTop = 1; newOverlay.TextFrame2.MarginBottom = 1;
                        newOverlay.TextFrame2.WordWrap = Office.MsoTriState.msoTrue;
                        newOverlay.TextFrame2.TextRange.Text = overlayText;
                        newOverlay.TextFrame2.TextRange.Font.Size = plan.FontSize;
                    }, overlay => overlay.Delete(), out var overlay, out var createFailure);
                    if (!created)
                    {
                        noteEntries.Add(createFailure + "\r" + overlayText);
                        continue;
                    }
                    if (overlay == null)
                    {
                        noteEntries.Add("创建原文覆盖框失败，已降级为旁注。\r" + overlayText);
                        continue;
                    }
                    if (!TextFitsBox(overlay, plan.FontSize))
                    {
                        try { overlay.Delete(); } catch { }
                        noteEntries.Add("覆盖框内原文按真实排版在最小字号下仍溢出，已降级为旁注。\r" + overlayText);
                        continue;
                    }
                    if (!MatchesPlan(overlay, imageLeft + plan.Left, imageTop + plan.Top, plan.Width, plan.Height))
                    {
                        try { overlay.Delete(); } catch { }
                        noteEntries.Add("覆盖框几何被宿主改动，已删除并降级为旁注。\r" + overlayText);
                        continue;
                    }
                }
                if (noteEntries.Count > 0)
                {
                    summary.RecordImageNeedsReview();
                    ImageOverlayNotes.Place(noteEntries, ImageOverlayNotes.NoteWidth(imageWidth), imageTop + imageHeight + 6F,
                        (chunk, top) => OcrPlaceCombinedNote(sheet, ownerId, ImageRecognitionText.NoteText(chunk), imageLeft, top, imageWidth),
                        (chunk, error) => summary.RecordImageNoteFailure(ownerId, ImageRecognitionText.NoteText(chunk), error.Message));
                }
            });
            summary.RecordImageRecognition(recognition);
        }

        private float OcrPlaceCombinedNote(Excel.Worksheet sheet, string ownerId, string noteText, float noteLeft, float noteTop, float imageWidth)
        {
            var noteWidth = ImageOverlayNotes.NoteWidth(imageWidth);
            var noteHeight = Math.Min(ImageOverlayLayout.MaxNoteHeightPt,
                ImageOverlayLayout.EstimateNoteHeight(noteWidth - 4F, noteText, ImageOverlayNotes.NoteFontSize));
            var note = sheet.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                noteLeft, noteTop, noteWidth, noteHeight);
            try
            {
                note.AlternativeText = ImageRecognitionTags.NoteFor(ownerId);
                note.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;
                note.Fill.Visible = Office.MsoTriState.msoTrue;
                note.Fill.ForeColor.RGB = 0xE1FFFF; // light yellow (BGR)
                note.Fill.Transparency = 0F;
                note.Line.Visible = Office.MsoTriState.msoFalse;
                note.TextFrame2.MarginLeft = 4; note.TextFrame2.MarginRight = 4;
                note.TextFrame2.MarginTop = 3; note.TextFrame2.MarginBottom = 3;
                note.TextFrame2.WordWrap = Office.MsoTriState.msoTrue;
                note.TextFrame2.TextRange.Text = noteText;
                note.TextFrame2.TextRange.Font.Size = ImageOverlayNotes.NoteFontSize;
                GrowNoteToFit(note);
                return note.Height;
            }
            catch
            {
                try { note.Delete(); } catch { }
                throw;
            }
        }

        private static void OcrRemovePreviousResults(Excel.Worksheet sheet, string ownerId)
        {
            var doomed = new List<Excel.Shape>();
            foreach (Excel.Shape shape in sheet.Shapes)
            {
                string alt;
                try { alt = shape.AlternativeText; } catch { continue; }
                if (!ImageRecognitionTags.WordExcelMarkerBelongsTo(alt, ownerId)) continue;
                doomed.Add(shape);
            }
            foreach (var shape in doomed)
            {
                try { shape.Delete(); } catch { }
            }
        }
    }
}
