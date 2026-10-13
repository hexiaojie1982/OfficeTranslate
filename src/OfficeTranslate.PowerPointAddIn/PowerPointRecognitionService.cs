using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeTranslate.PowerPointAddIn
{
    // Separate orchestration preserves the translation path byte-for-byte.
    // Capture primitives, planner, real-layout checks and note splitter are shared.
    internal sealed partial class PowerPointTranslationService
    {
        public bool HasSelectedOcrImage()
        {
            try { return ReadSelectedOcrImages().Count > 0; }
            catch (COMException) { return false; }
        }

        private List<PowerPoint.Shape> ReadSelectedOcrImages()
        {
            var result = new List<PowerPoint.Shape>();
            if (_powerPoint.ActiveWindow == null) return result;
            var selection = _powerPoint.ActiveWindow.Selection;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                foreach (PowerPoint.Shape shape in selection.ShapeRange) AddImageShape(shape, result);
            else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                foreach (PowerPoint.Slide slide in selection.SlideRange)
                    foreach (PowerPoint.Shape shape in slide.Shapes) AddImageShape(shape, result);
            return result;
        }

        public async Task<TranslationTaskSummary> RecognizeSelectionAsync(TranslationSettings settings,
            CancellationToken token, Action<string> progress, OfficeUiDispatcher ui)
        {
            Summary = null;

            var images = ui.Invoke(ReadSelectedOcrImages);
            if (images.Count == 0) throw new InvalidOperationException("请先选中一张图片");
            ImageOverlayDiagnostics.LogVersion("PowerPoint-OCR",
                typeof(PowerPointTranslationService).Assembly.GetName().Version?.ToString() ?? "unknown",
                typeof(PowerPointTranslationService).Module.ModuleVersionId.ToString());
            using (var client = new TranslationClient())
            {
                var summary = Summary = new TranslationTaskSummary(images.Count);
                for (var i = 0; i < images.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    progress($"OfficeTranslate：正在 OCR 识别 {i + 1}/{images.Count}");
                    await RecognizeSelectedImageAsync(images[i], client, settings, summary, token, ui);
                }
                return summary;
            }
        }

        private async Task RecognizeSelectedImageAsync(PowerPoint.Shape image, TranslationClient client, TranslationSettings settings, TranslationTaskSummary summary, CancellationToken token, OfficeUiDispatcher ui)
        {
            var path = Path.Combine(Path.GetTempPath(), "OfficeTranslate-" + Guid.NewGuid().ToString("N") + ".png");
            try
            {
                var preNetwork = await ui.InvokeAsync(() =>
                {
                    ui.LogProbe("capture_before");
                    token.ThrowIfCancellationRequested();
                    var slide = image.Parent as PowerPoint.Slide
                        ?? throw new InvalidOperationException("无法确定图片所在的幻灯片。");
                    byte[] bytes = Array.Empty<byte>();
                    var captureSource = "fallback";
                    var captureReason = "presentation-identity-not-proven";
                    var sourceWidth = 0; var sourceHeight = 0;
                    var original = false;
                    var presentation = slide.Parent as PowerPoint.Presentation;
                    if (presentation != null)
                    {
                        try
                        {
                            original = OriginalImageCapture.TryCaptureFile(presentation.Path.Length == 0 ? "" : presentation.FullName,
                                presentation.Saved == Office.MsoTriState.msoTrue, "PowerPoint", slide.SlideID.ToString(System.Globalization.CultureInfo.InvariantCulture),
                                image.Name, image.Width, image.Height, ImageCaptureBudget.MaxSinglePngBytes,
                                out bytes, out captureReason, out sourceWidth, out sourceHeight);
                        }
                        catch (ImageCaptureLimitException) { throw; }
                        catch (Exception ex) { captureReason = "ppt-package-" + ex.GetType().Name; }
                    }
                    if (original) captureSource = "original";
                    else
                    {
                        try
                        {
                            if (presentation == null) throw new InvalidOperationException("No presentation scale");
                            var scale = ImageCaptureResolution.PowerPointExportScale(image.Width, image.Height,
                                presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight);
                            image.Export(path, PowerPoint.PpShapeFormat.ppShapeFormatPNG, scale.Item1, scale.Item2, PowerPoint.PpExportMode.ppRelativeToSlide);
                            bytes = ImageCaptureBudget.ReadPngFile(path);
                            captureReason += "-scaled-export";
                        }
                        catch (ImageCaptureLimitException) { throw; }
                        catch (Exception ex)
                        {
                            if (File.Exists(path)) File.Delete(path);
                            image.Export(path, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                            bytes = ImageCaptureBudget.ReadPngFile(path);
                            captureReason += "-scaled-export-failed-" + ex.GetType().Name + "-default-export";
                        }
                    }
                    token.ThrowIfCancellationRequested();
                    ImageOverlayDiagnostics.LogOcrCapture("PowerPoint", captureSource, captureReason, sourceWidth, sourceHeight, bytes);
                    ui.LogProbe("capture_after");
                    var owner = ImageOverlayIdentity.ForNamedShape(
                        "pp", SafeString(() => image.Name), bytes);
                    return new
                    {
                        Bytes = bytes,
                        Owner = owner,
                        Slide = slide,
                        Left = image.Left,
                        Top = image.Top,
                        Width = image.Width,
                        Height = image.Height,
                        Rotation = SafeFloat(() => image.Rotation),
                        FlipH = SafeTriState(() => image.HorizontalFlip),
                        FlipV = SafeTriState(() => image.VerticalFlip),
                        CaptureSource = captureSource,
                        CaptureReason = captureReason,
                    };
                });
                var imageBytes = preNetwork.Bytes;
                var ownerId = preNetwork.Owner;
                var hasPixels = PngDimensions.TryRead(imageBytes, out var pixelWidth, out var pixelHeight);
                var recognition = await client.RecognizeImageAsync(imageBytes, settings, token, "PowerPoint", preNetwork.CaptureSource, preNetwork.CaptureReason);
                var regions = recognition.Regions;
                if (regions.Count == 0) { summary.RecordImageRecognition(recognition); return; }
                var slide = preNetwork.Slide;
                var imageLeft = preNetwork.Left; var imageTop = preNetwork.Top;
                var imageWidth = preNetwork.Width; var imageHeight = preNetwork.Height;
                var rotation = preNetwork.Rotation;
                var flipH = preNetwork.FlipH;
                var flipV = preNetwork.FlipV;
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
                            "无法读取导出图像的像素尺寸，坐标无法可靠换算。", 0F, 0F, 0F, 0F, 0F);
                    LogOverlay(region, pixelWidth, pixelHeight, imageLeft, imageTop,
                        imageWidth, imageHeight, rotation, flipH, flipV, plan, overlayText, ownerId);
                    plans.Add(Tuple.Create(region, plan, overlayText));
                }
                token.ThrowIfCancellationRequested();
                await ui.InvokeAsync(() =>
                {
                    token.ThrowIfCancellationRequested();
                    OcrRemovePreviousResults(slide, ownerId);

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
                            () => slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                                imageLeft + plan.Left, imageTop + plan.Top, plan.Width, plan.Height),
                            (PowerPoint.Shape newOverlay) =>
                        {
                            newOverlay.Tags.Add("OfficeTranslateOCR", ImageRecognitionTags.PowerPointOverlayFor(ownerId));
                            newOverlay.AlternativeText = ImageRecognitionTags.OverlayFor(ownerId);
                            newOverlay.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;
                            newOverlay.Fill.Visible = Office.MsoTriState.msoTrue; newOverlay.Fill.ForeColor.RGB = 0xFFFFFF; newOverlay.Fill.Transparency = 0F;
                            newOverlay.Line.Visible = Office.MsoTriState.msoFalse;
                            newOverlay.TextFrame2.MarginLeft = 2; newOverlay.TextFrame2.MarginRight = 2; newOverlay.TextFrame2.MarginTop = 1; newOverlay.TextFrame2.MarginBottom = 1;
                            newOverlay.TextFrame2.WordWrap = Office.MsoTriState.msoTrue;
                            newOverlay.TextFrame2.TextRange.Text = overlayText;
                            newOverlay.TextFrame2.TextRange.Font.Size = plan.FontSize;
                        }, overlay => overlay.Delete(), out var overlay, out var createFailure);
                        if (!created || overlay == null)
                        {
                            noteEntries.Add(createFailure + "\r" + overlayText);
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
                            (chunk, top) => OcrPlaceCombinedNote(slide, ownerId, ImageRecognitionText.NoteText(chunk), imageLeft, top, imageWidth),
                            (chunk, error) => summary.RecordImageNoteFailure(ownerId, ImageRecognitionText.NoteText(chunk), error.Message));
                    }
                });
                summary.RecordImageRecognition(recognition);
            }
            finally { try { if (File.Exists(path)) File.Delete(path); } catch { } }
        }

        private static float OcrPlaceCombinedNote(PowerPoint.Slide slide, string ownerId, string noteText, float noteLeft, float noteTop, float imageWidth)
        {
            var noteWidth = ImageOverlayNotes.NoteWidth(imageWidth);
            var noteHeight = Math.Min(ImageOverlayLayout.MaxNoteHeightPt,
                ImageOverlayLayout.EstimateNoteHeight(noteWidth - 4F, noteText, ImageOverlayNotes.NoteFontSize));
            var note = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                noteLeft, noteTop, noteWidth, noteHeight);
            try
            {
                note.Tags.Add("OfficeTranslateOCR", ImageRecognitionTags.PowerPointNoteFor(ownerId));
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

        private static void OcrRemovePreviousResults(PowerPoint.Slide slide, string ownerId)
        {
            var doomed = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                string tag;
                try { tag = shape.AlternativeText; } catch (COMException) { continue; }
                if (string.IsNullOrEmpty(tag)) continue;
                if (!ImageRecognitionTags.WordExcelMarkerBelongsTo(tag, ownerId)) continue;
                doomed.Add(shape);
            }
            foreach (var shape in doomed)
            {
                try { shape.Delete(); } catch { }
            }
        }
    }
}
