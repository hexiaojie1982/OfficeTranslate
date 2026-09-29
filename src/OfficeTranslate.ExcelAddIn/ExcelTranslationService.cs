using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeTranslate.ExcelAddIn
{
    internal sealed class ExcelTranslationService
    {
        private readonly Excel.Application _excel;
        public ExcelTranslationService(Excel.Application excel) => _excel = excel;

        public async Task<TranslationTaskSummary> TranslateAsync(bool wholeSheet, TranslationSettings settings, CancellationToken token, Action<string> progress)
        {
            var sheet = _excel.ActiveSheet as Excel.Worksheet ?? throw new InvalidOperationException("请先打开工作表。");
            var targets = wholeSheet ? ReadSheetTargets(sheet) : ReadSelectionTargets();
            var images = settings.ImageOcrEnabled ? (wholeSheet ? ReadSheetImages(sheet) : ReadSelectionImages()) : new List<Excel.Shape>();
            if (targets.Count == 0 && images.Count == 0) throw new InvalidOperationException("没有找到可翻译的单元格、图形文本或图片。");

            using (var client = new TranslationClient())
            {
                var total = targets.Count + images.Count;
                var summary = new TranslationTaskSummary(total);
                for (var i = 0; i < targets.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    progress($"OfficeTranslate：正在翻译 {i + 1}/{total}");
                    var result = await client.TranslateDetailedAsync(targets[i].Text, settings, token);
                    summary.Record(result);
                    token.ThrowIfCancellationRequested();
                    if (result.Changed)
                    {
                        var translated = result.Text.TrimEnd('\r', '\n');
                        targets[i].Write(settings.BilingualMode ? targets[i].Text + Environment.NewLine + translated : translated, settings.BilingualMode);
                    }
                    progress($"OfficeTranslate：已完成 {i + 1}/{total}");
                }
                for (var i = 0; i < images.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    var current = targets.Count + i + 1;
                    progress($"OfficeTranslate：正在识别图片 {current}/{total}");
                    await TranslateImageAsync(images[i], client, settings, summary, token);
                    progress($"OfficeTranslate：已完成 {current}/{total}");
                }
                return summary;
            }
        }

        private async Task TranslateImageAsync(Excel.Shape image, TranslationClient client, TranslationSettings settings, TranslationTaskSummary summary, CancellationToken token)
        {
            var sheet = _excel.ActiveSheet as Excel.Worksheet ?? throw new InvalidOperationException("无法确定图片所在的工作表。");
            // For a picture shape, Copy is more reliable than CopyPicture across
            // Excel builds and does not create the temporary chart that used to flash.
            // The copy itself runs inside CapturePng (after the user's clipboard is
            // saved), so a failed probe can never destroy clipboard content.
            // CapturePng also clears stale content first so a leftover image is
            // never mistaken for the shape.
            var imageBytes = ClipboardImageCapture.CapturePng(() =>
            {
                try { image.Copy(); }
                catch (COMException)
                {
                    try { image.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap); }
                    catch (COMException ex) { throw new InvalidOperationException("Excel 无法复制所选图片，请重新选择图片后再试。", ex); }
                }
            }, () => Application.DoEvents(), token);

            var hasPixels = PngDimensions.TryRead(imageBytes, out var pixelWidth, out var pixelHeight);
            var regions = await client.TranslateImageAsync(imageBytes, settings, token);
            if (regions.Count == 0) { summary.RecordImage(false); return; }
            summary.RecordImage(true);

            // Excel shapes live in a single sheet-points frame, so the overlay
            // rect is the image rect plus the planner's offsets directly.
            var imageLeft = image.Left; var imageTop = image.Top;
            var imageWidth = image.Width; var imageHeight = image.Height;
            var rotation = SafeFloat(() => image.Rotation);
            // Excel's object model exposes no flip-state property (only the Flip
            // method), so a flipped picture cannot be detected here.
            RemoveOverlaysForImage(sheet, imageLeft, imageTop, imageWidth, imageHeight);

            var needsReview = false;
            foreach (var region in regions)
            {
                var overlayText = settings.BilingualMode && !string.IsNullOrWhiteSpace(region.Source)
                    ? region.Source + "\r" + region.Translation
                    : region.Translation;
                var plan = hasPixels
                    ? ImageOverlayPlanner.Plan(
                        region.X1, region.Y1, region.X2, region.Y2,
                        pixelWidth, pixelHeight, imageWidth, imageHeight, overlayText,
                        rotation, false, false)
                    : new PlannedOverlay(ImageOverlayVerdict.SideNote,
                        "无法读取捕获图像的像素尺寸，坐标无法可靠换算。", 0F, 0F, 0F, 0F, 0F);
                LogOverlay(region, pixelWidth, pixelHeight, imageLeft, imageTop, imageWidth, imageHeight,
                    rotation, plan, overlayText);
                if (plan.Verdict == ImageOverlayVerdict.SideNote)
                {
                    PlaceSideNote(sheet, imageLeft, imageTop, imageWidth, imageHeight, plan.Reason, overlayText);
                    needsReview = true;
                    continue;
                }
                Excel.Shape? overlay = null;
                try
                {
                    // Exact region rect with an opaque cover: in non-bilingual
                    // mode the source text must not show through (the old 8%
                    // transparency did), and the box no longer grows downward.
                    overlay = sheet.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        imageLeft + plan.Left, imageTop + plan.Top, plan.Width, plan.Height);
                    overlay.AlternativeText = "OfficeTranslateOCR";
                    overlay.Fill.Visible = Office.MsoTriState.msoTrue; overlay.Fill.ForeColor.RGB = 0xFFFFFF; overlay.Fill.Transparency = 0F;
                    overlay.Line.Visible = Office.MsoTriState.msoFalse;
                    overlay.TextFrame2.MarginLeft = 2; overlay.TextFrame2.MarginRight = 2; overlay.TextFrame2.MarginTop = 1; overlay.TextFrame2.MarginBottom = 1;
                    overlay.TextFrame2.WordWrap = Office.MsoTriState.msoTrue;
                    overlay.TextFrame2.TextRange.Text = overlayText;
                    overlay.TextFrame2.TextRange.Font.Size = plan.FontSize;
                }
                catch (COMException ex)
                {
                    try { overlay?.Delete(); } catch { }
                    throw new InvalidOperationException("Excel 已完成图片识别，但创建译文覆盖框失败：" + ex.Message, ex);
                }
            }
            if (needsReview) summary.RecordImageNeedsReview();
        }

        private void PlaceSideNote(Excel.Worksheet sheet, float imageLeft, float imageTop, float imageWidth, float imageHeight, string reason, string text)
        {
            // A side-note explicitly does NOT claim positional coverage: it is
            // placed below the image.
            var noteWidth = Math.Min(420F, Math.Max(160F, imageWidth));
            var noteText = "【图片译文待检查】" + reason + "\r" + text;
            var noteHeight = ImageOverlayLayout.EstimateNoteHeight(noteWidth, noteText, 9F);
            var noteLeft = imageLeft;
            var noteTop = imageTop + imageHeight + 6F;
            RemoveNotesInZone(sheet, noteLeft, noteTop, noteWidth, noteHeight);
            var note = sheet.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                noteLeft, noteTop, noteWidth, noteHeight);
            note.AlternativeText = "OfficeTranslateOCR-Note";
            note.Fill.Visible = Office.MsoTriState.msoTrue;
            note.Fill.ForeColor.RGB = 0xE1FFFF; // light yellow (BGR)
            note.Fill.Transparency = 0F;
            note.Line.Visible = Office.MsoTriState.msoFalse;
            note.TextFrame2.MarginLeft = 4; note.TextFrame2.MarginRight = 4;
            note.TextFrame2.MarginTop = 3; note.TextFrame2.MarginBottom = 3;
            note.TextFrame2.WordWrap = Office.MsoTriState.msoTrue;
            note.TextFrame2.TextRange.Text = noteText;
            note.TextFrame2.TextRange.Font.Size = 9F;
        }

        private static void RemoveOverlaysForImage(Excel.Worksheet sheet, float left, float top, float width, float height)
        {
            // Re-translating the same image updates its overlays instead of stacking.
            var doomed = new List<Excel.Shape>();
            foreach (Excel.Shape shape in sheet.Shapes)
            {
                string alt;
                try { alt = shape.AlternativeText; } catch { continue; }
                if (alt != "OfficeTranslateOCR") continue;
                var centerX = shape.Left + shape.Width / 2F;
                var centerY = shape.Top + shape.Height / 2F;
                if (centerX >= left - 8F && centerX <= left + width + 8F &&
                    centerY >= top - 8F && centerY <= top + height + 8F)
                    doomed.Add(shape);
            }
            foreach (var shape in doomed)
            {
                try { shape.Delete(); } catch { }
            }
        }

        private static void RemoveNotesInZone(Excel.Worksheet sheet, float left, float top, float width, float height)
        {
            var doomed = new List<Excel.Shape>();
            foreach (Excel.Shape shape in sheet.Shapes)
            {
                string alt;
                try { alt = shape.AlternativeText; } catch { continue; }
                if (alt != "OfficeTranslateOCR-Note") continue;
                if (shape.Left < left + width && shape.Left + shape.Width > left &&
                    shape.Top < top + height && shape.Top + shape.Height > top)
                    doomed.Add(shape);
            }
            foreach (var shape in doomed)
            {
                try { shape.Delete(); } catch { }
            }
        }

        private static void LogOverlay(ImageTranslationRegion region,
            int pixelWidth, int pixelHeight, float imageLeft, float imageTop,
            float imageWidth, float imageHeight,
            float rotation, PlannedOverlay plan, string text)
        {
            ImageOverlayDiagnostics.Log(new ImageOverlayDiagnosticEntry
            {
                Host = "Excel",
                ImageKind = "Shape",
                PixelWidth = pixelWidth,
                PixelHeight = pixelHeight,
                BboxX1 = region.X1, BboxY1 = region.Y1, BboxX2 = region.X2, BboxY2 = region.Y2,
                ShapeLeft = imageLeft, ShapeTop = imageTop,
                ShapeWidth = imageWidth, ShapeHeight = imageHeight,
                FrameNote = "SheetPoints",
                RotationDegrees = rotation,
                FlipHorizontal = false,
                FlipVertical = false,
                OutLeft = imageLeft + plan.Left,
                OutTop = imageTop + plan.Top,
                OutWidth = plan.Width,
                OutHeight = plan.Height,
                FontSize = plan.FontSize,
                Verdict = plan.Verdict.ToString(),
                Reason = plan.Reason,
                TextLength = text == null ? 0 : text.Length
            });
        }

        private static float SafeFloat(Func<float> read)
        {
            try { return read(); } catch { return 0F; }
        }

        private List<Excel.Shape> ReadSelectionImages()
        {
            var result = new List<Excel.Shape>();
            try
            {
                var shapes = GetSelectedShapeRange(_excel.Selection); if (shapes == null) return result;
                for (var i = 1; i <= shapes.Count; i++) AddImageShape(shapes.Item(i), result);
            }
            catch (COMException) { } catch (TargetInvocationException) { }
            return result;
        }

        private static List<Excel.Shape> ReadSheetImages(Excel.Worksheet sheet)
        {
            var result = new List<Excel.Shape>(); foreach (Excel.Shape shape in sheet.Shapes) AddImageShape(shape, result); return result;
        }

        private static void AddImageShape(Excel.Shape shape, List<Excel.Shape> result)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup) { for (var i = 1; i <= shape.GroupItems.Count; i++) AddImageShape(shape.GroupItems.Item(i), result); return; }
            if ((shape.Type == Office.MsoShapeType.msoPicture || shape.Type == Office.MsoShapeType.msoLinkedPicture) && shape.AlternativeText != "OfficeTranslateOCR") result.Add(shape);
        }

        private List<Target> ReadSelectionTargets()
        {
            var range = _excel.Selection as Excel.Range;
            if (range != null) return ReadRangeTargets(range);

            var result = new List<Target>();
            try
            {
                var shapes = GetSelectedShapeRange(_excel.Selection);
                if (shapes == null) return result;
                for (var i = 1; i <= shapes.Count; i++) AddShape(shapes.Item(i), result);
            }
            catch (COMException) { }
            catch (InvalidCastException) { }
            catch (TargetInvocationException) { }
            return result;
        }

        private static Excel.ShapeRange? GetSelectedShapeRange(object selection)
        {
            if (selection is Excel.DrawingObjects drawingObjects) return drawingObjects.ShapeRange;

            // When editing a grouped flowchart, Excel returns a GroupObject or another
            // COM selection wrapper rather than DrawingObjects. Its ShapeRange property
            // is still available through IDispatch and contains only the selected child shapes.
            var value = selection.GetType().InvokeMember(
                "ShapeRange",
                BindingFlags.GetProperty,
                null,
                selection,
                null);
            return value as Excel.ShapeRange;
        }

        private static List<Target> ReadSheetTargets(Excel.Worksheet sheet)
        {
            var result = ReadRangeTargets(sheet.UsedRange);
            foreach (Excel.Shape shape in sheet.Shapes) AddShape(shape, result);
            return result;
        }

        private static List<Target> ReadRangeTargets(Excel.Range range)
        {
            var result = new List<Target>();
            // A selection can hold several Areas. Read each area's values and
            // formulas in bulk: a few COM calls per area instead of 3+
            // round-trips per cell. Write-back still touches only cells that
            // actually get translated.
            foreach (Excel.Range area in range.Areas)
            {
                var values = RangeGridHelper.Normalize(area.Value2);
                var formulas = RangeGridHelper.Normalize(area.Formula);
                var formulaFlags = RangeGridHelper.Normalize(area.HasFormula);
                var rows = values.GetLength(0);
                var cols = values.GetLength(1);
                for (var r = 0; r < rows; r++)
                    for (var c = 0; c < cols; c++)
                    {
                        if (RangeGridHelper.IsFormulaCell(formulaFlags, formulas, r, c)) continue;
                        if (values[r, c] is string value && !string.IsNullOrWhiteSpace(value))
                        {
                            // Merged areas expose their value only in the top-left
                            // cell; the rest read as empty, so no dedup is needed.
                            // Cells is 1-based relative to the area.
                            var targetCell = (Excel.Range)area.Cells[r + 1, c + 1];
                            result.Add(new Target(value, (text, bilingual) => { targetCell.Value2 = text; if (bilingual) targetCell.WrapText = true; }));
                        }
                    }
            }
            return result;
        }

        private static void AddShape(Excel.Shape shape, List<Target> result)
        {
            if (shape.AlternativeText != null && shape.AlternativeText.StartsWith("OfficeTranslateOCR", StringComparison.Ordinal)) return;
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                for (var i = 1; i <= shape.GroupItems.Count; i++) AddShape(shape.GroupItems.Item(i), result);
                return;
            }
            try
            {
                if (shape.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                {
                    var range = shape.TextFrame2.TextRange;
                    var text = range.Text?.TrimEnd('\r', '\n') ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(text))
                        result.Add(new Target(text, (value, bilingual) => range.Text = value));
                    return;
                }
            }
            catch (COMException) { }
        }

        private sealed class Target
        {
            private readonly Action<string, bool> _write;
            public Target(string text, Action<string, bool> write) { Text = text; _write = write; }
            public string Text { get; }
            public void Write(string value, bool bilingual) => _write(value, bilingual);
        }
    }
}
