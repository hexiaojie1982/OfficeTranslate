using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeTranslate.PowerPointAddIn
{
    internal sealed class PowerPointTranslationService
    {
        private readonly PowerPoint.Application _powerPoint;
        public PowerPointTranslationService(PowerPoint.Application powerPoint) => _powerPoint = powerPoint;

        public async Task<TranslationTaskSummary> TranslateAsync(bool wholePresentation, TranslationSettings settings, CancellationToken token, Action<string> progress)
        {
            if (_powerPoint.Presentations.Count == 0) throw new InvalidOperationException("请先打开演示文稿。");
            var targets = wholePresentation ? ReadPresentation() : ReadSelection();
            var images = settings.ImageOcrEnabled ? (wholePresentation ? ReadPresentationImages() : ReadSelectionImages()) : new List<PowerPoint.Shape>();
            if (targets.Count == 0 && images.Count == 0) throw new InvalidOperationException(wholePresentation ? "演示文稿中没有可翻译的文本或图片。" : "请先选择文本框、文字或图片。");
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
                        targets[i].Write(settings.BilingualMode ? targets[i].Text + "\r" + translated : translated);
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

        private async Task TranslateImageAsync(PowerPoint.Shape image, TranslationClient client, TranslationSettings settings, TranslationTaskSummary summary, CancellationToken token)
        {
            var path = Path.Combine(Path.GetTempPath(), "OfficeTranslate-" + Guid.NewGuid().ToString("N") + ".png");
            var recognized = false;
            var needsReview = false;
            try
            {
                image.Export(path, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                var imageBytes = File.ReadAllBytes(path);
                var hasPixels = PngDimensions.TryRead(imageBytes, out var pixelWidth, out var pixelHeight);
                var regions = await client.TranslateImageAsync(imageBytes, settings, token);
                recognized = regions.Count > 0;
                if (!recognized) return;
                var slide = image.Parent as PowerPoint.Slide ?? throw new InvalidOperationException("无法确定图片所在的幻灯片。");
                // Slide shapes live in a single slide-points frame.
                var imageLeft = image.Left; var imageTop = image.Top;
                var imageWidth = image.Width; var imageHeight = image.Height;
                var rotation = SafeFloat(() => image.Rotation);
                var flipH = SafeTriState(() => image.HorizontalFlip);
                var flipV = SafeTriState(() => image.VerticalFlip);
                // Re-translating the same image updates its overlays instead of stacking.
                RemoveOverlaysForImage(slide, imageLeft, imageTop, imageWidth, imageHeight);
                foreach (var region in regions)
                {
                    var overlayText = settings.BilingualMode && !string.IsNullOrWhiteSpace(region.Source)
                        ? region.Source + "\r" + region.Translation
                        : region.Translation;
                    var plan = hasPixels
                        ? ImageOverlayPlanner.Plan(
                            region.X1, region.Y1, region.X2, region.Y2,
                            pixelWidth, pixelHeight, imageWidth, imageHeight, overlayText,
                            rotation, flipH, flipV)
                        : new PlannedOverlay(ImageOverlayVerdict.SideNote,
                            "无法读取导出图像的像素尺寸，坐标无法可靠换算。", 0F, 0F, 0F, 0F, 0F);
                    LogOverlay(region, pixelWidth, pixelHeight, imageLeft, imageTop,
                        imageWidth, imageHeight, rotation, flipH, flipV, plan, overlayText);
                    if (plan.Verdict == ImageOverlayVerdict.SideNote)
                    {
                        PlaceSideNote(slide, imageLeft, imageTop, imageWidth, imageHeight, plan.Reason, overlayText);
                        needsReview = true;
                        continue;
                    }
                    // Exact region rect with an opaque cover: in non-bilingual
                    // mode the source text must not show through (the old 8%
                    // transparency did), and the box no longer grows downward.
                    var overlay = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        imageLeft + plan.Left, imageTop + plan.Top, plan.Width, plan.Height);
                    overlay.Tags.Add("OfficeTranslateOCR", "1");
                    overlay.Fill.Visible = Office.MsoTriState.msoTrue; overlay.Fill.ForeColor.RGB = 0xFFFFFF; overlay.Fill.Transparency = 0F;
                    overlay.Line.Visible = Office.MsoTriState.msoFalse;
                    overlay.TextFrame2.MarginLeft = 2; overlay.TextFrame2.MarginRight = 2; overlay.TextFrame2.MarginTop = 1; overlay.TextFrame2.MarginBottom = 1;
                    overlay.TextFrame2.WordWrap = Office.MsoTriState.msoTrue;
                    overlay.TextFrame2.TextRange.Text = overlayText;
                    overlay.TextFrame2.TextRange.Font.Size = plan.FontSize;
                }
            }
            finally { try { if (File.Exists(path)) File.Delete(path); } catch { } }
            summary.RecordImage(recognized);
            if (needsReview) summary.RecordImageNeedsReview();
        }

        private static void PlaceSideNote(PowerPoint.Slide slide, float imageLeft, float imageTop, float imageWidth, float imageHeight, string reason, string text)
        {
            // A side-note explicitly does NOT claim positional coverage: it is
            // placed below the image.
            var noteWidth = Math.Min(420F, Math.Max(160F, imageWidth));
            var noteText = "【图片译文待检查】" + reason + "\r" + text;
            var noteHeight = ImageOverlayLayout.EstimateNoteHeight(noteWidth, noteText, 9F);
            var noteLeft = imageLeft;
            var noteTop = imageTop + imageHeight + 6F;
            RemoveNotesInZone(slide, noteLeft, noteTop, noteWidth, noteHeight);
            var note = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                noteLeft, noteTop, noteWidth, noteHeight);
            note.Tags.Add("OfficeTranslateOCR", "Note");
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

        private static void RemoveOverlaysForImage(PowerPoint.Slide slide, float left, float top, float width, float height)
        {
            var doomed = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (!HasOcrTag(shape, "1")) continue;
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

        private static void RemoveNotesInZone(PowerPoint.Slide slide, float left, float top, float width, float height)
        {
            var doomed = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (!HasOcrTag(shape, "Note")) continue;
                if (shape.Left < left + width && shape.Left + shape.Width > left &&
                    shape.Top < top + height && shape.Top + shape.Height > top)
                    doomed.Add(shape);
            }
            foreach (var shape in doomed)
            {
                try { shape.Delete(); } catch { }
            }
        }

        private static bool HasOcrTag(PowerPoint.Shape shape, string value)
        {
            try { return shape.Tags["OfficeTranslateOCR"] == value; }
            catch (COMException) { return false; }
        }

        private static void LogOverlay(ImageTranslationRegion region,
            int pixelWidth, int pixelHeight, float imageLeft, float imageTop,
            float imageWidth, float imageHeight,
            float rotation, bool flipH, bool flipV,
            PlannedOverlay plan, string text)
        {
            ImageOverlayDiagnostics.Log(new ImageOverlayDiagnosticEntry
            {
                Host = "PowerPoint",
                ImageKind = "Shape",
                PixelWidth = pixelWidth,
                PixelHeight = pixelHeight,
                BboxX1 = region.X1, BboxY1 = region.Y1, BboxX2 = region.X2, BboxY2 = region.Y2,
                ShapeLeft = imageLeft, ShapeTop = imageTop,
                ShapeWidth = imageWidth, ShapeHeight = imageHeight,
                FrameNote = "SlidePoints",
                RotationDegrees = rotation,
                FlipHorizontal = flipH,
                FlipVertical = flipV,
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

        private static bool SafeTriState(Func<Office.MsoTriState> read)
        {
            try { return read() == Office.MsoTriState.msoTrue; }
            catch { return false; }
        }

        private List<PowerPoint.Shape> ReadPresentationImages()
        {
            var result = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Slide slide in _powerPoint.ActivePresentation.Slides)
                foreach (PowerPoint.Shape shape in slide.Shapes) AddImageShape(shape, result);
            return result;
        }

        private List<PowerPoint.Shape> ReadSelectionImages()
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

        private static void AddImageShape(PowerPoint.Shape shape, List<PowerPoint.Shape> result)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                dynamic groupItems = shape.GroupItems;
                for (var i = 1; i <= shape.GroupItems.Count; i++) AddImageShape((PowerPoint.Shape)groupItems.Item(i), result);
                return;
            }
            if (shape.Type == Office.MsoShapeType.msoPicture || shape.Type == Office.MsoShapeType.msoLinkedPicture) result.Add(shape);
        }

        private List<Target> ReadPresentation()
        {
            var result = new List<Target>();
            foreach (PowerPoint.Slide slide in _powerPoint.ActivePresentation.Slides)
                foreach (PowerPoint.Shape shape in slide.Shapes) AddShape(shape, result);
            return result;
        }

        private List<Target> ReadSelection()
        {
            var result = new List<Target>();
            if (_powerPoint.ActiveWindow == null) return result;
            var selection = _powerPoint.ActiveWindow.Selection;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                try
                {
                    if (selection.ShapeRange.Count > 0 && selection.ShapeRange[1].HasTable == Office.MsoTriState.msoTrue)
                    {
                        var selectedCells = AddSelectedTableCells(selection.ShapeRange[1], result);
                        if (selectedCells > 1) return result;
                        result.Clear();
                    }
                }
                catch (COMException) { }
                AddRange(selection.TextRange, result); return result;
            }
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTable == Office.MsoTriState.msoTrue && AddSelectedTableCells(shape, result) > 0) continue;
                    AddShape(shape, result);
                }
            }
            else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
            {
                foreach (PowerPoint.Slide slide in selection.SlideRange)
                    foreach (PowerPoint.Shape shape in slide.Shapes) AddShape(shape, result);
            }
            return result;
        }

        private static int AddSelectedTableCells(PowerPoint.Shape shape, List<Target> result)
        {
            if (shape.HasTable != Office.MsoTriState.msoTrue) return 0;
            var selected = 0;
            var seenCells = new HashSet<IntPtr>();
            for (var row = 1; row <= shape.Table.Rows.Count; row++)
                for (var column = 1; column <= shape.Table.Columns.Count; column++)
                {
                    try
                    {
                        var cell = shape.Table.Cell(row, column);
                        if (!cell.Selected) continue;
                        var cellShape = cell.Shape;
                        var identity = Marshal.GetIUnknownForObject(cellShape);
                        try { if (!seenCells.Add(identity)) continue; }
                        finally { Marshal.Release(identity); }
                        selected++;
                        if (cellShape.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                            AddRange(cellShape.TextFrame2.TextRange, result);
                    }
                    catch (COMException) { }
                }
            return selected;
        }

        private static void AddShape(PowerPoint.Shape shape, List<Target> result)
        {
            try { var tag = shape.Tags["OfficeTranslateOCR"]; if (!string.IsNullOrEmpty(tag)) return; } catch (COMException) { }
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                dynamic groupItems = shape.GroupItems;
                for (var i = 1; i <= shape.GroupItems.Count; i++) AddShape((PowerPoint.Shape)groupItems.Item(i), result);
                return;
            }
            if (shape.HasTable == Office.MsoTriState.msoTrue)
            {
                var seenCells = new HashSet<IntPtr>();
                for (var row = 1; row <= shape.Table.Rows.Count; row++)
                    for (var column = 1; column <= shape.Table.Columns.Count; column++)
                    {
                        try
                        {
                            var cellShape = shape.Table.Cell(row, column).Shape;
                            var identity = Marshal.GetIUnknownForObject(cellShape);
                            try { if (!seenCells.Add(identity)) continue; }
                            finally { Marshal.Release(identity); }
                            if (cellShape.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                                AddRange(cellShape.TextFrame2.TextRange, result);
                        }
                        catch (COMException) { }
                    }
                return;
            }
            if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                AddRange(shape.TextFrame.TextRange, result);
        }

        private static void AddRange(PowerPoint.TextRange range, List<Target> result)
        {
            var text = range.Text?.TrimEnd('\r', '\n') ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(text)) result.Add(new Target(text, value => range.Text = value));
        }

        private static void AddRange(Office.TextRange2 range, List<Target> result)
        {
            var text = range.Text?.TrimEnd('\r', '\n') ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(text)) result.Add(new Target(text, value => range.Text = value));
        }

        private sealed class Target
        {
            private readonly Action<string> _write;
            public Target(string text, Action<string> write) { Text = text; _write = write; }
            public string Text { get; }
            public void Write(string value) => _write(value);
        }
    }
}
