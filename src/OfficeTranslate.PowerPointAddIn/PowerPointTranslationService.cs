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
    internal sealed class PowerPointTranslationService
    {
        private readonly PowerPoint.Application _powerPoint;
        public PowerPointTranslationService(PowerPoint.Application powerPoint) => _powerPoint = powerPoint;

        public async Task<TranslationTaskSummary> TranslateAsync(bool wholePresentation, TranslationSettings settings, CancellationToken token, Action<string> progress, OfficeUiDispatcher ui)
        {
            // Menu acceptance must prove which DLL served the request: record
            // this add-in's assembly version so the reviewer can verify it in
            // the diagnostics log instead of inferring from registry keys.
            // The module MVID is unique per compilation and distinguishes
            // candidate builds that share the same assembly version.
            ImageOverlayDiagnostics.LogVersion("PowerPoint",
                typeof(PowerPointTranslationService).Assembly.GetName().Version?.ToString() ?? "unknown",
                typeof(PowerPointTranslationService).Module.ModuleVersionId.ToString());
            // M2: prove which thread the synchronous prefix runs on; every
            // COM/clipboard section below is dispatched explicitly.
            ui.LogProbe("translate_start");
            if (_powerPoint.Presentations.Count == 0) throw new InvalidOperationException("请先打开演示文稿。");
            var targets = wholePresentation ? ReadPresentation() : ReadSelection();
            var images = settings.ImageOcrEnabled ? (wholePresentation ? ReadPresentationImages() : ReadSelectionImages()) : new List<PowerPoint.Shape>();
            if (targets.Count == 0 && images.Count == 0) throw new InvalidOperationException(wholePresentation ? "演示文稿中没有可翻译的文本或图片。" : "请先选择文本框、文字或图片。");
            using (var client = new TranslationClient())
            {
                var total = targets.Count + images.Count;
                var summary = new TranslationTaskSummary(total);
                ui.LogProbe("text_loop_before");
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
                        // M2: the writeback touches the PowerPoint object
                        // model, so it runs on the Office UI (STA) thread
                        // explicitly instead of on whatever thread the network
                        // await resumed on.
                        var target = targets[i];
                        await ui.InvokeAsync(() => target.Write(settings.BilingualMode ? target.Text + "\r" + translated : translated));
                    }
                    progress($"OfficeTranslate：已完成 {i + 1}/{total}");
                }
                ui.LogProbe("text_loop_after");
                ui.LogProbe("image_loop_before");
                for (var i = 0; i < images.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    var current = targets.Count + i + 1;
                    progress($"OfficeTranslate：正在识别图片 {current}/{total}");
                    await TranslateImageAsync(images[i], client, settings, summary, token, ui);
                    progress($"OfficeTranslate：已完成 {current}/{total}");
                }
                ui.LogProbe("image_loop_after");
                return summary;
            }
        }

        private async Task TranslateImageAsync(PowerPoint.Shape image, TranslationClient client, TranslationSettings settings, TranslationTaskSummary summary, CancellationToken token, OfficeUiDispatcher ui)
        {
            var path = Path.Combine(Path.GetTempPath(), "OfficeTranslate-" + Guid.NewGuid().ToString("N") + ".png");
            try
            {
                // M2: the export and every shape read below run on the Office
                // UI (STA) thread explicitly.
                var preNetwork = await ui.InvokeAsync(() =>
                {
                    ui.LogProbe("capture_before");
                    image.Export(path, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                    ui.LogProbe("capture_after");
                    var bytes = File.ReadAllBytes(path);
                    // R5: owner identity is the slide-unique shape Name (persisted,
                    // stable across move/resize/reopen), falling back to a content
                    // hash when the name is unavailable.
                    var owner = ImageOverlayIdentity.ForNamedShape(
                        "pp", SafeString(() => image.Name), bytes);
                    var slide = image.Parent as PowerPoint.Slide
                        ?? throw new InvalidOperationException("无法确定图片所在的幻灯片。");
                    // Slide shapes live in a single slide-points frame.
                    // R3: flip state is recorded for diagnostics only. The exported
                    // PNG already shows the flipped rendering, so the model's bbox
                    // is in final display coordinates and must NOT be mirrored
                    // again by the planner.
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
                    };
                });
                var imageBytes = preNetwork.Bytes;
                var ownerId = preNetwork.Owner;
                var hasPixels = PngDimensions.TryRead(imageBytes, out var pixelWidth, out var pixelHeight);
                var regions = await client.TranslateImageAsync(imageBytes, settings, token);
                // R8: an empty OCR result is still an image outcome and must be
                // counted, exactly like the Word/Excel paths do.
                if (regions.Count == 0) { summary.RecordImage(false); return; }
                summary.RecordImage(true);
                var slide = preNetwork.Slide;
                var imageLeft = preNetwork.Left; var imageTop = preNetwork.Top;
                var imageWidth = preNetwork.Width; var imageHeight = preNetwork.Height;
                var rotation = preNetwork.Rotation;
                var flipH = preNetwork.FlipH;
                var flipV = preNetwork.FlipV;

                // Plan every region first; old results are replaced only after
                // planning succeeds, so a failure/cancel keeps the previous content.
                var plans = new List<Tuple<ImageTranslationRegion, PlannedOverlay, string>>();
                foreach (var region in regions)
                {
                    var overlayText = settings.BilingualMode && !string.IsNullOrWhiteSpace(region.Source)
                        ? region.Source + "\r" + region.Translation
                        : region.Translation;
                    var plan = hasPixels
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

                // R5: replace this image's previous results (overlays and notes).
                // Matching is by owner id only, never by region center, so an
                // overlapping image's cleanup cannot delete this image's boxes.
                // M2: all shape surgery below runs on the Office UI (STA) thread.
                // F4: one last cancellation check before the surgery phase, so
                // a cancelled run never reaches the shape mutations below.
                token.ThrowIfCancellationRequested();
                await ui.InvokeAsync(() =>
                {
                    RemovePreviousResults(slide, ownerId);

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
                        // Exact region rect with an opaque cover: in non-bilingual
                        // mode the source text must not show through (the old 8%
                        // transparency did), and the box no longer grows downward.
                        var overlay = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                            imageLeft + plan.Left, imageTop + plan.Top, plan.Width, plan.Height);
                        overlay.Tags.Add("OfficeTranslateOCR", ownerId);
                        // D2: the host must never resize the cover by itself. New
                        // textboxes default to auto-size-to-fit-text on some
                        // builds, which would let the box grow beyond the validated
                        // region and make the fit check measure the grown box. The
                        // box keeps the planner's fixed geometry; only the font
                        // size adapts.
                        overlay.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;
                        overlay.Fill.Visible = Office.MsoTriState.msoTrue; overlay.Fill.ForeColor.RGB = 0xFFFFFF; overlay.Fill.Transparency = 0F;
                        overlay.Line.Visible = Office.MsoTriState.msoFalse;
                        overlay.TextFrame2.MarginLeft = 2; overlay.TextFrame2.MarginRight = 2; overlay.TextFrame2.MarginTop = 1; overlay.TextFrame2.MarginBottom = 1;
                        overlay.TextFrame2.WordWrap = Office.MsoTriState.msoTrue;
                        overlay.TextFrame2.TextRange.Text = overlayText;
                        overlay.TextFrame2.TextRange.Font.Size = plan.FontSize;
                        // W1: shrink the font (never the box) until the real
                        // laid-out text height fits the fixed box; a box that still
                        // overflows at the floor is deleted and the region degrades
                        // to a side-note instead of showing clipped text.
                        if (!TextFitsBox(overlay, plan.FontSize))
                        {
                            try { overlay.Delete(); } catch { }
                            noteEntries.Add("覆盖框内译文按真实排版在最小字号下仍溢出，已降级为旁注。\r" + overlayText);
                            continue;
                        }
                        // D2: re-read the geometry and confirm the host did not
                        // move or resize the cover behind our back; it must still
                        // match the planner's fixed rect.
                        if (!MatchesPlan(overlay, imageLeft + plan.Left, imageTop + plan.Top, plan.Width, plan.Height))
                        {
                            try { overlay.Delete(); } catch { }
                            noteEntries.Add("覆盖框几何被宿主改动，已删除并降级为旁注。\r" + overlayText);
                            continue;
                        }
                    }
                    // R2: one image gets ONE combined side-note, so no region's
                    // translation is lost when several regions degrade.
                    if (noteEntries.Count > 0)
                    {
                        PlaceCombinedNote(slide, ownerId, ImageOverlayNotes.Combine(noteEntries),
                            imageLeft, imageTop + imageHeight + 6F, imageWidth);
                        summary.RecordImageNeedsReview();
                    }
                });
            }
            finally { try { if (File.Exists(path)) File.Delete(path); } catch { } }
        }

        private static void PlaceCombinedNote(PowerPoint.Slide slide, string ownerId, string noteText, float noteLeft, float noteTop, float imageWidth)
        {
            // A side-note explicitly does NOT claim positional coverage: it is
            // placed below the image.
            var noteWidth = Math.Min(420F, Math.Max(160F, imageWidth));
            // D3: clamp the initial height to the cap up front; an estimate
            // that already exceeds the cap must go through the bounded growth
            // check instead of returning success immediately.
            var noteHeight = Math.Min(ImageOverlayLayout.MaxNoteHeightPt,
                ImageOverlayLayout.EstimateNoteHeight(noteWidth, noteText, 9F));
            var note = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                noteLeft, noteTop, noteWidth, noteHeight);
            note.Tags.Add("OfficeTranslateOCR", "Note:" + ownerId);
            // D2/D3: no host auto-growth for notes either; W2 grows the note
            // manually under the cap so the limit cannot be bypassed.
            note.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;
            note.Fill.Visible = Office.MsoTriState.msoTrue;
            note.Fill.ForeColor.RGB = 0xE1FFFF; // light yellow (BGR)
            note.Fill.Transparency = 0F;
            note.Line.Visible = Office.MsoTriState.msoFalse;
            note.TextFrame2.MarginLeft = 4; note.TextFrame2.MarginRight = 4;
            note.TextFrame2.MarginTop = 3; note.TextFrame2.MarginBottom = 3;
            note.TextFrame2.WordWrap = Office.MsoTriState.msoTrue;
            note.TextFrame2.TextRange.Text = noteText;
            note.TextFrame2.TextRange.Font.Size = 9F;
            // W2: grow the note downward until the real laid-out text height
            // fits. Bounded; a note that still overflows at the cap is deleted
            // and reported explicitly instead of being kept clipped.
            GrowNoteToFit(note);
        }

        // W1: real-layout fit check for PowerPoint overlays.
        // TextRange2.BoundHeight is the actual rendered text height with word
        // wrap applied. The font shrinks stepwise to the 8pt floor; the box
        // geometry never expands.
        private static bool TextFitsBox(PowerPoint.Shape box, float startSize)
        {
            const float floorSize = 8F;
            var size = Math.Min(startSize, 18F);
            try { box.TextFrame2.TextRange.Font.Size = size; }
            catch { return false; }
            for (var i = 0; i < 12; i++)
            {
                float boundHeight, availHeight;
                try
                {
                    boundHeight = box.TextFrame2.TextRange.BoundHeight;
                    availHeight = box.Height - box.TextFrame2.MarginTop - box.TextFrame2.MarginBottom;
                }
                catch { return false; }
                if (boundHeight <= availHeight + 0.5F) return true;
                if (size <= floorSize) return false;
                size = Math.Max(floorSize, size - 1F);
                try { box.TextFrame2.TextRange.Font.Size = size; }
                catch { return false; }
            }
            return false;
        }

        // W2: grows the side-note downward until its real laid-out height fits.
        // D3: the cap is enforced BEFORE the fit check on every pass, so an
        // over-cap box can never return success, however it got that tall
        // (estimator overshoot or host auto-growth). Text that cannot fit
        // inside the cap is deleted and reported explicitly instead of being
        // kept silently clipped.
        private static void GrowNoteToFit(PowerPoint.Shape note)
        {
            for (var i = 0; i < 12; i++)
            {
                float height;
                try { height = note.Height; }
                catch (COMException ex)
                {
                    try { note.Delete(); } catch { }
                    throw new InvalidOperationException("已完成图片识别，但无法读取译文旁注的尺寸，已删除旁注避免显示不全。", ex);
                }
                if (height > ImageOverlayLayout.MaxNoteHeightPt)
                {
                    try { note.Height = ImageOverlayLayout.MaxNoteHeightPt; }
                    catch (COMException ex)
                    {
                        try { note.Delete(); } catch { }
                        throw new InvalidOperationException("已完成图片识别，但无法约束译文旁注的高度，已删除避免显示不全。", ex);
                    }
                    height = ImageOverlayLayout.MaxNoteHeightPt;
                }
                float boundHeight, availHeight;
                try
                {
                    boundHeight = note.TextFrame2.TextRange.BoundHeight;
                    availHeight = note.Height - note.TextFrame2.MarginTop - note.TextFrame2.MarginBottom;
                }
                catch (COMException ex)
                {
                    try { note.Delete(); } catch { }
                    throw new InvalidOperationException("已完成图片识别，但无法校验译文旁注的排版溢出，已删除旁注避免显示不全。", ex);
                }
                if (boundHeight <= availHeight + 0.5F) return;
                if (height >= ImageOverlayLayout.MaxNoteHeightPt)
                {
                    string text;
                    try { text = note.TextFrame2.TextRange.Text; } catch { text = string.Empty; }
                    try { note.Delete(); } catch { }
                    throw new InvalidOperationException(
                        "译文旁注过长，增高到上限仍无法完整显示，已删除避免静默裁切。译文（截断）：" +
                        ImageOverlayText.Truncate(text));
                }
                try { note.Height = Math.Min(ImageOverlayLayout.MaxNoteHeightPt, height * 1.5F + 12F); }
                catch (COMException ex)
                {
                    try { note.Delete(); } catch { }
                    throw new InvalidOperationException("已完成图片识别，但无法增高译文旁注，已删除避免显示不全。", ex);
                }
            }
            string leftover;
            try { leftover = note.TextFrame2.TextRange.Text; } catch { leftover = string.Empty; }
            try { note.Delete(); } catch { }
            throw new InvalidOperationException(
                "译文旁注排版校验未收敛，已删除避免静默裁切。译文（截断）：" +
                ImageOverlayText.Truncate(leftover));
        }

        // D2: confirms a box still matches the planner's fixed rect. Guards
        // against host auto-size behaviors that would silently expand the
        // cover beyond the validated region.
        private static bool MatchesPlan(PowerPoint.Shape box, float left, float top, float width, float height)
        {
            const float epsilon = 0.5F;
            try
            {
                return Math.Abs(box.Left - left) <= epsilon
                    && Math.Abs(box.Top - top) <= epsilon
                    && Math.Abs(box.Width - width) <= epsilon
                    && Math.Abs(box.Height - height) <= epsilon;
            }
            catch { return false; }
        }

        private static void RemovePreviousResults(PowerPoint.Slide slide, string ownerId)
        {
            // Deletes only shapes that provably belong to this image (by owner
            // id). Legacy id-less tags are retained by default (C4): they
            // cannot be attributed to an image, and deleting them here would
            // destroy other images' translations. Never touches other images'
            // results, even when their rects overlap.
            var doomed = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                string tag;
                try { tag = shape.Tags["OfficeTranslateOCR"]; } catch (COMException) { continue; }
                if (string.IsNullOrEmpty(tag)) continue;
                if (!ImageOverlayTags.PowerPointTagBelongsTo(tag, ownerId)) continue;
                doomed.Add(shape);
            }
            foreach (var shape in doomed)
            {
                try { shape.Delete(); } catch { }
            }
        }

        private static string SafeString(Func<string> read)
        {
            try { return read() ?? string.Empty; } catch { return string.Empty; }
        }

        private static void LogOverlay(ImageTranslationRegion region,
            int pixelWidth, int pixelHeight, float imageLeft, float imageTop,
            float imageWidth, float imageHeight,
            float rotation, bool flipH, bool flipV,
            PlannedOverlay plan, string text, string ownerId)
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
                TextLength = text == null ? 0 : text.Length,
                OwnerId = ownerId
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
