using Excel = Microsoft.Office.Interop.Excel;
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
    internal sealed class ExcelTranslationService
    {
        private readonly Excel.Application _excel;
        public ExcelTranslationService(Excel.Application excel) => _excel = excel;

        public async Task<TranslationTaskSummary> TranslateAsync(bool wholeSheet, TranslationSettings settings, CancellationToken token, Action<string> progress, OfficeUiDispatcher ui)
        {
            // Menu acceptance must prove which DLL served the request: record
            // this add-in's assembly version so the reviewer can verify it in
            // the diagnostics log instead of inferring from registry keys.
            // The module MVID is unique per compilation and distinguishes
            // candidate builds that share the same assembly version.
            ImageOverlayDiagnostics.LogVersion("Excel",
                typeof(ExcelTranslationService).Assembly.GetName().Version?.ToString() ?? "unknown",
                typeof(ExcelTranslationService).Module.ModuleVersionId.ToString());
            // M2: prove which thread the synchronous prefix runs on; every
            // COM/clipboard section below is dispatched explicitly.
            ui.LogProbe("translate_start");
            var sheet = _excel.ActiveSheet as Excel.Worksheet ?? throw new InvalidOperationException("请先打开工作表。");
            var targets = wholeSheet ? ReadSheetTargets(sheet) : ReadSelectionTargets();
            var images = settings.ImageOcrEnabled ? (wholeSheet ? ReadSheetImages(sheet) : ReadSelectionImages()) : new List<Excel.Shape>();
            if (targets.Count == 0 && images.Count == 0) throw new InvalidOperationException("没有找到可翻译的单元格、图形文本或图片。");

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
                        // M2: the writeback touches the Excel object model, so
                        // it runs on the Office UI (STA) thread explicitly
                        // instead of on whatever thread the network await
                        // resumed on.
                        await ui.InvokeAsync(() =>
                        {
                            // N3: re-check inside the queued UI callback. The
                            // token may have been cancelled after the
                            // pre-queue check but before this callback ran.
                            token.ThrowIfCancellationRequested();
                            targets[i].Write(settings.BilingualMode ? targets[i].Text + Environment.NewLine + translated : translated, settings.BilingualMode);
                        });
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
                    await TranslateImageAsync(images[i], i + 1, current, client, settings, summary, token, ui);
                    progress($"OfficeTranslate：已完成 {current}/{total}");
                }
                ui.LogProbe("image_loop_after");
                return summary;
            }
        }

        private async Task TranslateImageAsync(Excel.Shape image, int imageOrdinal, int taskIndex, TranslationClient client, TranslationSettings settings, TranslationTaskSummary summary, CancellationToken token, OfficeUiDispatcher ui)
        {
            // M2: clipboard capture runs on the Office UI (STA) thread
            // explicitly -- later menu runs failed here with "Current thread
            // must be set to single thread apartment (STA) mode before OLE
            // calls can be made". The copy itself runs inside CapturePng
            // (after the user's clipboard is saved), so a failed probe can
            // never destroy clipboard content. CapturePng also clears stale
            // content first so a leftover image is never mistaken for the
            // shape.
            var preNetwork = await ui.InvokeAsync(() =>
            {
                ui.LogProbe("capture_before");
                // For a picture shape, Copy is more reliable than CopyPicture across
                // Excel builds and does not create the temporary chart that used to flash.
                // Excel failure log (not only the popup): image number, shape
                // identity, activation state, and the original COM errors, so
                // the next review can tell a stale object from an activation
                // or clipboard issue.
                byte[] bytes;
                try
                {
                    bytes = ClipboardImageCapture.CapturePng(_ =>
                    {
                        // Excel copy diagnostics: keep the ORIGINAL COM HResult/type of
                        // both the Copy attempt and the CopyPicture fallback, plus the
                        // target identity and activation state. A previous wrapper
                        // surfaced only the managed 0x80131509 and hid the real
                        // failure, which made a first-round copy failure unlocatable.
                        // Only COMException triggers the CopyPicture fallback; other
                        // exceptions propagate unchanged (no blanket retry).
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
                // R5: owner identity is the sheet-unique shape Name (persisted,
                // stable across move/resize/reopen), falling back to a content
                // hash when the name is unavailable.
                var owner = ImageOverlayIdentity.ForNamedShape(
                    "xl", SafeString(() => image.Name), bytes);
                return Tuple.Create(bytes, owner);
            });
            var imageBytes = preNetwork.Item1;
            var ownerId = preNetwork.Item2;

            var hasPixels = PngDimensions.TryRead(imageBytes, out var pixelWidth, out var pixelHeight);
            var regions = await client.TranslateImageAsync(imageBytes, settings, token);
            if (regions.Count == 0) { summary.RecordImage(false); return; }
            summary.RecordImage(true);

            // M2: every shape access below runs on the Office UI (STA) thread.
            // The sheet is resolved here (not earlier) so the "no sheet"
            // error is raised on the same thread that performs the surgery.
            // F4: one last cancellation check before the surgery phase, so
            // a cancelled run never reaches the shape mutations below.
            token.ThrowIfCancellationRequested();
            await ui.InvokeAsync(() =>
            {
                // N3: re-check inside the queued UI callback -- see above.
                // Once this callback starts, its shape mutations run to
                // completion for this image.
                token.ThrowIfCancellationRequested();
                var sheet = _excel.ActiveSheet as Excel.Worksheet
                    ?? throw new InvalidOperationException("无法确定图片所在的工作表。");
                // Excel shapes live in a single sheet-points frame, so the overlay
                // rect is the image rect plus the planner's offsets directly.
                var imageLeft = image.Left; var imageTop = image.Top;
                var imageWidth = image.Width; var imageHeight = image.Height;
                var rotation = SafeFloat(() => image.Rotation);
                // Excel's object model exposes no flip-state property (only the Flip
                // method), so a flipped picture cannot be detected here.

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
                            "无法读取捕获图像的像素尺寸，坐标无法可靠换算。", 0F, 0F, 0F, 0F, 0F);
                    LogOverlay(region, pixelWidth, pixelHeight, imageLeft, imageTop, imageWidth, imageHeight,
                        rotation, plan, overlayText, ownerId);
                    plans.Add(Tuple.Create(region, plan, overlayText));
                }

                // R5: replace this image's previous results (overlays and notes).
                // Matching is by owner id only, never by region center, so an
                // overlapping image's cleanup cannot delete this image's boxes.
                RemovePreviousResults(sheet, ownerId);

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
                    Excel.Shape? overlay = null;
                    try
                    {
                        // Exact region rect with an opaque cover: in non-bilingual
                        // mode the source text must not show through (the old 8%
                        // transparency did), and the box no longer grows downward.
                        overlay = sheet.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                            imageLeft + plan.Left, imageTop + plan.Top, plan.Width, plan.Height);
                        overlay.AlternativeText = ImageOverlayTags.OverlayFor(ownerId);
                        // D2: the host must never resize the cover by itself. Some
                        // builds default a new textbox to auto-size-to-fit-text,
                        // which would let the box grow beyond the validated region
                        // and make the fit check measure the grown box. The box
                        // keeps the planner's fixed geometry; only the font adapts.
                        overlay.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;
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
                    // W1: Excel's TextFrame has no Overflowing property; the Office
                    // TextRange2.BoundHeight reports the real laid-out text height.
                    // Shrink the font (never the box) until it fits the fixed box;
                    // a box that still overflows at the floor is deleted and the
                    // region degrades to a side-note instead of showing clipped
                    // text as a successful overlay.
                    if (overlay == null)
                    {
                        noteEntries.Add("创建译文覆盖框失败，已降级为旁注。\r" + overlayText);
                        continue;
                    }
                    if (!TextFitsBox(overlay, plan.FontSize))
                    {
                        try { overlay.Delete(); } catch { }
                        noteEntries.Add("覆盖框内译文按真实排版在最小字号下仍溢出，已降级为旁注。\r" + overlayText);
                        continue;
                    }
                    // D2: re-read the geometry and confirm the host did not move
                    // or resize the cover behind our back; it must still match the
                    // planner's fixed rect.
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
                    PlaceCombinedNote(sheet, ownerId, ImageOverlayNotes.Combine(noteEntries),
                        imageLeft, imageTop + imageHeight + 6F, imageWidth);
                    summary.RecordImageNeedsReview();
                }
            });
        }

        private void PlaceCombinedNote(Excel.Worksheet sheet, string ownerId, string noteText, float noteLeft, float noteTop, float imageWidth)
        {
            // A side-note explicitly does NOT claim positional coverage: it is
            // placed below the image.
            var noteWidth = Math.Min(420F, Math.Max(160F, imageWidth));
            // D3: clamp the initial height to the cap up front; an estimate
            // that already exceeds the cap must go through the bounded growth
            // check instead of returning success immediately.
            var noteHeight = Math.Min(ImageOverlayLayout.MaxNoteHeightPt,
                ImageOverlayLayout.EstimateNoteHeight(noteWidth, noteText, 9F));
            var note = sheet.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                noteLeft, noteTop, noteWidth, noteHeight);
            note.AlternativeText = ImageOverlayTags.NoteFor(ownerId);
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

        // W1: real-layout fit check for Excel overlays. TextRange2.BoundHeight
        // is the actual rendered text height with word wrap applied. The font
        // shrinks stepwise to the 8pt floor; the box geometry never expands.
        private static bool TextFitsBox(Excel.Shape box, float startSize)
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
        private static void GrowNoteToFit(Excel.Shape note)
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
        private static bool MatchesPlan(Excel.Shape box, float left, float top, float width, float height)
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

        private static void RemovePreviousResults(Excel.Worksheet sheet, string ownerId)
        {
            // Deletes only shapes that provably belong to this image (by owner
            // id). Legacy id-less markers are retained by default (C4): they
            // cannot be attributed to an image, and deleting them here would
            // destroy other images' translations. Never touches other images'
            // results, even when their rects overlap.
            var doomed = new List<Excel.Shape>();
            foreach (Excel.Shape shape in sheet.Shapes)
            {
                string alt;
                try { alt = shape.AlternativeText; } catch { continue; }
                if (!ImageOverlayTags.WordExcelMarkerBelongsTo(alt, ownerId)) continue;
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

        // Excel copy diagnostics (review bc69f03): keep the original COM
        // HResult/type of a failed Copy/CopyPicture instead of surfacing
        // only the managed wrapper. Newlines are flattened so the text is
        // safe for both the popup and the single-line capture log.
        private static string DescribeComError(string op, COMException ex)
        {
            return op + ": " + ex.GetType().Name + " HResult=0x" +
                ex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) + ": " +
                Flatten(ex.Message);
        }

        private static string ShapeId(Excel.Shape image, int imageOrdinal)
        {
            var name = SafeString(() => image.Name);
            return "#" + imageOrdinal + (string.IsNullOrEmpty(name) ? string.Empty : " " + name);
        }

        // R2: the sheet identity combines workbook + sheet names. Comparing
        // sheet names alone misreports "active" when two open workbooks
        // have a sheet with the same name. When a name cannot be read, the
        // state is "unknown" -- two unreadable names must not be treated
        // as equal.
        private string ShapeState(Excel.Shape image)
        {
            try
            {
                var sheet = image.Parent as Excel.Worksheet;
                var id = SheetId(sheet);
                var active = _excel.ActiveSheet as Excel.Worksheet;
                var activeId = SheetId(active);
                if (id == "?" || activeId == "?") return "sheet=" + id + ",unknown";
                return "sheet=" + id +
                    (string.Equals(activeId, id, StringComparison.Ordinal) ? ",active" : ",inactive");
            }
            catch { return "sheet=?"; }
        }

        private static string SheetId(Excel.Worksheet? sheet)
        {
            if (sheet == null) return "?";
            var book = SafeString(() =>
            {
                var b = sheet.Parent as Excel.Workbook;
                return b != null ? b.Name : "?";
            });
            var name = SafeString(() => sheet.Name);
            if (string.IsNullOrEmpty(book) || book == "?" || string.IsNullOrEmpty(name)) return "?";
            return book + "!" + name;
        }

        private static string Flatten(string? value)
        {
            return (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ");
        }

        private static void LogOverlay(ImageTranslationRegion region,
            int pixelWidth, int pixelHeight, float imageLeft, float imageTop,
            float imageWidth, float imageHeight,
            float rotation, PlannedOverlay plan, string text, string ownerId)
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
                TextLength = text == null ? 0 : text.Length,
                OwnerId = ownerId
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
            if ((shape.Type == Office.MsoShapeType.msoPicture || shape.Type == Office.MsoShapeType.msoLinkedPicture) && !ImageOverlayTags.IsOwnMarker(SafeString(() => shape.AlternativeText))) result.Add(shape);
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
