using Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;
using WordApplication = Microsoft.Office.Interop.Word.Application;

namespace OfficeTranslate.WordAddIn
{
    internal sealed class WordTranslationService
    {
        private readonly WordApplication _word;
        public WordTranslationService(WordApplication word) => _word = word;

        public async Task<TranslationTaskSummary> TranslateAsync(bool wholeDocument, bool bilingual, TranslationSettings settings, CancellationToken token, Action<string> progress)
        {
            // Menu acceptance must prove which DLL served the request: record
            // this add-in's assembly version so the reviewer can verify it in
            // the diagnostics log instead of inferring from registry keys.
            // The module MVID is unique per compilation and distinguishes
            // candidate builds that share the same assembly version.
            ImageOverlayDiagnostics.LogVersion("Word",
                typeof(WordTranslationService).Assembly.GetName().Version?.ToString() ?? "unknown",
                typeof(WordTranslationService).Module.ModuleVersionId.ToString());
            var targets = wholeDocument ? ReadDocumentParagraphs() : ReadSelection();
            var images = settings.ImageOcrEnabled ? (wholeDocument ? ReadDocumentImages() : ReadSelectionImages()) : new List<ImageTarget>();
            if (targets.Count == 0 && images.Count == 0) throw new InvalidOperationException(wholeDocument ? "文档中没有可翻译的正文或图片。" : "请先选择需要翻译的文字或图片。");
            using (var client = new TranslationClient())
            {
                var total = targets.Count + images.Count;
                var summary = new TranslationTaskSummary(total);
                // Word represents a selected inline picture with a non-printing object
                // character. Image-only selections are filtered below, and images are
                // handled before ordinary text so OCR is always the first real request.
                for (var i = 0; i < images.Count; i++)
                {
                    token.ThrowIfCancellationRequested(); var current = i + 1;
                    progress($"OfficeTranslate：正在识别图片 {current}/{total}");
                    await TranslateImageAsync(images[i], client, settings, summary, token);
                    progress($"OfficeTranslate：已完成 {current}/{total}");
                }
                for (var i = 0; i < targets.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    var current = images.Count + i + 1;
                    progress($"OfficeTranslate：正在翻译 {current}/{total}");
                    var sourceText = CleanWordObjectMarkers(targets[i].Text);
                    if (!HasTranslatableText(sourceText))
                    {
                        summary.RecordSkipped();
                        progress($"OfficeTranslate：已完成 {current}/{total}");
                        continue;
                    }
                    var result = await client.TranslateDetailedAsync(sourceText, settings, token);
                    summary.Record(result);
                    token.ThrowIfCancellationRequested();

                    if (!result.Changed)
                    {
                        progress($"OfficeTranslate：已完成 {current}/{total}");
                        continue;
                    }

                    var translatedText = result.Text.TrimEnd('\r', '\a');
                    progress($"OfficeTranslate：正在写回 {i + 1}/{targets.Count}");

                    var undo = _word.UndoRecord;
                    undo.StartCustomRecord(bilingual ? $"OfficeTranslate 双语排版 {i + 1}/{targets.Count}" : $"OfficeTranslate 翻译 {i + 1}/{targets.Count}");
                    try
                    {
                        if (bilingual)
                        {
                            var insertion = targets[i].Range.Duplicate;
                            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
                            insertion.InsertAfter("\r" + translatedText);
                        }
                        else
                            targets[i].Range.Text = translatedText;
                    }
                    finally { undo.EndCustomRecord(); }
                    progress($"OfficeTranslate：已完成 {current}/{total}");
                }
                return summary;
            }
        }

        private async Task TranslateImageAsync(ImageTarget image, TranslationClient client, TranslationSettings settings, TranslationTaskSummary summary, CancellationToken token)
        {
            // Clipboard round-trip is the only way to rasterize a Word shape.
            // CapturePng saves/restores the user's clipboard and clears stale
            // content first so a leftover image is never mistaken for the shape.
            var bytes = ClipboardImageCapture.CapturePng(image.CopyAsPicture, token);
            var hasPixels = PngDimensions.TryRead(bytes, out var pixelWidth, out var pixelHeight);
            // R5/C2: owner identity for cleanup. Floating shapes use their Name
            // (unique per document, persisted, stable across move/resize/
            // reopen). Inline shapes have no Name: each instance gets its own
            // persistent GUID in a document bookmark, so two identical
            // pictures never share an owner.
            var ownerId = image.Kind == "Inline"
                ? GetOrCreateInlineOwnerId(image.Anchor, bytes)
                : ImageOverlayIdentity.ForNamedShape("wd", image.ShapeName, bytes);
            var regions = await client.TranslateImageAsync(bytes, settings, token);
            if (regions.Count == 0) { summary.RecordImage(false); return; }
            summary.RecordImage(true);

            // R4: a floating picture positioned by alignment (center/right/...)
            // reports a WdShapePosition sentinel (e.g. -999995) as Left/Top,
            // which is not a coordinate. It cannot be used for overlays, and
            // the side-note itself must use a real coordinate, so the note is
            // placed at the anchor's page position instead.
            var originIsSentinel = IsAlignmentSentinel(image.Left) || IsAlignmentSentinel(image.Top);
            var sentinelReason = "图片使用对齐定位（如居中/靠右），无法解析实际显示坐标，已降级为旁注。";

            // Plan every region first; old results are replaced only after
            // planning succeeds, so a failure/cancel keeps the previous content.
            var plans = new List<Tuple<ImageTranslationRegion, PlannedOverlay, string>>();
            foreach (var region in regions)
            {
                var bilingualImage = settings.BilingualMode && !string.IsNullOrWhiteSpace(region.Source);
                var overlayText = bilingualImage ? region.Source + "\r" + region.Translation : region.Translation;
                PlannedOverlay plan;
                if (originIsSentinel)
                    plan = new PlannedOverlay(ImageOverlayVerdict.SideNote, sentinelReason, 0F, 0F, 0F, 0F, 0F);
                else if (hasPixels)
                    plan = ImageOverlayPlanner.Plan(
                        region.X1, region.Y1, region.X2, region.Y2,
                        pixelWidth, pixelHeight, image.Width, image.Height, overlayText,
                        image.Rotation);
                else
                    plan = new PlannedOverlay(ImageOverlayVerdict.SideNote,
                        "无法读取捕获图像的像素尺寸，坐标无法可靠换算。", 0F, 0F, 0F, 0F, 0F);
                LogOverlay(image, region, pixelWidth, pixelHeight, plan, overlayText, ownerId);
                plans.Add(Tuple.Create(region, plan, overlayText));
            }

            // R5: replace this image's previous results (overlays and notes).
            // Matching is by owner id only, never by region center or anchor
            // character offsets, so overlapping images and anchor shifts after
            // adding shapes cannot cause wrong deletes or missed deletes.
            RemovePreviousResults(image, ownerId);

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
                string frameFailure;
                if (!PlaceOverlay(image, plan, overlayText, ownerId, out frameFailure))
                    noteEntries.Add(frameFailure + "\r" + overlayText);
            }
            // R2: one image gets ONE combined side-note, placed after all
            // regions are processed, so no region's translation is lost.
            if (noteEntries.Count > 0)
            {
                // C3: coordinates and their reference frame travel together.
                // The anchor page position is only meaningful in the page
                // frame; the image's own numbers are only meaningful in the
                // image's frame. Mixing them misplaces the note.
                float noteLeft, noteTop;
                WdRelativeHorizontalPosition noteRelH;
                WdRelativeVerticalPosition noteRelV;
                if (originIsSentinel)
                {
                    // The sentinel Left/Top must never be reused for the note.
                    if (TryGetAnchorPagePosition(image, out var anchorX, out var anchorY))
                    {
                        noteLeft = anchorX;
                        noteTop = anchorY + 6F;
                    }
                    else
                    {
                        // Explicit, valid fallback: a fixed page position in
                        // the page frame, recorded in the note itself.
                        noteEntries.Insert(0, "无法取得图片锚点的页面坐标，旁注放在页面左上固定位置。");
                        noteLeft = 72F;
                        noteTop = 78F;
                    }
                    noteRelH = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    noteRelV = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                }
                else
                {
                    noteLeft = image.Left;
                    noteTop = image.Top + image.Height + 6F;
                    noteRelH = image.RelativeHorizontalPosition;
                    noteRelV = image.RelativeVerticalPosition;
                }
                PlaceCombinedNote(image, ownerId, ImageOverlayNotes.Combine(noteEntries),
                    noteLeft, noteTop, noteRelH, noteRelV);
                summary.RecordImageNeedsReview();
            }
        }

        // Word reports WdShapePosition alignment constants (wdShapeCenter =
        // -999995, wdShapeLeft = -999998, wdShapeRight = -999997,
        // wdShapeInside = -999999, wdShapeOutside = -999996) as Left/Top for
        // floating pictures positioned by alignment. These are not points.
        private static bool IsAlignmentSentinel(float value)
        {
            return value <= -999990F;
        }

        // C2: every inline picture instance owns a persistent GUID, stored in
        // a document bookmark named OTImg_<32 hex> (38 chars, under Word's
        // 40-char bookmark limit). A content hash alone cannot distinguish
        // two identical pictures, so it must never be the primary owner id.
        // The bookmark persists across save/reopen, tracks the picture through
        // edits, and never modifies the picture's own description. The
        // content hash is only an auxiliary fallback, still position-qualified
        // so identical pictures at different positions never share an owner.
        private string GetOrCreateInlineOwnerId(Range imageRange, byte[] pngBytes)
        {
            const string prefix = "OTImg_";
            int start, end;
            try { start = imageRange.Start; end = imageRange.End; }
            catch { start = -1; end = -1; }
            try
            {
                var doc = _word.ActiveDocument;
                foreach (Bookmark bookmark in doc.Bookmarks)
                {
                    string name;
                    try { name = bookmark.Name; }
                    catch { continue; }
                    if (string.IsNullOrEmpty(name) || !name.StartsWith(prefix, StringComparison.Ordinal)) continue;
                    int bStart, bEnd, shapes;
                    try { bStart = bookmark.Range.Start; bEnd = bookmark.Range.End; shapes = bookmark.Range.InlineShapes.Count; }
                    catch { continue; }
                    // The bookmark must still sit on a live inline picture;
                    // an orphaned bookmark (picture deleted) is not reused.
                    if (shapes > 0 && bStart == start && bEnd == end)
                        return "wdi:guid-" + name.Substring(prefix.Length);
                }
                var guid = Guid.NewGuid().ToString("N");
                object rangeObject = imageRange;
                doc.Bookmarks.Add(prefix + guid, ref rangeObject);
                return "wdi:guid-" + guid;
            }
            catch
            {
                return "wdi:img-" + ImageOverlayIdentity.ContentHash(pngBytes) + "-pos" + start;
            }
        }

        // C3: returns false when the anchor position is unavailable. Word's
        // Information[] returns -1 when the position cannot be determined
        // (e.g. the anchor is not visible); -1, NaN and Infinity are rejected
        // and never used as coordinates.
        private static bool TryGetAnchorPagePosition(ImageTarget image, out float x, out float y)
        {
            x = 72F; y = 72F;
            try
            {
                var px = Convert.ToSingle(image.Anchor.Information[WdInformation.wdHorizontalPositionRelativeToPage]);
                var py = Convert.ToSingle(image.Anchor.Information[WdInformation.wdVerticalPositionRelativeToPage]);
                if (ImageOverlayGeometry.IsUsablePageCoordinate(px) &&
                    ImageOverlayGeometry.IsUsablePageCoordinate(py))
                {
                    x = px; y = py;
                    return true;
                }
            }
            catch { }
            return false;
        }

        // Returns false when the reference frame could not be assigned: the
        // overlay is deleted and the caller degrades the region to a side-note
        // instead of drawing a box that claims a trustworthy position.
        private bool PlaceOverlay(ImageTarget image, PlannedOverlay plan, string text, string ownerId, out string failureReason)
        {
            failureReason = string.Empty;
            object anchor = image.Anchor.Duplicate;
            var overlay = _word.ActiveDocument.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                image.Left, image.Top, plan.Width, plan.Height, ref anchor);
            // Inherit the source image's reference frame. A floating picture's
            // own Left/Top may be paragraph/margin-relative rather than
            // page-relative, so forcing page-relative here was the offset bug.
            // Geometry is assigned AFTER the frame so Office interprets
            // Left/Top in the correct coordinate system. If the frame cannot
            // be assigned, the position is untrustworthy: delete the box.
            try
            {
                overlay.RelativeHorizontalPosition = image.RelativeHorizontalPosition;
                overlay.RelativeVerticalPosition = image.RelativeVerticalPosition;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                try { overlay.Delete(); } catch { }
                failureReason = "无法设置覆盖框的定位参考系，已降级为旁注。";
                return false;
            }
            overlay.AlternativeText = ImageOverlayTags.OverlayFor(ownerId);
            overlay.Left = image.Left + plan.Left;
            overlay.Top = image.Top + plan.Top;
            overlay.Width = plan.Width;
            overlay.Height = plan.Height;
            overlay.WrapFormat.Type = WdWrapType.wdWrapFront;
            // Opaque cover over the validated region: in non-bilingual mode the
            // source text must not show through (the old 8% transparency did).
            overlay.Fill.Visible = Office.MsoTriState.msoTrue;
            overlay.Fill.ForeColor.RGB = 0xFFFFFF;
            overlay.Fill.Transparency = 0F;
            overlay.Line.Visible = Office.MsoTriState.msoFalse;
            overlay.TextFrame.MarginLeft = 2; overlay.TextFrame.MarginRight = 2;
            overlay.TextFrame.MarginTop = 1; overlay.TextFrame.MarginBottom = 1;
            // R1: Word.TextFrame.WordWrap is int in the interop assembly, not MsoTriState.
            overlay.TextFrame.WordWrap = (int)Office.MsoTriState.msoTrue;
            // D2: the host must never resize the cover by itself. Some builds
            // default a new textbox to auto-size-to-fit-text, which would let
            // the box grow beyond the validated region. The box keeps the
            // planner's fixed geometry; only the font size adapts.
            overlay.TextFrame.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;
            overlay.TextFrame.TextRange.Text = text;
            // Word does not consistently support msoAutoSizeTextToFitShape. Some
            // desktop builds reject it with "value out of range", so the font
            // size is fitted inside the region by ImageOverlayPlanner instead
            // of growing the box downward.
            overlay.TextFrame.TextRange.Font.Size = plan.FontSize;
            // W1: the planner's character-width/line-height estimate cannot
            // replace Word's real layout. Shrink the font (never the box)
            // until the text actually fits; if it still overflows at the
            // floor, the box is deleted and the caller degrades the region
            // to a side-note instead of showing clipped text as success.
            if (!ShrinkFontToFit(overlay, plan.FontSize))
            {
                try { overlay.Delete(); } catch { }
                failureReason = "覆盖框内译文按真实排版在最小字号下仍溢出，已降级为旁注。";
                return false;
            }
            // D2: re-read the geometry and confirm the host did not move or
            // resize the cover behind our back; it must still match the
            // planner's fixed rect.
            if (!MatchesPlan(overlay, image.Left + plan.Left, image.Top + plan.Top, plan.Width, plan.Height))
            {
                try { overlay.Delete(); } catch { }
                failureReason = "覆盖框几何被宿主改动，已删除并降级为旁注。";
                return false;
            }
            return true;
        }

        // D2: confirms a box still matches the planner's fixed rect. Guards
        // against host auto-size behaviors that would silently expand the
        // cover beyond the validated region.
        private static bool MatchesPlan(Shape box, float left, float top, float width, float height)
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

        // W1: real-layout overflow check. TextFrame.Overflowing reports whether
        // the current text actually fits the fixed box (after Repaginate).
        // The font shrinks stepwise from the planned size to the 8pt floor
        // (the planner's minimum); the box geometry is never expanded.
        private bool ShrinkFontToFit(Shape box, float startSize)
        {
            const float floorSize = 8F;
            var size = Math.Min(startSize, 18F);
            try { box.TextFrame.TextRange.Font.Size = size; }
            catch { return false; }
            for (var i = 0; i < 12; i++)
            {
                bool overflowing;
                try
                {
                    _word.ActiveDocument.Repaginate();
                    // D1: Word's TextFrame.Overflowing is bool in this interop
                    // assembly, not MsoTriState; read it directly.
                    overflowing = box.TextFrame.Overflowing;
                }
                catch { return false; }
                if (!overflowing) return true;
                if (size <= floorSize) return false;
                size = Math.Max(floorSize, size - 1F);
                try { box.TextFrame.TextRange.Font.Size = size; }
                catch { return false; }
            }
            return false;
        }

        private void PlaceCombinedNote(ImageTarget image, string ownerId, string noteText,
            float noteLeft, float noteTop,
            WdRelativeHorizontalPosition relH, WdRelativeVerticalPosition relV)
        {
            // A side-note explicitly does NOT claim positional coverage: it is
            // placed below the image (or at a fixed page position when the
            // image uses alignment positioning), in the frame its coordinates
            // were computed in (C3: coordinates and frame travel together).
            var noteWidth = Math.Min(420F, Math.Max(160F, image.Width));
            // D3: clamp the initial height to the cap up front; an estimate
            // that already exceeds the cap must go through the bounded growth
            // check instead of returning success immediately.
            var noteHeight = Math.Min(ImageOverlayLayout.MaxNoteHeightPt,
                ImageOverlayLayout.EstimateNoteHeight(noteWidth, noteText, 9F));
            object anchor = image.Anchor.Duplicate;
            var note = _word.ActiveDocument.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                noteLeft, noteTop, noteWidth, noteHeight, ref anchor);
            note.AlternativeText = ImageOverlayTags.NoteFor(ownerId);
            // D2/D3: no host auto-growth for notes either; W2 grows the note
            // manually under the cap so the limit cannot be bypassed.
            note.TextFrame.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;
            // The frame is assigned BEFORE Left/Top so Office interprets the
            // numbers in the intended system. C3: a note whose frame cannot be
            // set is deleted and reported instead of being kept in an unknown
            // frame, which would silently misplace the translations.
            try
            {
                note.RelativeHorizontalPosition = relH;
                note.RelativeVerticalPosition = relV;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                try { note.Delete(); } catch { }
                throw new InvalidOperationException(
                    "已完成图片识别，但无法设置译文旁注的定位参考系，已删除错位旁注避免误导。译文（截断）：" +
                    ImageOverlayText.Truncate(noteText), ex);
            }
            note.Left = noteLeft; note.Top = noteTop; note.Width = noteWidth; note.Height = noteHeight;
            note.WrapFormat.Type = WdWrapType.wdWrapFront;
            note.Fill.Visible = Office.MsoTriState.msoTrue;
            note.Fill.ForeColor.RGB = 0xE1FFFF; // light yellow (BGR)
            note.Fill.Transparency = 0F;
            note.Line.Visible = Office.MsoTriState.msoFalse;
            note.TextFrame.MarginLeft = 4; note.TextFrame.MarginRight = 4;
            note.TextFrame.MarginTop = 3; note.TextFrame.MarginBottom = 3;
            // R1: Word.TextFrame.WordWrap is int in the interop assembly, not MsoTriState.
            note.TextFrame.WordWrap = (int)Office.MsoTriState.msoTrue;
            note.TextFrame.TextRange.Text = noteText;
            note.TextFrame.TextRange.Font.Size = 9F;
            // W2: "no height cap" is not "fully visible". Grow the note
            // downward until Word's real layout reports no overflow. Bounded;
            // a note that still overflows at the cap is deleted and reported
            // explicitly instead of being kept clipped.
            GrowNoteToFit(note);
        }

        // W2: grows the side-note downward until its real layout fits.
        // D3: the cap is enforced BEFORE the fit check on every pass, so an
        // over-cap box can never return success, however it got that tall
        // (estimator overshoot or host auto-growth). Text that cannot fit
        // inside the cap is deleted and reported explicitly instead of being
        // kept silently clipped.
        private void GrowNoteToFit(Shape note)
        {
            for (var i = 0; i < 12; i++)
            {
                float height;
                try { height = note.Height; }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    try { note.Delete(); } catch { }
                    throw new InvalidOperationException(
                        "已完成图片识别，但无法读取译文旁注的尺寸，已删除旁注避免显示不全。", ex);
                }
                if (height > ImageOverlayLayout.MaxNoteHeightPt)
                {
                    try { note.Height = ImageOverlayLayout.MaxNoteHeightPt; }
                    catch (System.Runtime.InteropServices.COMException ex)
                    {
                        try { note.Delete(); } catch { }
                        throw new InvalidOperationException(
                            "已完成图片识别，但无法约束译文旁注的高度，已删除避免显示不全。", ex);
                    }
                    height = ImageOverlayLayout.MaxNoteHeightPt;
                }
                bool overflowing;
                try
                {
                    _word.ActiveDocument.Repaginate();
                    // D1: bool, not MsoTriState (see ShrinkFontToFit).
                    overflowing = note.TextFrame.Overflowing;
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    try { note.Delete(); } catch { }
                    throw new InvalidOperationException(
                        "已完成图片识别，但无法校验译文旁注的排版溢出，已删除旁注避免显示不全。", ex);
                }
                if (!overflowing) return;
                if (height >= ImageOverlayLayout.MaxNoteHeightPt)
                {
                    string text;
                    try { text = note.TextFrame.TextRange.Text; } catch { text = string.Empty; }
                    try { note.Delete(); } catch { }
                    throw new InvalidOperationException(
                        "译文旁注过长，增高到上限仍无法完整显示，已删除避免静默裁切。译文（截断）：" +
                        ImageOverlayText.Truncate(text));
                }
                try { note.Height = Math.Min(ImageOverlayLayout.MaxNoteHeightPt, height * 1.5F + 12F); }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    try { note.Delete(); } catch { }
                    throw new InvalidOperationException(
                        "已完成图片识别，但无法增高译文旁注，已删除避免显示不全。", ex);
                }
            }
            string leftover;
            try { leftover = note.TextFrame.TextRange.Text; } catch { leftover = string.Empty; }
            try { note.Delete(); } catch { }
            throw new InvalidOperationException(
                "译文旁注排版校验未收敛，已删除避免静默裁切。译文（截断）：" +
                ImageOverlayText.Truncate(leftover));
        }

        private void RemovePreviousResults(ImageTarget image, string ownerId)
        {
            // Deletes only shapes that provably belong to this image (by owner
            // id). Legacy id-less markers are retained by default (C4): they
            // cannot be attributed to an image, and deleting them here would
            // destroy other images' translations. Never touches other images'
            // results, even when their rects overlap.
            var doomed = new List<Shape>();
            foreach (Shape shape in _word.ActiveDocument.Shapes)
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

        private void LogOverlay(ImageTarget image, ImageTranslationRegion region,
            int pixelWidth, int pixelHeight, PlannedOverlay plan, string text, string ownerId)
        {
            ImageOverlayDiagnostics.Log(new ImageOverlayDiagnosticEntry
            {
                Host = "Word",
                ImageKind = image.Kind,
                PixelWidth = pixelWidth,
                PixelHeight = pixelHeight,
                BboxX1 = region.X1, BboxY1 = region.Y1, BboxX2 = region.X2, BboxY2 = region.Y2,
                ShapeLeft = image.Left, ShapeTop = image.Top,
                ShapeWidth = image.Width, ShapeHeight = image.Height,
                FrameNote = image.Kind + "(RelH=" + image.RelativeHorizontalPosition +
                    ",RelV=" + image.RelativeVerticalPosition +
                    ",crop=" + image.CropLeft + "/" + image.CropTop + "/" + image.CropRight + "/" + image.CropBottom + ";FlipStateUnknown)",
                RotationDegrees = image.Rotation,
                FlipHorizontal = image.FlipHorizontal,
                FlipVertical = image.FlipVertical,
                OutLeft = image.Left + plan.Left,
                OutTop = image.Top + plan.Top,
                OutWidth = plan.Width,
                OutHeight = plan.Height,
                FontSize = plan.FontSize,
                Verdict = plan.Verdict.ToString(),
                Reason = plan.Reason,
                TextLength = text == null ? 0 : text.Length,
                OwnerId = ownerId
            });
        }

        private List<ImageTarget> ReadDocumentImages()
        {
            var result = new List<ImageTarget>();
            foreach (InlineShape shape in _word.ActiveDocument.InlineShapes) AddInlineImage(shape, result);
            foreach (Shape shape in _word.ActiveDocument.Shapes) AddFloatingImage(shape, result);
            return result;
        }

        private List<ImageTarget> ReadSelectionImages()
        {
            var result = new List<ImageTarget>(); var selection = _word.Selection;
            foreach (InlineShape shape in selection.Range.InlineShapes) AddInlineImage(shape, result);
            try { foreach (Shape shape in selection.ShapeRange) AddFloatingImage(shape, result); } catch (System.Runtime.InteropServices.COMException) { }
            return result;
        }

        private static void AddInlineImage(InlineShape shape, List<ImageTarget> result)
        {
            if (shape.Type != WdInlineShapeType.wdInlineShapePicture && shape.Type != WdInlineShapeType.wdInlineShapeLinkedPicture) return;
            var range = shape.Range.Duplicate; var left = Convert.ToSingle(range.Information[WdInformation.wdHorizontalPositionRelativeToPage]); var top = Convert.ToSingle(range.Information[WdInformation.wdVerticalPositionRelativeToPage]);
            var target = new ImageTarget(range, left, top, shape.Width, shape.Height, () => range.CopyAsPicture());
            target.Kind = "Inline";
            // Inline pictures cannot be rotated in Word, and InlineShape does
            // not expose crop; the planner's aspect-ratio check remains the guard.
            target.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            target.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            result.Add(target);
        }

        private static void AddFloatingImage(Shape shape, List<ImageTarget> result)
        {
            if (ImageOverlayTags.IsOwnMarker(SafeGet(() => shape.AlternativeText, string.Empty))) return;
            if (shape.Type == Office.MsoShapeType.msoGroup) { for (var i = 1; i <= shape.GroupItems.Count; i++) AddFloatingImage(shape.GroupItems[i], result); return; }
            if (shape.Type != Office.MsoShapeType.msoPicture && shape.Type != Office.MsoShapeType.msoLinkedPicture) return;
            var anchor = shape.Anchor.Duplicate;
            var target = new ImageTarget(anchor, shape.Left, shape.Top, shape.Width, shape.Height, () => { object replace = true; shape.Select(ref replace); shape.Anchor.Application.Selection.CopyAsPicture(); });
            target.Kind = "Floating";
            target.Rotation = SafeFloat(() => shape.Rotation);
            // Word's object model exposes no flip-state property (only the Flip
            // method), so a flipped floating picture cannot be detected here.
            // This is a known limitation, recorded in the diagnostics as false.
            var format = SafeGet(() => shape.PictureFormat, null);
            if (format != null)
            {
                target.CropLeft = SafeFloat(() => format.CropLeft);
                target.CropTop = SafeFloat(() => format.CropTop);
                target.CropRight = SafeFloat(() => format.CropRight);
                target.CropBottom = SafeFloat(() => format.CropBottom);
            }
            target.RelativeHorizontalPosition = SafeGet(() => shape.RelativeHorizontalPosition, WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage);
            target.RelativeVerticalPosition = SafeGet(() => shape.RelativeVerticalPosition, WdRelativeVerticalPosition.wdRelativeVerticalPositionPage);
            target.ShapeName = SafeGet(() => shape.Name, string.Empty);
            result.Add(target);
        }

        private static float SafeFloat(Func<float> read)
        {
            try { return read(); } catch { return 0F; }
        }

        private static T SafeGet<T>(Func<T> read, T fallback)
        {
            try { return read(); } catch { return fallback; }
        }

        private List<Target> ReadSelection()
        {
            var selection = _word.Selection;
            if (selection == null || selection.Range.Start == selection.Range.End || !HasTranslatableText(selection.Range.Text)) return new List<Target>();
            return new List<Target> { ToTarget(selection.Range) };
        }

        private List<Target> ReadDocumentParagraphs()
        {
            var result = new List<Target>();
            foreach (Paragraph paragraph in _word.ActiveDocument.StoryRanges[WdStoryType.wdMainTextStory].Paragraphs)
            {
                var target = ToTarget(paragraph.Range);
                if (HasTranslatableText(target.Text)) result.Add(target);
            }
            return result;
        }

        private static Target ToTarget(Range range)
        {
            var raw = range.Text ?? string.Empty;
            var contentLength = raw.Length;
            while (contentLength > 0 && (raw[contentLength - 1] == '\r' || raw[contentLength - 1] == '\a')) contentLength--;
            var liveRange = range.Duplicate;
            liveRange.End = liveRange.Start + contentLength;
            return new Target(liveRange, raw.Substring(0, contentLength));
        }

        private static string StripMarks(string text) => text.TrimEnd('\r', '\a');

        private static bool HasTranslatableText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            foreach (var character in CleanWordObjectMarkers(text))
                if (char.IsLetterOrDigit(character)) return true;
            return false;
        }

        private static string CleanWordObjectMarkers(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var output = new System.Text.StringBuilder(text.Length);
            foreach (var character in StripMarks(text))
            {
                // Word uses several low control characters for inline pictures,
                // fields and anchors. Preserve only normal text plus useful layout.
                if (character == '\uFFFC') continue;
                if (char.IsControl(character) && character != '\r' && character != '\n' &&
                    character != '\v' && character != '\f' && character != '\t') continue;
                output.Append(character);
            }
            return output.ToString();
        }

        private sealed class Target
        {
            public Target(Range range, string text) { Range = range; Text = text; }
            public Range Range { get; }
            public string Text { get; }
        }

        private sealed class ImageTarget
        {
            private readonly Action _selectOrCopy;
            public ImageTarget(Range anchor, float left, float top, float width, float height, Action action) { Anchor = anchor; Left = left; Top = top; Width = width; Height = height; _selectOrCopy = action; }
            public Range Anchor { get; } public float Left { get; } public float Top { get; } public float Width { get; } public float Height { get; }
            public string Kind { get; set; } = "Unknown";
            public string ShapeName { get; set; } = string.Empty;
            public float Rotation { get; set; }
            public bool FlipHorizontal { get; set; }
            public bool FlipVertical { get; set; }
            public float CropLeft { get; set; }
            public float CropTop { get; set; }
            public float CropRight { get; set; }
            public float CropBottom { get; set; }
            public WdRelativeHorizontalPosition RelativeHorizontalPosition { get; set; } = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            public WdRelativeVerticalPosition RelativeVerticalPosition { get; set; } = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            public void CopyAsPicture() => _selectOrCopy();
        }
    }
}
