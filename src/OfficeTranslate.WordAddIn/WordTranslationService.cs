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

        public async Task<TranslationTaskSummary> TranslateAsync(bool wholeDocument, bool bilingual, TranslationSettings settings, CancellationToken token, Action<string> progress, OfficeUiDispatcher ui)
        {
            // Menu acceptance must prove which DLL served the request: record
            // this add-in's assembly version so the reviewer can verify it in
            // the diagnostics log instead of inferring from registry keys.
            // The module MVID is unique per compilation and distinguishes
            // candidate builds that share the same assembly version.
            ImageOverlayDiagnostics.LogVersion("Word",
                typeof(WordTranslationService).Assembly.GetName().Version?.ToString() ?? "unknown",
                typeof(WordTranslationService).Module.ModuleVersionId.ToString());
            // M2: prove which thread the synchronous prefix runs on; every
            // COM/clipboard section below is dispatched explicitly.
            ui.LogProbe("translate_start");
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
                ui.LogProbe("image_loop_before");
                for (var i = 0; i < images.Count; i++)
                {
                    token.ThrowIfCancellationRequested(); var current = i + 1;
                    progress($"OfficeTranslate：正在识别图片 {current}/{total}");
                    await TranslateImageAsync(images[i], client, settings, summary, token, ui);
                    progress($"OfficeTranslate：已完成 {current}/{total}");
                }
                ui.LogProbe("image_loop_after");
                ui.LogProbe("text_loop_before");
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

                    // M2: the writeback touches the Word object model, so it
                    // runs on the Office UI (STA) thread explicitly instead of
                    // on whatever thread the network await resumed on.
                    await ui.InvokeAsync(() =>
                    {
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
                    });
                    progress($"OfficeTranslate：已完成 {current}/{total}");
                }
                ui.LogProbe("text_loop_after");
                return summary;
            }
        }

        private async Task TranslateImageAsync(ImageTarget image, TranslationClient client, TranslationSettings settings, TranslationTaskSummary summary, CancellationToken token, OfficeUiDispatcher ui)
        {
            // M2: clipboard capture and owner-identity bookkeeping touch COM
            // and the clipboard, so they run on the Office UI (STA) thread
            // explicitly. Async continuations are not guaranteed to resume
            // there (a later menu run used to fail here with an STA/OLE error).
            var preNetwork = await ui.InvokeAsync(() =>
            {
                ui.LogProbe("capture_before");
                // Clipboard round-trip is the only way to rasterize a Word shape.
                // CapturePng saves/restores the user's clipboard and clears stale
                // content first so a leftover image is never mistaken for the shape.
                var captured = ClipboardImageCapture.CapturePng(image.CopyAsPicture, token);
                ui.LogProbe("capture_after");
                // R5/C2: owner identity for cleanup. Floating shapes use their Name
                // (unique per document, persisted, stable across move/resize/
                // reopen). Inline shapes have no Name: each instance gets its own
                // persistent GUID in a document bookmark, so two identical
                // pictures never share an owner.
                // M1: ResolveInlineOwnerId only PROVES identity here; bookmark
                // deletion and re-tightening are committed after the OCR plan
                // succeeds, so a cancelled/failed/empty run never mutates the
                // document.
                InlineOwnerResolution? inlineResolution = null;
                string owner;
                if (image.Kind == "Inline")
                {
                    inlineResolution = ResolveInlineOwnerId(image.Anchor);
                    owner = inlineResolution.OwnerId;
                }
                else
                {
                    owner = ImageOverlayIdentity.ForNamedShape("wd", image.ShapeName, captured);
                }
                return Tuple.Create(captured, owner, inlineResolution);
            });
            var bytes = preNetwork.Item1;
            var ownerId = preNetwork.Item2;
            var inlineResolution = preNetwork.Item3;
            var hasPixels = PngDimensions.TryRead(bytes, out var pixelWidth, out var pixelHeight);
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
            // M1: commit owner bookkeeping only after planning succeeded.
            // Duplicate bookmarks are deleted and the canonical bookmark
            // re-tightened here; their stale boxes are removed below by the
            // absorbed owner ids.
            // M2: all shape surgery below runs on the Office UI (STA) thread.
            await ui.InvokeAsync(() =>
            {
                if (inlineResolution != null)
                    CommitInlineOwnerResolution(image.Anchor, inlineResolution);
                RemovePreviousResults(image, ownerId);
                if (inlineResolution != null)
                    foreach (var absorbed in inlineResolution.AbsorbedOwnerIds)
                        RemovePreviousResults(image, absorbed);

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
            });
        }

        // Word reports WdShapePosition alignment constants (wdShapeCenter =
        // -999995, wdShapeLeft = -999998, wdShapeRight = -999997,
        // wdShapeInside = -999999, wdShapeOutside = -999996) as Left/Top for
        // floating pictures positioned by alignment. These are not points.
        private static bool IsAlignmentSentinel(float value)
        {
            return value <= -999990F;
        }

        // C2/M1: every inline picture instance owns a persistent GUID, stored in
        // a document bookmark named OTImg_<32 hex> (38 chars, under Word's
        // 40-char bookmark limit). A content hash alone cannot distinguish
        // two identical pictures, so it must never be the primary owner id.
        // The bookmark persists across save/reopen, tracks the picture through
        // edits, and never modifies the picture's own description. The
        // content hash is only an auxiliary fallback, still position-qualified
        // so identical pictures at different positions never share an owner.
        //
        // M1: a bookmark's Range EXPANDS when the document is edited (text
        // writeback, overlay insertion), so the bookmark's outer bounds are no
        // longer a reliable identity test -- the old strict bounds comparison
        // minted a fresh GUID on every re-translation and orphaned the
        // previous boxes. Identity is now proven by the bookmark's range
        // containing exactly ONE inline shape whose range and story match the
        // image. A bookmark spanning zero or several inline shapes is
        // ambiguous (or orphaned) and is never reused: no content-hash or
        // center-point guessing.
        //
        // Two phases, separated by the network call:
        //   - ResolveInlineOwnerId (pre-network) only READS: it proves which
        //     bookmarks name this exact instance, picks the canonical one by
        //     a deterministic order (bookmark start, end, name), and reports
        //     the rest as duplicates. It changes nothing except creating the
        //     first bookmark when none exists.
        //   - CommitInlineOwnerResolution (post-plan, UI thread) deletes the
        //     duplicate bookmarks and re-tightens the canonical bookmark to
        //     the image range. A cancelled/failed/empty OCR run therefore
        //     never mutates the document.
        // Re-tightening matters because an expanded bookmark range could pick
        // up an unrelated nearby picture later and turn ambiguous.
        private sealed class InlineOwnerResolution
        {
            public string OwnerId = string.Empty;
            public string CanonicalBookmark = string.Empty;
            public readonly List<string> DuplicateBookmarks = new List<string>();
            public readonly List<string> AbsorbedOwnerIds = new List<string>();
        }

        private InlineOwnerResolution ResolveInlineOwnerId(Range imageRange)
        {
            const string prefix = "OTImg_";
            var result = new InlineOwnerResolution();
            int start, end;
            WdStoryType story;
            try { start = imageRange.Start; end = imageRange.End; story = imageRange.StoryType; }
            catch { start = -1; end = -1; story = WdStoryType.wdMainTextStory; }
            try
            {
                var doc = _word.ActiveDocument;
                var matches = new List<Tuple<string, int, int>>();
                foreach (Bookmark bookmark in doc.Bookmarks)
                {
                    string name;
                    try { name = bookmark.Name; }
                    catch { continue; }
                    if (string.IsNullOrEmpty(name) || !name.StartsWith(prefix, StringComparison.Ordinal)) continue;
                    Range shapeRange;
                    int bStart, bEnd;
                    try
                    {
                        var bRange = bookmark.Range;
                        // Exactly one inline shape: zero means the picture is
                        // gone (orphaned bookmark), several means ambiguous.
                        // Both are rejected, never guessed.
                        if (bRange.InlineShapes.Count != 1) continue;
                        bStart = bRange.Start; bEnd = bRange.End;
                        shapeRange = bRange.InlineShapes[1].Range;
                    }
                    catch { continue; }
                    int sStart, sEnd;
                    WdStoryType sStory;
                    try { sStart = shapeRange.Start; sEnd = shapeRange.End; sStory = shapeRange.StoryType; }
                    catch { continue; }
                    // The single contained inline shape must be exactly this
                    // image: same range in the same story. Range offsets alone
                    // are not enough because different stories can share them,
                    // and two identical pictures at different positions have
                    // different ranges, so they never merge.
                    if (sStart == start && sEnd == end && sStory == story)
                        matches.Add(Tuple.Create(name, bStart, bEnd));
                }
                // Deterministic canonical choice: bookmark start, then end,
                // then name. Bookmark enumeration order is not specified, so
                // it must not decide. The pure selection lives in Core so it
                // is unit-testable without COM.
                if (matches.Count >= 1)
                {
                    var chosen = ImageOverlayIdentity.ChooseCanonicalBookmark(matches);
                    result.CanonicalBookmark = chosen.Item1;
                    result.OwnerId = "wdi:guid-" + chosen.Item1.Substring(prefix.Length);
                    foreach (var duplicate in chosen.Item2)
                    {
                        result.DuplicateBookmarks.Add(duplicate);
                        result.AbsorbedOwnerIds.Add("wdi:guid-" + duplicate.Substring(prefix.Length));
                    }
                    return result;
                }
                var guid = Guid.NewGuid().ToString("N");
                object rangeObject = imageRange;
                doc.Bookmarks.Add(prefix + guid, ref rangeObject);
                result.CanonicalBookmark = prefix + guid;
                result.OwnerId = "wdi:guid-" + guid;
                return result;
            }
            catch
            {
                // Identity unprovable: refuse to guess. A fresh GUID never
                // merges two images; worst case it orphans this run's boxes,
                // exactly like the pre-M1 behavior, but only on this
                // exceptional path. No content-hash fallback: two identical
                // pictures must never share an owner.
                result.OwnerId = "wdi:guid-" + Guid.NewGuid().ToString("N");
                result.CanonicalBookmark = string.Empty;
                return result;
            }
        }

        private void CommitInlineOwnerResolution(Range imageRange, InlineOwnerResolution resolution)
        {
            if (string.IsNullOrEmpty(resolution.CanonicalBookmark)) return;
            Document doc;
            try { doc = _word.ActiveDocument; }
            catch { return; }
            // Re-tighten first: if this fails, the duplicates are still
            // reported and absorbed on the next run.
            TightenBookmark(doc, resolution.CanonicalBookmark, imageRange);
            foreach (var duplicate in resolution.DuplicateBookmarks)
                try { doc.Bookmarks[duplicate].Delete(); } catch { }
        }

        // Re-tightens an expanded bookmark to the image range without ever
        // losing the only bookmark: the tight range is added under a
        // temporary name FIRST; the original is deleted only after the tight
        // bookmark exists. Every step is best-effort -- identity does not
        // depend on tightening, the next run simply retries.
        private static void TightenBookmark(Document doc, string name, Range imageRange)
        {
            int start, end;
            try { start = imageRange.Start; end = imageRange.End; }
            catch { return; }
            Bookmark canonical;
            try { canonical = doc.Bookmarks[name]; }
            catch { return; }
            int bStart, bEnd;
            try { bStart = canonical.Range.Start; bEnd = canonical.Range.End; }
            catch { return; }
            if (bStart == start && bEnd == end) return;
            var tmp = name + "_tighten";
            try { doc.Bookmarks[tmp].Delete(); } catch { }
            try
            {
                object tightRange = imageRange;
                doc.Bookmarks.Add(tmp, ref tightRange);
            }
            catch { return; } // original untouched
            Range newRange;
            try { newRange = doc.Bookmarks[tmp].Range; }
            catch { return; } // original untouched; tmp is absorbed next run
            try { doc.Bookmarks[name].Delete(); }
            catch { try { doc.Bookmarks[tmp].Delete(); } catch { } return; }
            try
            {
                object r = newRange;
                doc.Bookmarks.Add(name, ref r);
                try { doc.Bookmarks[tmp].Delete(); } catch { }
            }
            catch
            {
                // Original lost but tmp still proves identity; retry the
                // restore once, otherwise the next run absorbs tmp.
                try
                {
                    object r2 = newRange;
                    doc.Bookmarks.Add(name, ref r2);
                    try { doc.Bookmarks[tmp].Delete(); } catch { }
                }
                catch { }
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
            // Word.TextFrame.AutoSize is int in the interop assembly, not
            // MsoAutoSize (same quirk as WordWrap, see R1).
            overlay.TextFrame.AutoSize = (int)Office.MsoAutoSize.msoAutoSizeNone;
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
            // Word.TextFrame.AutoSize is int in the interop assembly (see R1).
            note.TextFrame.AutoSize = (int)Office.MsoAutoSize.msoAutoSizeNone;
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
