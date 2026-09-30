using Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
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
            // R1/P2: save the user's selection once at task entry, before
            // the first selection-changing COM op. The per-image restore in
            // SelectAndCopyAsPicture only covers the copy phase; overlay
            // box/bookmark creation moves the selection again afterwards,
            // so the authoritative restore runs in the finally below, at
            // the end of the whole task. The duplicate carries the original
            // Document/Story identity (header/footer/footnote safe).
            Range? taskSelection = null;
            int taskGeneration = 0;
            try
            {
                // N1: every task takes a new selection generation, shared
                // across service instances (static). A restore ticket from
                // an older task must never overwrite a newer task's (or
                // the user's) selection.
                ui.Invoke(() =>
                {
                    taskGeneration = Interlocked.Increment(ref _selectionGeneration);
                    taskSelection = SaveSelectionDuplicate();
                });
            }
            catch { taskSelection = null; }
            try
            {
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
                    await TranslateImageAsync(images[i], current, client, settings, summary, token, ui);
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
                        // N3: re-check inside the queued UI callback. The
                        // token may have been cancelled after the pre-queue
                        // check above but before this callback ran; a
                        // cancelled run must not write back. The cancel
                        // boundary is "the item whose commit has started":
                        // once mutations begin they run to completion.
                        token.ThrowIfCancellationRequested();
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
            finally
            {
                // N1: the 2s bound below is a WAIT boundary, not a guarantee
                // on the COM call itself. A queued restore callback that only
                // runs AFTER the wait gave up must not clobber the user's
                // newer selection, so every restore carries a one-shot
                // ticket: pending -> running is claimed atomically when the
                // callback starts on the UI thread, and pending -> expired
                // is set atomically when the wait gives up. An expired
                // ticket makes the late callback a no-op. A callback that
                // already started running cannot be interrupted; its own
                // generation/staleness checks still apply inside.
                if (taskSelection != null)
                {
                    var ticket = new RestoreTicket(taskSelection, taskGeneration,
                        Stopwatch.GetTimestamp());
                    try
                    {
                        var restore = ui.InvokeAsync(() => ClaimAndRestore(ticket));
                        var finished = await Task.WhenAny(restore, Task.Delay(SelectionRestoreWaitMs));
                        if (!ReferenceEquals(finished, restore))
                        {
                            // Wait gave up first: expire the ticket so the
                            // still-queued callback becomes a no-op instead
                            // of overwriting the user's newer selection.
                            // Atomic: only pending -> expired. A callback
                            // that already claimed running is untouched.
                            if (Interlocked.CompareExchange(ref ticket.State,
                                    RestoreTicket.Expired, RestoreTicket.Pending) == RestoreTicket.Pending)
                            {
                                try { ImageOverlayDiagnostics.LogCaptureFailure("Word", "selection_restore_expired"); }
                                catch { }
                            }
                        }
                    }
                    catch { }
                }
            }
        }

        private async Task TranslateImageAsync(ImageTarget image, int imageNumber, TranslationClient client, TranslationSettings settings, TranslationTaskSummary summary, CancellationToken token, OfficeUiDispatcher ui)
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
                // S1: pump the UI message loop while capturing (like the Excel
                // host): Office delayed clipboard rendering may need its
                // messages dispatched, and a transiently busy clipboard is
                // retried a bounded number of times inside CapturePng.
                byte[] captured;
                try
                {
                    captured = ClipboardImageCapture.CapturePng(
                        image.CopyAsPicture,
                        () => System.Windows.Forms.Application.DoEvents(),
                        token);
                }
                catch (Exception ex) when (!(ex is OperationCanceledException))
                {
                    // S2/D1: capture failures must reach the diagnostics log
                    // file, not just the popup. Metadata only: image index,
                    // kind, range start/end, story type, live inline-shape
                    // count. No document content, clipboard text, or keys.
                    try
                    {
                        var detail = "image=" + imageNumber
                            + " kind=" + image.Kind
                            + " range=" + SafeGet(() => image.Anchor.Start, -1)
                            + "-" + SafeGet(() => image.Anchor.End, -1)
                            + " story=" + SafeGet(() => (int)image.Anchor.StoryType, -1)
                            + " inlineShapes=" + SafeGet(() => _word.ActiveDocument.InlineShapes.Count, -1)
                            + " error=" + ex.Message;
                        ImageOverlayDiagnostics.LogCaptureFailure("Word", detail);
                    }
                    catch { }
                    throw new InvalidOperationException(
                        "第 " + imageNumber + " 张图片捕获失败：" + ex.Message, ex);
                }
                ui.LogProbe("capture_after");
                // R5/C2: owner identity for cleanup. Floating shapes use their Name
                // (unique per document, persisted, stable across move/resize/
                // reopen). Inline shapes have no Name: each instance gets its own
                // persistent GUID in a document bookmark, so two identical
                // pictures never share an owner.
                // M1/F4: ResolveInlineOwnerId only PROVES identity here; the
                // first bookmark creation, duplicate deletion and
                // re-tightening are committed after the OCR plan succeeds,
                // so a cancelled/failed/empty run never mutates the document.
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
            // F4: one last cancellation check before the commit phase, so a
            // cancelled run never reaches the bookmark/shape mutations below.
            token.ThrowIfCancellationRequested();
            await ui.InvokeAsync(() =>
            {
                // N3: re-check inside the queued UI callback -- see the text
                // writeback above. Once this callback starts, its bookmark
                // and shape mutations run to completion for this image.
                token.ThrowIfCancellationRequested();
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
        //     the rest as duplicates. For a first sighting it records the
        //     pending bookmark name but creates nothing: the document is not
        //     modified.
        //   - CommitInlineOwnerResolution (post-plan, UI thread) creates the
        //     pending first-sighting bookmark, deletes the duplicate
        //     bookmarks and re-tightens the canonical bookmark to the image
        //     range. A cancelled/failed/empty OCR run returns before commit
        //     and therefore never mutates the document.
        // Re-tightening matters because an expanded bookmark range could pick
        // up an unrelated nearby picture later and turn ambiguous.
        private sealed class InlineOwnerResolution
        {
            public string OwnerId = string.Empty;
            public string CanonicalBookmark = string.Empty;
            public readonly List<string> DuplicateBookmarks = new List<string>();
            public readonly List<string> AbsorbedOwnerIds = new List<string>();
            // F4: first sighting. The bookmark does NOT exist yet; Resolve
            // only records the name. Commit creates it after the OCR plan
            // succeeds, so a cancelled/empty/failed run never mutates the
            // document. Empty when the canonical bookmark already existed.
            public string PendingBookmark = string.Empty;
            // R1: staging adoption. When the pre-network resolve proves a
            // verifiable OTTmp_<32hex> staging belongs to this image and no
            // canonical does, the resolution ADOPTS that staging's identity
            // instead of minting a new owner: AdoptedStaging names the
            // staging bookmark, and Commit restores its canonical from it
            // (RecoverInterruptedTightens). RejectedStagings are proven
            // stagings whose guid lost the canonical choice: their identity
            // was absorbed into the chosen one, so the commit-phase recovery
            // drops them WITHOUT resurrecting their canonical.
            public string AdoptedStaging = string.Empty;
            public readonly List<string> RejectedStagings = new List<string>();
        }

        // R1: one bookmark (canonical or verifiable staging) whose strict
        // ownership proof succeeded for the image being resolved.
        private sealed class ProvenIdentity
        {
            public readonly string BookmarkName;
            public readonly string Guid;
            public readonly int BStart;
            public readonly int BEnd;
            public readonly bool IsStaging;
            public ProvenIdentity(string bookmarkName, string guid, int bStart, int bEnd, bool isStaging)
            {
                BookmarkName = bookmarkName;
                Guid = guid;
                BStart = bStart;
                BEnd = bEnd;
                IsStaging = isStaging;
            }
        }

        private InlineOwnerResolution ResolveInlineOwnerId(Range imageRange)
        {
            var result = new InlineOwnerResolution();
            int start, end;
            WdStoryType story;
            try { start = imageRange.Start; end = imageRange.End; story = imageRange.StoryType; }
            catch { start = -1; end = -1; story = WdStoryType.wdMainTextStory; }
            try
            {
                var doc = _word.ActiveDocument;
                // R1: recognize verifiable staging bookmarks (OTTmp_<32hex>)
                // in addition to canonical ones. A staging left by an
                // interrupted tighten carries the full identity in its name;
                // when it provably belongs to this image (exactly the same
                // strict proof as a canonical) the resolution ADOPTS its guid
                // instead of minting a new one. A name alone never proves
                // ownership: unproven stagings are ignored here, exactly as
                // before, and a cancelled/empty/failed run still returns
                // before Commit and never mutates the document.
                var proven = new List<ProvenIdentity>();
                foreach (Bookmark bookmark in doc.Bookmarks)
                {
                    string name;
                    try { name = bookmark.Name; }
                    catch { continue; }
                    if (string.IsNullOrEmpty(name)) continue;
                    bool isCanonical = name.StartsWith(ImageOverlayIdentity.CanonicalBookmarkPrefix, StringComparison.Ordinal);
                    bool isStaging = !isCanonical && ImageOverlayIdentity.IsVerifiableStagingName(name);
                    if (!isCanonical && !isStaging) continue;
                    string guid = isCanonical
                        ? name.Substring(ImageOverlayIdentity.CanonicalBookmarkPrefix.Length)
                        : name.Substring(ImageOverlayIdentity.StagingBookmarkPrefix.Length);
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
                        proven.Add(new ProvenIdentity(name, guid, bStart, bEnd, isStaging));
                }
                // Deterministic canonical choice: bookmark start, then end,
                // then name. Bookmark enumeration order is not specified, so
                // it must not decide. The pure selection lives in Core so it
                // is unit-testable without COM. One identity per guid: a
                // canonical proof beats a staging proof for the same guid
                // (the staging is then just a leftover that the commit-phase
                // recovery drops once the canonical is confirmed).
                if (proven.Count >= 1)
                {
                    var byGuid = new Dictionary<string, ProvenIdentity>(StringComparer.Ordinal);
                    foreach (var p in proven)
                    {
                        ProvenIdentity existing;
                        if (!byGuid.TryGetValue(p.Guid, out existing) || (!p.IsStaging && existing.IsStaging))
                            byGuid[p.Guid] = p;
                    }
                    var matches = new List<Tuple<string, int, int>>();
                    foreach (var p in byGuid.Values)
                        matches.Add(Tuple.Create(p.BookmarkName, p.BStart, p.BEnd));
                    var chosen = ImageOverlayIdentity.ChooseCanonicalBookmark(matches);
                    var byName = new Dictionary<string, ProvenIdentity>(StringComparer.Ordinal);
                    foreach (var p in byGuid.Values) byName[p.BookmarkName] = p;
                    if (!byName.TryGetValue(chosen.Item1, out var winner))
                        throw new InvalidOperationException("内部错误：未能确定图片身份。");
                    result.CanonicalBookmark = ImageOverlayIdentity.CanonicalBookmarkPrefix + winner.Guid;
                    result.OwnerId = "wdi:guid-" + winner.Guid;
                    if (winner.IsStaging)
                    {
                        // R1: adopt the interrupted identity. The canonical
                        // bookmark does not exist yet; Commit restores it
                        // from this staging (RecoverInterruptedTightens)
                        // instead of creating a second owner. PendingBookmark
                        // reuses the existing creation path; the recovery
                        // reports the restored names so the Add is skipped
                        // when the restore already succeeded, and a failed
                        // restore falls through to creating the canonical at
                        // the current image range -- with the SAME adopted
                        // guid, never a second identity.
                        result.PendingBookmark = result.CanonicalBookmark;
                        result.AdoptedStaging = winner.BookmarkName;
                    }
                    foreach (var duplicate in chosen.Item2)
                    {
                        if (!byName.TryGetValue(duplicate, out var dup)) continue;
                        if (dup.IsStaging)
                            result.RejectedStagings.Add(dup.BookmarkName);
                        else
                            result.DuplicateBookmarks.Add(dup.BookmarkName);
                        result.AbsorbedOwnerIds.Add("wdi:guid-" + dup.Guid);
                    }
                    return result;
                }
                // F4: first sighting. Do NOT create the bookmark here: record
                // the pending name only. CommitInlineOwnerResolution creates
                // it after the OCR plan succeeds, so a cancelled/empty/
                // failed run never mutates the document.
                var newGuid = Guid.NewGuid().ToString("N");
                result.PendingBookmark = ImageOverlayIdentity.CanonicalBookmarkPrefix + newGuid;
                result.CanonicalBookmark = ImageOverlayIdentity.CanonicalBookmarkPrefix + newGuid;
                result.OwnerId = "wdi:guid-" + newGuid;
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
            Document doc;
            try { doc = _word.ActiveDocument; }
            catch { return; }
            // N2: recover interrupted tightens FIRST. A staging bookmark
            // "OTTmp_" + 32 hex maps deterministically to its canonical
            // "OTImg_" + 32 hex. If the canonical is missing, the previous
            // run died between deleting the canonical and re-adding it:
            // restore the canonical at the staged (tight) range, then drop
            // the staging. If the canonical exists, the staging is a
            // post-success leftover: drop it. A staging whose restore fails
            // is kept for the next run; only bookmarks with a verifiable
            // 38-char hex shape are touched, never a blind prefix sweep.
            // R1: rejected stagings (proven for this image but absorbed into
            // the chosen identity) are dropped WITHOUT resurrecting their
            // canonical -- restoring them would recreate the second owner
            // this fix removes. Returns the canonical names this call
            // restored, so the pending creation below does not re-add them.
            // This is independent of this image's resolution, so it runs
            // before the empty-canonical early return.
            var restored = RecoverInterruptedTightens(doc, resolution.RejectedStagings);
            if (string.IsNullOrEmpty(resolution.CanonicalBookmark)) return;
            // N2: a first sighting creates its owner bookmark only here,
            // after the plan succeeded. Cancelled/empty/failed runs return
            // before this point and never touch the document. Creation
            // failure is NOT swallowed: without a persisted bookmark this
            // run would report success while the next run mints a new owner
            // and orphans these shapes, so the image aborts with an
            // explicit error before any shape is created or removed.
            // R1: skip the Add when the recovery above already restored this
            // canonical from its adopted staging (same identity, already
            // persisted). When the restore failed, fall through and create
            // the canonical at the current image range -- with the adopted
            // guid, never a second identity; the kept staging is cleaned up
            // by the next run once the canonical is confirmed.
            if (!string.IsNullOrEmpty(resolution.PendingBookmark) && !restored.Contains(resolution.PendingBookmark))
            {
                try
                {
                    object rangeObject = imageRange;
                    doc.Bookmarks.Add(resolution.PendingBookmark, ref rangeObject);
                }
                catch { /* verified below */ }
                bool exists = false;
                try { exists = doc.Bookmarks[resolution.PendingBookmark] != null; }
                catch { exists = false; }
                if (!exists)
                    throw new InvalidOperationException("图片身份书签创建失败：无法持久保存图片身份，本图片的译文未写入，文档未被修改。");
            }
            // Re-tighten first: if this throws, the failure is visible to
            // the caller instead of being swallowed as a success, and the
            // canonical bookmark is never lost (see TightenBookmark).
            TightenBookmark(doc, resolution.CanonicalBookmark, imageRange);
            foreach (var duplicate in resolution.DuplicateBookmarks)
                try { doc.Bookmarks[duplicate].Delete(); } catch { }
        }

        // N2: repairs tightens interrupted by a crash between staging and
        // canonical restore. The staging name carries the full identity
        // ("OTTmp_" + 32 hex -> "OTImg_" + 32 hex), so the mapping is
        // verifiable from the name alone. Bookmarks that do not have the
        // exact 38-char hex shape are left alone, including user bookmarks
        // that merely share the prefix. R1: rejectedStagings are proven
        // stagings whose identity was absorbed into another owner during
        // resolve -- they are dropped WITHOUT resurrecting their canonical.
        // Returns the canonical names this call restored.
        private static HashSet<string> RecoverInterruptedTightens(Document doc, List<string> rejectedStagings)
        {
            var restored = new HashSet<string>(StringComparer.Ordinal);
            var rejected = new HashSet<string>(StringComparer.Ordinal);
            if (rejectedStagings != null)
                foreach (var r in rejectedStagings) rejected.Add(r);
            var stagings = new List<Tuple<string, string>>();
            var toDrop = new List<string>();
            try
            {
                foreach (Bookmark bookmark in doc.Bookmarks)
                {
                    string name;
                    try { name = bookmark.Name; }
                    catch { continue; }
                    if (string.IsNullOrEmpty(name)) continue;
                    if (!ImageOverlayIdentity.IsVerifiableStagingName(name)) continue;
                    if (rejected.Contains(name))
                    {
                        toDrop.Add(name);
                        continue;
                    }
                    stagings.Add(Tuple.Create(name, ImageOverlayIdentity.CanonicalNameForStaging(name)));
                }
            }
            catch { return restored; }
            foreach (var name in toDrop)
                try { doc.Bookmarks[name].Delete(); } catch { }
            foreach (var pair in stagings)
            {
                bool canonicalExists = false;
                try { canonicalExists = doc.Bookmarks[pair.Item2] != null; }
                catch { canonicalExists = false; }
                if (!canonicalExists)
                {
                    // The previous run died mid-migration: restore the
                    // canonical name at the staged (tight) range. If this
                    // fails too, the staging bookmark is the ONLY identity
                    // record left -- keep it; the next run retries.
                    bool ok = false;
                    try
                    {
                        Range stagedRange = doc.Bookmarks[pair.Item1].Range;
                        object r = stagedRange;
                        doc.Bookmarks.Add(pair.Item2, ref r);
                        ok = true;
                    }
                    catch { /* staging kept; retry next run */ }
                    if (!ok) continue;
                    restored.Add(pair.Item2);
                }
                try { doc.Bookmarks[pair.Item1].Delete(); } catch { }
            }
            return restored;
        }

        // N2: derives the staging name for a canonical bookmark. For names
        // we minted ("OTImg_" + 32 hex) the staging is "OTTmp_" + the same
        // 32 hex = 38 chars: deterministic and self-describing, so an
        // interrupted tighten can be recovered without any side channel
        // (see RecoverInterruptedTightens and the R1 adoption in
        // ResolveInlineOwnerId), and it can never parse as an OTImg_
        // owner. For foreign bookmark shapes inside our prefix no
        // verifiable mapping is possible; keep the old short random
        // staging (14 chars, never an OTImg_ owner).
        private static string StagingNameFor(string canonicalName)
        {
            var mapped = ImageOverlayIdentity.StagingNameForCanonical(canonicalName);
            if (!string.IsNullOrEmpty(mapped)) return mapped;
            return "OTTmp_" + Guid.NewGuid().ToString("N").Substring(0, 8);
        }

        // F3/N2: re-tightens an expanded bookmark to the image range without
        // ever losing the only bookmark. Word limits bookmark names to 40
        // characters
        // (https://learn.microsoft.com/en-us/office/vba/api/word.bookmarks.add).
        // The tight range is staged FIRST; the canonical bookmark is deleted
        // only after staging succeeded. A tightening failure is never
        // swallowed: it throws a descriptive error so the caller cannot
        // mistake it for success. If the canonical re-add AND its restore
        // both fail, the staging bookmark is KEPT as the only remaining
        // identity record (the next run recovers it); it is never deleted
        // on that path.
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
            if (bStart == start && bEnd == end) return; // already tight
            var tmp = StagingNameFor(name);
            try
            {
                object tightRange = imageRange;
                doc.Bookmarks.Add(tmp, ref tightRange);
            }
            catch
            {
                // Staging failed: the canonical bookmark is untouched.
                throw new InvalidOperationException("图片书签收紧失败：无法创建临时书签，原书签保持不变。");
            }
            Range newRange;
            try { newRange = doc.Bookmarks[tmp].Range; }
            catch
            {
                try { doc.Bookmarks[tmp].Delete(); } catch { }
                throw new InvalidOperationException("图片书签收紧失败：无法读取临时书签范围，原书签保持不变。");
            }
            try { doc.Bookmarks[name].Delete(); }
            catch
            {
                try { doc.Bookmarks[tmp].Delete(); } catch { }
                throw new InvalidOperationException("图片书签收紧失败：无法删除原书签，原书签保持不变。");
            }
            try
            {
                object r = newRange;
                doc.Bookmarks.Add(name, ref r);
            }
            catch
            {
                // The old canonical is gone and the re-add failed: attempt
                // to restore the canonical name at the staged tight range.
                // If the restore ALSO fails, the staging bookmark is the
                // only identity record left and MUST be kept so the next
                // run can recover it (RecoverInterruptedTightens).
                // Deleting it here would lose the owner permanently.
                bool restored = false;
                try
                {
                    object r2 = newRange;
                    doc.Bookmarks.Add(name, ref r2);
                    restored = true;
                }
                catch { /* staging kept; the next run retries the recovery */ }
                if (restored)
                {
                    try { doc.Bookmarks[tmp].Delete(); } catch { }
                }
                throw new InvalidOperationException("图片书签收紧失败：原书签未能恢复，请重试翻译。");
            }
            try { doc.Bookmarks[tmp].Delete(); } catch { }
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

        private void AddInlineImage(InlineShape shape, List<ImageTarget> result)
        {
            if (shape.Type != WdInlineShapeType.wdInlineShapePicture && shape.Type != WdInlineShapeType.wdInlineShapeLinkedPicture) return;
            var range = shape.Range.Duplicate; var left = Convert.ToSingle(range.Information[WdInformation.wdHorizontalPositionRelativeToPage]); var top = Convert.ToSingle(range.Information[WdInformation.wdVerticalPositionRelativeToPage]);
            // S2: select the image explicitly, then copy via the Selection
            // (the same select-then-copy the floating path already uses). A
            // menu run showed the second inline picture capturing fine only
            // when it was the current selection, while Range.CopyAsPicture()
            // on the cached range silently left no image on the clipboard in
            // the multi-image flow. The user's selection is saved and
            // restored best-effort; a stale range now fails visibly at
            // Select() instead of silently copying nothing.
            var target = new ImageTarget(range, left, top, shape.Width, shape.Height, attempt => SelectAndCopyAsPicture(range, attempt));
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
            var target = new ImageTarget(anchor, shape.Left, shape.Top, shape.Width, shape.Height, _ => { object replace = true; shape.Select(ref replace); shape.Anchor.Application.Selection.CopyAsPicture(); });
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

        // D1: the capture attempt is passed explicitly by the capture loop
        // (1-based, same numbering as the "attempt N/3" message). It is NOT
        // counted here: a clear-phase failure retries without invoking the
        // delegate at all, so a local counter would undercount.
        // S2: select-then-copy for inline pictures. Saves the user's current
        // selection, selects the image range, copies via Selection, then
        // restores the selection best-effort. Runs synchronously on the UI
        // thread inside the capture's copy phase (before the read poll), so
        // message-pump reentrancy during the later poll cannot disturb the
        // copy that already happened. Cancellation is honored by the
        // capture loop around this call; the select/copy itself is fast.
        // P2: the saved selection is a duplicated Range, not start/end
        // numbers. The duplicate carries the original Document and Story
        // identity, so a selection in a header/footer/footnote restores to
        // the same story instead of being rebuilt in the main text via
        // ActiveDocument.Range (which only addresses the main story).
        private void SelectAndCopyAsPicture(Range range, int attempt)
        {
            Range? savedSelection = null;
            Selection? beforeSelection = null;
            try
            {
                var selection = _word.Selection;
                if (selection != null)
                {
                    savedSelection = selection.Range.Duplicate;
                    beforeSelection = selection;
                }
            }
            catch { savedSelection = null; }
            // S1: per-attempt condition snapshot BEFORE Select, so the next
            // review can compare the select/copy preconditions of a failing
            // attempt against succeeding ones (first image vs second image,
            // attempt 1 vs attempt 2).
            string beforeState = SnapshotSelectionState(beforeSelection);
            try
            {
                // S1: keep Select and CopyAsPicture as separate guarded
                // steps. 0x800A11FD ("command not available") is NOT
                // clipboard-busy and must never be retried blanket-style;
                // the diagnosis below records which step failed and the
                // target vs. actual selection identity, so the next review
                // can tell whether Select() actually landed on the image.
                // The original exception is rethrown unchanged: the capture
                // loop's transient/no-image classification must keep seeing
                // the real HResult.
                try { range.Select(); }
                catch (Exception ex) { LogCopyDiagnosis(range, "select", attempt, beforeState, null, ex); throw; }
                string afterSelectState = SnapshotSelectionState(SafeGet<Selection?>(() => _word.Selection, null));
                try { _word.Selection.CopyAsPicture(); }
                catch (Exception ex) { LogCopyDiagnosis(range, "copy", attempt, beforeState, afterSelectState, ex); throw; }
                // S1: also record the succeeding attempts' full
                // before/after conditions; without them a failing attempt
                // cannot be compared against anything.
                LogCopyDiagnosis(range, "copied", attempt, beforeState, afterSelectState, null);
            }
            finally
            {
                // Immediate restore after the copy only; the authoritative
                // restore happens at the end of the whole task (see
                // TranslateAsync), because overlay/box/bookmark work moves
                // the selection again afterwards.
                TryRestoreSelection(savedSelection);
            }
        }

        private Range? SaveSelectionDuplicate()
        {
            try { return _word.Selection?.Range.Duplicate; }
            catch { return null; }
        }

        // R1/P2: best-effort selection restore. The duplicate carries the
        // original Document/Story identity; never rebuild from
        // ActiveDocument by numbers (main-story only). On failure, log a
        // content-free diagnosis instead of guessing across
        // stories/documents.
        private void TryRestoreSelection(Range? savedSelection)
        {
            if (savedSelection == null) return;
            try { savedSelection.Select(); }
            catch (Exception ex)
            {
                try
                {
                    ImageOverlayDiagnostics.LogCaptureFailure("Word",
                        "selection_restore_failed error=" + ex.GetType().Name);
                }
                catch { }
            }
        }

        // N1: one-shot restore ticket for the task-end restore. State is
        // claimed with Interlocked so exactly one of "callback started" /
        // "wait gave up" wins; an expired ticket makes a late callback a
        // no-op instead of clobbering the user's newer selection.
        private sealed class RestoreTicket
        {
            public const int Pending = 0;
            public const int Running = 1;
            public const int Expired = 2;
            public int State = Pending;
            public readonly Range? Selection;
            public readonly int Generation;
            public readonly long PostedTicks;
            public RestoreTicket(Range? selection, int generation, long postedTicks)
            {
                Selection = selection;
                Generation = generation;
                PostedTicks = postedTicks;
            }
        }

        private const int SelectionRestoreWaitMs = 2000;

        // N1: shared across task instances (a new service is created per
        // task). Bumped at every task entry on the UI thread; a restore
        // ticket from an older generation must never overwrite a newer
        // task's (or the user's own) selection.
        private static int _selectionGeneration;

        // S1: metadata-only snapshot of a restore target range. Never throws.
        private string SnapshotRangeState(Range? range)
        {
            if (range == null) return "none";
            return "doc=" + SafeGet(() => (range.Parent as Document)?.Name ?? "?", "?")
                + " story=" + SafeGet(() => ((int)range.StoryType).ToString(), "?")
                + " range=" + SafeGet(() => range.Start, -1) + "-" + SafeGet(() => range.End, -1);
        }

        // S1: does the live selection actually match the restore target
        // (document, story, start, end)? Best-effort; never throws.
        private bool SelectionMatchesRange(Selection? sel, Range target)
        {
            try
            {
                if (sel == null) return false;
                Selection s = sel;
                string selDoc = SafeGet(() => (s.Range.Parent as Document)?.Name ?? "?", "?");
                string tgtDoc = SafeGet(() => (target.Parent as Document)?.Name ?? "?", "?");
                if (!string.Equals(selDoc, tgtDoc, StringComparison.Ordinal)) return false;
                if (SafeGet(() => (int)s.Range.StoryType, -1) != SafeGet(() => (int)target.StoryType, -1)) return false;
                if (SafeGet(() => s.Range.Start, -1) != SafeGet(() => target.Start, -1)) return false;
                if (SafeGet(() => s.Range.End, -1) != SafeGet(() => target.End, -1)) return false;
                return true;
            }
            catch { return false; }
        }

        // S1 error-path gap: the task-end restore can run Select() without
        // throwing yet leave the selection on the wrong story/range
        // (observed: dual-image 0x800A11FD left MainText 13-14 instead of
        // the header 0-6, with no restore log line at all). Record target
        // vs actual before and after, plus whether they match, so the next
        // review can tell a silent no-op from a real restore. Metadata
        // only; never throws; does NOT re-post anything (that would revive
        // N1).
        private void RestoreSelectionWithDiag(Range? target, string scope)
        {
            string targetState = SnapshotRangeState(target);
            string beforeState = SnapshotSelectionState(SafeGet<Selection?>(() => _word.Selection, null));
            string outcome;
            try
            {
                if (target == null) outcome = "no-target";
                else { target.Select(); outcome = "select-ok"; }
            }
            catch (Exception ex) { outcome = "throw:" + ex.GetType().Name; }
            Selection? afterSel = SafeGet<Selection?>(() => _word.Selection, null);
            string afterState = SnapshotSelectionState(afterSel);
            bool matched = target != null && SelectionMatchesRange(afterSel, target);
            try
            {
                ImageOverlayDiagnostics.LogCaptureFailure("Word",
                    "selection_restore_diag scope=" + scope
                    + " target=[" + targetState + "]"
                    + " before=[" + beforeState + "]"
                    + " after=[" + afterState + "]"
                    + " outcome=" + outcome
                    + " matched=" + (matched ? "1" : "0"));
            }
            catch { }
        }

        private void ClaimAndRestore(RestoreTicket ticket)
        {
            // N1: runs on the UI thread. Atomically claims the ticket, then
            // refuses a restore that is expired or superseded before
            // touching COM.
            if (Interlocked.CompareExchange(ref ticket.State,
                    RestoreTicket.Running, RestoreTicket.Pending) != RestoreTicket.Pending)
                return; // expired (wait already gave up) or double claim: no-op
            if (ticket.Generation != _selectionGeneration)
            {
                LogRestoreSkipped("superseded");
                return;
            }
            if (ElapsedMs(ticket.PostedTicks) > SelectionRestoreWaitMs)
            {
                // Backstop for the corner where the wait path itself never
                // ran: a restore landing this late would clobber newer
                // user intent, so it is dropped like an expired ticket.
                LogRestoreSkipped("stale");
                return;
            }
            // S1: instrumented restore (target vs actual before/after); see
            // RestoreSelectionWithDiag. Still best-effort and never throws.
            RestoreSelectionWithDiag(ticket.Selection, "task");
        }

        private static long ElapsedMs(long startTicks)
        {
            return (Stopwatch.GetTimestamp() - startTicks) * 1000 / Stopwatch.Frequency;
        }

        private static void LogRestoreSkipped(string reason)
        {
            try { ImageOverlayDiagnostics.LogCaptureFailure("Word", "selection_restore_skipped reason=" + reason); }
            catch { }
        }

        // S1 diagnostics for the Word select -> copy phase. Metadata only:
        // document/story names and numbers, selection identity, view type,
        // window caption/counts, inline-shape count, exception type/HResult/
        // message. No document text, no OCR text, no clipboard content. All
        // reads best-effort; never throws, and the caller rethrows the
        // original exception. ex == null marks a succeeding attempt
        // ("copied"): its before/after conditions are the comparison base
        // for failing attempts.
        // D1: attempt is the capture loop's real 1-based attempt number,
        // passed explicitly by the caller (same as the "attempt N/3"
        // message). It is not counted locally.
        private void LogCopyDiagnosis(Range range, string stage, int attempt, string beforeState, string? afterSelectState, Exception? ex)
        {
            try
            {
                var detail = "select_copy_diag stage=" + stage
                    + " attempt=" + attempt
                    + " targetDoc=" + SafeGet(() => (range.Parent as Document)?.Name, "?")
                    + " targetStory=" + SafeGet(() => (int)range.StoryType, -1)
                    + " targetRange=" + SafeGet(() => range.Start, -1) + "-" + SafeGet(() => range.End, -1)
                    + " before=[" + beforeState + "]"
                    + (afterSelectState != null ? " afterSelect=[" + afterSelectState + "]" : "")
                    + " inlineShapes=" + SafeGet(() => _word.ActiveDocument.InlineShapes.Count, -1)
                    + (ex != null
                        ? " error=" + ex.GetType().Name
                            + " hresult=0x" + SafeGet(() => Marshal.GetHRForException(ex).ToString("X8"), "?")
                            + " msg=" + SafeGet(() => FlattenMessage(ex.Message), "?")
                        : " ok=1");
                ImageOverlayDiagnostics.LogCaptureFailure("Word", detail);
            }
            catch { }
        }

        // S1: compact metadata-only snapshot of the current selection plus
        // the active window state. Used for the before-select / after-select
        // per-attempt comparison. Never throws.
        private string SnapshotSelectionState(Selection? sel)
        {
            try
            {
                string doc = "?", story = "?", range = "?-?", type = "?";
                if (sel != null)
                {
                    Selection s = sel;
                    doc = SafeGet(() => (s.Range.Parent as Document)?.Name ?? "?", "?");
                    story = SafeGet(() => ((int)s.Range.StoryType).ToString(), "?");
                    range = SafeGet(() => s.Range.Start, -1) + "-" + SafeGet(() => s.Range.End, -1);
                    type = SafeGet(() => s.Type.ToString(), "?");
                }
                return "doc=" + doc + " story=" + story + " range=" + range + " type=" + type
                    + " view=" + SafeGet(() => _word.ActiveWindow.View.Type.ToString(), "?")
                    + " winCap=" + SafeGet(() => FlattenMessage(_word.ActiveWindow.Caption), "?")
                    + " wins=" + SafeGet(() => _word.Windows.Count, -1)
                    + " docs=" + SafeGet(() => _word.Documents.Count, -1);
            }
            catch { return "?"; }
        }

        private static string FlattenMessage(string message)
        {
            if (string.IsNullOrEmpty(message)) return "?";
            return message.Replace('\r', ' ').Replace('\n', ' ');
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
            private readonly Action<int> _selectOrCopy;
            public ImageTarget(Range anchor, float left, float top, float width, float height, Action<int> action) { Anchor = anchor; Left = left; Top = top; Width = width; Height = height; _selectOrCopy = action; }
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
            public void CopyAsPicture(int attempt) => _selectOrCopy(attempt);
        }
    }
}
