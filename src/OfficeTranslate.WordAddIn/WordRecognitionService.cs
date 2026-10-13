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
    // Separate orchestration preserves the translation path byte-for-byte.
    // Capture primitives, planner, real-layout checks and note splitter are shared.
    internal sealed partial class WordTranslationService
    {
        public bool HasSelectedOcrImage()
        {
            try { return ReadSelectedOcrImages().Count > 0; }
            catch (COMException) { return false; }
        }

        private List<ImageTarget> ReadSelectedOcrImages() => ReadSelectionImages();

        public async Task<TranslationTaskSummary> RecognizeSelectionAsync(TranslationSettings settings,
            CancellationToken token, Action<string> progress, OfficeUiDispatcher ui)
        {
            Summary = null;
            ui.Invoke(() =>
            {
                _taskGeneration = Interlocked.Increment(ref _selectionGeneration);
                _taskSelection = SaveSelectionDuplicate();
                _taskViewType = SafeGet(() => (int)_word.ActiveWindow.View.Type, -1);
                _taskWindowHwnd = SafeGet(() => _word.ActiveWindow.Hwnd, 0);
                _taskDocument = SafeGet(() => _word.ActiveDocument, (Document?)null);
            });
            var images = ui.Invoke(ReadSelectedOcrImages);
            if (images.Count == 0) throw new InvalidOperationException("请先选中一张图片");
            ImageOverlayDiagnostics.LogVersion("Word-OCR",
                typeof(WordTranslationService).Assembly.GetName().Version?.ToString() ?? "unknown",
                typeof(WordTranslationService).Module.ModuleVersionId.ToString());
            using (var client = new TranslationClient())
            {
                var summary = Summary = new TranslationTaskSummary(images.Count);
                var captured = await ImagePreCapture.CaptureAllAsync(images,
                    (image, current, remaining) => CaptureImageAsync(image, current, token, ui, remaining), token);
                for (var i = 0; i < captured.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    progress($"OfficeTranslate：正在 OCR 识别 {i + 1}/{images.Count}");
                    await RecognizeSelectedImageAsync(captured[i].Item1, i + 1, captured[i].Item2,
                        client, settings, summary, token, ui);
                }
                return summary;
            }
        }

        private async Task RecognizeSelectedImageAsync(ImageTarget image, int imageNumber, byte[] bytes, TranslationClient client, TranslationSettings settings, TranslationTaskSummary summary, CancellationToken token, OfficeUiDispatcher ui)
        {
            var resolved = await ui.InvokeAsync(() =>
            {
                InlineOwnerResolution? inlineResolution = null;
                string owner;
                if (image.Kind == "Inline")
                {
                    inlineResolution = ResolveInlineOwnerId(image.Anchor);
                    owner = inlineResolution.OwnerId;
                }
                else
                {
                    owner = ImageOverlayIdentity.ForNamedShape("wd", image.ShapeName, bytes);
                }
                return Tuple.Create(owner, inlineResolution);
            });
            var ownerId = resolved.Item1;
            var inlineResolution = resolved.Item2;
            var hasPixels = PngDimensions.TryRead(bytes, out var pixelWidth, out var pixelHeight);
            var recognition = await client.RecognizeImageAsync(bytes, settings, token, "Word", image.CaptureSource, image.CaptureReason);
            var regions = recognition.Regions;
            if (regions.Count == 0) { summary.RecordImageRecognition(recognition); return; }
            var originIsSentinel = IsAlignmentSentinel(image.Left) || IsAlignmentSentinel(image.Top);
            var sentinelReason = "图片使用对齐定位（如居中/靠右），无法解析实际显示坐标，已降级为旁注。";
            var plans = new List<Tuple<ImageTranslationRegion, PlannedOverlay, string>>();
            foreach (var region in regions)
            {
                var overlayText = region.Source;
                PlannedOverlay plan;
                if (region.RequiresReview)
                    plan = new PlannedOverlay(ImageOverlayVerdict.SideNote, region.ReviewReason, 0F, 0F, 0F, 0F, 0F);
                else if (originIsSentinel)
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
            token.ThrowIfCancellationRequested();
            await ui.InvokeAsync(() =>
            {
                token.ThrowIfCancellationRequested();
                if (inlineResolution != null)
                    CommitInlineOwnerResolution(image.Anchor, inlineResolution);
                OcrRemovePreviousResults(image, ownerId);
                if (inlineResolution != null)
                    foreach (var absorbed in inlineResolution.AbsorbedOwnerIds)
                        OcrRemovePreviousResults(image, absorbed);

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
                    if (!OcrPlaceOverlay(image, plan, overlayText, ownerId, out frameFailure))
                    {
                        LogOverlay(image, item.Item1, pixelWidth, pixelHeight,
                            new PlannedOverlay(ImageOverlayVerdict.SideNote, frameFailure,
                                plan.Left, plan.Top, plan.Width, plan.Height, plan.FontSize),
                            overlayText, ownerId, "writeback");
                        noteEntries.Add(frameFailure + "\r" + overlayText);
                    }
                }
                if (noteEntries.Count > 0)
                {
                    float noteLeft, noteTop;
                    WdRelativeHorizontalPosition noteRelH;
                    WdRelativeVerticalPosition noteRelV;
                    if (originIsSentinel)
                    {
                        if (TryGetAnchorPagePosition(image, out var anchorX, out var anchorY))
                        {
                            noteLeft = anchorX;
                            noteTop = anchorY + 6F;
                        }
                        else
                        {
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
                    summary.RecordImageNeedsReview();
                    ImageOverlayNotes.Place(noteEntries, ImageOverlayNotes.NoteWidth(image.Width), noteTop,
                        (chunk, top) => OcrPlaceCombinedNote(image, ownerId, ImageRecognitionText.NoteText(chunk),
                            noteLeft, top, noteRelH, noteRelV),
                        (chunk, error) => summary.RecordImageNoteFailure(ownerId, ImageRecognitionText.NoteText(chunk), error.Message));
                }
            });
            summary.RecordImageRecognition(recognition);
        }

        private float OcrPlaceCombinedNote(ImageTarget image, string ownerId, string noteText,
            float noteLeft, float noteTop,
            WdRelativeHorizontalPosition relH, WdRelativeVerticalPosition relV)
        {
            var noteWidth = ImageOverlayNotes.NoteWidth(image.Width);
            var noteHeight = Math.Min(ImageOverlayLayout.MaxNoteHeightPt,
                ImageOverlayLayout.EstimateNoteHeight(noteWidth - 4F, noteText, ImageOverlayNotes.NoteFontSize));
            object anchor = image.Anchor.Duplicate;
            var note = TaskDocument().Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                noteLeft, noteTop, noteWidth, noteHeight, ref anchor);
            try
            {
                note.AlternativeText = ImageRecognitionTags.NoteFor(ownerId);
                note.TextFrame.AutoSize = (int)Office.MsoAutoSize.msoAutoSizeNone;
                try
                {
                    note.RelativeHorizontalPosition = relH;
                    note.RelativeVerticalPosition = relV;
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    try { note.Delete(); } catch { }
                    throw new InvalidOperationException(
                        "已完成图片识别，但无法设置原文旁注的定位参考系，已删除错位旁注避免误导。原文（截断）：" +
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
                note.TextFrame.WordWrap = (int)Office.MsoTriState.msoTrue;
                note.TextFrame.TextRange.Text = noteText;
                note.TextFrame.TextRange.Font.Size = ImageOverlayNotes.NoteFontSize;
                GrowNoteToFit(note);
                return note.Height;
            }
            catch
            {
                try { note.Delete(); } catch { }
                throw;
            }
        }

        private void OcrRemovePreviousResults(ImageTarget image, string ownerId)
        {
            var doomed = new List<Shape>();
            foreach (Shape shape in TaskDocument().Shapes)
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

        private bool OcrPlaceOverlay(ImageTarget image, PlannedOverlay plan, string text, string ownerId, out string failureReason)
        {
            failureReason = string.Empty;
            object anchor = image.Anchor.Duplicate;
            var overlay = TaskDocument().Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                image.Left, image.Top, plan.Width, plan.Height, ref anchor);
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
            overlay.AlternativeText = ImageRecognitionTags.OverlayFor(ownerId);
            overlay.Left = image.Left + plan.Left;
            overlay.Top = image.Top + plan.Top;
            overlay.Width = plan.Width;
            overlay.Height = plan.Height;
            overlay.WrapFormat.Type = WdWrapType.wdWrapFront;
            overlay.Fill.Visible = Office.MsoTriState.msoTrue;
            overlay.Fill.ForeColor.RGB = 0xFFFFFF;
            overlay.Fill.Transparency = 0F;
            overlay.Line.Visible = Office.MsoTriState.msoFalse;
            overlay.TextFrame.MarginLeft = 2; overlay.TextFrame.MarginRight = 2;
            overlay.TextFrame.MarginTop = 1; overlay.TextFrame.MarginBottom = 1;
            overlay.TextFrame.WordWrap = (int)Office.MsoTriState.msoTrue;
            overlay.TextFrame.AutoSize = (int)Office.MsoAutoSize.msoAutoSizeNone;
            overlay.TextFrame.TextRange.Text = text;
            overlay.TextFrame.TextRange.Font.Size = plan.FontSize;
            if (!ShrinkFontToFit(overlay, plan.FontSize, out failureReason))
            {
                try { overlay.Delete(); } catch { }
                return false;
            }
            if (!MatchesPlan(overlay, image.Left + plan.Left, image.Top + plan.Top, plan.Width, plan.Height))
            {
                try { overlay.Delete(); } catch { }
                failureReason = "覆盖框几何被宿主改动，已删除并降级为旁注。";
                return false;
            }
            return true;
        }
    }
}
