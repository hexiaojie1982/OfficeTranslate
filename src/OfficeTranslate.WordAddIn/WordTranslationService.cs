using Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
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
            var regions = await client.TranslateImageAsync(bytes, settings, token);
            if (regions.Count == 0) { summary.RecordImage(false); return; }
            summary.RecordImage(true);

            // Re-translating the same image updates its overlays instead of stacking new ones.
            RemoveOverlaysForImage(image);

            var needsReview = false;
            foreach (var region in regions)
            {
                var bilingualImage = settings.BilingualMode && !string.IsNullOrWhiteSpace(region.Source);
                var overlayText = bilingualImage ? region.Source + "\r" + region.Translation : region.Translation;
                var plan = hasPixels
                    ? ImageOverlayPlanner.Plan(
                        region.X1, region.Y1, region.X2, region.Y2,
                        pixelWidth, pixelHeight, image.Width, image.Height, overlayText,
                        image.Rotation, image.FlipHorizontal, image.FlipVertical)
                    : new PlannedOverlay(ImageOverlayVerdict.SideNote,
                        "无法读取捕获图像的像素尺寸，坐标无法可靠换算。", 0F, 0F, 0F, 0F, 0F);
                LogOverlay(image, region, pixelWidth, pixelHeight, plan, overlayText);
                if (plan.Verdict == ImageOverlayVerdict.SideNote)
                {
                    PlaceSideNote(image, plan.Reason, overlayText);
                    needsReview = true;
                    continue;
                }
                PlaceOverlay(image, plan, overlayText);
            }
            if (needsReview) summary.RecordImageNeedsReview();
        }

        private void PlaceOverlay(ImageTarget image, PlannedOverlay plan, string text)
        {
            object anchor = image.Anchor.Duplicate;
            var overlay = _word.ActiveDocument.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                image.Left, image.Top, plan.Width, plan.Height, ref anchor);
            overlay.AlternativeText = "OfficeTranslateOCR";
            // Inherit the source image's reference frame. A floating picture's
            // own Left/Top may be paragraph/margin-relative rather than
            // page-relative, so forcing page-relative here was the offset bug.
            // Geometry is assigned AFTER the frame so Office interprets
            // Left/Top in the correct coordinate system.
            try
            {
                overlay.RelativeHorizontalPosition = image.RelativeHorizontalPosition;
                overlay.RelativeVerticalPosition = image.RelativeVerticalPosition;
            }
            catch (System.Runtime.InteropServices.COMException) { }
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
            overlay.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
            overlay.TextFrame.TextRange.Text = text;
            // Word does not consistently support msoAutoSizeTextToFitShape. Some
            // desktop builds reject it with "value out of range", so the font
            // size is fitted inside the region by ImageOverlayPlanner instead
            // of growing the box downward.
            overlay.TextFrame.TextRange.Font.Size = plan.FontSize;
        }

        private void PlaceSideNote(ImageTarget image, string reason, string text)
        {
            // A side-note explicitly does NOT claim positional coverage: it is
            // placed below the image, in the image's own reference frame.
            var noteWidth = Math.Min(420F, Math.Max(160F, image.Width));
            var noteText = "【图片译文待检查】" + reason + "\r" + text;
            var noteHeight = ImageOverlayLayout.EstimateNoteHeight(noteWidth, noteText, 9F);
            var noteLeft = image.Left;
            var noteTop = image.Top + image.Height + 6F;
            RemoveNotesInZone(image, noteLeft, noteTop, noteWidth, noteHeight);
            object anchor = image.Anchor.Duplicate;
            var note = _word.ActiveDocument.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                noteLeft, noteTop, noteWidth, noteHeight, ref anchor);
            note.AlternativeText = "OfficeTranslateOCR-Note";
            try
            {
                note.RelativeHorizontalPosition = image.RelativeHorizontalPosition;
                note.RelativeVerticalPosition = image.RelativeVerticalPosition;
            }
            catch (System.Runtime.InteropServices.COMException) { }
            note.Left = noteLeft; note.Top = noteTop; note.Width = noteWidth; note.Height = noteHeight;
            note.WrapFormat.Type = WdWrapType.wdWrapFront;
            note.Fill.Visible = Office.MsoTriState.msoTrue;
            note.Fill.ForeColor.RGB = 0xE1FFFF; // light yellow (BGR)
            note.Fill.Transparency = 0F;
            note.Line.Visible = Office.MsoTriState.msoFalse;
            note.TextFrame.MarginLeft = 4; note.TextFrame.MarginRight = 4;
            note.TextFrame.MarginTop = 3; note.TextFrame.MarginBottom = 3;
            note.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
            note.TextFrame.TextRange.Text = noteText;
            note.TextFrame.TextRange.Font.Size = 9F;
        }

        private void RemoveOverlaysForImage(ImageTarget image)
        {
            var doomed = new List<Shape>();
            foreach (Shape shape in _word.ActiveDocument.Shapes)
            {
                string alt;
                try { alt = shape.AlternativeText; } catch { continue; }
                if (alt != "OfficeTranslateOCR") continue;
                if (!SameAnchorAndFrame(shape, image)) continue;
                var centerX = shape.Left + shape.Width / 2F;
                var centerY = shape.Top + shape.Height / 2F;
                if (centerX >= image.Left - 8F && centerX <= image.Left + image.Width + 8F &&
                    centerY >= image.Top - 8F && centerY <= image.Top + image.Height + 8F)
                    doomed.Add(shape);
            }
            foreach (var shape in doomed)
            {
                try { shape.Delete(); } catch { }
            }
        }

        private void RemoveNotesInZone(ImageTarget image, float left, float top, float width, float height)
        {
            var doomed = new List<Shape>();
            foreach (Shape shape in _word.ActiveDocument.Shapes)
            {
                string alt;
                try { alt = shape.AlternativeText; } catch { continue; }
                if (alt != "OfficeTranslateOCR-Note") continue;
                if (!SameAnchorAndFrame(shape, image)) continue;
                if (shape.Left < left + width && shape.Left + shape.Width > left &&
                    shape.Top < top + height && shape.Top + shape.Height > top)
                    doomed.Add(shape);
            }
            foreach (var shape in doomed)
            {
                try { shape.Delete(); } catch { }
            }
        }

        private static bool SameAnchorAndFrame(Shape shape, ImageTarget image)
        {
            try
            {
                var anchor = shape.Anchor;
                return anchor.Start == image.Anchor.Start && anchor.End == image.Anchor.End &&
                    shape.RelativeHorizontalPosition == image.RelativeHorizontalPosition &&
                    shape.RelativeVerticalPosition == image.RelativeVerticalPosition;
            }
            catch { return false; }
        }

        private void LogOverlay(ImageTarget image, ImageTranslationRegion region,
            int pixelWidth, int pixelHeight, PlannedOverlay plan, string text)
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
                TextLength = text == null ? 0 : text.Length
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
            if (shape.AlternativeText == "OfficeTranslateOCR") return;
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
