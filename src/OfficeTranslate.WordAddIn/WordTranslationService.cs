using Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;
using WordApplication = Microsoft.Office.Interop.Word.Application;

namespace OfficeTranslate.WordAddIn
{
    internal sealed class WordTranslationService
    {
        private readonly WordApplication _word;
        public WordTranslationService(WordApplication word) => _word = word;

        public async Task TranslateAsync(bool wholeDocument, bool bilingual, TranslationSettings settings, CancellationToken token, Action<string> progress)
        {
            var targets = wholeDocument ? ReadDocumentParagraphs() : ReadSelection();
            var images = settings.ImageOcrEnabled ? (wholeDocument ? ReadDocumentImages() : ReadSelectionImages()) : new List<ImageTarget>();
            if (targets.Count == 0 && images.Count == 0) throw new InvalidOperationException(wholeDocument ? "文档中没有可翻译的正文或图片。" : "请先选择需要翻译的文字或图片。");
            using (var client = new TranslationClient())
            {
                var total = targets.Count + images.Count;
                // Word represents a selected inline picture with a non-printing object
                // character. Image-only selections are filtered below, and images are
                // handled before ordinary text so OCR is always the first real request.
                for (var i = 0; i < images.Count; i++)
                {
                    token.ThrowIfCancellationRequested(); var current = i + 1;
                    progress($"OfficeTranslate：正在识别图片 {current}/{total}");
                    await TranslateImageAsync(images[i], client, settings, token);
                    progress($"OfficeTranslate：已完成 {current}/{total}");
                }
                for (var i = 0; i < targets.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    var current = images.Count + i + 1;
                    progress($"OfficeTranslate：正在翻译 {current}/{total}");
                    var sourceText = CleanWordObjectMarkers(targets[i].Text);
                    if (!HasTranslatableText(sourceText)) continue;
                    var text = await client.TranslateAsync(sourceText, settings, token);
                    token.ThrowIfCancellationRequested();

                    var translatedText = text.TrimEnd('\r', '\a');
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
            }
        }

        private async Task TranslateImageAsync(ImageTarget image, TranslationClient client, TranslationSettings settings, CancellationToken token)
        {
            image.CopyAsPicture();
            System.Drawing.Image? clipboardImage = null;
            for (var attempt = 0; attempt < 10 && clipboardImage == null; attempt++)
            {
                if (Clipboard.ContainsImage()) clipboardImage = Clipboard.GetImage();
                if (clipboardImage == null) Thread.Sleep(50);
            }
            if (clipboardImage == null) throw new InvalidOperationException("无法从 Word 图片获取可识别图像。");
            byte[] bytes;
            using (clipboardImage) using (var stream = new MemoryStream()) { clipboardImage.Save(stream, ImageFormat.Png); bytes = stream.ToArray(); }
            var regions = await client.TranslateImageAsync(bytes, settings, token);
            foreach (var region in regions)
            {
                var left = image.Left + image.Width * region.X1 / 1000F; var top = image.Top + image.Height * region.Y1 / 1000F;
                var width = Math.Max(24F, image.Width * (region.X2 - region.X1) / 1000F); var originalHeight = Math.Max(10F, image.Height * (region.Y2 - region.Y1) / 1000F);
                var bilingualImage = settings.BilingualMode && !string.IsNullOrWhiteSpace(region.Source);
                var overlayText = bilingualImage ? region.Source + "\r" + region.Translation : region.Translation;
                var layout = ImageOverlayLayout.Calculate(width, originalHeight, overlayText);
                object anchor = image.Anchor.Duplicate;
                var overlay = _word.ActiveDocument.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, layout.Height, ref anchor);
                overlay.AlternativeText = "OfficeTranslateOCR"; overlay.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage; overlay.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                overlay.WrapFormat.Type = WdWrapType.wdWrapFront; overlay.Fill.Visible = Office.MsoTriState.msoTrue; overlay.Fill.ForeColor.RGB = 0xFFFFFF; overlay.Fill.Transparency = 0.08F; overlay.Line.Visible = Office.MsoTriState.msoFalse;
                overlay.TextFrame.MarginLeft = 2; overlay.TextFrame.MarginRight = 2; overlay.TextFrame.MarginTop = 1; overlay.TextFrame.MarginBottom = 1;
                overlay.TextFrame.TextRange.Text = overlayText;
                // Word does not consistently support msoAutoSizeTextToFitShape. Some
                // desktop builds reject it with "value out of range", so calculate a
                // conservative font size without using that COM property.
                overlay.TextFrame.TextRange.Font.Size = layout.FontSize;
            }
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
            result.Add(new ImageTarget(range, left, top, shape.Width, shape.Height, () => range.CopyAsPicture()));
        }

        private static void AddFloatingImage(Shape shape, List<ImageTarget> result)
        {
            if (shape.AlternativeText == "OfficeTranslateOCR") return;
            if (shape.Type == Office.MsoShapeType.msoGroup) { for (var i = 1; i <= shape.GroupItems.Count; i++) AddFloatingImage(shape.GroupItems[i], result); return; }
            if (shape.Type != Office.MsoShapeType.msoPicture && shape.Type != Office.MsoShapeType.msoLinkedPicture) return;
            var anchor = shape.Anchor.Duplicate;
            result.Add(new ImageTarget(anchor, shape.Left, shape.Top, shape.Width, shape.Height, () => { object replace = true; shape.Select(ref replace); shape.Anchor.Application.Selection.CopyAsPicture(); }));
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
                if (char.IsControl(character) && character != '\r' && character != '\n' && character != '\t') continue;
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
            public void CopyAsPicture() => _selectOrCopy();
        }
    }
}
