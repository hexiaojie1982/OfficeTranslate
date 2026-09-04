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

        public async Task TranslateAsync(bool wholeSheet, TranslationSettings settings, CancellationToken token, Action<string> progress)
        {
            var sheet = _excel.ActiveSheet as Excel.Worksheet ?? throw new InvalidOperationException("请先打开工作表。");
            var targets = wholeSheet ? ReadSheetTargets(sheet) : ReadSelectionTargets();
            var images = settings.ImageOcrEnabled ? (wholeSheet ? ReadSheetImages(sheet) : ReadSelectionImages()) : new List<Excel.Shape>();
            if (targets.Count == 0 && images.Count == 0) throw new InvalidOperationException("没有找到可翻译的单元格、图形文本或图片。");

            using (var client = new TranslationClient())
            {
                var total = targets.Count + images.Count;
                for (var i = 0; i < targets.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    progress($"OfficeTranslate：正在翻译 {i + 1}/{total}");
                    var translated = (await client.TranslateAsync(targets[i].Text, settings, token)).TrimEnd('\r', '\n');
                    token.ThrowIfCancellationRequested();
                    targets[i].Write(settings.BilingualMode ? targets[i].Text + Environment.NewLine + translated : translated, settings.BilingualMode);
                    progress($"OfficeTranslate：已完成 {i + 1}/{total}");
                }
                for (var i = 0; i < images.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    var current = targets.Count + i + 1;
                    progress($"OfficeTranslate：正在识别图片 {current}/{total}");
                    await TranslateImageAsync(images[i], client, settings, token);
                    progress($"OfficeTranslate：已完成 {current}/{total}");
                }
            }
        }

        private async Task TranslateImageAsync(Excel.Shape image, TranslationClient client, TranslationSettings settings, CancellationToken token)
        {
            var sheet = _excel.ActiveSheet as Excel.Worksheet ?? throw new InvalidOperationException("无法确定图片所在的工作表。");
            byte[] imageBytes;
            try
            {
                // For a picture shape, Copy is more reliable than CopyPicture across
                // Excel builds and does not create the temporary chart that used to flash.
                image.Copy();
            }
            catch (COMException firstError)
            {
                try { image.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap); }
                catch (COMException) { throw new InvalidOperationException("Excel 无法复制所选图片，请重新选择图片后再试。", firstError); }
            }
            System.Drawing.Image? clipboardImage = null;
            for (var attempt = 0; attempt < 20 && clipboardImage == null; attempt++)
            {
                token.ThrowIfCancellationRequested();
                Application.DoEvents();
                if (Clipboard.ContainsImage()) clipboardImage = Clipboard.GetImage();
                if (clipboardImage == null) Thread.Sleep(50);
            }
            if (clipboardImage == null) throw new InvalidOperationException("无法从 Excel 图片获取可识别图像。");
            using (clipboardImage) using (var stream = new MemoryStream()) { clipboardImage.Save(stream, ImageFormat.Png); imageBytes = stream.ToArray(); }

            var regions = await client.TranslateImageAsync(imageBytes, settings, token);
            foreach (var region in regions)
            {
                var left = image.Left + image.Width * region.X1 / 1000F;
                var top = image.Top + image.Height * region.Y1 / 1000F;
                var width = Math.Max(12F, image.Width * (region.X2 - region.X1) / 1000F);
                var height = Math.Max(10F, image.Height * (region.Y2 - region.Y1) / 1000F);
                Excel.Shape? overlay = null;
                try
                {
                    overlay = sheet.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
                    overlay.AlternativeText = "OfficeTranslateOCR";
                    overlay.Fill.Visible = Office.MsoTriState.msoTrue; overlay.Fill.ForeColor.RGB = 0xFFFFFF; overlay.Fill.Transparency = 0.08F;
                    overlay.Line.Visible = Office.MsoTriState.msoFalse;
                    overlay.TextFrame2.MarginLeft = 2; overlay.TextFrame2.MarginRight = 2; overlay.TextFrame2.MarginTop = 1; overlay.TextFrame2.MarginBottom = 1;
                    var overlayText = settings.BilingualMode && !string.IsNullOrWhiteSpace(region.Source) ? region.Source + "\r" + region.Translation : region.Translation;
                    overlay.TextFrame2.TextRange.Text = overlayText;
                    overlay.TextFrame2.TextRange.Font.Size = CalculateOverlayFontSize(width, height, overlayText);
                }
                catch (COMException ex)
                {
                    try { overlay?.Delete(); } catch { }
                    throw new InvalidOperationException("Excel 已完成图片识别，但创建译文覆盖框失败：" + ex.Message, ex);
                }
            }
        }

        private static float CalculateOverlayFontSize(float width, float height, string text)
        {
            var lines = (text ?? string.Empty).Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var lineCount = Math.Max(1, lines.Length);
            var longest = 1;
            foreach (var line in lines) longest = Math.Max(longest, line.Length);
            var byHeight = Math.Max(1F, height - 2F) / (lineCount * 1.25F);
            var byWidth = Math.Max(1F, width - 4F) / (longest * 0.75F);
            return Math.Max(6F, Math.Min(24F, Math.Min(byHeight, byWidth)));
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
            var result = new List<Target>(); var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Excel.Range item in range.Cells)
            {
                var cell = item;
                if ((bool)cell.MergeCells) cell = (Excel.Range)cell.MergeArea.Cells[1, 1];
                var address = cell.Address[false, false, Excel.XlReferenceStyle.xlA1];
                if (!seen.Add(address) || (bool)cell.HasFormula) continue;
                var value = cell.Value2 as string;
                if (value != null && !string.IsNullOrWhiteSpace(value))
                {
                    var targetCell = cell;
                    result.Add(new Target(value, (text, bilingual) => { targetCell.Value2 = text; if (bilingual) targetCell.WrapText = true; }));
                }
            }
            return result;
        }

        private static void AddShape(Excel.Shape shape, List<Target> result)
        {
            if (shape.AlternativeText == "OfficeTranslateOCR") return;
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
