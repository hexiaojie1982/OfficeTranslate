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

        public async Task<TranslationTaskSummary> TranslateAsync(bool wholeSheet, TranslationSettings settings, CancellationToken token, Action<string> progress)
        {
            var sheet = _excel.ActiveSheet as Excel.Worksheet ?? throw new InvalidOperationException("请先打开工作表。");
            var targets = wholeSheet ? ReadSheetTargets(sheet) : ReadSelectionTargets();
            var images = settings.ImageOcrEnabled ? (wholeSheet ? ReadSheetImages(sheet) : ReadSelectionImages()) : new List<Excel.Shape>();
            if (targets.Count == 0 && images.Count == 0) throw new InvalidOperationException("没有找到可翻译的单元格、图形文本或图片。");

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
                        targets[i].Write(settings.BilingualMode ? targets[i].Text + Environment.NewLine + translated : translated, settings.BilingualMode);
                    }
                    progress($"OfficeTranslate：已完成 {i + 1}/{total}");
                }
                for (var i = 0; i < images.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    var current = targets.Count + i + 1;
                    progress($"OfficeTranslate：正在识别图片 {current}/{total}");
                    summary.RecordImage(await TranslateImageAsync(images[i], client, settings, token));
                    progress($"OfficeTranslate：已完成 {current}/{total}");
                }
                return summary;
            }
        }

        private async Task<bool> TranslateImageAsync(Excel.Shape image, TranslationClient client, TranslationSettings settings, CancellationToken token)
        {
            var sheet = _excel.ActiveSheet as Excel.Worksheet ?? throw new InvalidOperationException("无法确定图片所在的工作表。");
            // For a picture shape, Copy is more reliable than CopyPicture across
            // Excel builds and does not create the temporary chart that used to flash.
            // The copy itself runs inside CapturePng (after the user's clipboard is
            // saved), so a failed probe can never destroy clipboard content.
            // CapturePng also clears stale content first so a leftover image is
            // never mistaken for the shape.
            var imageBytes = ClipboardImageCapture.CapturePng(() =>
            {
                try { image.Copy(); }
                catch (COMException)
                {
                    try { image.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap); }
                    catch (COMException ex) { throw new InvalidOperationException("Excel 无法复制所选图片，请重新选择图片后再试。", ex); }
                }
            }, () => Application.DoEvents(), token);

            var regions = await client.TranslateImageAsync(imageBytes, settings, token);
            foreach (var region in regions)
            {
                var left = image.Left + image.Width * region.X1 / 1000F;
                var top = image.Top + image.Height * region.Y1 / 1000F;
                var width = Math.Max(24F, image.Width * (region.X2 - region.X1) / 1000F);
                var originalHeight = Math.Max(10F, image.Height * (region.Y2 - region.Y1) / 1000F);
                var overlayText = settings.BilingualMode && !string.IsNullOrWhiteSpace(region.Source) ? region.Source + "\r" + region.Translation : region.Translation;
                var layout = ImageOverlayLayout.Calculate(width, originalHeight, overlayText);
                Excel.Shape? overlay = null;
                try
                {
                    overlay = sheet.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, layout.Height);
                    overlay.AlternativeText = "OfficeTranslateOCR";
                    overlay.Fill.Visible = Office.MsoTriState.msoTrue; overlay.Fill.ForeColor.RGB = 0xFFFFFF; overlay.Fill.Transparency = 0.08F;
                    overlay.Line.Visible = Office.MsoTriState.msoFalse;
                    overlay.TextFrame2.MarginLeft = 2; overlay.TextFrame2.MarginRight = 2; overlay.TextFrame2.MarginTop = 1; overlay.TextFrame2.MarginBottom = 1;
                    overlay.TextFrame2.TextRange.Text = overlayText;
                    overlay.TextFrame2.TextRange.Font.Size = layout.FontSize;
                }
                catch (COMException ex)
                {
                    try { overlay?.Delete(); } catch { }
                    throw new InvalidOperationException("Excel 已完成图片识别，但创建译文覆盖框失败：" + ex.Message, ex);
                }
            }
            return regions.Count > 0;
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
            var result = new List<Target>();
            // Two bulk COM calls for the whole range instead of 3+ round-trips per
            // cell (MergeCells, HasFormula, Value2, Address). On a 10k-cell sheet
            // this turns minutes of COM chatter into two array transfers.
            // Write-back still touches only cells that actually get translated.
            var values = ToCellArray(range.Value2);
            var hasFormula = ToCellArray(range.HasFormula);
            var rows = values.GetLength(0);
            var cols = values.GetLength(1);
            for (var r = 1; r <= rows; r++)
                for (var c = 1; c <= cols; c++)
                {
                    if (hasFormula[r, c] is bool isFormula && isFormula) continue;
                    if (values[r, c] is string value && !string.IsNullOrWhiteSpace(value))
                    {
                        // Array indexes are 1-based relative to the range, matching
                        // Range.Cells[row, column]. Merged areas expose their value
                        // only in the top-left cell, so no address dedup is needed.
                        var targetCell = (Excel.Range)range.Cells[r, c];
                        result.Add(new Target(value, (text, bilingual) => { targetCell.Value2 = text; if (bilingual) targetCell.WrapText = true; }));
                    }
                }
            return result;
        }

        // Excel returns a 1-based object[,] for multi-cell ranges but a scalar
        // for a single cell. Normalize to a 1-based 2-D array either way.
        private static object[,] ToCellArray(object raw)
        {
            if (raw is object[,] grid) return grid;
            var single = new object[2, 2];
            single[1, 1] = raw;
            return single;
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
