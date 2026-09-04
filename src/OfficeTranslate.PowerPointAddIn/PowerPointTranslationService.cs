using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeTranslate.PowerPointAddIn
{
    internal sealed class PowerPointTranslationService
    {
        private readonly PowerPoint.Application _powerPoint;
        public PowerPointTranslationService(PowerPoint.Application powerPoint) => _powerPoint = powerPoint;

        public async Task TranslateAsync(bool wholePresentation, TranslationSettings settings, CancellationToken token, Action<string> progress)
        {
            if (_powerPoint.Presentations.Count == 0) throw new InvalidOperationException("请先打开演示文稿。");
            var targets = wholePresentation ? ReadPresentation() : ReadSelection();
            var images = settings.ImageOcrEnabled ? (wholePresentation ? ReadPresentationImages() : ReadSelectionImages()) : new List<PowerPoint.Shape>();
            if (targets.Count == 0 && images.Count == 0) throw new InvalidOperationException(wholePresentation ? "演示文稿中没有可翻译的文本或图片。" : "请先选择文本框、文字或图片。");
            using (var client = new TranslationClient())
            {
                var total = targets.Count + images.Count;
                for (var i = 0; i < targets.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    progress($"OfficeTranslate：正在翻译 {i + 1}/{total}");
                    var translated = (await client.TranslateAsync(targets[i].Text, settings, token)).TrimEnd('\r', '\n');
                    token.ThrowIfCancellationRequested();
                    targets[i].Write(settings.BilingualMode ? targets[i].Text + "\r" + translated : translated);
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

        private async Task TranslateImageAsync(PowerPoint.Shape image, TranslationClient client, TranslationSettings settings, CancellationToken token)
        {
            var path = Path.Combine(Path.GetTempPath(), "OfficeTranslate-" + Guid.NewGuid().ToString("N") + ".png");
            try
            {
                image.Export(path, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                var regions = await client.TranslateImageAsync(File.ReadAllBytes(path), settings, token);
                var slide = image.Parent as PowerPoint.Slide ?? throw new InvalidOperationException("无法确定图片所在的幻灯片。");
                foreach (var region in regions)
                {
                    var left = image.Left + image.Width * region.X1 / 1000F;
                    var top = image.Top + image.Height * region.Y1 / 1000F;
                    var width = Math.Max(12F, image.Width * (region.X2 - region.X1) / 1000F);
                    var height = Math.Max(10F, image.Height * (region.Y2 - region.Y1) / 1000F);
                    var overlay = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
                    overlay.Tags.Add("OfficeTranslateOCR", "1");
                    overlay.Fill.Visible = Office.MsoTriState.msoTrue; overlay.Fill.ForeColor.RGB = 0xFFFFFF; overlay.Fill.Transparency = 0.08F;
                    overlay.Line.Visible = Office.MsoTriState.msoFalse;
                    overlay.TextFrame2.MarginLeft = 2; overlay.TextFrame2.MarginRight = 2; overlay.TextFrame2.MarginTop = 1; overlay.TextFrame2.MarginBottom = 1;
                    overlay.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeTextToFitShape;
                    overlay.TextFrame2.TextRange.Text = settings.BilingualMode && !string.IsNullOrWhiteSpace(region.Source)
                        ? region.Source + "\r" + region.Translation
                        : region.Translation;
                    overlay.TextFrame2.TextRange.Font.Size = Math.Max(8F, Math.Min(24F, height * 0.55F));
                }
            }
            finally { try { if (File.Exists(path)) File.Delete(path); } catch { } }
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
            try { if (shape.Tags["OfficeTranslateOCR"] == "1") return; } catch (COMException) { }
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
