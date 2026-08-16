using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
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
            if (targets.Count == 0) throw new InvalidOperationException(wholePresentation ? "演示文稿中没有可翻译的文本。" : "请先选择文本框或文字。");
            using (var client = new TranslationClient())
            {
                for (var i = 0; i < targets.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    progress($"OfficeTranslate：正在翻译 {i + 1}/{targets.Count}");
                    var translated = (await client.TranslateAsync(targets[i].Text, settings, token)).TrimEnd('\r', '\n');
                    token.ThrowIfCancellationRequested();
                    targets[i].Write(settings.BilingualMode ? targets[i].Text + "\r" + translated : translated);
                    progress($"OfficeTranslate：已完成 {i + 1}/{targets.Count}");
                }
            }
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
                        AddShape(selection.ShapeRange[1], result); return result;
                    }
                }
                catch (COMException) { }
                AddRange(selection.TextRange, result); return result;
            }
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in selection.ShapeRange) AddShape(shape, result);
            }
            else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
            {
                foreach (PowerPoint.Slide slide in selection.SlideRange)
                    foreach (PowerPoint.Shape shape in slide.Shapes) AddShape(shape, result);
            }
            return result;
        }

        private static void AddShape(PowerPoint.Shape shape, List<Target> result)
        {
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
