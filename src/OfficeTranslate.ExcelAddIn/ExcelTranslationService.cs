using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

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
            if (targets.Count == 0) throw new InvalidOperationException("没有找到可翻译的单元格或可编辑图形文本。图片中的文字需要先进行 OCR。");

            using (var client = new TranslationClient())
            {
                for (var i = 0; i < targets.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    progress($"OfficeTranslate：正在翻译 {i + 1}/{targets.Count}");
                    var translated = (await client.TranslateAsync(targets[i].Text, settings, token)).TrimEnd('\r', '\n');
                    token.ThrowIfCancellationRequested();
                    targets[i].Write(settings.BilingualMode ? targets[i].Text + Environment.NewLine + translated : translated, settings.BilingualMode);
                    progress($"OfficeTranslate：已完成 {i + 1}/{targets.Count}");
                }
            }
        }

        private List<Target> ReadSelectionTargets()
        {
            var range = _excel.Selection as Excel.Range;
            if (range != null) return ReadRangeTargets(range);

            var result = new List<Target>();
            try
            {
                var drawingObjects = (Excel.DrawingObjects)_excel.Selection;
                Excel.ShapeRange shapes = drawingObjects.ShapeRange;
                for (var i = 1; i <= shapes.Count; i++) AddShape(shapes.Item(i), result);
            }
            catch (COMException) { }
            catch (InvalidCastException) { }
            return result;
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
