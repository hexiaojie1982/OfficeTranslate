using Excel = Microsoft.Office.Interop.Excel;
using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
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
            var source = wholeSheet ? sheet.UsedRange : _excel.Selection as Excel.Range;
            if (source == null) throw new InvalidOperationException("请先选择需要翻译的单元格。");
            var targets = ReadTargets(source);
            if (targets.Count == 0) throw new InvalidOperationException("没有找到可翻译的文本单元格。");

            using (var client = new TranslationClient())
            {
                for (var i = 0; i < targets.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    progress($"OfficeTranslate：正在翻译 {i + 1}/{targets.Count}");
                    var translated = (await client.TranslateAsync(targets[i].Text, settings, token)).TrimEnd('\r', '\n');
                    token.ThrowIfCancellationRequested();
                    var cell = sheet.Range[targets[i].Address];
                    cell.Value2 = settings.BilingualMode ? targets[i].Text + Environment.NewLine + translated : translated;
                    if (settings.BilingualMode) cell.WrapText = true;
                    progress($"OfficeTranslate：已完成 {i + 1}/{targets.Count}");
                }
            }
        }

        private static List<Target> ReadTargets(Excel.Range range)
        {
            var result = new List<Target>(); var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Excel.Range item in range.Cells)
            {
                var cell = item;
                if ((bool)cell.MergeCells) cell = (Excel.Range)cell.MergeArea.Cells[1, 1];
                var address = cell.Address[false, false, Excel.XlReferenceStyle.xlA1];
                if (!seen.Add(address) || (bool)cell.HasFormula) continue;
                var value = cell.Value2 as string;
                if (value != null && !string.IsNullOrWhiteSpace(value)) result.Add(new Target(address, value));
            }
            return result;
        }

        private sealed class Target
        {
            public Target(string address, string text) { Address = address; Text = text; }
            public string Address { get; } public string Text { get; }
        }
    }
}
