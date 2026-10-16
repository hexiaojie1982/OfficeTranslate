using System;
using System.Collections.Generic;
using System.Text;
using OfficeTranslate.Core;

namespace OfficeTranslate.WordAddIn
{
    // All methods execute on Office's STA thread. Each callback closes over the originating window/document.
    internal static class DocumentChatHost
    {
        internal static ChatContext Capture(dynamic window)
        {
#if WORD
            dynamic document = window.Document;
            dynamic selection = window.Selection;
            dynamic range = selection.Range.Duplicate;
            string selected = range.Start != range.End ? Convert.ToString(range.Text) ?? "" : "";
            if (selected.Replace("\r", "").Replace("\a", "").Replace("\u0001", "").Trim().Length == 0) selected = "";
            var canReplaceSelection = !string.IsNullOrWhiteSpace(selected);
            // Preserve a partial text selection inside a shape; only expand an actual shape selection.
            try
            {
                if (string.IsNullOrWhiteSpace(selected) || (int)selection.Type == 8)
                {
                    var shapes = new StringBuilder();
                    for (var i = 1; i <= (int)selection.ShapeRange.Count; i++) AppendWordShape(shapes, selection.ShapeRange[i]);
                    if (!string.IsNullOrWhiteSpace(shapes.ToString()))
                    {
                        selected = shapes.ToString().TrimStart('\n');
                        canReplaceSelection = (int)selection.ShapeRange.Count == 1 && selection.ShapeRange[1].Type != 6;
                        if (canReplaceSelection) { range = selection.ShapeRange[1].TextFrame.TextRange.Duplicate; selected = Convert.ToString(range.Text) ?? ""; }
                    }
                }
            }
            catch { /* No shape selection is normal. */ }
            string original = selected;
            dynamic target = range;
            var end = document.Content.End - 1;
            return ChatContext.Capture(selected, () => WordText(document), (action, text) =>
            {
                if (!Same(window.Document, document)) throw new InvalidOperationException();
                if (action == ChatWriteAction.ReplaceSelection)
                { if (Convert.ToString(target.Text) != original) throw new InvalidOperationException(); target.Text = text; }
                else if (action == ChatWriteAction.InsertAtStart) document.Range(0, 0).InsertBefore(text + "\r");
                else if (!string.IsNullOrWhiteSpace(original) && canReplaceSelection)
                {
                    if (Convert.ToString(target.Text) != original) throw new InvalidOperationException();
                    dynamic insertion = target.Duplicate; insertion.Collapse(0); insertion.InsertAfter("\r" + text);
                }
                else { if (document.Content.End - 1 != end) throw new InvalidOperationException(); document.Range(end, end).InsertAfter("\r" + text); }
            }, canReplaceSelection, DocumentName((object)document), "Word");
#elif EXCEL
            dynamic sheet = window.ActiveSheet; dynamic workbook = sheet.Parent;
            // Shapes are selected through Application.Selection; only consult it when this is the active window.
            dynamic shape = null!; string selectedShape = ""; var hasShapeSelection = false;
            try
            {
                if (Same(window.Application.ActiveWindow, window))
                {
                    dynamic shapes = window.Application.Selection.ShapeRange;
                    hasShapeSelection = (int)shapes.Count > 0;
                    var texts = new List<string>();
                    for (var i = 1; i <= (int)shapes.Count; i++) texts.Add(ShapeText(shapes.Item(i)));
                    selectedShape = string.Join("\n", texts);
                    if ((int)shapes.Count == 1 && (int)shapes.Item(1).Type != 6) shape = shapes.Item(1);
                }
            }
            catch { }
            if (!string.IsNullOrWhiteSpace(selectedShape))
            {
                string original = selectedShape; dynamic target = shape!; bool replaceShape = (object?)shape != null;
                return new ChatContext(original, true, (action, text) =>
                {
                    if (!Same(window.ActiveSheet.Parent, workbook)) throw new InvalidOperationException();
                    if (replaceShape && Convert.ToString(target.TextFrame2.TextRange.Text) != original) throw new InvalidOperationException();
                    if (action == ChatWriteAction.ReplaceSelection) target.TextFrame2.TextRange.Text = text;
                    else if (action == ChatWriteAction.InsertAtStart) ExcelStart(sheet, text);
                    else ExcelBelow(sheet, replaceShape ? (int)target.BottomRightCell.Row : (int)sheet.UsedRange.Row + (int)sheet.UsedRange.Rows.Count - 1, text);
                }, replaceShape, DocumentName((object)workbook), "Excel");
            }
            dynamic range = window.RangeSelection;
            // Intersect with UsedRange before materializing arrays (whole-column/whole-sheet selections).
            dynamic intersection = window.Application.Intersect(range, sheet.UsedRange);
            string selected = hasShapeSelection || intersection == null ? "" : ExcelRangeText(intersection);
            bool canReplace = !hasShapeSelection && Convert.ToDouble(range.CountLarge) == 1 && !Convert.ToBoolean(range.MergeCells);
            string originalValue = canReplace ? Convert.ToString(range.Formula) ?? "" : "";
            string originalCellValue = canReplace ? Convert.ToString(range.Value2) ?? "" : "";
            var bottom = (int)range.Row + (int)range.Rows.Count - 1;
            return ChatContext.Capture(selected, () => ExcelWorkbookText(workbook), (action, text) =>
            {
                if (!Same(window.ActiveSheet.Parent, workbook)) throw new InvalidOperationException();
                // Also validates the originating worksheet still exists.
                var name = sheet.Name;
                if (action == ChatWriteAction.ReplaceSelection)
                { if (!canReplace || (Convert.ToString(range.Formula) ?? "") != originalValue || (Convert.ToString(range.Value2) ?? "") != originalCellValue) throw new InvalidOperationException(); ExcelLiteral(range, text); }
                else if (action == ChatWriteAction.InsertAtStart) ExcelStart(sheet, text);
                else ExcelBelow(sheet, string.IsNullOrWhiteSpace(selected) ? (int)sheet.UsedRange.Row + (int)sheet.UsedRange.Rows.Count - 1 : bottom, text);
            }, canReplace, DocumentName((object)workbook), "Excel");
#else
            dynamic presentation = window.Presentation;
            dynamic selection = window.Selection;
            var selected = ""; dynamic target = null!; dynamic slide = null!;
            if ((int)selection.Type == 3) { target = selection.TextRange; selected = Convert.ToString(target.Text) ?? ""; }
            else if ((int)selection.Type == 2)
            {
                var texts = new List<string>();
                for (var i = 1; i <= (int)selection.ShapeRange.Count; i++) texts.Add(ShapeText(selection.ShapeRange[i]));
                selected = string.Join("\n", texts);
                if ((int)selection.ShapeRange.Count == 1 && selection.ShapeRange[1].HasTextFrame != 0) target = selection.ShapeRange[1].TextFrame.TextRange;
            }
            else if ((int)selection.Type == 1)
            {
                var texts = new List<string>();
                for (var i = 1; i <= (int)selection.SlideRange.Count; i++) texts.Add(SlideText(selection.SlideRange[i]));
                selected = string.Join("\n", texts);
                slide = selection.SlideRange[selection.SlideRange.Count];
            }
            try { if ((object?)slide == null) slide = window.View.Slide; } catch { }
            string original = target == null ? "" : Convert.ToString(target.Text) ?? "";
            dynamic capturedTarget = target!; dynamic capturedSlide = slide!;
            return ChatContext.Capture(selected, () => PresentationText(presentation), (action, text) =>
            {
                if (!Same(window.Presentation, presentation)) throw new InvalidOperationException();
                if (action == ChatWriteAction.ReplaceSelection)
                { if ((object?)capturedTarget == null) throw new InvalidOperationException();
                    if (Convert.ToString(capturedTarget!.Text) != original) throw new InvalidOperationException(); capturedTarget!.Text = text; }
                else if (action == ChatWriteAction.InsertAtStart)
                { dynamic first = presentation.Slides.Add(1, 12); AddTextBox(first, text, (float)presentation.PageSetup.SlideWidth, (float)presentation.PageSetup.SlideHeight); }
                else if (capturedTarget != null)
                { if (Convert.ToString(capturedTarget.Text) != original) throw new InvalidOperationException(); capturedTarget.InsertAfter("\r" + text); }
                else
                {
                    var index = capturedSlide == null ? (int)presentation.Slides.Count : (int)capturedSlide.SlideIndex;
                    dynamic added = presentation.Slides.Add(index + 1, 12); AddTextBox(added, text, (float)presentation.PageSetup.SlideWidth, (float)presentation.PageSetup.SlideHeight);
                }
            }, (object?)target != null && !string.IsNullOrWhiteSpace(selected), DocumentName((object)presentation), "PowerPoint");
#endif
        }
        private static bool Same(object a, object b) => ReferenceEquals(a, b) || a.Equals(b);
        private static string DocumentName(dynamic document)
        { try { return Convert.ToString(document.Name) ?? ""; } catch { return ""; } }
#if WORD
        private static string WordText(dynamic document)
        {
            var text = new StringBuilder(Convert.ToString(document.Content.Text) ?? "");
            // Floating body shapes are outside Document.Content.Text.
            for (var i = 1; i <= (int)document.Shapes.Count; i++) AppendWordShape(text, document.Shapes[i]);
            return text.ToString();
        }
        private static void AppendWordShape(StringBuilder text, dynamic shape)
        {
            if ((int)shape.Type == 6) { for (var i = 1; i <= (int)shape.GroupItems.Count; i++) AppendWordShape(text, shape.GroupItems[i]); }
            else if (shape.TextFrame.HasText != 0) text.Append('\n').Append(Convert.ToString(shape.TextFrame.TextRange.Text));
        }
#elif EXCEL
        private static string ExcelRangeText(dynamic range)
        {
            var text = new StringBuilder();
            for (var a = 1; a <= (int)range.Areas.Count; a++)
            {
                object value = range.Areas[a].Value2;
                if (value is Array values)
                {
                    for (var r = values.GetLowerBound(0); r <= values.GetUpperBound(0); r++)
                    { for (var c = values.GetLowerBound(1); c <= values.GetUpperBound(1); c++) { if (c > values.GetLowerBound(1)) text.Append('\t'); text.Append(Convert.ToString(values.GetValue(r, c))); } text.Append('\n'); }
                }
                else text.Append(Convert.ToString(value)).Append('\n');
            }
            return text.ToString().Trim();
        }
        private static string ExcelWorkbookText(dynamic workbook)
        {
            var text = new StringBuilder();
            for (var i = 1; i <= (int)workbook.Worksheets.Count; i++)
            {
                dynamic sheet = workbook.Worksheets[i];
                var body = new StringBuilder((string)ExcelRangeText(sheet.UsedRange));
                for (var j = 1; j <= (int)sheet.Shapes.Count; j++) body.Append('\n').Append(ShapeText(sheet.Shapes.Item(j)));
                if (!string.IsNullOrWhiteSpace(body.ToString())) text.Append('[').Append(Convert.ToString(sheet.Name)).Append("]\n").Append(body).Append('\n');
            }
            return text.ToString().Trim();
        }
        private static void ExcelLiteral(dynamic cell, string text)
        {
            if (text.Length > 32767) throw new InvalidOperationException();
            // Leading apostrophe stores text literally, including strings beginning with =,+,-,@.
            cell.Value2 = "'" + text;
        }
        private static void ExcelStart(dynamic sheet, string text)
        { if (text.Length > 32767) throw new InvalidOperationException(); sheet.Rows[1].Insert(); ExcelLiteral(sheet.Cells[1, 1], text); }
        private static void ExcelBelow(dynamic sheet, int row, string text)
        { if (text.Length > 32767 || row >= (int)sheet.Rows.Count) throw new InvalidOperationException(); sheet.Rows[row + 1].Insert(); ExcelLiteral(sheet.Cells[row + 1, 1], text); }
#else
        private static string PresentationText(dynamic presentation)
        { var text = new StringBuilder(); for (var i = 1; i <= (int)presentation.Slides.Count; i++) text.Append(SlideText(presentation.Slides[i])).Append('\n'); return text.ToString(); }
        private static string SlideText(dynamic slide)
        { var text = new StringBuilder(); for (var i = 1; i <= (int)slide.Shapes.Count; i++) text.Append(ShapeText(slide.Shapes[i])).Append('\n'); return text.ToString(); }
        private static void AddTextBox(dynamic slide, string text, float width, float height)
        { dynamic box = slide.Shapes.AddTextbox(1, 24f, 24f, width - 48f, height - 48f); box.TextFrame.TextRange.Text = text; box.TextFrame.TextRange.Font.Size = 18; box.TextFrame.WordWrap = -1; box.TextFrame2.AutoSize = 2; }
#endif
#if !WORD
        private static string ShapeText(dynamic shape)
        {
            var text = new StringBuilder();
            if ((int)shape.Type == 6) { for (var i = 1; i <= (int)shape.GroupItems.Count; i++) text.Append(ShapeText(shape.GroupItems[i])).Append('\n'); }
#if EXCEL
            else { try { text.Append(Convert.ToString(shape.TextFrame2.TextRange.Text)); } catch { } }
#else
            else if (shape.HasTable != 0)
            { for (var r = 1; r <= (int)shape.Table.Rows.Count; r++) for (var c = 1; c <= (int)shape.Table.Columns.Count; c++) text.Append(Convert.ToString(shape.Table.Cell(r,c).Shape.TextFrame.TextRange.Text)).Append('\n'); }
            else if (shape.HasTextFrame != 0) text.Append(Convert.ToString(shape.TextFrame.TextRange.Text));
#endif
            return text.ToString();
        }
#endif
    }
}
