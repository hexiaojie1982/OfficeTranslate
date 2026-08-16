using Microsoft.Office.Interop.Word;
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

        public async Task TranslateAsync(bool wholeDocument, bool bilingual, TranslationSettings settings, CancellationToken token, Action<string> progress)
        {
            var targets = wholeDocument ? ReadDocumentParagraphs() : ReadSelection();
            if (targets.Count == 0) throw new InvalidOperationException(wholeDocument ? "文档中没有可翻译的正文。" : "请先选择需要翻译的文字。");
            using (var client = new TranslationClient())
            {
                for (var i = 0; i < targets.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    progress($"OfficeTranslate：正在翻译 {i + 1}/{targets.Count}");
                    var text = await client.TranslateAsync(targets[i].Text, settings, token);
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
                    progress($"OfficeTranslate：已完成 {i + 1}/{targets.Count}");
                }
            }
        }

        private List<Target> ReadSelection()
        {
            var selection = _word.Selection;
            if (selection == null || selection.Range.Start == selection.Range.End || string.IsNullOrWhiteSpace(StripMarks(selection.Range.Text))) return new List<Target>();
            return new List<Target> { ToTarget(selection.Range) };
        }

        private List<Target> ReadDocumentParagraphs()
        {
            var result = new List<Target>();
            foreach (Paragraph paragraph in _word.ActiveDocument.StoryRanges[WdStoryType.wdMainTextStory].Paragraphs)
            {
                var target = ToTarget(paragraph.Range);
                if (!string.IsNullOrWhiteSpace(target.Text)) result.Add(target);
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

        private sealed class Target
        {
            public Target(Range range, string text) { Range = range; Text = text; }
            public Range Range { get; }
            public string Text { get; }
        }
    }
}
