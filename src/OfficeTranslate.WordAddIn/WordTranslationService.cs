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
            var document = _word.ActiveDocument;
            var positionOffset = 0;
            using (var client = new TranslationClient())
            {
                for (var i = 0; i < targets.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    progress($"OfficeTranslate：正在翻译 {i + 1}/{targets.Count}");
                    var text = await client.TranslateAsync(targets[i].Text, settings, token);
                    token.ThrowIfCancellationRequested();

                    var translatedText = text.TrimEnd('\r', '\a');
                    var start = targets[i].Start + positionOffset;
                    var end = targets[i].End + positionOffset;
                    var contentEndBefore = document.Content.End;
                    progress($"OfficeTranslate：正在写回 {i + 1}/{targets.Count}");

                    var undo = _word.UndoRecord;
                    undo.StartCustomRecord(bilingual ? $"OfficeTranslate 双语排版 {i + 1}/{targets.Count}" : $"OfficeTranslate 翻译 {i + 1}/{targets.Count}");
                    try
                    {
                        if (bilingual)
                            document.Range(end, end).InsertAfter("\r" + translatedText);
                        else
                            document.Range(start, end).Text = translatedText;
                    }
                    finally { undo.EndCustomRecord(); }

                    // Every write can change all following Word character positions.
                    // Use Word's actual content length instead of estimating from strings,
                    // which also handles paragraph and table-cell markers correctly.
                    positionOffset += document.Content.End - contentEndBefore;
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
            return new Target(range.Start, range.Start + contentLength, raw.Substring(0, contentLength));
        }

        private static string StripMarks(string text) => text.TrimEnd('\r', '\a');

        private sealed class Target
        {
            public Target(int start, int end, string text) { Start = start; End = end; Text = text; }
            public int Start { get; }
            public int End { get; }
            public string Text { get; }
        }
    }
}
