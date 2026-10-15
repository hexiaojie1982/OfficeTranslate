using System;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using OfficeTranslate.WordAddIn;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ChatTranscriptExportTests
    {
        [DataTestMethod]
        [DataRow("zh-CN", true)] [DataRow("en", true)]
        [DataRow("zh-CN", false)] [DataRow("en", false)]
        public void Format_PreservesActualMultilineUnicodeContent_WithLocalizedHeadingsAndTimestamp(string language, bool markdown)
        {
            string T(string key) => UiText.Get(language, key);
            var time = new DateTimeOffset(2026, 10, 15, 22, 13, 14, TimeSpan.FromHours(8));
            var history = new[] { new ChatMessage("user", "问题😀\r\n\r\n# heading\n> quote", time), new ChatMessage("assistant", "回答\n- item\rfinal", time) };
            var expected = markdown
                ? "# " + T("ChatExportHeader") + "\r\n\r\n## " + T("ChatExportTimestamp") + ": 2026-10-15 22:13:14 +08:00\r\n\r\n### " + T("ChatExportQuestion") + "\r\n> 问题😀\r\n> \r\n> # heading\r\n> > quote\r\n\r\n### " + T("ChatExportAnswer") + "\r\n> 回答\r\n> - item\r\n> final\r\n\r\n"
                : T("ChatExportHeader") + "\r\n\r\n----------------------------------------\r\n" + T("ChatExportTimestamp") + ": 2026-10-15 22:13:14 +08:00\r\n\r\n" + T("ChatExportQuestion") + "\r\n问题😀\r\n\r\n# heading\n> quote\r\n\r\n" + T("ChatExportAnswer") + "\r\n回答\n- item\rfinal\r\n\r\n";
            Assert.AreEqual(expected, ChatTranscriptExport.Format(history, markdown, T));
            Assert.AreEqual("问题😀\r\n\r\n# heading\n> quote", history[0].Content);
        }

        [TestMethod]
        public async Task Export_UsesTenCommittedTurns_EvenWhenRequestBudgetTrimsHistory_AndClearRemovesTimestamps()
        {
            var session = new DocumentChatSession();
            var settings = new TranslationSettings { UiLanguage = "en", ChatContextTokens = 2048, ChatOutputTokens = 256 };
            var before = DateTimeOffset.Now;
            for (var i = 0; i < 12; i++)
                await session.SendAsync(new ChatContext("PRIVATE DOCUMENT CONTEXT", false), "question-" + i, ChatMode.Question, settings,
                    (m, s, n, o, t) => Task.FromResult(new ChatReply("answer-" + i + new string('a', 400), false)), _ => { }, (i1, n) => { }, CancellationToken.None);
            var snapshot = session.History;
            Assert.AreEqual(20, snapshot.Count);
            Assert.AreEqual("question-2", snapshot[0].Content);
            for (var i = 0; i < snapshot.Count; i += 2)
            {
                Assert.AreEqual(snapshot[i].Timestamp, snapshot[i + 1].Timestamp);
                Assert.IsTrue(snapshot[i].Timestamp >= before && snapshot[i].Timestamp <= DateTimeOffset.Now);
            }
            var export = ChatTranscriptExport.Format(snapshot, false, k => UiText.Get("en", k));
            Assert.AreEqual(10, export.Split(new[] { "----------------------------------------" }, StringSplitOptions.None).Length - 1);
            Assert.IsFalse(export.Contains("PRIVATE DOCUMENT CONTEXT"));
            Assert.IsTrue(export.IndexOf("question-2", StringComparison.Ordinal) < export.IndexOf("question-11", StringComparison.Ordinal));
            Assert.AreEqual(20, session.History.Count);
            session.Clear();
            Assert.AreEqual("", ChatTranscriptExport.Format(session.History, true, k => UiText.Get("en", k)));
            Assert.AreEqual(20, snapshot.Count, "An export snapshot must remain immutable when the session is cleared.");
        }

        [DataTestMethod]
        [DataRow("zh-CN")] [DataRow("zh-HK")] [DataRow("ko")] [DataRow("ja")] [DataRow("en")]
        public void ExportStrings_AreLocalized_InEveryUiLanguage(string language)
        {
            var keys = new[] { "ChatExport", "ChatExportTitle", "ChatExportFilter", "ChatExportFileName", "ChatExportHeader", "ChatExportTimestamp", "ChatExportQuestion", "ChatExportAnswer", "ChatExportEmpty", "ChatExportSuccess", "ChatExportFailed", "ChatExportAccessError", "ChatExportIoError", "ChatExportPathError", "ChatExportUnknownError" };
            foreach (var key in keys)
            {
                var value = UiText.Get(language, key);
                Assert.IsFalse(string.IsNullOrWhiteSpace(value)); Assert.AreNotEqual(key, value);
                if (language != "en") Assert.AreNotEqual(UiText.Get("en", key), value, key);
            }
            var filter = UiText.Get(language, "ChatExportFilter").Split('|');
            Assert.AreEqual(4, filter.Length); Assert.AreEqual("*.md", filter[1]); Assert.AreEqual("*.txt", filter[3]);
            var name = string.Format(CultureInfo.InvariantCulture, UiText.Get(language, "ChatExportFileName"), new DateTime(2026, 10, 15, 22, 13, 14));
            StringAssert.EndsWith(name, "_20261015_221314.md");
            if (language == "zh-CN") Assert.AreEqual("文档内Chat_20261015_221314.md", name);
        }
    }
}
