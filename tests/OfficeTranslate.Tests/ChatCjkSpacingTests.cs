using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using OfficeTranslate.WordAddIn;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ChatCjkSpacingTests
    {
        private const BindingFlags Hidden = BindingFlags.Instance | BindingFlags.NonPublic;
        private static object Field(object o, string name) => o.GetType().GetField(name, Hidden)!.GetValue(o)!;
        private static void Call(object o, string name, params object[] args) => o.GetType().GetMethod(name, Hidden)!.Invoke(o, args);
        private static IEnumerable<Control> Descendants(Control c) => c.Controls.Cast<Control>().SelectMany(x => new[] { x }.Concat(Descendants(x)));

        [TestMethod]
        public void HanExtensions_Fullwidth_Kana_AndUnicodeEdges_AreDeterministic()
        {
            var points = new[] { 0x3400, 0x4DBF, 0x4E00, 0x9FFF, 0xF900, 0xFAFF,
                0x20000, 0x2A6DF, 0x2A700, 0x2B73F, 0x2B740, 0x2B81F, 0x2B820, 0x2CEAF,
                0x2CEB0, 0x2EBEF, 0x2EBF0, 0x2EE5F, 0x2F800, 0x2FA1F,
                0x30000, 0x3134F, 0x31350, 0x323AF, 0x323B0, 0x3347F,
                0x3041, 0x309F, 0x30A0, 0x30FF, 0x3001, 0x303F, 0xFF01, 0xFFEF };
            foreach (var point in points)
            {
                var c = char.ConvertFromUtf32(point);
                Assert.AreEqual(c + "中" + c, ChatCjkSpacing.Clean(c + " \t\u3000中 " + c), "U+" + point.ToString("X"));
            }
            foreach (var boundary in new[] { "A", "1", "😀", "한", "ᄒ", "ㅎ", "\ud800", "\udc00" })
                Assert.AreEqual("中 " + boundary + " 文", ChatCjkSpacing.Clean("中 " + boundary + " 文"));
            const string sample = "中 文 `中 文`\r\n```\r\n中 文\r\n```\n中\u3000文";
            var clean = ChatCjkSpacing.Clean(sample);
            Assert.AreEqual(clean, ChatCjkSpacing.Clean(clean));
        }

        [DataTestMethod]
        [DataRow("韩 国 翻 译", "韩国翻译")]
        [DataRow("韩\t \u3000国\u3000翻\t译", "韩国翻译")]
        [DataRow("我在用 iPhone 16", "我在用 iPhone 16")]
        [DataRow("2024 年 10 月", "2024 年 10 月")]
        [DataRow("A 中 文 B", "A 中文 B")]
        [DataRow("한국어 번역 입니다", "한국어 번역 입니다")]
        [DataRow("한 글 中 文 한 글", "한 글 中文 한 글")]
        [DataRow("English words  and\ttabs", "English words  and\ttabs")]
        [DataRow("中\r\n文\n字\r句", "中\r\n文\n字\r句")]
        [DataRow(" 中 文  ", " 中文  ")]
        [DataRow("中\u00a0文\u2009字", "中\u00a0文\u2009字")]
        [DataRow("你 好 ， 世 界 ！", "你好，世界！")]
        [DataRow("ひ ら が な カ タ カ ナ", "ひらがなカタカナ")]
        [DataRow("\u3400 \u4dbf \uf900 \ufaff", "\u3400\u4dbf\uf900\ufaff")]
        [DataRow("𠀀 中 𪜀 𰀀 𱍐", "𠀀中𪜀𰀀𱍐")]
        [DataRow("中 😀 文", "中 😀 文")]
        [DataRow("韩 国 `中 文\t字` 翻 译", "韩国 `中 文\t字` 翻译")]
        [DataRow("中 ``中 ` 文`` 字", "中 ``中 ` 文`` 字")]
        [DataRow("中 `中\n文 字` 字", "中 `中\n文 字` 字")]
        [DataRow("中 `未 闭 合", "中 `未闭合")]
        [DataRow("中 \\`文 字", "中 \\`文字")]
        [DataRow("中 `文\\` 字 句", "中 `文\\` 字句")]
        [DataRow("中 文\r\n```cs\r\n中 文\t字\u3000句\r\n```\r\n翻 译", "中文\r\n```cs\r\n中 文\t字\u3000句\r\n```\r\n翻译")]
        [DataRow("中 文\n````\n中 文\n```\n字 句\n````\n翻 译", "中文\n````\n中 文\n```\n字 句\n````\n翻译")]
        [DataRow("中 文\n   ```中 文\n中 文\n   ```\n翻 译", "中文\n   ```中 文\n中 文\n   ```\n翻译")]
        [DataRow("中 文\n```\n中 文", "中文\n```\n中 文")]
        [DataRow("中 文\n~~~\n中 文\n~~~\n翻 译", "中文\n~~~\n中 文\n~~~\n翻译")]
        [DataRow("中 文\n```\n中 文\n``` invalid\n字 句\n```\n翻 译", "中文\n```\n中 文\n``` invalid\n字 句\n```\n翻译")]
        [DataRow("中 ``文 字` 不 同", "中 ``文字` 不同")]
        [DataRow("", "")]
        public async Task CompletedAnswer_CleansOnlyUnprotectedCjkGaps(string source, string expected)
        {
            var session = new DocumentChatSession();
            var snapshots = new List<string>();
            var reply = await session.SendAsync(new ChatContext("context", false), "用 户", ChatMode.Question, new TranslationSettings(),
                (m, s, n, o, t) => { o(source); return Task.FromResult(new ChatReply(source, true)); }, snapshots.Add, (i, n) => { }, CancellationToken.None);
            Assert.AreEqual(expected, reply.Text); Assert.IsTrue(reply.UsedFallback);
            Assert.AreEqual(expected, session.History[1].Content);
            Assert.AreEqual("用 户", session.History[0].Content);
            Assert.AreEqual(source, snapshots[0], "Streaming snapshots must remain unchanged.");
        }

        [DataTestMethod]
        [DataRow(ChatMode.Question)] [DataRow(ChatMode.Summarize)] [DataRow(ChatMode.Rewrite)] [DataRow(ChatMode.Translate)]
        public async Task CompletedAnswer_IsCleanBeforeFailedSaveRetry_AndNextRequest(ChatMode mode)
        {
            using (var sandbox = new HistorySandbox())
            {
                File.WriteAllText(sandbox.DirectoryPath, "block directory");
                var session = new DocumentChatSession(sandbox.Store);
                await session.SendAsync(new ChatContext("context", false), "用 户", mode, new TranslationSettings(),
                    (m, s, n, o, t) => Task.FromResult(new ChatReply("韩 国 翻 译", false)), _ => { }, (i, n) => { }, CancellationToken.None);
                Assert.IsTrue(session.HistorySaveFailed);
                File.Delete(sandbox.DirectoryPath); Assert.IsTrue(session.TrySaveHistory());
                var record = sandbox.Store.Read(sandbox.Store.List(out _).Single().SessionId)!;
                Assert.AreEqual("韩国翻译", record.Turns.Single().Answer);
                Assert.AreEqual("用 户", record.Turns.Single().Question);
                await session.SendAsync(new ChatContext("context", false), "next", mode, new TranslationSettings(),
                    (m, s, n, o, t) => { Assert.AreEqual("韩国翻译", m.Single(x => x.Role == "assistant").Content); return Task.FromResult(new ChatReply("answer", false)); },
                    _ => { }, (i, n) => { }, CancellationToken.None);
            }
        }

        [DataTestMethod]
        [DataRow("Word", false)] [DataRow("Word", true)]
        [DataRow("Excel", false)] [DataRow("Excel", true)]
        [DataRow("PowerPoint", false)] [DataRow("PowerPoint", true)]
        public void Pane_FinalBubble_DiskHistory_Exports_AndCopyShareCleanText(string host, bool translationSettings)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var sandbox = new HistorySandbox())
                using (var pane = (UserControl)Activator.CreateInstance(DocumentChatHostTests.Host(host).GetType("OfficeTranslate.WordAddIn.DocumentChatPane", true)!)!)
                using (var window = new Form { ClientSize = new Size(440, 760) })
                {
                    const string question = "用 户 原 文";
                    const string source = "韩 国 翻 译\n我在用 iPhone 16，2024 年\n`中 文`\n```\n中 文\t字\n```\n한국어 답변 입니다";
                    const string expected = "韩国翻译\n我在用 iPhone 16，2024 年\n`中 文`\n```\n中 文\t字\n```\n한국어 답변 입니다";
                    pane.GetType().GetProperty("HistoryStore", Hidden)!.SetValue(pane, sandbox.Store);
                    var settings = new TranslationSettings { ChatUseTranslationSettings = translationSettings };
                    Call(pane, "Initialize", (Func<ChatContext>)(() => new ChatContext("context", false, host: host)), (Func<TranslationSettings>)(() => settings));
                    pane.Dock = DockStyle.Fill; window.Controls.Add(pane); window.Show(); Application.DoEvents();
                    Action<string> output = null!; var completion = new TaskCompletionSource<ChatReply>();
                    pane.GetType().GetProperty("Transport", Hidden)!.SetValue(pane, (ChatTransport)((m, s, n, o, t) => { output = o; return completion.Task; }));
                    ((TextBox)Field(pane, "_input")).Text = question; Call(pane, "Start", ChatMode.Question);
                    output(source); Call(pane, "FlushOutput");
                    Assert.AreEqual(source, Field(pane, "_renderedOutput"));
                    Assert.AreEqual(0, sandbox.Store.List(out _).Count);
                    completion.SetResult(new ChatReply(source, false));
                    var deadline = DateTime.UtcNow.AddSeconds(8);
                    while (FieldOrNull(pane, "_pending") != null && DateTime.UtcNow < deadline) { Application.DoEvents(); Thread.Sleep(5); }
                    Assert.IsNull(FieldOrNull(pane, "_pending"));
                    Assert.AreEqual(expected, Field(pane, "_result"));
                    var record = sandbox.Store.Read(sandbox.Store.List(out _).Single().SessionId)!;
                    Assert.AreEqual(expected, record.Turns.Single().Answer); Assert.AreEqual(question, record.Turns.Single().Question);
                    // Read the JSON itself as well as the public store API.
                    var json = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<ChatHistoryRecord>(File.ReadAllText(Directory.GetFiles(sandbox.DirectoryPath, "*.json").Single()));
                    Assert.AreEqual(expected, json.Turns.Single().Answer);
                    var rows = Descendants((Control)Field(pane, "_transcript")).Where(c => c.GetType().Name == "ChatBubbleRow").ToArray();
                    Assert.AreEqual(2, rows.Length);
                    Assert.AreEqual(expected, Field(rows[1], "_source"));
                    for (var i = 0; i < rows.Length; i++)
                    {
                        var copy = (Control)rows[i].GetType().GetProperty("CopyButton", Hidden)!.GetValue(rows[i])!;
                        Assert.IsTrue(copy.Visible); string? copied = null;
                        copy.GetType().GetProperty("ClipboardWriter", Hidden)!.SetValue(copy, (Action<string>)(text => copied = text));
                        Call(copy, "OnClick", EventArgs.Empty); Assert.AreEqual(i == 0 ? question : expected, copied);
                    }
                    foreach (var markdown in new[] { true, false })
                    {
                        var path = Path.Combine(sandbox.DirectoryPath, markdown ? "chat.md" : "chat.txt");
                        pane.GetType().GetProperty("ExportDialog", Hidden)!.SetValue(pane, (Func<SaveFileDialog, DialogResult>)(d => { d.FileName = path; d.FilterIndex = markdown ? 1 : 2; return DialogResult.OK; }));
                        Call(pane, "Export"); var exported = File.ReadAllText(path);
                        StringAssert.Contains(exported, "韩国翻译"); Assert.IsFalse(exported.Contains("韩 国 翻 译"));
                        StringAssert.Contains(exported, question); StringAssert.Contains(exported, "中 文\t字");
                        StringAssert.Contains(exported, "我在用 iPhone 16，2024 年");
                    }
                    using (var history = (Form)Activator.CreateInstance(DocumentChatHostTests.Host(host).GetType("OfficeTranslate.WordAddIn.ChatHistoryForm", true)!, sandbox.Store, "en")!)
                    {
                        history.Show(); Application.DoEvents();
                        ((ListView)Field(history, "_list")).Items[0].Selected = true; Application.DoEvents();
                        var savedRows = Descendants((Control)Field(history, "_details")).Where(c => c.GetType().Name == "ChatBubbleRow").ToArray();
                        Assert.AreEqual(expected, Field(savedRows[1], "_source"));
                        var copy = (Control)savedRows[1].GetType().GetProperty("CopyButton", Hidden)!.GetValue(savedRows[1])!;
                        string? copied = null;
                        copy.GetType().GetProperty("ClipboardWriter", Hidden)!.SetValue(copy, (Action<string>)(text => copied = text));
                        Call(copy, "OnClick", EventArgs.Empty); Assert.AreEqual(expected, copied);
                    }
                }
            });
        }
        private static object? FieldOrNull(object o, string name) => o.GetType().GetField(name, Hidden)!.GetValue(o);
    }
}
