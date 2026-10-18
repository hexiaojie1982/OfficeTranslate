using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using OfficeTranslate.WordAddIn;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ChatTranslateTests
    {
        private const BindingFlags Hidden = BindingFlags.Instance | BindingFlags.NonPublic;
        private static object? Field(object target, string name) => target.GetType().GetField(name, Hidden)!.GetValue(target);
        private static object? Call(object target, string name, params object[] args) => target.GetType().GetMethod(name, Hidden)!.Invoke(target, args);
        private static IEnumerable<Control> All(Control parent) => parent.Controls.Cast<Control>().SelectMany(c => new[] { c }.Concat(All(c)));
        private static Button Button(Control parent, string key) => All(parent).OfType<Button>().Single(b => (string?)b.Tag == key);
        private static void Pump(Func<bool> done)
        {
            var end = DateTime.UtcNow.AddSeconds(5);
            while (!done() && DateTime.UtcNow < end) { Application.DoEvents(); Thread.Sleep(5); }
            Assert.IsTrue(done(), "Translate UI operation timed out.");
        }

        [DataTestMethod]
        [DataRow("Word", "zh-CN")] [DataRow("Word", "zh-HK")] [DataRow("Word", "en")] [DataRow("Word", "ko")] [DataRow("Word", "ja")]
        [DataRow("Excel", "zh-CN")] [DataRow("Excel", "zh-HK")] [DataRow("Excel", "en")] [DataRow("Excel", "ko")] [DataRow("Excel", "ja")]
        [DataRow("PowerPoint", "zh-CN")] [DataRow("PowerPoint", "zh-HK")] [DataRow("PowerPoint", "en")] [DataRow("PowerPoint", "ko")] [DataRow("PowerPoint", "ja")]
        public void Pane_OrdersButtons_UsesLatestTargetAndContext_StreamsAndGatesWrites(string host, string language)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var history = new HistorySandbox())
                using (var pane = (UserControl)Activator.CreateInstance(DocumentChatHostTests.Host(host).GetType("OfficeTranslate.WordAddIn.DocumentChatPane", true)!)!)
                using (var window = new Form { ClientSize = new Size(440, 760) })
                {
                    var settings = new TranslationSettings();
                    var applied = new List<ChatWriteAction>();
                    var context = new ChatContext("initial preview", false);
                    pane.GetType().GetProperty("HistoryStore", Hidden)!.SetValue(pane, history.Store);
                    pane.GetType().GetProperty("ConfirmWrite", Hidden)!.SetValue(pane, (Func<ChatWriteAction, bool>)(_ => true));
                    Call(pane, "Initialize", (Func<ChatContext>)(() => context), (Func<TranslationSettings>)(() => settings));
                    pane.Dock = DockStyle.Fill; window.Controls.Add(pane); window.Show();
                    settings.UiLanguage = language; Call(pane, "RefreshLanguage");
                    var translate = Button(pane, "ChatTranslate");
                    CollectionAssert.AreEqual(new[] { "ChatTranslate", "ChatSummary", "ChatRewrite", "ChatSettings" },
                        translate.Parent!.Controls.Cast<Control>().Select(c => (string)c.Tag).ToArray());
                    Assert.AreEqual(UiText.Get(language, "ChatTranslate"), translate.Text);
                    Assert.AreEqual(Color.FromArgb(45, 108, 223), translate.BackColor);
                    foreach (var width in new[] { 320, 440, 600 })
                    {
                        window.ClientSize = new Size(width, 760); Application.DoEvents();
                        foreach (Control button in translate.Parent.Controls)
                        { Assert.IsTrue(button.Right <= button.Parent!.ClientSize.Width); Assert.IsTrue(button.Bottom <= button.Parent.ClientSize.Height); }
                    }
                    // Refresh preview only once: each click must capture the new selection/full text.
                    foreach (var action in new[] { ChatWriteAction.ReplaceSelection, ChatWriteAction.InsertBelow, ChatWriteAction.InsertBelow })
                    {
                        var selected = applied.Count < 2;
                        settings.TargetLanguage = selected ? "英语" : "한국어";
                        var fullReads = 0;
                        context = ChatContext.Capture(selected ? "new selected text" : "", () => { fullReads++; return "new whole document"; },
                            (a, text) => { Assert.AreEqual("中文 **AI**", text); applied.Add(a); }, true, "fixture", host);
                        Assert.AreEqual(selected ? 0 : 1, fullReads);
                        var completion = new TaskCompletionSource<ChatReply>(TaskCreationOptions.RunContinuationsAsynchronously);
                        ChatTransport transport = (messages, s, n, output, token) =>
                        {
                            Assert.AreSame(settings, s);
                            Assert.AreEqual(string.Format(UiText.Get(language, "ChatTranslateRequest"), settings.TargetLanguage) +
                                "\n<document>\n" + context.Text + "\n</document>", messages.Last().Content);
                            StringAssert.StartsWith(messages[0].Content, settings.ChatTranslateSystemPrompt);
                            output("stream snapshot"); return completion.Task;
                        };
                        pane.GetType().GetProperty("Transport", Hidden)!.SetValue(pane, transport);
                        translate.PerformClick();
                        var transcript = (Control)Field(pane, "_transcript")!;
                        Pump(() => transcript.Text.Contains("stream snapshot"));
                        var writes = (FlowLayoutPanel)Field(pane, "_writes")!;
                        Assert.AreEqual(0, writes.Controls.Count);
                        Assert.IsFalse(translate.Enabled);
                        completion.SetResult(new ChatReply("中 文 **AI**", false));
                        Pump(() => Field(pane, "_pending") == null);
                        CollectionAssert.AreEqual(new[] { "ChatReplace", "ChatInsertBelow" }, writes.Controls.Cast<Control>().Select(c => (string)c.Tag).ToArray());
                        Assert.AreEqual(selected, Button(writes, "ChatReplace").Enabled);
                        Assert.IsTrue(Button(writes, "ChatInsertBelow").Enabled);
                        if (!selected) StringAssert.Contains(((Label)Field(pane, "_status")!).Text, UiText.Get(language, "ChatReplaceUnavailable"));
                        var session = (DocumentChatSession)Field(pane, "_session")!;
                        Assert.AreEqual("中文 **AI**", session.History.Last().Content, "Completed translations share CJK cleanup and preserve raw Markdown.");
                        Button(writes, action == ChatWriteAction.ReplaceSelection ? "ChatReplace" : "ChatInsertBelow").PerformClick();
                        Assert.AreEqual(action, applied.Last()); Assert.IsFalse(writes.Enabled);
                        Button(writes, "ChatInsertBelow").PerformClick();
                    }
                    Assert.AreEqual(3, applied.Count);
                    Assert.AreEqual(3, history.Store.Read(history.Store.List(out var unreadable).Single().SessionId)!.Turns.Count);
                    Assert.AreEqual(0, unreadable);
                    window.ClientSize = new Size(440, 760); Application.DoEvents();
                    var dir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../../../TestResults/chat-translate-2.3.8"));
                    Directory.CreateDirectory(dir);
                    using (var bitmap = new Bitmap(pane.Width, pane.Height)) { pane.DrawToBitmap(bitmap, pane.ClientRectangle); bitmap.Save(Path.Combine(dir, host + "-" + language + ".png")); }
                }
            });
        }

        [DataTestMethod]
        [DataRow(true, true)] [DataRow(true, false)] [DataRow(false, true)] [DataRow(false, false)]
        public void WritePolicy_MatchesRewrite_RejectsStartEmptyAndMissingWriter(bool selected, bool canReplace)
        {
            var calls = 0;
            var context = new ChatContext("source", selected, (a, text) => calls++, canReplace);
            foreach (ChatWriteAction action in Enum.GetValues(typeof(ChatWriteAction)))
            {
                Assert.AreEqual(context.Allows(ChatMode.Rewrite, action, "result"), context.Allows(ChatMode.Translate, action, "result"));
                Assert.IsFalse(context.Allows(ChatMode.Translate, action, "  "));
                Assert.IsFalse(new ChatContext("source", selected).Allows(ChatMode.Translate, action, "result"));
            }
            Assert.ThrowsException<InvalidOperationException>(() => context.Apply(ChatMode.Translate, ChatWriteAction.InsertAtStart, "result"));
            Assert.AreEqual(0, calls);
            context.Apply(ChatMode.Translate, ChatWriteAction.InsertBelow, "result"); Assert.AreEqual(1, calls);
        }

        [DataTestMethod]
        [DataRow("Word", "zh-CN")] [DataRow("Word", "zh-HK")] [DataRow("Word", "en")] [DataRow("Word", "ko")] [DataRow("Word", "ja")]
        [DataRow("Excel", "zh-CN")] [DataRow("Excel", "zh-HK")] [DataRow("Excel", "en")] [DataRow("Excel", "ko")] [DataRow("Excel", "ja")]
        [DataRow("PowerPoint", "zh-CN")] [DataRow("PowerPoint", "zh-HK")] [DataRow("PowerPoint", "en")] [DataRow("PowerPoint", "ko")] [DataRow("PowerPoint", "ja")]
        public void SettingsForms_RoundTripTranslatePrompt_AndPreserveOtherSettings(string host, string language)
        {
            DocumentChatHostTests.Sta(() =>
            {
                var path = Path.Combine(Path.GetTempPath(), "OfficeTranslate-Translate-" + Guid.NewGuid() + ".json");
                var store = new SettingsStore(); typeof(SettingsStore).GetField("_path", Hidden)!.SetValue(store, path);
                var custom = "Translate prompt\r\n" + new string('文', 6000);
                try
                {
                    File.WriteAllText(path, "{}");
                    Assert.AreEqual(ChatPrompts.Default("zh-CN", ChatMode.Translate), store.Load().ChatTranslateSystemPrompt);
                    var settings = new TranslationSettings { UiLanguage = language, ChatUseTranslationSettings = true, Glossary = "term=fixture" };
                    store.Save(settings);
                    var assembly = DocumentChatHostTests.Host(host);
                    using (var form = (Form)Activator.CreateInstance(assembly.GetType("OfficeTranslate.WordAddIn.ChatSettingsForm", true)!, store)!)
                    {
                        form.Show(); Application.DoEvents();
                        var box = All(form).OfType<TextBox>().Single(b => (string?)b.Tag == "ChatTranslatePrompt");
                        Assert.AreEqual(ChatPrompts.Default(language, ChatMode.Translate), box.Text);
                        Assert.AreEqual(0, box.MaxLength);
                        Assert.IsTrue(All(form).OfType<Label>().Any(l => l.Text == UiText.Get(language, "ChatTranslatePrompt")));
                        box.Text = custom;
                        ((ScrollableControl)form.Controls[0]).ScrollControlIntoView((Button)form.AcceptButton!);
                        ((Button)form.AcceptButton!).PerformClick();
                    }
                    Assert.AreEqual(custom, store.Load().ChatTranslateSystemPrompt);
                    using (var form = (Form)Activator.CreateInstance(assembly.GetType("OfficeTranslate.WordAddIn.SettingsForm", true)!, store)!)
                        store.Save((TranslationSettings)Call(form, "BuildSettings", true)!);
                    var restored = store.Load();
                    Assert.AreEqual(custom, restored.ChatTranslateSystemPrompt); Assert.IsTrue(restored.ChatUseTranslationSettings);
                    Assert.AreEqual(settings.Glossary, restored.Glossary); Assert.AreEqual(settings.ChatSummarySystemPrompt, restored.ChatSummarySystemPrompt);
                    Assert.AreEqual(settings.ChatRewriteSystemPrompt, restored.ChatRewriteSystemPrompt);
                    using (var form = (Form)Activator.CreateInstance(assembly.GetType("OfficeTranslate.WordAddIn.ChatSettingsForm", true)!, store)!)
                        Assert.AreEqual(custom, All(form).OfType<TextBox>().Single(b => (string?)b.Tag == "ChatTranslatePrompt").Text);
                }
                finally { File.Delete(path); }
            });
        }

        [DataTestMethod]
        [DataRow("zh-CN")] [DataRow("zh-HK")] [DataRow("en")] [DataRow("ko")] [DataRow("ja")]
        public void Defaults_LocalizeButDoNotForceUiAnswerLanguage_AndCustomPromptsSurvive(string language)
        {
            foreach (var key in new[] { "ChatTranslate", "ChatTranslatePrompt", "ChatTranslateRequest", "TranslateChangelog" })
                Assert.AreNotEqual(key, UiText.Get(language, key));
            StringAssert.Contains(UiText.Get(language, "ChatTranslateRequest"), "{0}");
            var json = new JavaScriptSerializer();
            var settings = new TranslationSettings { UiLanguage = language };
            var restored = json.Deserialize<TranslationSettings>(json.Serialize(settings)); restored.UiLanguage = language == "en" ? "ja" : "en";
            Assert.AreEqual(ChatPrompts.Default(restored.UiLanguage, ChatMode.Translate), restored.ChatTranslateSystemPrompt);
            Assert.AreNotEqual(ChatPrompts.Safety(language), ChatPrompts.Safety(language, ChatMode.Translate));
            settings.ChatTranslateSystemPrompt = "My custom prompt"; settings.UiLanguage = restored.UiLanguage;
            Assert.AreEqual("My custom prompt", json.Deserialize<TranslationSettings>(json.Serialize(settings)).ChatTranslateSystemPrompt);
        }

        [TestMethod]
        public void ErrorDetails_RedactCustomTranslatePromptAndIndividualLines()
        {
            var settings = new TranslationSettings { ChatTranslateSystemPrompt = "private first line\nprivate second line" };
            var details = ChatErrors.Describe(new IOException("failure: private second line"), settings);
            Assert.IsFalse(details.Contains("private second line")); StringAssert.Contains(details, "[redacted]");
        }

        private sealed class SseHandler : HttpMessageHandler
        {
            internal readonly List<Dictionary<string, object>> Requests = new List<Dictionary<string, object>>();
            internal bool Interrupt;
            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
            {
                Requests.Add((Dictionary<string, object>)new JavaScriptSerializer().DeserializeObject(await request.Content!.ReadAsStringAsync()));
                var stream = (bool)Requests.Last()["stream"];
                var body = stream ? "data: {\"choices\":[{\"delta\":{\"content\":\"translated\"}}]}\n\n" + (Interrupt ? "" : "data: [DONE]\n\n") :
                    "{\"choices\":[{\"message\":{\"content\":\"recovered translation\"},\"finish_reason\":\"stop\"}]}";
                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(body, Encoding.UTF8, stream ? "text/event-stream" : "application/json") };
            }
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public async Task Translate_UsesRealSseTransportAndFallback_PreservesRequestAndHistory(bool interrupt)
        {
            var handler = new SseHandler { Interrupt = interrupt };
            var settings = new TranslationSettings { Provider = ProviderKind.OpenAiCompatible, BaseUrl = "http://fixture/v1", Model = "fixture", TargetLanguage = "English" };
            var session = new DocumentChatSession(); var output = new List<string>();
            var question = string.Format(UiText.Get("en", "ChatTranslateRequest"), settings.TargetLanguage);
            using (var client = new TranslationClient(handler))
            {
                var reply = await session.SendAsync(new ChatContext("source", true), question, ChatMode.Translate, settings,
                    client.ChatAsync, output.Add, (i, n) => { }, CancellationToken.None);
                Assert.AreEqual(interrupt, reply.UsedFallback); Assert.AreEqual(interrupt ? 2 : 1, handler.Requests.Count);
                Assert.AreEqual(interrupt ? "recovered translation" : "translated", reply.Text);
                Assert.AreEqual(reply.Text, output.Last()); Assert.AreEqual(reply.Text, session.History.Last().Content);
            }
            foreach (var body in handler.Requests)
            {
                var messages = (object[])body["messages"];
                StringAssert.StartsWith((string)((Dictionary<string, object>)messages.Last())["content"], question);
                StringAssert.StartsWith((string)((Dictionary<string, object>)messages.First())["content"], settings.ChatTranslateSystemPrompt);
            }
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public async Task Translate_LongContextReusesExistingReductionAndTranslationSettings(bool enabled)
        {
            var settings = new TranslationSettings { ChatContextTokens = 4096, ChatOutputTokens = 512, ChatUseTranslationSettings = enabled,
                TranslationStyle = "Technology", CustomInstructions = "fixture-extra", Glossary = "fixture-term" };
            var maps = 0; var finals = 0;
            var reply = await new DocumentChatSession().SendAsync(new ChatContext(new string('文', 5000), false), "Translate into English", ChatMode.Translate, settings,
                (messages, s, n, output, token) =>
                {
                    Assert.IsTrue(ChatBudget.Cost(messages) + n <= settings.ChatContextTokens);
                    if (n == 256) { maps++; StringAssert.StartsWith(messages[0].Content, ChatPrompts.Map(settings.UiLanguage)); }
                    else { finals++; StringAssert.StartsWith(messages[0].Content, settings.ChatTranslateSystemPrompt); }
                    Assert.AreEqual(enabled, messages[0].Content.Contains(settings.CustomInstructions));
                    Assert.AreEqual(enabled, messages[0].Content.Contains(settings.Glossary));
                    Assert.AreEqual(enabled, messages[0].Content.Contains(TranslationStyleCatalog.GetPrompt(settings.TranslationStyle)));
                    return Task.FromResult(new ChatReply("short result", false));
                }, _ => { }, (i, n) => { }, CancellationToken.None);
            Assert.IsTrue(maps > 1); Assert.AreEqual(1, finals); Assert.IsTrue(reply.UsedSummaries);
        }
    }
}
