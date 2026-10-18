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
    public class ChatWritebackFixTests
    {
        private const BindingFlags Hidden = BindingFlags.Instance | BindingFlags.NonPublic;
        private static object? Field(object target, string name) => target.GetType().GetField(name, Hidden)!.GetValue(target);
        private static void Call(object target, string method, params object[] args) => target.GetType().GetMethod(method, Hidden)!.Invoke(target, args);
        private static IEnumerable<Control> All(Control parent) => parent.Controls.Cast<Control>().SelectMany(c => new[] { c }.Concat(All(c)));
        private static Button Button(Control parent, string key) => All(parent).OfType<Button>().Single(b => (string?)b.Tag == key);
        private static void Pump(Func<bool> done)
        {
            var deadline = DateTime.UtcNow.AddSeconds(8);
            while (!done() && DateTime.UtcNow < deadline) { Application.DoEvents(); Thread.Sleep(5); }
            Assert.IsTrue(done(), "Chat operation timed out.");
        }

        private sealed class PaneFixture : IDisposable
        {
            private readonly HistorySandbox _history = new HistorySandbox();
            internal readonly UserControl Pane;
            private readonly Form _window = new Form { ClientSize = new Size(440, 760) };
            internal ChatContext Context = new ChatContext("selected source", true, (a, t) => { }, true);
            internal PaneFixture(string host, string language)
            {
                Pane = (UserControl)Activator.CreateInstance(DocumentChatHostTests.Host(host).GetType("OfficeTranslate.WordAddIn.DocumentChatPane", true)!)!;
                Pane.GetType().GetProperty("HistoryStore", Hidden)!.SetValue(Pane, _history.Store);
                var settings = new TranslationSettings { UiLanguage = language };
                Call(Pane, "Initialize", (Func<ChatContext>)(() => Context), (Func<TranslationSettings>)(() => settings));
                Pane.Dock = DockStyle.Fill; _window.Controls.Add(Pane); _window.Show(); Application.DoEvents();
            }
            internal TextBox Input => (TextBox)Field(Pane, "_input")!;
            internal FlowLayoutPanel Writes => (FlowLayoutPanel)Field(Pane, "_writes")!;
            internal void Transport(ChatTransport transport) => Pane.GetType().GetProperty("Transport", Hidden)!.SetValue(Pane, transport);
            public void Dispose() { _window.Close(); _window.Dispose(); Pane.Dispose(); _history.Dispose(); }
        }

        [DataTestMethod]
        [DataRow("Word", "zh-CN")] [DataRow("Word", "zh-HK")] [DataRow("Word", "en")] [DataRow("Word", "ko")] [DataRow("Word", "ja")]
        [DataRow("Excel", "zh-CN")] [DataRow("Excel", "zh-HK")] [DataRow("Excel", "en")] [DataRow("Excel", "ko")] [DataRow("Excel", "ja")]
        [DataRow("PowerPoint", "zh-CN")] [DataRow("PowerPoint", "zh-HK")] [DataRow("PowerPoint", "en")] [DataRow("PowerPoint", "ko")] [DataRow("PowerPoint", "ja")]
        public void Question_SendClearsBeforeStream_CompletionPreservesNextDraft_AndBothWritesWork(string host, string language)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new PaneFixture(host, language))
                {
                    foreach (var action in new[] { ChatWriteAction.ReplaceSelection, ChatWriteAction.InsertBelow })
                    {
                        var applied = new List<ChatWriteAction>();
                        f.Context = new ChatContext("selected source", true, (a, text) =>
                        { Assert.AreEqual("韩国翻译 **AI**", text); applied.Add(a); }, true);
                        const string question = "请 翻 译\r\n保留原意";
                        f.Input.Text = question;
                        Action<string> output = null!;
                        var completion = new TaskCompletionSource<ChatReply>(TaskCreationOptions.RunContinuationsAsynchronously);
                        var called = false;
                        f.Transport((messages, settings, n, o, token) =>
                        {
                            called = true;
                            Assert.AreEqual("", f.Input.Text, "Input must clear before the transport starts.");
                            StringAssert.StartsWith(messages.Last().Content, question + "\n<document>");
                            output = o; return completion.Task;
                        });
                        Button(f.Pane, "ChatSend").PerformClick();
                        Assert.IsTrue(called); Assert.AreEqual("", f.Input.Text);
                        Assert.IsFalse(completion.Task.IsCompleted);
                        output("韩 国 翻 译"); Call(f.Pane, "FlushOutput");
                        Assert.AreEqual("", f.Input.Text, "Input must stay empty during streaming.");
                        Assert.AreEqual(0, f.Writes.Controls.Count);
                        f.Input.Text = "draft for next question";
                        completion.SetResult(new ChatReply("韩 国 翻 译 **AI**", false));
                        Pump(() => Field(f.Pane, "_pending") == null);
                        Assert.AreEqual("draft for next question", f.Input.Text, "Completion must not erase a new draft.");
                        CollectionAssert.AreEqual(new[] { "ChatReplace", "ChatInsertBelow" }, f.Writes.Controls.Cast<Control>().Select(c => (string)c.Tag).ToArray());
                        Assert.IsTrue(f.Writes.Controls.Cast<Control>().All(b => b.Visible && b.Enabled));
                        Assert.AreEqual(UiText.Get(language, "ChatReplace"), Button(f.Writes, "ChatReplace").Text);
                        Assert.AreEqual(UiText.Get(language, "ChatInsertBelow"), Button(f.Writes, "ChatInsertBelow").Text);
                        Assert.AreEqual(UiText.Get(language, "ChatDone"), ((Label)Field(f.Pane, "_status")!).Text);
                        var session = (DocumentChatSession)Field(f.Pane, "_session")!;
                        Assert.AreEqual(question, session.History[session.History.Count - 2].Content);
                        Assert.AreEqual("韩国翻译 **AI**", session.History.Last().Content);
                        if (action == ChatWriteAction.ReplaceSelection)
                        {
                            var directory = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../../../TestResults/chat-writeback-2.3.8"));
                            Directory.CreateDirectory(directory);
                            using (var bitmap = new Bitmap(f.Pane.Width, f.Pane.Height))
                            { f.Pane.DrawToBitmap(bitmap, f.Pane.ClientRectangle); bitmap.Save(Path.Combine(directory, host + "-" + language + ".png")); }
                        }
                        Assert.AreEqual(0, applied.Count, "Completion must never write automatically.");
                        var confirm = false; var confirmations = 0;
                        f.Pane.GetType().GetProperty("ConfirmWrite", Hidden)!.SetValue(f.Pane, (Func<ChatWriteAction, bool>)(a =>
                        { Assert.AreEqual(action, a); confirmations++; return confirm; }));
                        var write = Button(f.Writes, action == ChatWriteAction.ReplaceSelection ? "ChatReplace" : "ChatInsertBelow");
                        write.PerformClick(); Assert.AreEqual(0, applied.Count);
                        confirm = true; write.PerformClick();
                        CollectionAssert.AreEqual(new[] { action }, applied);
                        Assert.AreEqual(2, confirmations); Assert.IsFalse(f.Writes.Enabled);
                        Button(f.Writes, "ChatInsertBelow").PerformClick(); Assert.AreEqual(1, applied.Count);
                    }
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word", ChatMode.Question)] [DataRow("Word", ChatMode.Summarize)] [DataRow("Word", ChatMode.Rewrite)] [DataRow("Word", ChatMode.Translate)]
        [DataRow("Excel", ChatMode.Question)] [DataRow("Excel", ChatMode.Summarize)] [DataRow("Excel", ChatMode.Rewrite)] [DataRow("Excel", ChatMode.Translate)]
        [DataRow("PowerPoint", ChatMode.Question)] [DataRow("PowerPoint", ChatMode.Summarize)] [DataRow("PowerPoint", ChatMode.Rewrite)] [DataRow("PowerPoint", ChatMode.Translate)]
        public void AllModes_ClearAtSend_AndFilterButtonsWithReplacementAndWriterGuards(string host, ChatMode mode)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new PaneFixture(host, "en"))
                {
                    foreach (var context in new[] {
                        new ChatContext("selection", true, (a, t) => { }, true),
                        new ChatContext("selection", true, (a, t) => { }, false),
                        new ChatContext("whole document", false, (a, t) => { }, true),
                        new ChatContext("no writer", true, canReplace: true) })
                    {
                        f.Context = context; f.Input.Text = "pending instruction";
                        var completion = new TaskCompletionSource<ChatReply>(TaskCreationOptions.RunContinuationsAsynchronously);
                        var called = false;
                        f.Transport((m, s, n, o, t) => { called = true; return completion.Task; });
                        var send = mode == ChatMode.Question ? "ChatSend" : mode == ChatMode.Summarize ? "ChatSummary" : mode == ChatMode.Rewrite ? "ChatRewrite" : "ChatTranslate";
                        Button(f.Pane, send).PerformClick();
                        Assert.IsTrue(called); Assert.AreEqual("", f.Input.Text);
                        Assert.IsFalse(completion.Task.IsCompleted); Assert.AreEqual(0, f.Writes.Controls.Count);
                        completion.SetResult(new ChatReply("answer", false));
                        Pump(() => Field(f.Pane, "_pending") == null);
                        var summary = mode == ChatMode.Summarize;
                        CollectionAssert.AreEqual(summary ? new[] { "ChatInsertStart" } : new[] { "ChatReplace", "ChatInsertBelow" },
                            f.Writes.Controls.Cast<Control>().Select(c => (string)c.Tag).ToArray());
                        if (summary) Assert.AreEqual(context.Write != null, Button(f.Writes, "ChatInsertStart").Enabled);
                        else
                        {
                            Assert.AreEqual(context.Write != null && context.CanReplace, Button(f.Writes, "ChatReplace").Enabled);
                            Assert.AreEqual(context.Write != null, Button(f.Writes, "ChatInsertBelow").Enabled);
                            if (!context.CanReplace) StringAssert.Contains(((Label)Field(f.Pane, "_status")!).Text, UiText.Get("en", "ChatReplaceUnavailable"));
                            Assert.IsFalse(context.Allows(mode, ChatWriteAction.InsertAtStart, "answer"));
                        }
                        foreach (ChatWriteAction action in Enum.GetValues(typeof(ChatWriteAction)))
                        {
                            Assert.IsFalse(context.Allows(mode, action, " "));
                            Assert.IsFalse(new ChatContext("no writer", true, canReplace: true).Allows(mode, action, "answer"));
                        }
                    }
                }
            });
        }
    }
}
