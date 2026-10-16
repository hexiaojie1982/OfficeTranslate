using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using OfficeTranslate.WordAddIn;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ChatUiRegressionTests
    {
        private const BindingFlags Hidden = BindingFlags.NonPublic | BindingFlags.Instance;
        private static object Field(object target, string name) => target.GetType().GetField(name, Hidden)!.GetValue(target)!;
        private static void Call(object target, string method, params object[] args) => target.GetType().GetMethod(method, Hidden)!.Invoke(target, args);
        private static IEnumerable<Control> Descendants(Control parent) => parent.Controls.Cast<Control>().SelectMany(c => new[] { c }.Concat(Descendants(c)));
        private static Button Button(Control parent, string key) => Descendants(parent).OfType<Button>().Single(c => (string?)c.Tag == key);
        private static void Pump(Func<bool> done)
        {
            var end = DateTime.UtcNow.AddSeconds(5);
            while (!done() && DateTime.UtcNow < end) { Application.DoEvents(); Thread.Sleep(5); }
            Assert.IsTrue(done(), "UI operation timed out");
        }
        private sealed class PaneFixture : IDisposable
        {
            internal readonly TranslationSettings Settings;
            internal readonly UserControl Pane;
            internal readonly Form Window = new Form { ClientSize = new Size(440, 760) };
            internal ChatContext Context = new ChatContext("document", true, (a, t) => { }, true);
            internal int Reads;
            internal PaneFixture(string host, string language)
            {
                Settings = new TranslationSettings { UiLanguage = language };
                Pane = (UserControl)Activator.CreateInstance(DocumentChatHostTests.Host(host).GetType("OfficeTranslate.WordAddIn.DocumentChatPane", true)!)!;
                Call(Pane, "Initialize", (Func<ChatContext>)(() => { Reads++; return Context; }), (Func<TranslationSettings>)(() => Settings));
                Pane.Dock = DockStyle.Fill; Window.Controls.Add(Pane); Window.Show(); Application.DoEvents();
            }
            internal Control Transcript => (Control)Field(Pane, "_transcript");
            internal RichTextBox Selected = null!;
            internal void SelectText(string text)
            {
                Selected = Descendants(Transcript).OfType<RichTextBox>().Last(b => b.Text.Contains(text));
                Selected.Select(Selected.Text.IndexOf(text, StringComparison.Ordinal), text.Length);
            }
            internal void Transport(ChatTransport transport) => Pane.GetType().GetProperty("Transport", Hidden)!.SetValue(Pane, transport);
            public void Dispose() { Window.Close(); Window.Dispose(); Pane.Dispose(); }
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void Streaming_BatchesBackgroundTokens_OnlyChangesTranscript_AndReplacesFallbackTail(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new PaneFixture(host, "en"))
                {
                    Action<string> output = null!;
                    var completion = new TaskCompletionSource<ChatReply>();
                    f.Transport((m, s, n, o, t) => { output = o; return completion.Task; });
                    Call(f.Pane, "Start", ChatMode.Rewrite);
                    var prefix = f.Transcript.Text;
                    var changes = 0; var layouts = 0; var invalidations = 0;
                    f.Transcript.TextChanged += (s, e) => changes++;
                    f.Pane.Layout += (s, e) => layouts++;
                    f.Pane.Controls[0].Layout += (s, e) => layouts++;
                    f.Pane.Controls[0].Invalidated += (s, e) => invalidations++;
                    var commands = (Control)Field(f.Pane, "_commands");
                    commands.Invalidated += (s, e) => invalidations++;
                    ((Control)Field(f.Pane, "_context")).Invalidated += (s, e) => invalidations++;
                    Task.Run(() => { for (var i = 1; i <= 600; i++) output(new string('文', i) + "😀"); }).GetAwaiter().GetResult();
                    Assert.AreEqual(0, changes, "Token callbacks must not mutate any UI, even from a background thread.");
                    Call(f.Pane, "FlushOutput");
                    Assert.AreEqual(1, changes); Assert.AreEqual(prefix + new string('文', 600) + "😀", f.Transcript.Text);
                    output(new string('文', 600) + "😀\r\nnext"); Call(f.Pane, "FlushOutput");
                    Assert.AreEqual(prefix + new string('文', 600) + "😀\nnext", f.Transcript.Text);
                    Assert.AreEqual(0, layouts); Assert.AreEqual(0, invalidations, "Streaming must not invalidate the command row or context preview.");
                    output(""); output("recovered"); Call(f.Pane, "FlushOutput");
                    Assert.AreEqual(prefix + "recovered", f.Transcript.Text);
                    Assert.AreEqual(UiText.Get("en", "ChatRecovering"), ((Label)Field(f.Pane, "_status")).Text);
                    completion.SetResult(new ChatReply("final answer without a final token callback", true));
                    Pump(() => FieldOrNull(f.Pane, "_pending") == null);
                    StringAssert.Contains(f.Transcript.Text, "final answer without a final token callback");
                    Assert.IsFalse(f.Transcript.Text.Contains("recovered"));
                    Assert.IsFalse(((System.Windows.Forms.Timer)Field(f.Pane, "_streamTimer")).Enabled);
                    var finished = f.Transcript.Text; output("late callback"); Call(f.Pane, "FlushOutput");
                    Assert.AreEqual(finished, f.Transcript.Text);
                    Console.WriteLine(host + ": 600 background snapshots -> 1 transcript update; sibling invalidations=0, pane layouts=0.");
                }
            });
        }
        private static object? FieldOrNull(object target, string name) => target.GetType().GetField(name, Hidden)!.GetValue(target);

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void Markdown_StreamsBeforeCompletion_KeepsQuestionPlain_AndExportsOriginalSource(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new PaneFixture(host, "en"))
                {
                    const string question = "## user **question** `literal`";
                    const string answer = "## 标题\n**bold** *italic* `inline`\n\n- item\n1. numbered\n> quote\n```cs\nvar x = 1;\n```\nfinal 😀";
                    ((TextBox)Field(f.Pane, "_input")).Text = question;
                    Action<string> output = null!; var completion = new TaskCompletionSource<ChatReply>();
                    f.Transport((m, s, n, o, t) => { output = o; return completion.Task; });
                    Call(f.Pane, "Start", ChatMode.Question);
                    Assert.AreEqual(50, ((System.Windows.Forms.Timer)Field(f.Pane, "_streamTimer")).Interval);
                    var changes = 0; var layouts = 0; var invalidations = 0;
                    f.Transcript.TextChanged += (s, e) => changes++;
                    f.Pane.Controls[0].Layout += (s, e) => layouts++;
                    f.Pane.Controls[0].Invalidated += (s, e) => invalidations++;
                    Task.Run(() => { for (var i = 0; i < 600; i++) output("## 标题\n**bo"); }).GetAwaiter().GetResult();
                    Assert.AreEqual(0, changes); Call(f.Pane, "FlushOutput"); Assert.AreEqual(1, changes);
                    StringAssert.Contains(f.Transcript.Text, "**bo");
                    output("## 标题\n**bold**");
                    Pump(() => f.Transcript.Text.EndsWith("bold", StringComparison.Ordinal));
                    Assert.IsFalse(completion.Task.IsCompleted, "Rendering must be progressive, before the request completes.");
                    f.SelectText("bold");
                    Assert.IsTrue(f.Selected.SelectionFont!.Bold);
                    output(answer); Call(f.Pane, "FlushOutput");
                    Assert.AreEqual(0, layouts); Assert.AreEqual(0, invalidations);
                    StringAssert.Contains(f.Transcript.Text, question);
                    f.SelectText("**question**");
                    Assert.IsFalse(f.Selected.SelectionFont!.Bold);
                    f.SelectText("var x");
                    Assert.AreEqual(Color.FromArgb(242, 244, 247), f.Selected.SelectionBackColor);
                    output(""); output("fallback plain"); Call(f.Pane, "FlushOutput");
                    Assert.IsFalse(f.Transcript.Text.Contains("var x"));
                    f.SelectText("fallback");
                    Assert.IsFalse(f.Selected.SelectionFont!.Bold); Assert.AreEqual(0, f.Selected.SelectionIndent);
                    completion.SetResult(new ChatReply(answer, true)); Pump(() => FieldOrNull(f.Pane, "_pending") == null);
                    Assert.IsFalse(f.Transcript.Text.Contains("fallback plain"));
                    foreach (var language in UiText.LanguageCodes)
                    {
                        f.Settings.UiLanguage = language; Call(f.Pane, "RefreshLanguage");
                        StringAssert.Contains(f.Transcript.Text, question);
                        f.SelectText("bold");
                        Assert.IsTrue(f.Selected.SelectionFont!.Bold);
                    }
                    SaveImage(f, host + "-markdown");
                    var history = ((DocumentChatSession)Field(f.Pane, "_session")).History;
                    Assert.AreEqual(answer, history[1].Content);
                    foreach (var markdown in new[] { true, false })
                    {
                        var path = Path.GetTempFileName();
                        try
                        {
                            ExportDialog(f, d => { d.FileName = path; d.FilterIndex = markdown ? 1 : 2; return DialogResult.OK; });
                            Button(f.Pane, "ChatExport").PerformClick();
                            var exported = File.ReadAllText(path);
                            Assert.AreEqual(ChatTranscriptExport.Format(history, markdown, k => UiText.Get(f.Settings.UiLanguage, k)), exported);
                            StringAssert.Contains(exported, "**bold** *italic* `inline`"); StringAssert.Contains(exported, "```cs");
                            Assert.IsFalse(exported.Contains(@"{\rtf"));
                        }
                        finally { File.Delete(path); }
                    }
                    Button(f.Pane, "ChatClear").PerformClick(); Assert.AreEqual(UiText.Get(f.Settings.UiLanguage, "ChatEmpty"), f.Transcript.Text);
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void Streaming_TimerDisplaysProgress_AndCancellationCannotCommitOrReviveOutput(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new PaneFixture(host, "zh-CN"))
                {
                    var count = 0; Action<string> output = null!;
                    f.Transport((m, s, n, o, token) => Task.Run(async () =>
                    {
                        output = o;
                        for (var i = 1; i <= 80; i++) { token.ThrowIfCancellationRequested(); o("逐步输出" + new string('字', i)); await Task.Delay(8); }
                        await Task.Delay(5000, token); return new ChatReply("not expected", false);
                    }));
                    f.Transcript.TextChanged += (s, e) => count++;
                    Call(f.Pane, "Start", ChatMode.Summarize);
                    Pump(() => f.Transcript.Text.Contains("逐步输出"));
                    ((Button)Field(f.Pane, "_cancel")).PerformClick();
                    Pump(() => FieldOrNull(f.Pane, "_pending") == null);
                    Assert.AreEqual(UiText.Get("zh-CN", "ChatCancelled"), ((Label)Field(f.Pane, "_status")).Text);
                    Assert.AreEqual(0, ((DocumentChatSession)Field(f.Pane, "_session")).History.Count);
                    Assert.AreEqual(0, ((FlowLayoutPanel)Field(f.Pane, "_writes")).Controls.Count);
                    var cancelled = f.Transcript.Text; output("stale"); Call(f.Pane, "FlushOutput");
                    Assert.AreEqual(cancelled, f.Transcript.Text); Assert.IsTrue(count >= 2);
                    f.Settings.UiLanguage = "en"; Call(f.Pane, "RefreshLanguage");
                    StringAssert.Contains(f.Transcript.Text, "Assistant\n");
                    Assert.AreEqual(UiText.Get("en", "ChatCancelled"), ((Label)Field(f.Pane, "_status")).Text);
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word", "zh-CN")] [DataRow("Word", "en")]
        [DataRow("Excel", "zh-CN")] [DataRow("Excel", "en")]
        [DataRow("PowerPoint", "zh-CN")] [DataRow("PowerPoint", "en")]
        public void Buttons_AreGroupedAlignedAndUnclipped_ContextPreviewIsExplained_AndSendRereads(string host, string language)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new PaneFixture(host, language))
                {
                    foreach (var width in new[] { 320, 440, 600 })
                    {
                        f.Window.ClientSize = new Size(width, 760); Application.DoEvents();
                        var summary = Button(f.Pane, "ChatSummary"); var rewrite = Button(f.Pane, "ChatRewrite");
                        Assert.AreSame(summary.Parent, rewrite.Parent);
                        Assert.AreNotSame(summary.Parent, Button(f.Pane, "ChatClear").Parent);
                        foreach (var button in Descendants(f.Pane).OfType<Button>())
                        {
                            Assert.AreEqual(32, button.Height); Assert.AreEqual(FlatStyle.Flat, button.FlatStyle);
                            Assert.IsTrue(button.Right <= button.Parent!.ClientSize.Width, host + " " + language + " width=" + width + " " + button.Text + " overflows");
                            Assert.IsTrue(button.Bottom <= button.Parent.ClientSize.Height);
                            Assert.AreEqual(summary.Font.Name, button.Font.Name);
                        }
                        Assert.AreEqual(Color.FromArgb(45, 108, 223), summary.BackColor);
                        Assert.AreEqual(Color.FromArgb(231, 236, 243), Button(f.Pane, "ChatClear").BackColor);
                        SaveImage(f, host + "-" + language + "-" + width);
                    }
                    var tips = (ToolTip)Field(f.Pane, "_tips");
                    Assert.AreEqual(UiText.Get(language, "ChatRefreshTip"), tips.GetToolTip(Button(f.Pane, "ChatRefresh")));
                    var before = f.Reads;
                    f.Context = new ChatContext("updated preview", true);
                    Button(f.Pane, "ChatRefresh").PerformClick(); Assert.AreEqual(before + 1, f.Reads);
                    StringAssert.Contains(((Label)Field(f.Pane, "_context")).Text, "15");
                    f.Context = new ChatContext("changed after preview", false);
                    string sent = "";
                    f.Transport((m, s, n, o, t) => { sent = m.Last().Content; return Task.FromResult(new ChatReply("done", false)); });
                    Call(f.Pane, "Start", ChatMode.Summarize); Pump(() => FieldOrNull(f.Pane, "_pending") == null);
                    StringAssert.Contains(sent, "changed after preview"); Assert.IsFalse(sent.Contains("updated preview"));
                }
            });
        }
        private static void SaveImage(PaneFixture f, string name)
        {
            var dir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../../../TestResults/chat-bubbles-2.3.2/screenshots"));
            Directory.CreateDirectory(dir);
            using (var bitmap = new Bitmap(f.Pane.Width, f.Pane.Height)) { f.Pane.DrawToBitmap(bitmap, f.Pane.ClientRectangle); bitmap.Save(Path.Combine(dir, name + ".png")); }
        }

        private static Control[] Bubbles(PaneFixture f) => Descendants(f.Transcript).Where(c => c.GetType().Name == "ChatBubbleRow").ToArray();
        private static object Property(object target, string name) => target.GetType().GetProperty(name, Hidden)!.GetValue(target)!;
        private static RichTextBox BubbleContent(Control row) => (RichTextBox)Property(row, "Content");

        [DataTestMethod]
        [DataRow("Word", "zh-CN")] [DataRow("Excel", "zh-CN")] [DataRow("PowerPoint", "zh-CN")]
        [DataRow("Word", "en")] [DataRow("Excel", "ko")] [DataRow("PowerPoint", "ja")]
        public void Bubbles_RenderAlignedColorsTimesAndMarkdown_AndWrapWithinNarrowViewports(string host, string language)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new PaneFixture(host, language))
                {
                    const string question = "**hi**";
                    var answer = "# 标题\n**bold**\n> quote\n```cs\n" + new string('x', 180) + "\n```\n" + string.Join("\n", Enumerable.Repeat("中文 English 😀 long text wraps automatically.", 35)) + "\nlast line";
                    ((TextBox)Field(f.Pane, "_input")).Text = question;
                    f.Transport((m, s, n, o, t) => Task.FromResult(new ChatReply(answer, false)));
                    Call(f.Pane, "Start", ChatMode.Question); Pump(() => FieldOrNull(f.Pane, "_pending") == null);
                    var rows = Bubbles(f); Assert.AreEqual(2, rows.Length);
                    var user = BubbleContent(rows[0]); var ai = BubbleContent(rows[1]);
                    Assert.AreEqual(question, user.Text); user.Select(0, 6); Assert.IsFalse(user.SelectionFont!.Bold);
                    Assert.AreEqual(Color.White, user.SelectionColor);
                    Assert.AreEqual(Color.FromArgb(45, 108, 223), user.BackColor);
                    Assert.AreEqual(user.BackColor, user.SelectionBackColor, "Plain user text must not acquire white paragraph shading.");
                    Assert.AreEqual(Color.White, ai.BackColor);
                    ai.Select(ai.Text.IndexOf("bold", StringComparison.Ordinal), 4); Assert.IsTrue(ai.SelectionFont!.Bold);
                    Assert.AreEqual(Color.FromArgb(32, 55, 88), ai.SelectionColor);
                    Assert.AreEqual(ai.BackColor, ai.SelectionBackColor, "Plain AI text must not acquire dark paragraph shading.");
                    Assert.AreEqual(Color.FromArgb(240, 244, 250), f.Pane.BackColor);
                    foreach (var width in new[] { 260, 320, 440, 600 })
                    {
                        f.Window.ClientSize = new Size(width, 760); Application.DoEvents();
                        foreach (var row in rows)
                        {
                            var rect = (Rectangle)Property(row, "BubbleBounds"); var body = BubbleContent(row);
                            Assert.IsTrue(rect.Width <= row.Width * .8, "Bubble exceeds 80% at " + width);
                            Assert.IsTrue(rect.Left >= 0 && rect.Right <= row.Width);
                            Assert.IsTrue(body.WordWrap); Assert.AreEqual(RichTextBoxScrollBars.None, body.ScrollBars);
                            body.Select(0, 0); body.ScrollToCaret();
                            var last = body.GetPositionFromCharIndex(body.TextLength - 1);
                            Assert.IsTrue(last.Y >= 0 && last.Y + body.Font.Height <= body.Height, "Final line clipped at " + width + ": " + last + " / " + body.Size);
                            var time = row.Controls.OfType<Label>().Single();
                            var stamp = (DateTimeOffset)Property(row, "Timestamp");
                            Assert.AreEqual(stamp.ToLocalTime().ToString("HH:mm", System.Globalization.CultureInfo.InvariantCulture), time.Text);
                            Assert.IsTrue(time.Font.SizeInPoints < body.Font.SizeInPoints);
                            Assert.IsTrue(time.Top >= rect.Bottom); Assert.IsTrue(time.Bottom <= row.Height);
                            Assert.AreEqual(UiText.Get(language, (bool)Property(row, "IsUser") ? "ChatYou" : "ChatAssistant"), body.AccessibleName);
                        }
                        var u = (Rectangle)Property(rows[0], "BubbleBounds"); var a = (Rectangle)Property(rows[1], "BubbleBounds");
                        Assert.IsTrue(u.Left > a.Left); Assert.AreEqual(rows[0].Width - a.Left, u.Right, "User and AI must anchor to opposite avatar columns.");
                        Assert.IsTrue(user.Width < ai.Width, "Short text should shrink to its content.");
                        ((ScrollableControl)Field(f.Transcript, "_viewport")).AutoScrollPosition = Point.Empty;
                        SaveImage(f, host + "-bubbles-" + language + "-" + width);
                    }
                    var original = (DateTimeOffset)Property(rows[0], "Timestamp");
                    foreach (var lang in UiText.LanguageCodes)
                    {
                        f.Settings.UiLanguage = lang; Call(f.Pane, "RefreshLanguage");
                        Assert.AreEqual(original, Property(rows[0], "Timestamp"));
                        Assert.AreEqual(UiText.Get(lang, "ChatUserAvatar"), Field(rows[0], "_avatar"));
                        Assert.AreEqual(UiText.Get(lang, "ChatAiAvatar"), Field(rows[1], "_avatar"));
                        Assert.AreEqual(question, user.Text);
                    }
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void Bubbles_StreamOnlyIntoActiveAnswer_KeepHistoryAndScrollPosition_AndDisposeOnClear(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new PaneFixture(host, "en"))
                {
                    SeedExport(f); Call(f.Pane, "RenderTranscript");
                    var committed = Bubbles(f); var historyLayout = 0; var historyText = 0;
                    foreach (var row in committed)
                    { row.Layout += (s, e) => historyLayout++; BubbleContent(row).TextChanged += (s, e) => historyText++; }
                    Action<string> output = null!; var completion = new TaskCompletionSource<ChatReply>();
                    f.Transport((m, s, n, o, t) => { output = o; return completion.Task; });
                    ((TextBox)Field(f.Pane, "_input")).Text = "next";
                    Call(f.Pane, "Start", ChatMode.Question);
                    var rows = Bubbles(f); Assert.AreEqual(4, rows.Length);
                    var timestamp = Property(rows[3], "Timestamp");
                    historyLayout = historyText = 0;
                    output(string.Join("\n", Enumerable.Repeat("line 中文😀", 80))); Call(f.Pane, "FlushOutput");
                    Assert.AreSame(committed[0], Bubbles(f)[0]);
                    Assert.AreEqual(0, historyLayout); Assert.AreEqual(0, historyText);
                    var viewport = (ScrollableControl)Field(f.Transcript, "_viewport");
                    Assert.IsTrue(-viewport.AutoScrollPosition.Y > 0, "Streaming should follow new text at the bottom.");
                    viewport.AutoScrollPosition = Point.Empty;
                    output(string.Join("\n", Enumerable.Repeat("line 中文😀", 90))); Call(f.Pane, "FlushOutput");
                    Assert.AreEqual(0, viewport.AutoScrollPosition.Y, "Reading history must not be interrupted by streaming.");
                    Assert.AreEqual(timestamp, Property(rows[3], "Timestamp"));
                    output(""); output("**final**"); Call(f.Pane, "FlushOutput");
                    Assert.AreEqual("final", BubbleContent(rows[3]).Text);
                    completion.SetResult(new ChatReply("**final**", false)); Pump(() => FieldOrNull(f.Pane, "_pending") == null);
                    var finalRows = Bubbles(f); Button(f.Pane, "ChatClear").PerformClick();
                    Assert.AreEqual(0, Bubbles(f).Length); Assert.IsTrue(finalRows.All(r => r.IsDisposed));
                    Assert.AreEqual(UiText.Get("en", "ChatEmpty"), f.Transcript.Text);
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void LanguageChanges_UpdateExistingPaneControlsHistoryStatusAndRibbon(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new PaneFixture(host, "zh-CN"))
                {
                    f.Transport((m, s, n, o, t) => Task.FromResult(new ChatReply("answer", false)));
                    Call(f.Pane, "Start", ChatMode.Rewrite); Pump(() => FieldOrNull(f.Pane, "_pending") == null);
                    foreach (var language in UiText.LanguageCodes)
                    {
                        f.Settings.UiLanguage = language; Call(f.Pane, "RefreshLanguage");
                        foreach (var c in Descendants(f.Pane).Where(c => c.Tag is string)) Assert.AreEqual(UiText.Get(language, (string)c.Tag), c.Text);
                        StringAssert.Contains(f.Transcript.Text, UiText.Get(language, "ChatAssistant"));
                        Assert.AreEqual(UiText.Get(language, "ChatDone"), ((Label)Field(f.Pane, "_status")).Text);
                        Assert.AreEqual(UiText.Get(language, "ChatInputHint"), ((TextBox)Field(f.Pane, "_input")).AccessibleName);
                        var assembly = DocumentChatHostTests.Host(host);
                        var addinType = assembly.GetType("OfficeTranslate." + host + "AddIn.AddIn", true)!;
                        var addin = Activator.CreateInstance(addinType)!;
                        addinType.GetField("_settings", Hidden)!.SetValue(addin, f.Settings);
                        var office = Assembly.Load(addinType.Assembly.GetReferencedAssemblies().Single(a => a.Name == "office"));
                        var ribbon = office.GetType("Microsoft.Office.Core.IRibbonControl", true)!;
                        var control = new Proxy(ribbon, c => c.MethodName == "get_Id" ? "OfficeTranslate.Chat" : null).GetTransparentProxy();
                        var label = (string)addinType.GetMethod("GetRibbonLabel")!.Invoke(addin, new[] { control })!;
                        Assert.AreEqual(UiText.Get(language, "Chat"), label.Replace("\u2060", ""));
                        Assert.AreEqual(UiText.Get(language, "ChatRibbonTip"), addinType.GetMethod("GetChatTip")!.Invoke(addin, new object?[] { null }));
                        using (var stream = assembly.GetManifestResourceStream("OfficeTranslate." + host + "AddIn.Ribbon.xml"))
                        {
                            var chat = XDocument.Load(stream!).Descendants().Single(e => (string?)e.Attribute("id") == "OfficeTranslate.Chat");
                            Assert.AreEqual("GetRibbonLabel", (string?)chat.Attribute("getLabel"));
                            Assert.AreEqual("GetChatTip", (string?)chat.Attribute("getScreentip"));
                            Assert.IsNull(chat.Attribute("label"));
                        }
                    }
                    f.Settings.UiLanguage = "zh-CN"; Call(f.Pane, "RefreshLanguage");
                    f.Transport((m, s, n, o, t) => throw new IOException("fixture error"));
                    Call(f.Pane, "Start", ChatMode.Summarize); Pump(() => FieldOrNull(f.Pane, "_pending") == null);
                    f.Settings.UiLanguage = "en"; Call(f.Pane, "RefreshLanguage");
                    StringAssert.Contains(f.Transcript.Text, UiText.Get("en", "ChatErrorDetails"));
                    Assert.IsFalse(f.Transcript.Text.Contains(UiText.Get("zh-CN", "ChatErrorDetails")));
                    Assert.AreEqual(UiText.Get("en", "ChatNetworkError"), ((Label)Field(f.Pane, "_status")).Text);
                }
            });
        }
        private static void ExportDialog(PaneFixture fixture, Func<SaveFileDialog, DialogResult> show) =>
            fixture.Pane.GetType().GetProperty("ExportDialog", Hidden)!.SetValue(fixture.Pane, show);

        private static void SeedExport(PaneFixture fixture)
        {
            ((DocumentChatSession)Field(fixture.Pane, "_session")).SendAsync(
                new ChatContext("PRIVATE CONTEXT", false), "question😀\nline two", ChatMode.Question, fixture.Settings,
                (m, s, n, o, t) => Task.FromResult(new ChatReply("answer😀\nnext line", false)),
                _ => { }, (i, n) => { }, CancellationToken.None).GetAwaiter().GetResult();
        }

        [DataTestMethod]
        [DataRow("Word", "zh-CN")] [DataRow("Word", "en")]
        [DataRow("Excel", "zh-CN")] [DataRow("Excel", "en")]
        [DataRow("PowerPoint", "zh-CN")] [DataRow("PowerPoint", "en")]
        public void Export_WritesBothFormats_WithLocalizedSaveDialog_AndOnlyReadsMemory(string host, string language)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new PaneFixture(host, language == "en" ? "zh-CN" : "en"))
                {
                    SeedExport(f); f.Settings.UiLanguage = language;
                    var reads = f.Reads; var writes = 0;
                    f.Context = new ChatContext("PRIVATE CONTEXT", true, (a, s) => writes++, true);
                    foreach (var markdown in new[] { true, false })
                    {
                        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + (markdown ? ".md" : ".txt"));
                        try
                        {
                            ExportDialog(f, dialog =>
                            {
                                Assert.AreEqual(UiText.Get(language, "ChatExportTitle"), dialog.Title);
                                Assert.AreEqual(UiText.Get(language, "ChatExportFilter"), dialog.Filter);
                                Assert.AreEqual("md", dialog.DefaultExt); Assert.AreEqual(1, dialog.FilterIndex);
                                Assert.IsTrue(dialog.AddExtension); Assert.IsTrue(dialog.OverwritePrompt); Assert.IsTrue(dialog.RestoreDirectory);
                                var template = UiText.Get(language, "ChatExportFileName");
                                var prefix = template.Substring(0, template.IndexOf('{'));
                                StringAssert.Matches(dialog.FileName, new System.Text.RegularExpressions.Regex("^" + System.Text.RegularExpressions.Regex.Escape(prefix) + @"\d{8}_\d{6}\.md$"));
                                dialog.FilterIndex = markdown ? 1 : 2; dialog.FileName = path; return DialogResult.OK;
                            });
                            Button(f.Pane, "ChatExport").PerformClick();
                            Assert.AreEqual(UiText.Get(language, "ChatExport"), Button(f.Pane, "ChatExport").Text);
                            var session = (DocumentChatSession)Field(f.Pane, "_session");
                            Assert.AreEqual(ChatTranscriptExport.Format(session.History, markdown, k => UiText.Get(language, k)), File.ReadAllText(path));
                            Assert.AreEqual(2, session.History.Count); Assert.AreEqual(reads, f.Reads); Assert.AreEqual(0, writes);
                            Assert.IsFalse(File.ReadAllText(path).Contains("PRIVATE CONTEXT"));
                            Assert.AreEqual(UiText.Get(language, "ChatExportSuccess"), ((Label)Field(f.Pane, "_status")).Text);
                        }
                        finally { if (File.Exists(path)) File.Delete(path); }
                    }
                    var export = Button(f.Pane, "ChatExport"); var clear = Button(f.Pane, "ChatClear");
                    Assert.AreSame(clear.Parent, export.Parent); Assert.AreEqual(clear.BackColor, export.BackColor);
                    Assert.AreEqual(clear.Height, export.Height); Assert.AreEqual(clear.FlatStyle, export.FlatStyle);
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word", "zh-CN")] [DataRow("Word", "en")]
        [DataRow("Excel", "zh-CN")] [DataRow("Excel", "en")]
        [DataRow("PowerPoint", "zh-CN")] [DataRow("PowerPoint", "en")]
        public void Export_EmptyAndClearedSessions_ShowLocalizedHint_WithoutDialogOrDocumentRead(string host, string language)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new PaneFixture(host, language))
                {
                    var dialogs = 0; var reads = f.Reads;
                    ExportDialog(f, d => { dialogs++; return DialogResult.Cancel; });
                    for (var i = 0; i < 2; i++)
                    {
                        Button(f.Pane, "ChatExport").PerformClick();
                        Assert.AreEqual(UiText.Get(language, "ChatExportEmpty"), ((Label)Field(f.Pane, "_status")).Text);
                        Assert.AreEqual(0, dialogs); Assert.AreEqual(reads, f.Reads);
                        if (i == 0) { SeedExport(f); Button(f.Pane, "ChatClear").PerformClick(); }
                    }
                    f.Settings.UiLanguage = language == "en" ? "zh-CN" : "en"; Call(f.Pane, "RefreshLanguage");
                    Assert.AreEqual(UiText.Get(f.Settings.UiLanguage, "ChatExportEmpty"), ((Label)Field(f.Pane, "_status")).Text);
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void Export_CancelAndFailures_KeepSessionAndDocumentUntouched_AndNeverEchoRawError(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new PaneFixture(host, "en"))
                {
                    SeedExport(f); var reads = f.Reads; var status = ((Label)Field(f.Pane, "_status")).Text;
                    var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".md");
                    ExportDialog(f, d => { d.FileName = path; return DialogResult.Cancel; });
                    Button(f.Pane, "ChatExport").PerformClick(); Assert.IsFalse(File.Exists(path));
                    Assert.AreEqual(status, ((Label)Field(f.Pane, "_status")).Text);
                    var exceptions = new Exception[] { new UnauthorizedAccessException("secret-user/path question😀"), new IOException("secret-key answer😀"), new ArgumentException("private-path"), new InvalidOperationException("private-message") };
                    var keys = new[] { "ChatExportAccessError", "ChatExportIoError", "ChatExportPathError", "ChatExportUnknownError" };
                    for (var i = 0; i < exceptions.Length; i++)
                    {
                        ExportDialog(f, d => throw exceptions[i]); Button(f.Pane, "ChatExport").PerformClick();
                        foreach (var language in new[] { "zh-CN", "en" })
                        {
                            f.Settings.UiLanguage = language; Call(f.Pane, "RefreshLanguage");
                            Assert.AreEqual(UiText.Get(language, "ChatExportFailed") + " " + UiText.Get(language, keys[i]), ((Label)Field(f.Pane, "_status")).Text);
                        }
                    }
                    // Exercise a real filesystem failure, beyond dialog exceptions.
                    ExportDialog(f, d => { d.FileName = Path.Combine(path, "missing", "file.md"); return DialogResult.OK; });
                    Button(f.Pane, "ChatExport").PerformClick();
                    StringAssert.Contains(((Label)Field(f.Pane, "_status")).Text, UiText.Get("en", "ChatExportIoError"));
                    Assert.AreEqual(reads, f.Reads); Assert.AreEqual(2, ((DocumentChatSession)Field(f.Pane, "_session")).History.Count);
                }
            });
        }

        private sealed class Proxy : System.Runtime.Remoting.Proxies.RealProxy
        {
            private readonly Func<System.Runtime.Remoting.Messaging.IMethodCallMessage, object?> _call;
            internal Proxy(Type type, Func<System.Runtime.Remoting.Messaging.IMethodCallMessage, object?> call) : base(type) { _call = call; }
            public override System.Runtime.Remoting.Messaging.IMessage Invoke(System.Runtime.Remoting.Messaging.IMessage msg)
            { var call = (System.Runtime.Remoting.Messaging.IMethodCallMessage)msg; return new System.Runtime.Remoting.Messaging.ReturnMessage(_call(call), null, 0, call.LogicalCallContext, call); }
        }

        [DataTestMethod]
        [DataRow("zh-CN", "请用中文回答")] [DataRow("en", "Please answer in English")]
        [DataRow("zh-HK", "請用繁體中文回答")] [DataRow("ko", "한국어로 답하세요")] [DataRow("ja", "日本語で回答してください")]
        public async Task DefaultsAndSerializedLegacyPrompts_UseUiLanguageForEveryMode(string language, string requirement)
        {
            const string legacy = "{\"ChatQuestionSystemPrompt\":\"Answer questions using the document. Distinguish facts from inference; say when the document does not contain the answer.\",\"ChatSummarySystemPrompt\":\"Summarize the document clearly and concisely. Preserve key facts, numbers, conclusions and caveats.\",\"ChatRewriteSystemPrompt\":\"Rewrite the supplied text for clarity while preserving its meaning, facts and language. Return only the rewritten text.\"}";
            var json = new JavaScriptSerializer();
            var settings = json.Deserialize<TranslationSettings>(legacy); settings.UiLanguage = language;
            foreach (ChatMode mode in Enum.GetValues(typeof(ChatMode)))
            {
                var session = new DocumentChatSession();
                await session.SendAsync(new ChatContext("A document in a different language. 文書の言語", false), "Q", mode, settings,
                    (m, s, n, o, t) => { StringAssert.StartsWith(m[0].Content, ChatPrompts.Default(language, mode)); StringAssert.Contains(m[0].Content, requirement);
                        Assert.IsFalse(m[0].Content.Contains("user's language")); Assert.IsFalse(m[0].Content.Contains("facts and language")); return Task.FromResult(new ChatReply("ok", false)); }, _ => { }, (i, n) => { }, CancellationToken.None);
            }
            var restored = json.Deserialize<TranslationSettings>(json.Serialize(settings)); restored.UiLanguage = language == "en" ? "zh-CN" : "en";
            Assert.AreEqual(ChatPrompts.Default(restored.UiLanguage, ChatMode.Question), restored.ChatQuestionSystemPrompt);
            settings.ChatQuestionSystemPrompt = "My custom prompt. Répondez en français.";
            settings.UiLanguage = "en";
            Assert.AreEqual("My custom prompt. Répondez en français.", settings.ChatQuestionSystemPrompt);
        }

        [DataTestMethod]
        [DataRow("zh-CN")] [DataRow("en")]
        public async Task LongDocuments_UseUiLanguageForIntermediateSummariesAndFinalPrompt(string language)
        {
            var settings = new TranslationSettings { UiLanguage = language, ChatContextTokens = 4096, ChatOutputTokens = 512 };
            var maps = 0; var finals = 0;
            await new DocumentChatSession().SendAsync(new ChatContext(new string('文', 5000), false), "Q", ChatMode.Summarize, settings,
                (m, s, n, o, t) =>
                {
                    if (n == 256) { maps++; Assert.AreEqual(ChatPrompts.Map(language), m[0].Content); }
                    else { finals++; StringAssert.StartsWith(m[0].Content, UiText.Get(language, "ChatSummaryDefaultPrompt")); }
                    Assert.IsFalse(m[0].Content.Contains("original language"));
                    return Task.FromResult(new ChatReply("short summary", false));
                }, _ => { }, (i, n) => { }, CancellationToken.None);
            Assert.IsTrue(maps > 1); Assert.AreEqual(1, finals);
        }

        [DataTestMethod]
        [DataRow("zh-CN")] [DataRow("en")]
        public void ChatSettings_LabelsAndDefaultPromptTextUseUiLanguage_WithoutTouchingRealSettings(string language)
        {
            DocumentChatHostTests.Sta(() =>
            {
                var path = Path.Combine(Path.GetTempPath(), "OfficeTranslate-ChatUi-" + Guid.NewGuid() + ".json");
                var store = new SettingsStore(); typeof(SettingsStore).GetField("_path", Hidden)!.SetValue(store, path);
                try
                {
                    store.Save(new TranslationSettings { UiLanguage = language });
                    var type = DocumentChatHostTests.Host("Word").GetType("OfficeTranslate.WordAddIn.ChatSettingsForm", true)!;
                    using (var form = (Form)Activator.CreateInstance(type, new object[] { store })!)
                    {
                        form.Show(); Application.DoEvents();
                        Assert.AreEqual(UiText.Get(language, "ChatSettings"), form.Text);
                        var all = Descendants(form).ToArray();
                        foreach (var key in new[] { "ChatQuestionPrompt", "ChatSummaryPrompt", "ChatRewritePrompt", "ChatBudget", "ChatOutputBudget", "Save" }) Assert.IsTrue(all.Any(c => c.Text == UiText.Get(language, key)), key);
                        var boxes = all.OfType<TextBox>().Where(b => b.Multiline).ToArray();
                        Assert.AreEqual(3, boxes.Length);
                        for (var i = 0; i < 3; i++) Assert.AreEqual(ChatPrompts.Default(language, (ChatMode)i), boxes[i].Text);
                        form.Close();
                    }
                }
                finally { File.Delete(path); }
            });
        }
    }
}
