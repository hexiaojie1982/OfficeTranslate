using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using OfficeTranslate.WordAddIn;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ChatCopyButtonTests
    {
        [DllImport("user32.dll")] private static extern bool OpenClipboard(IntPtr owner);
        [DllImport("user32.dll")] private static extern bool CloseClipboard();
        private const BindingFlags Hidden = BindingFlags.Instance | BindingFlags.NonPublic;
        private static object Field(object o, string name) => o.GetType().GetField(name, Hidden)!.GetValue(o)!;
        private static object Property(object o, string name) => o.GetType().GetProperty(name, Hidden)!.GetValue(o)!;
        private static void Call(object o, string name, params object[] args) => o.GetType().GetMethod(name, Hidden)!.Invoke(o, args);
        private static Control Copy(Control row) => (Control)Property(row, "CopyButton");
        private static Control[] Rows(Control view) => Descendants(view).Where(c => c.GetType().Name == "ChatBubbleRow").ToArray();
        private static IEnumerable<Control> Descendants(Control c) => c.Controls.Cast<Control>().SelectMany(x => new[] { x }.Concat(Descendants(x)));
        private static void Writer(Control copy, Action<string> writer) => copy.GetType().GetProperty("ClipboardWriter", Hidden)!.SetValue(copy, writer);
        private static string Tip(Control copy) => ((ToolTip)Field(copy, "_tips")).GetToolTip(copy);
        private static void SaveWindow(Form window, string name)
        {
            var evidence = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../../../TestResults/copy-button-2.3.5")); Directory.CreateDirectory(evidence);
            using (var bitmap = new Bitmap(window.Width, window.Height))
            { window.DrawToBitmap(bitmap, new Rectangle(Point.Empty, window.Size)); bitmap.Save(Path.Combine(evidence, name + ".png")); }
        }
        private static void Pump(Func<bool> done, int milliseconds = 6000)
        {
            var end = DateTime.UtcNow.AddMilliseconds(milliseconds);
            while (!done() && DateTime.UtcNow < end) { Application.DoEvents(); Thread.Sleep(5); }
            Assert.IsTrue(done(), "UI operation timed out");
        }
        private sealed class Fixture : IDisposable
        {
            private readonly HistorySandbox _history = new HistorySandbox();
            internal readonly UserControl Pane;
            internal readonly Form Window = new Form { ClientSize = new Size(440, 760) };
            internal readonly TranslationSettings Settings = new TranslationSettings { UiLanguage = "en" };
            internal Fixture(string host)
            {
                Pane = (UserControl)Activator.CreateInstance(DocumentChatHostTests.Host(host).GetType("OfficeTranslate.WordAddIn.DocumentChatPane", true)!)!;
                Pane.GetType().GetProperty("HistoryStore", Hidden)!.SetValue(Pane, _history.Store);
                Call(Pane, "Initialize", (Func<ChatContext>)(() => new ChatContext("context", false, host: host)), (Func<TranslationSettings>)(() => Settings));
                Pane.Dock = DockStyle.Fill; Window.Controls.Add(Pane); Window.Show(); Application.DoEvents();
            }
            internal Control View => (Control)Field(Pane, "_transcript");
            internal void Send(string question, ChatTransport transport)
            {
                Pane.GetType().GetProperty("Transport", Hidden)!.SetValue(Pane, transport);
                ((TextBox)Field(Pane, "_input")).Text = question; Call(Pane, "Start", ChatMode.Question);
            }
            internal void Done() => Pump(() => Pane.GetType().GetField("_pending", Hidden)!.GetValue(Pane) == null);
            public void Dispose() { Window.Dispose(); Pane.Dispose(); _history.Dispose(); }
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void Copy_UsesExactUserAndMarkdownSource_IncludingLongUnicodeAndLineEndings(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new Fixture(host))
                {
                    const string question = "## user **literal**\r\n`code` 😀";
                    var answer = "  ## title\r\n**bold**\n```cs\r\nx++;\r\n```\n" + new string('文', 100000) + "😀\r\n  ";
                    f.Send(question, (m, s, n, o, t) => Task.FromResult(new ChatReply("**initial**", false))); f.Done();
                    // Test clipboard payload length independently of RichEdit's
                    // expensive natural-height measurement of a 100k-char bubble.
                    Call(Rows(f.View)[1], "UpdateOutput", answer, "");
                    var expected = new[] { question, answer }; var rows = Rows(f.View);
                    for (var i = 0; i < rows.Length; i++)
                    {
                        var copy = Copy(rows[i]); Assert.IsTrue(copy.Visible);
                        string? captured = null; Writer(copy, text => captured = text);
                        Call(copy, "OnClick", EventArgs.Empty); Assert.AreEqual(expected[i], captured);
                    }
                    Assert.AreEqual(2, rows.Length);
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void Copy_HidesDuringStreamingAndForEmptyMessages_ThenAppearsAfterCompletion(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new Fixture(host))
                {
                    Action<string> output = null!; var completion = new TaskCompletionSource<ChatReply>();
                    f.Send("question", (m, s, n, o, t) => { output = o; return completion.Task; });
                    var rows = Rows(f.View); Assert.IsTrue(Copy(rows[0]).Visible); Assert.IsFalse(Copy(rows[1]).Visible);
                    output("**partial**"); Call(f.Pane, "FlushOutput"); Assert.IsFalse(Copy(rows[1]).Visible);
                    f.Settings.UiLanguage = "ko"; Call(f.Pane, "RefreshLanguage"); Assert.IsFalse(Copy(rows[1]).Visible);
                    completion.SetResult(new ChatReply("**final**", false)); f.Done();
                    Assert.IsTrue(Copy(Rows(f.View)[1]).Visible);
                    foreach (var row in Rows(f.View))
                    {
                        foreach (var empty in new[] { "", " \r\n\t" })
                        {
                            Call(row, "UpdateOutput", empty, "diagnostic only");
                            Assert.IsFalse(Copy(row).Visible); var calls = 0; Writer(Copy(row), _ => calls++);
                            Call(Copy(row), "Copy"); Assert.AreEqual(0, calls);
                        }
                        Call(row, "UpdateOutput", "again", ""); Assert.IsTrue(Copy(row).Visible);
                    }
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void Copy_FailureIsContained_RetryAndRepeatedClickResetThe1500msFeedback(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new Fixture(host))
                {
                    f.Send("question", (m, s, n, o, t) => Task.FromResult(new ChatReply("answer", false))); f.Done();
                    var copy = Copy(Rows(f.View)[1]);
                    Writer(copy, _ => throw new ExternalException("private clipboard content must not appear"));
                    Call(copy, "OnClick", EventArgs.Empty); Assert.AreEqual("Copy failed", Tip(copy));
                    Assert.IsFalse((bool)Property(copy, "IsCopied"));
                    var timer = (System.Windows.Forms.Timer)Field(copy, "_feedbackTimer"); Assert.AreEqual(1500, timer.Interval);
                    Pump(() => Tip(copy) == "Copy");
                    Writer(copy, _ => { }); Call(copy, "OnClick", EventArgs.Empty);
                    Assert.IsTrue((bool)Property(copy, "IsCopied")); Assert.AreEqual("Copied", Tip(copy));
                    var until = DateTime.UtcNow.AddMilliseconds(900); Pump(() => DateTime.UtcNow >= until);
                    Call(copy, "OnClick", EventArgs.Empty); until = DateTime.UtcNow.AddMilliseconds(900); Pump(() => DateTime.UtcNow >= until);
                    Assert.AreEqual("Copied", Tip(copy)); Pump(() => Tip(copy) == "Copy");
                    Assert.IsFalse((bool)Property(copy, "IsCopied"));
                    Call(copy, "Copy"); f.Window.Dispose(); Assert.IsTrue(copy.IsDisposed); Assert.IsFalse(timer.Enabled);
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void Copy_BackgroundInvocationDispatchesClipboardToOwningStaThread(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new Fixture(host))
                {
                    f.Send("question", (m, s, n, o, t) => Task.FromResult(new ChatReply("answer", false))); f.Done();
                    var ui = Thread.CurrentThread.ManagedThreadId; var actual = 0; var apartment = ApartmentState.Unknown;
                    var copy = Copy(Rows(f.View)[1]); Writer(copy, _ => { actual = Thread.CurrentThread.ManagedThreadId; apartment = Thread.CurrentThread.GetApartmentState(); });
                    var task = Task.Run(() => Call(copy, "Copy")); Pump(() => task.IsCompleted && actual != 0); task.GetAwaiter().GetResult();
                    Assert.AreEqual(ui, actual); Assert.AreEqual(ApartmentState.STA, apartment);
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void Copy_RealClipboardLockShowsFailure_ThenRecoversAfterRelease(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new Fixture(host))
                using (var ready = new ManualResetEventSlim())
                using (var release = new ManualResetEventSlim())
                {
                    f.Send("question", (m, s, n, o, t) => Task.FromResult(new ChatReply("**answer**", false))); f.Done();
                    var copy = Copy(Rows(f.View)[1]); var locked = false;
                    var holder = new Thread(() =>
                    {
                        locked = OpenClipboard(IntPtr.Zero); ready.Set();
                        try { release.Wait(5000); } finally { if (locked) CloseClipboard(); }
                    }) { IsBackground = true };
                    holder.Start(); Assert.IsTrue(ready.Wait(2000));
                    try
                    {
                        Assert.IsTrue(locked); Call(copy, "OnClick", EventArgs.Empty);
                        Assert.AreEqual("Copy failed", Tip(copy)); Assert.IsFalse((bool)Property(copy, "IsCopied"));
                    }
                    finally { release.Set(); Assert.IsTrue(holder.Join(2000)); }
                    Call(copy, "Copy"); Assert.AreEqual("Copied", Tip(copy)); Assert.AreEqual("**answer**", Clipboard.GetText());
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void Copy_CancelledAndFailedStreamsCopyPartialSourceWithoutDiagnosticSuffix(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new Fixture(host))
                {
                    foreach (var cancel in new[] { true, false })
                    {
                        Action<string> output = null!; var completion = new TaskCompletionSource<ChatReply>();
                        f.Send("question", (m, s, n, o, t) => { output = o; return completion.Task; });
                        output("**partial**\r\nraw"); Call(f.Pane, "FlushOutput");
                        var copy = Copy(Rows(f.View).Last()); Assert.IsFalse(copy.Visible);
                        if (cancel) ((Button)Field(f.Pane, "_cancel")).PerformClick();
                        completion.SetException(new IOException("fixture error")); f.Done();
                        Assert.IsTrue(copy.Visible); string? captured = null; Writer(copy, text => captured = text);
                        Call(copy, "Copy"); Assert.AreEqual("**partial**\r\nraw", captured);
                        Assert.AreEqual(0, ((DocumentChatSession)Field(f.Pane, "_session")).History.Count);
                        Call(f.Pane, "Clear");
                    }
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void Copy_VectorIconsDarkenOnHover_AndChangeToCheckForBothRoles(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new Fixture(host))
                {
                    f.Send("question", (m, s, n, o, t) => Task.FromResult(new ChatReply("answer", false))); f.Done();
                    SaveWindow(f.Window, host + "-preview");
                    var evidence = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../../../TestResults/copy-button-2.3.5")); Directory.CreateDirectory(evidence);
                    foreach (var row in Rows(f.View))
                    {
                        var copy = Copy(row); var user = (bool)Property(row, "IsUser");
                        Bitmap Snapshot(string state)
                        {
                            var bitmap = new Bitmap(copy.Width, copy.Height); copy.DrawToBitmap(bitmap, copy.ClientRectangle);
                            bitmap.Save(Path.Combine(evidence, host + (user ? "-user-" : "-ai-") + state + ".png")); return bitmap;
                        }
                        Call(copy, "OnMouseLeave", EventArgs.Empty);
                        using (var normal = Snapshot("normal"))
                        {
                            Call(copy, "OnMouseEnter", EventArgs.Empty);
                            using (var hover = Snapshot("hover"))
                            {
                                double normalLight = 0, hoverLight = 0;
                                for (var y = 0; y < normal.Height; y++) for (var x = 0; x < normal.Width; x++)
                                { normalLight += normal.GetPixel(x, y).GetBrightness(); hoverLight += hover.GetPixel(x, y).GetBrightness(); }
                                Assert.IsTrue(hoverLight < normalLight, "Both role icons must darken on hover.");
                                Writer(copy, _ => { }); Call(copy, "OnClick", EventArgs.Empty);
                                using (var check = Snapshot("copied"))
                                {
                                    var differences = 0;
                                    for (var y = 0; y < check.Height; y++) for (var x = 0; x < check.Width; x++)
                                        if (check.GetPixel(x, y) != hover.GetPixel(x, y)) differences++;
                                    Assert.IsTrue(differences > 20, "Copied state must draw a different vector glyph.");
                                }
                            }
                        }
                    }
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void Copy_RealUnicodeClipboardRoundTripsLongRawMarkdown(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new Fixture(host))
                {
                    var answer = "## raw\r\n**bold**\n```\n" + new string('한', 70000) + "😀\n```\r\n  ";
                    f.Send("question", (m, s, n, o, t) => Task.FromResult(new ChatReply("**initial**", false))); f.Done();
                    Call(Rows(f.View)[1], "UpdateOutput", answer, "");
                    var copy = Copy(Rows(f.View)[1]); Call(copy, "OnClick", EventArgs.Empty);
                    Assert.AreEqual("Copied", Tip(copy)); Assert.AreEqual(answer, Clipboard.GetText(TextDataFormat.UnicodeText));
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word", "zh-CN", "复制", "已复制", "复制失败")]
        [DataRow("Excel", "zh-HK", "複製", "已複製", "複製失敗")]
        [DataRow("PowerPoint", "en", "Copy", "Copied", "Copy failed")]
        [DataRow("Word", "ko", "복사", "복사됨", "복사 실패")]
        [DataRow("Excel", "ja", "コピー", "コピーしました", "コピーに失敗しました")]
        public void Copy_TooltipsAndFeedbackFollowUiText_AndRefreshExistingMessages(string host, string language, string normal, string success, string failure)
        {
            Assert.AreEqual(normal, UiText.Get(language, "ChatCopy")); Assert.AreEqual(success, UiText.Get(language, "ChatCopied")); Assert.AreEqual(failure, UiText.Get(language, "ChatCopyFailed"));
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new Fixture(host))
                {
                    f.Send("question", (m, s, n, o, t) => Task.FromResult(new ChatReply("**answer**", false))); f.Done();
                    f.Settings.UiLanguage = language; Call(f.Pane, "RefreshLanguage"); var copy = Copy(Rows(f.View)[1]);
                    Assert.AreEqual(normal, Tip(copy)); Assert.AreEqual(normal, copy.AccessibleName);
                    Writer(copy, _ => { }); Call(copy, "Copy"); Assert.AreEqual(success, Tip(copy));
                    f.Settings.UiLanguage = "zh-CN"; Call(f.Pane, "RefreshLanguage"); Assert.AreEqual("已复制", Tip(copy));
                    Writer(copy, _ => throw new InvalidOperationException("private failure")); Call(copy, "Copy"); Assert.AreEqual("复制失败", Tip(copy));
                    f.Settings.UiLanguage = language; Call(f.Pane, "RefreshLanguage"); Assert.AreEqual(failure, Tip(copy));
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void History_EachReadOnlyMessageCanCopyExactSource_WithoutChangingPersistence(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var sandbox = new HistorySandbox())
                {
                    var id = Guid.NewGuid().ToString("N"); var now = DateTimeOffset.Now;
                    var expected = new[] { "user\r\n**literal**", "## heading\n**raw**\r\n```cs\nx++;\n```  ", "next", "final" };
                    sandbox.Store.Append(id, now, "Unsaved1", new[] {
                        new ChatHistoryTurn { Timestamp = now.ToString("O"), Question = expected[0], Answer = expected[1] },
                        new ChatHistoryTurn { Timestamp = now.ToString("O"), Question = expected[2], Answer = expected[3] } }, host);
                    var path = Path.Combine(sandbox.DirectoryPath, id + ".json"); var original = File.ReadAllBytes(path);
                    using (var form = (Form)Activator.CreateInstance(DocumentChatHostTests.Host(host).GetType("OfficeTranslate.WordAddIn.ChatHistoryForm", true)!, sandbox.Store, "en")!)
                    {
                        form.Show(); Application.DoEvents(); ((ListView)Field(form, "_list")).Items[0].Selected = true; Application.DoEvents();
                        var rows = Rows((Control)Field(form, "_details")); Assert.AreEqual(4, rows.Length);
                        Assert.AreEqual(Point.Empty, ((ScrollableControl)Field(Field(form, "_details"), "_viewport")).AutoScrollPosition);
                        for (var i = 0; i < rows.Length; i++)
                        {
                            Assert.IsTrue(((RichTextBox)Property(rows[i], "Content")).ReadOnly);
                            var copy = Copy(rows[i]); Assert.IsTrue(copy.Visible); string? captured = null;
                            Writer(copy, text => captured = text); Call(copy, "OnClick", EventArgs.Empty); Assert.AreEqual(expected[i], captured);
                        }
                        SaveWindow(form, host + "-history");
                    }
                    CollectionAssert.AreEqual(original, File.ReadAllBytes(path));
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void Copy_StaysInsideBottomRightWithoutTextOverlap_AndPreservesAdaptiveWidths(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new Fixture(host))
                {
                    f.Send("short", (m, s, n, o, t) => Task.FromResult(new ChatReply("**answer**\n" + new string('文', 300), false))); f.Done();
                    foreach (var width in new[] { 260, 320, 440, 600 })
                    {
                        f.Window.ClientSize = new Size(width, 760); Application.DoEvents();
                        var rows = Rows(f.View);
                        foreach (var row in rows)
                        {
                            var bubble = (Rectangle)Property(row, "BubbleBounds"); var copy = Copy(row); var content = (Control)Property(row, "Content");
                            Assert.IsTrue(bubble.Contains(copy.Bounds)); Assert.IsTrue(copy.Top >= content.Bottom);
                            Assert.IsTrue(copy.Right >= bubble.Right - 16); Assert.IsTrue(copy.Bottom >= bubble.Bottom - 12);
                            Assert.IsTrue(bubble.Width <= row.Width * .8 + 1); Assert.IsTrue(copy.Visible);
                            Assert.IsTrue(((Control)Field(row, "_time")).Top >= bubble.Bottom);
                        }
                        Assert.IsTrue(((Rectangle)Property(rows[0], "BubbleBounds")).Width < ((Rectangle)Property(rows[1], "BubbleBounds")).Width);
                        SaveWindow(f.Window, host + "-" + width);
                    }
                }
            });
        }
    }
}
