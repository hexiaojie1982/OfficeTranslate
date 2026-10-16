using System;
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
    internal sealed class HistorySandbox : IDisposable
    {
        internal readonly string DirectoryPath = Path.Combine(Path.GetTempPath(), "OfficeTranslate-History-" + Guid.NewGuid().ToString("N"));
        internal ChatHistoryStore Store => new ChatHistoryStore(DirectoryPath);
        public void Dispose() { if (Directory.Exists(DirectoryPath)) Directory.Delete(DirectoryPath, true); }
    }
    [TestClass]
    public class ChatHistoryTests
    {
        private static readonly DateTimeOffset Time = new DateTimeOffset(2026, 10, 16, 12, 30, 0, TimeSpan.FromHours(8));
        private static ChatHistoryTurn Turn(string question = "问题😀\r\nline", string answer = "## 回答\n**raw**\n```cs\nx++;\n```") =>
            new ChatHistoryTurn { Timestamp = Time.ToString("O"), Question = question, Answer = answer };
        private static string Add(ChatHistoryStore store, DateTimeOffset time, string question = "first")
        { var id = Guid.NewGuid().ToString("N"); store.Append(id, time, "文档.docx", new[] { Turn(question) }); return id; }
        private static Task<ChatReply> Send(DocumentChatSession session, string question = "Q", ChatTransport? transport = null, CancellationToken token = default) =>
            session.SendAsync(new ChatContext("PRIVATE DOCUMENT CONTEXT", false, documentName: "测试.docx"), question, ChatMode.Question,
                new TranslationSettings(), transport ?? ((m, s, n, o, t) => Task.FromResult(new ChatReply("**raw**\r\nanswer😀", false))), _ => { }, (i, n) => { }, token);

        [TestMethod]
        public void Store_RoundTripsRawTextAndOffset_UpdatesSameFile_AndUsesAppData()
        {
            using (var sandbox = new HistorySandbox())
            {
                var store = sandbox.Store; var id = Guid.NewGuid().ToString("N"); var turn = Turn();
                store.Append(id, Time, "报告😀.docx", new[] { turn });
                store.Append(id, Time, "renamed.docx", new[] { Turn("second") });
                var reloaded = sandbox.Store.Read(id)!;
                Assert.AreEqual(1, reloaded.SchemaVersion); Assert.AreEqual(id, reloaded.SessionId);
                Assert.AreEqual(Time.ToString("O"), reloaded.CreatedAt); Assert.AreEqual("报告😀.docx", reloaded.DocumentName);
                Assert.AreEqual(2, reloaded.Turns.Count); Assert.AreEqual(turn.Timestamp, reloaded.Turns[0].Timestamp);
                Assert.AreEqual(turn.Question, reloaded.Turns[0].Question); Assert.AreEqual(turn.Answer, reloaded.Turns[0].Answer);
                CollectionAssert.AreEqual(new[] { id + ".json" }, Directory.GetFiles(sandbox.DirectoryPath).Select(Path.GetFileName).ToArray());
            }
            var directory = (string)typeof(ChatHistoryStore).GetField("_directory", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(new ChatHistoryStore())!;
            Assert.AreEqual(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OfficeTranslate", "ChatHistory"), directory);
        }

        [TestMethod]
        public async Task Session_PersistsAllTurns_ButKeepsOnlyTenInMemory_AndClearStartsNewFile()
        {
            using (var sandbox = new HistorySandbox())
            {
                var session = new DocumentChatSession(sandbox.Store);
                for (var i = 0; i < 13; i++) await Send(session, "Q" + i);
                Assert.AreEqual(20, session.History.Count); Assert.AreEqual("Q3", session.History[0].Content);
                var summaries = sandbox.Store.List(out var unreadable); Assert.AreEqual(0, unreadable); Assert.AreEqual(1, summaries.Count);
                var record = sandbox.Store.Read(summaries[0].SessionId)!;
                Assert.AreEqual(13, record.Turns.Count); Assert.AreEqual("Q0", record.Turns[0].Question); Assert.AreEqual("测试.docx", record.DocumentName);
                Assert.IsFalse(File.ReadAllText(Path.Combine(sandbox.DirectoryPath, record.SessionId + ".json")).Contains("PRIVATE DOCUMENT CONTEXT"));
                Assert.IsTrue(session.Clear()); Assert.AreEqual(0, session.History.Count);
                await Send(session, "new session");
                Assert.AreEqual(2, sandbox.Store.List(out _).Count); Assert.AreEqual(13, sandbox.Store.Read(record.SessionId)!.Turns.Count);
            }
        }

        [TestMethod]
        public async Task Session_CancellationAndTransportFailure_DoNotCreateHistory()
        {
            using (var sandbox = new HistorySandbox())
            using (var cancellation = new CancellationTokenSource())
            {
                var session = new DocumentChatSession(sandbox.Store);
                await Assert.ThrowsExceptionAsync<IOException>(() => Send(session, transport: (m, s, n, o, t) => throw new IOException("failure")));
                var completion = new TaskCompletionSource<ChatReply>();
                var pending = Send(session, transport: (m, s, n, o, t) => completion.Task, token: cancellation.Token);
                cancellation.Cancel(); completion.SetResult(new ChatReply("partial", false));
                try { await pending; Assert.Fail("Expected cancellation"); } catch (OperationCanceledException) { }
                completion = new TaskCompletionSource<ChatReply>();
                pending = Send(session, transport: (m, s, n, o, t) => completion.Task);
                Assert.IsTrue(session.Clear()); completion.SetResult(new ChatReply("stale", false));
                try { await pending; Assert.Fail("Expected generation cancellation"); } catch (OperationCanceledException) { }
                Assert.AreEqual(0, sandbox.Store.List(out _).Count); Assert.AreEqual(0, session.History.Count);
            }
        }

        [TestMethod]
        public void Store_ListSortsCreationTimeAndPreviews_SkipsCorruption_AndDeletesOnlyHistory()
        {
            using (var sandbox = new HistorySandbox())
            {
                Assert.AreEqual(0, sandbox.Store.List(out _).Count); sandbox.Store.DeleteAll();
                var newer = Add(sandbox.Store, Time.AddDays(1), new string('问', 110) + "\nlast");
                var older = Add(sandbox.Store, Time, "first\r\nquestion");
                File.SetLastWriteTimeUtc(Path.Combine(sandbox.DirectoryPath, older + ".json"), DateTime.UtcNow.AddDays(3));
                File.WriteAllText(Path.Combine(sandbox.DirectoryPath, "broken.json"), "{");
                var foreign = Path.Combine(sandbox.DirectoryPath, "keep.txt"); File.WriteAllText(foreign, "keep");
                var list = sandbox.Store.List(out var skipped);
                Assert.AreEqual(1, skipped); CollectionAssert.AreEqual(new[] { newer, older }, list.Select(r => r.SessionId).ToArray());
                Assert.AreEqual(101, list[0].Preview.Length); Assert.AreEqual("first  question", list[1].Preview);
                sandbox.Store.Delete(newer); Assert.IsNull(sandbox.Store.Read(newer)); Assert.IsNotNull(sandbox.Store.Read(older));
                sandbox.Store.Delete(newer); // Idempotent when another window already deleted it.
                Assert.ThrowsException<ArgumentException>(() => sandbox.Store.Delete("../outside"));
                sandbox.Store.DeleteAll(); Assert.AreEqual(0, sandbox.Store.List(out skipped).Count); Assert.AreEqual(0, skipped);
                Assert.IsTrue(File.Exists(foreign));
            }
        }

        [TestMethod]
        public async Task Session_FailedAtomicSaveRetainsOldFileAndPendingTurns_ClearRetriesWithoutDuplicates()
        {
            using (var sandbox = new HistorySandbox())
            {
                var session = new DocumentChatSession(sandbox.Store); await Send(session, "first");
                var id = sandbox.Store.List(out _).Single().SessionId;
                var path = Path.Combine(sandbox.DirectoryPath, id + ".json"); var original = File.ReadAllText(path);
                using (var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    await Send(session, "second"); Assert.IsTrue(session.HistorySaveFailed); Assert.AreEqual(4, session.History.Count);
                    Assert.IsFalse(session.Clear()); Assert.AreEqual(4, session.History.Count);
                    Assert.AreEqual(original, File.ReadAllText(path));
                    Assert.AreEqual(0, Directory.GetFiles(sandbox.DirectoryPath, "*.tmp").Length);
                }
                Assert.IsTrue(session.TrySaveHistory()); Assert.IsFalse(session.HistorySaveFailed);
                Assert.AreEqual(2, sandbox.Store.Read(id)!.Turns.Count); Assert.IsTrue(session.TrySaveHistory());
                Assert.IsTrue(session.Clear()); Assert.AreEqual(2, sandbox.Store.Read(id)!.Turns.Count);
            }
        }

        [TestMethod]
        public async Task DeletedActiveHistory_IsNotRestoredByNextTurn()
        {
            using (var sandbox = new HistorySandbox())
            {
                var session = new DocumentChatSession(sandbox.Store); await Send(session, "deleted question"); sandbox.Store.DeleteAll();
                await Send(session, "future question");
                var record = sandbox.Store.Read(sandbox.Store.List(out _).Single().SessionId)!;
                Assert.AreEqual(1, record.Turns.Count); Assert.AreEqual("future question", record.Turns[0].Question);
            }
        }

        [DataTestMethod]
        [DataRow("Word", "zh-CN")] [DataRow("Excel", "zh-CN")] [DataRow("PowerPoint", "zh-CN")]
        [DataRow("Word", "en")] [DataRow("Excel", "en")] [DataRow("PowerPoint", "en")]
        public void Dialog_ListsAndShowsRawReadOnlyDetails_ConfirmsDeletes_AndLocalizesEmptyState(string host, string language)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var sandbox = new HistorySandbox())
                {
                    var older = Add(sandbox.Store, Time); var newer = Add(sandbox.Store, Time.AddHours(1), "new question");
                    var type = DocumentChatHostTests.Host(host).GetType("OfficeTranslate.WordAddIn.ChatHistoryForm", true)!;
                    using (var form = (Form)Activator.CreateInstance(type, sandbox.Store, language)!)
                    {
                        const BindingFlags flags = BindingFlags.Instance | BindingFlags.NonPublic;
                        object Field(string name) => type.GetField(name, flags)!.GetValue(form)!;
                        void Call(string method) => type.GetMethod(method, flags)!.Invoke(form, null);
                        form.Show(); Application.DoEvents();
                        Assert.AreEqual(UiText.Get(language, "ChatHistoryTitle"), form.Text);
                        var list = (ListView)Field("_list"); var details = (TextBox)Field("_details");
                        Assert.AreEqual(2, list.Items.Count); Assert.AreEqual(newer, list.Items[0].Tag);
                        CollectionAssert.AreEqual(new[] { "ChatHistoryTime", "ChatHistoryDocument", "ChatHistoryQuestion" }.Select(k => UiText.Get(language, k)).ToArray(), list.Columns.Cast<ColumnHeader>().Select(c => c.Text).ToArray());
                        list.Items[0].Selected = true; Application.DoEvents(); Assert.IsTrue(details.ReadOnly);
                        StringAssert.Contains(details.Text, "new question"); StringAssert.Contains(details.Text, "**raw**");
                        StringAssert.Contains(details.Text, "## 回答\r\n**raw**\r\n```cs\r\nx++;\r\n```");
                        StringAssert.Contains(details.Text, UiText.Get(language, "ChatExportAnswer")); StringAssert.Contains(details.Text, Time.ToString("O"));
                        var evidence = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../../../TestResults/chat-history-2.3.3"));
                        Directory.CreateDirectory(evidence);
                        using (var bitmap = new System.Drawing.Bitmap(form.ClientSize.Width, form.ClientSize.Height)) { form.Controls[0].DrawToBitmap(bitmap, form.ClientRectangle); bitmap.Save(Path.Combine(evidence, host + "-" + language + "-details.png")); }
                        string confirmedKey = ""; bool allow = false;
                        type.GetProperty("ConfirmDelete", flags)!.SetValue(form, (Func<string, bool>)(key => { confirmedKey = key; return allow; }));
                        Call("DeleteSelected"); Assert.AreEqual("ChatHistoryDeleteConfirm", confirmedKey); Assert.AreEqual(2, sandbox.Store.List(out _).Count);
                        allow = true; ((Button)Field("_delete")).PerformClick(); Assert.AreEqual(1, list.Items.Count); Assert.AreEqual(older, list.Items[0].Tag); Assert.AreEqual("", details.Text);
                        allow = false; ((Button)Field("_deleteAll")).PerformClick(); Assert.AreEqual("ChatHistoryDeleteAllConfirm", confirmedKey); Assert.AreEqual(1, list.Items.Count);
                        allow = true; ((Button)Field("_deleteAll")).PerformClick(); Assert.AreEqual(0, list.Items.Count);
                        Assert.AreEqual(UiText.Get(language, "ChatHistoryEmpty"), ((Label)Field("_status")).Text);
                        Assert.IsFalse(((Button)Field("_delete")).Enabled); Assert.IsFalse(((Button)Field("_deleteAll")).Enabled);
                        using (var bitmap = new System.Drawing.Bitmap(form.ClientSize.Width, form.ClientSize.Height)) { form.Controls[0].DrawToBitmap(bitmap, form.ClientRectangle); bitmap.Save(Path.Combine(evidence, host + "-" + language + "-empty.png")); }
                    }
                }
            });
        }
    }
}
