using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using OfficeTranslate.WordAddIn;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class WordParallelTranslationTests
    {
        private static TaskCompletionSource<T> Gate<T>() => new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        private static string Format(string key, int current, int total) => string.Format(UiText.Get("zh-CN", key), current, total);
        private static Task Run(IReadOnlyList<WordTextTarget> targets, TranslationTaskSummary summary,
            Func<string, CancellationToken, Task<TranslationResult>> translate, OfficeUiDispatcher ui,
            int limit = 4, CancellationToken token = default, Action<string>? progress = null, int images = 0) =>
            WordTextTranslation.RunAsync(targets, new TranslationSettings { MaxParallelTextBlocks = limit },
                summary, images, translate, token, progress ?? (_ => { }), ui, Format);

        [DataTestMethod]
        [DataRow(1)]
        [DataRow(4)]
        [DataRow(32)]
        [Timeout(20000)]
        public async Task ReverseCompletion_BoundedWindow_OrderedIncrementalCommits(int limit)
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            using (var started = new SemaphoreSlim(0))
            using (var committed = new SemaphoreSlim(0))
            {
                var count = limit + 1;
                var gates = Enumerable.Range(0, count).Select(_ => Gate<TranslationResult>()).ToArray();
                var writes = new List<int>();
                var reports = new List<string>();
                var calls = 0;
                var active = 0;
                var peak = 0;
                var summary = new TranslationTaskSummary(count);
                var targets = Enumerable.Range(0, count).Select(i => new WordTextTarget(i.ToString(), value =>
                {
                    Assert.IsTrue(pump.Dispatcher.IsUiThread);
                    Assert.AreEqual("译文" + i, value);
                    writes.Add(i); committed.Release();
                })).ToArray();
                var run = Run(targets, summary, async (text, token) =>
                {
                    var now = Interlocked.Increment(ref active);
                    int previous;
                    do { previous = Volatile.Read(ref peak); }
                    while (now > previous && Interlocked.CompareExchange(ref peak, now, previous) != previous);
                    Interlocked.Increment(ref calls); started.Release();
                    try { return await gates[int.Parse(text)].Task; }
                    finally { Interlocked.Decrement(ref active); }
                }, pump.Dispatcher, limit, progress: message =>
                { Assert.IsTrue(pump.Dispatcher.IsUiThread); reports.Add(message); });
                try
                {
                    for (var i = 0; i < limit; i++) Assert.IsTrue(await started.WaitAsync(5000));
                    Assert.AreEqual(limit, Volatile.Read(ref calls));
                    // Later responses finish first; no write is allowed past the missing head.
                    for (var i = limit - 1; i > 0; i--) gates[i].SetResult(TranslationResult.Translated("译文" + i));
                    pump.Dispatcher.Invoke(() => Assert.AreEqual(0, writes.Count));
                    gates[0].SetResult(TranslationResult.Translated("译文0\r\a"));
                    for (var i = 0; i < limit; i++) Assert.IsTrue(await committed.WaitAsync(5000));
                    Assert.IsFalse(run.IsCompleted, "First wave must be visible before the last request finishes.");
                    Assert.IsTrue(await started.WaitAsync(5000));
                    gates[limit].SetResult(TranslationResult.Translated("译文" + limit));
                    await run;
                    Assert.AreEqual(limit, peak);
                    CollectionAssert.AreEqual(Enumerable.Range(0, count).ToArray(), writes.ToArray());
                    Assert.AreEqual(count, summary.Translated);
                    Assert.AreEqual(count * 3, reports.Count);
                    CollectionAssert.AreEqual(Enumerable.Range(1, count).Select(i => Format("TextCompleted", i, count)).ToArray(),
                        reports.Where(m => m.Contains("已完成")).ToArray());
                    if (limit == 1)
                        CollectionAssert.AreEqual(Enumerable.Range(1, count).SelectMany(i => new[] {
                            Format("TextTranslating", i, count), Format("TextWriting", i, count), Format("TextCompleted", i, count)
                        }).ToArray(), reports.ToArray(), "Concurrency 1 must retain the serial event sequence.");
                }
                finally { foreach (var gate in gates) gate.TrySetResult(TranslationResult.Skipped("")); await run; }
            }
        }

        [TestMethod]
        [Timeout(15000)]
        public async Task CancelAfterFirstCommit_PreservesIt_CancelsOutstanding_StopsProducer()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            using (var cancel = new CancellationTokenSource())
            using (var started = new SemaphoreSlim(0))
            {
                var first = Gate<TranslationResult>();
                var canceled = 0;
                var calls = 0;
                var writes = new List<int>();
                var summary = new TranslationTaskSummary(10);
                var run = Run(Enumerable.Range(0, 10).Select(i => new WordTextTarget(i.ToString(), _ => writes.Add(i))).ToArray(),
                    summary, async (text, token) =>
                    {
                        Interlocked.Increment(ref calls); started.Release();
                        if (text == "0") return await first.Task;
                        try { await Task.Delay(Timeout.Infinite, token); return TranslationResult.Translated("unreachable"); }
                        catch (OperationCanceledException) { Interlocked.Increment(ref canceled); throw; }
                    }, pump.Dispatcher, token: cancel.Token, progress: message =>
                    { if (message == Format("TextCompleted", 1, 10)) cancel.Cancel(); });
                for (var i = 0; i < 4; i++) Assert.IsTrue(await started.WaitAsync(5000));
                first.SetResult(TranslationResult.Translated("first"));
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(async () => await run);
                CollectionAssert.AreEqual(new[] { 0 }, writes.ToArray());
                Assert.AreEqual(4, calls); Assert.AreEqual(3, canceled);
                Assert.AreEqual(1, summary.Translated); Assert.AreEqual(0, summary.TextBlockFailures.Count);
            }
        }

        [TestMethod]
        [Timeout(15000)]
        public async Task CancelWhileWriteQueued_DrainsWorkerIgnoringToken_NoLateMutation()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            using (var cancel = new CancellationTokenSource())
            using (var entered = new ManualResetEventSlim())
            using (var release = new ManualResetEventSlim())
            using (var started = new SemaphoreSlim(0))
            {
                // Inline completion proves the first consumer queued its commit before cancel.
                var first = new TaskCompletionSource<TranslationResult>();
                var second = Gate<TranslationResult>();
                var writes = 0;
                var summary = new TranslationTaskSummary(2);
                var run = Run(new[] { new WordTextTarget("0", _ => writes++), new WordTextTarget("1", _ => writes++) }, summary,
                    (text, _) => { started.Release(); return text == "0" ? first.Task : second.Task; },
                    pump.Dispatcher, token: cancel.Token);
                for (var i = 0; i < 2; i++) Assert.IsTrue(await started.WaitAsync(5000));
                var blocker = pump.Dispatcher.InvokeAsync(() => { entered.Set(); release.Wait(10000); });
                try
                {
                    Assert.IsTrue(entered.Wait(5000));
                    await Task.Run(() => first.SetResult(TranslationResult.Translated("first")));
                    cancel.Cancel(); release.Set(); await blocker;
                    Assert.IsFalse(run.IsCompleted, "An uncooperative request must drain before return.");
                    second.SetResult(TranslationResult.Translated("second"));
                    await Assert.ThrowsExceptionAsync<OperationCanceledException>(async () => await run);
                    Assert.AreEqual(0, writes); Assert.AreEqual(0, summary.Translated);
                }
                finally { release.Set(); first.TrySetCanceled(); second.TrySetCanceled(); await blocker; }
            }
        }

        [TestMethod]
        public async Task TranslationAndWriteFailures_Continue_OrderedSafeErrorSummary()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            {
                var writes = new List<int>();
                var reports = new List<string>();
                var summary = new TranslationTaskSummary(7);
                var targets = Enumerable.Range(0, 7).Select(i => new WordTextTarget(i.ToString(), _ =>
                { if (i == 3) throw new IOException("private document text"); writes.Add(i); })).ToArray();
                await Run(targets, summary, (text, token) =>
                {
                    if (text == "1") throw new InvalidOperationException("secret provider echo");
                    if (text == "4") return Task.FromException<TranslationResult>(new TaskCanceledException("provider timeout"));
                    return Task.FromResult(TranslationResult.Translated("translated"));
                }, pump.Dispatcher, progress: reports.Add);
                CollectionAssert.AreEqual(new[] { 0, 2, 5, 6 }, writes.ToArray());
                CollectionAssert.AreEqual(new[] { 2, 4, 5 }, summary.TextBlockFailures.Select(f => f.BlockNumber).ToArray());
                CollectionAssert.AreEqual(new[] { false, true, false }, summary.TextBlockFailures.Select(f => f.DuringWrite).ToArray());
                Assert.AreEqual("TaskCanceledException", summary.TextBlockFailures[2].ErrorType);
                Assert.AreEqual(4, summary.Translated); Assert.AreEqual(3, summary.NeedsReview);
                Assert.AreEqual(3, summary.TextNeedsReview); Assert.AreEqual(0, summary.TextUnchangedNeedsReview);
                Assert.AreEqual(7, reports.Count(m => m.Contains("已完成")));
            }
        }

        [TestMethod]
        public async Task MixedResults_SkipChangedCacheAndFallback_ExactProgressWithImageOffset()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            {
                var writes = new List<string>();
                var reports = new List<string>();
                var calls = 0;
                var sources = new[] { "  \t", "!?", "unchanged", "cached", "partial", "skipped" };
                var summary = new TranslationTaskSummary(8);
                await Run(sources.Select(text => new WordTextTarget(text, writes.Add)).ToArray(), summary, (text, token) =>
                {
                    Interlocked.Increment(ref calls);
                    return Task.FromResult(text == "unchanged" ? TranslationResult.Unchanged(text, 2) :
                        text == "cached" ? TranslationResult.Cached("cache") : text == "partial" ?
                        TranslationResult.Translated("partial", true, 1, 1) : TranslationResult.Skipped(text));
                }, pump.Dispatcher, progress: reports.Add, images: 2);
                Assert.AreEqual(4, calls); CollectionAssert.AreEqual(new[] { "cache", "partial" }, writes.ToArray());
                Assert.AreEqual(1, summary.Translated); Assert.AreEqual(1, summary.CacheHits);
                Assert.AreEqual(3, summary.SkippedNoSource); Assert.AreEqual(2, summary.NeedsReview);
                Assert.AreEqual(1, summary.TextUnchangedNeedsReview); Assert.AreEqual(1, summary.TextPartiallyTranslated);
                Assert.AreEqual(1, summary.FallbackTranslatedParagraphs); Assert.AreEqual(3, summary.FallbackReviewParagraphs);
                Assert.AreEqual(14, reports.Count);
                CollectionAssert.AreEqual(Enumerable.Range(3, 6).Select(i => Format("TextCompleted", i, 8)).ToArray(),
                    reports.Where(m => m.Contains("已完成")).ToArray());
                CollectionAssert.AreEqual(new[] { Format("TextWriting", 4, 6), Format("TextWriting", 5, 6) },
                    reports.Where(m => m.Contains("正在写回")).ToArray());
            }
        }

        [DataTestMethod]
        [DataRow(0)]
        [DataRow(1)]
        public async Task EmptyOrNoTranslatableText_NoNetworkOrWrite(int count)
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            {
                var reports = new List<string>();
                var summary = new TranslationTaskSummary(count);
                await Run(Enumerable.Range(0, count).Select(_ => new WordTextTarget(" \t!?", __ => Assert.Fail())).ToArray(),
                    summary, (_, __) => throw new AssertFailedException("Must skip network"), pump.Dispatcher, progress: reports.Add);
                Assert.AreEqual(count, summary.SkippedNoSource); Assert.AreEqual(count * 2, reports.Count);
            }
        }

        [TestMethod]
        public async Task PreCanceled_DoesNotStartOrReport()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            using (var cancel = new CancellationTokenSource())
            {
                cancel.Cancel();
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => Run(new[] { new WordTextTarget("text", _ => Assert.Fail()) },
                    new TranslationTaskSummary(1), (_, __) => throw new AssertFailedException(), pump.Dispatcher,
                    token: cancel.Token, progress: _ => Assert.Fail()));
            }
        }

        [TestMethod]
        public async Task CancellationInsideStartedCommit_FinishesBlockAndRecordsIt_ThenStops()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            using (var cancel = new CancellationTokenSource())
            {
                var state = "original";
                var summary = new TranslationTaskSummary(2);
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => Run(new[] {
                    new WordTextTarget("first", value => { state = "commit started"; cancel.Cancel(); state = value; }),
                    new WordTextTarget("second", _ => Assert.Fail("No following write is allowed"))
                }, summary, (_, __) => Task.FromResult(TranslationResult.Translated("complete")), pump.Dispatcher, token: cancel.Token));
                Assert.AreEqual("complete", state); Assert.AreEqual(1, summary.Translated);
                Assert.AreEqual(0, summary.TextBlockFailures.Count);
            }
        }

        [TestMethod]
        public async Task CancellationFromStartingProgress_DoesNotSendRequest()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            using (var cancel = new CancellationTokenSource())
            {
                var summary = new TranslationTaskSummary(1);
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => Run(
                    new[] { new WordTextTarget("source", _ => Assert.Fail("No write after cancellation")) }, summary,
                    (_, __) => throw new AssertFailedException("No request after progress cancels"), pump.Dispatcher,
                    token: cancel.Token, progress: _ => cancel.Cancel()));
                Assert.AreEqual(0, summary.Translated);
                Assert.AreEqual(0, summary.TextBlockFailures.Count);
            }
        }

        [TestMethod]
        [Timeout(15000)]
        public async Task ProgressFailure_DrainsWorkers_WithoutMaskingOriginalError()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            using (var entered = new ManualResetEventSlim())
            using (var release = new ManualResetEventSlim())
            {
                var blocker = pump.Dispatcher.InvokeAsync(() => { entered.Set(); release.Wait(10000); });
                try
                {
                    Assert.IsTrue(entered.Wait(5000));
                    var run = Run(new[] { new WordTextTarget("first", _ => Assert.Fail()),
                        new WordTextTarget("second", _ => Assert.Fail()) }, new TranslationTaskSummary(2),
                        (_, __) => Task.FromResult(TranslationResult.Translated("result")), pump.Dispatcher,
                        progress: message =>
                        {
                            if (message == Format("TextTranslating", 2, 2)) throw new IOException("secondary worker callback");
                            if (message == Format("TextWriting", 1, 2)) throw new InvalidOperationException("original consumer callback");
                        });
                    release.Set();
                    var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(async () => await run);
                    Assert.AreEqual("original consumer callback", error.Message);
                }
                finally { release.Set(); await blocker; }
            }
        }

        [TestMethod]
        public void Settings_DefaultLegacyJsonAndRoundTrip()
        {
            var json = new JavaScriptSerializer();
            Assert.AreEqual(4, new TranslationSettings().MaxParallelTextBlocks);
            Assert.AreEqual(4, json.Deserialize<TranslationSettings>("{\"MaxParallelCells\":8}").MaxParallelTextBlocks);
            var copy = json.Deserialize<TranslationSettings>(json.Serialize(new TranslationSettings { MaxParallelTextBlocks = 1, MaxParallelCells = 8 }));
            Assert.AreEqual(1, copy.MaxParallelTextBlocks); Assert.AreEqual(8, copy.MaxParallelCells);
        }

        [DataTestMethod]
        [DataRow(0)]
        [DataRow(-1)]
        [DataRow(33)]
        public void Settings_RejectInvalidConcurrency(int value) => Assert.ThrowsException<InvalidOperationException>(() =>
            new TranslationSettings { MaxParallelTextBlocks = value }.Validate());

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public async Task OversizeSingleBlock_ExistingClientSplitting_WholeBlockCommitOrFailure(bool failSecondChunk)
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            using (var handler = new LongBlockHandler(failSecondChunk))
            using (var client = new TranslationClient(handler))
            {
                var settings = new TranslationSettings { Provider = ProviderKind.OpenAiCompatible, BaseUrl = "https://fixture.invalid/v1",
                    Model = Guid.NewGuid().ToString(), SourceLanguage = "自动检测", TargetLanguage = "简体中文", MaxCharactersPerChunk = 500 };
                var writes = new List<string>();
                var summary = new TranslationTaskSummary(1);
                await WordTextTranslation.RunAsync(new[] { new WordTextTarget(new string('x', 1600), writes.Add) },
                    settings, summary, 0, (text, token) => client.TranslateDetailedAsync(text, settings, token),
                    CancellationToken.None, _ => { }, pump.Dispatcher, Format);
                Assert.IsTrue(handler.Sizes.All(n => n <= 500));
                if (failSecondChunk)
                {
                    Assert.AreEqual(0, writes.Count, "A partial long-block response must never reach Word.");
                    Assert.AreEqual(1, summary.TextBlockFailures.Count); Assert.AreEqual(0, summary.Translated);
                }
                else
                {
                    CollectionAssert.AreEqual(new[] { 500, 500, 500, 100 }, handler.Sizes.ToArray());
                    Assert.AreEqual(1, writes.Count); Assert.AreEqual(1, summary.Translated);
                    Assert.AreEqual(string.Concat(Enumerable.Repeat(LongBlockHandler.Translation, 4)), writes.Single());
                }
            }
        }

        private sealed class LongBlockHandler : HttpMessageHandler
        {
            public const string Translation = "这是一段用于验证超长文本完整翻译行为的中文结果。";
            private readonly bool _fail;
            public readonly List<int> Sizes = new List<int>();
            public LongBlockHandler(bool fail) => _fail = fail;
            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
            {
                var json = new JavaScriptSerializer();
                var payload = (Dictionary<string, object>)json.DeserializeObject(await request.Content!.ReadAsStringAsync());
                var messages = (object[])payload["messages"];
                Sizes.Add(((string)((Dictionary<string, object>)messages[1])["content"]).Length);
                if (_fail && Sizes.Count >= 2) return new HttpResponseMessage(HttpStatusCode.BadRequest)
                    { Content = new StringContent("fixture rejected second chunk") };
                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(json.Serialize(new {
                    choices = new[] { new { message = new { content = Translation }, finish_reason = "stop" } }
                })) };
            }
        }

        [DataTestMethod]
        [DataRow("Word")]
        [DataRow("Excel")]
        [DataRow("PowerPoint")]
        public void SettingsForm_LoadEditPersist_PreservesIndependentExcelLimit(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                var configuration = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent!.Name;
                var assembly = Assembly.LoadFrom(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "../../../../../src/OfficeTranslate." + host + "AddIn/bin/" + configuration + "/net48/OfficeTranslate." + host + "AddIn.dll")));
                var file = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".json");
                var store = new SettingsStore();
                typeof(SettingsStore).GetField("_path", BindingFlags.NonPublic | BindingFlags.Instance)!.SetValue(store, file);
                try
                {
                    store.Save(new TranslationSettings { MaxParallelTextBlocks = 7, MaxParallelCells = 9 });
                    using (var form = (Form)Activator.CreateInstance(assembly.GetType("OfficeTranslate.WordAddIn.SettingsForm", true)!, store)!)
                    {
                        var field = (NumericUpDown)form.GetType().GetField("_textConcurrency", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(form)!;
                        Assert.AreEqual(7m, field.Value); Assert.AreEqual(1m, field.Minimum); Assert.AreEqual(32m, field.Maximum);
                        field.Value = 2;
                        var settings = (TranslationSettings)form.GetType().GetMethod("BuildSettings", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(form, new object[] { false })!;
                        store.Save(settings);
                        Assert.AreEqual(2, store.Load().MaxParallelTextBlocks); Assert.AreEqual(9, store.Load().MaxParallelCells);
                    }
                }
                finally { if (File.Exists(file)) File.Delete(file); }
            });
        }

        [DataTestMethod]
        [DataRow("zh-CN")]
        [DataRow("zh-HK")]
        [DataRow("ko")]
        [DataRow("ja")]
        [DataRow("en")]
        public void FailuresVisibleLocalizedCopyable_NoAutoClose_AlsoOnCancellation(string language)
        {
            DocumentChatHostTests.Sta(() =>
            {
                var summary = new TranslationTaskSummary(2);
                summary.Record(TranslationResult.Translated("ok"));
                summary.RecordTextFailure(2, false, new InvalidOperationException("PRIVATE-SOURCE-SECRET"));
                foreach (var canceled in new[] { false, true })
                {
                    if (canceled)
                    {
                        using (var cancel = new CancellationTokenSource())
                        {
                            cancel.Cancel();
                            TranslationErrorDialog.ShowTaskFailure(new OperationCanceledException(cancel.Token), summary, cancel.Token, language, null);
                        }
                    }
                    else TranslationResultForm.ShowResult(language, summary, IntPtr.Zero);
                    var form = Application.OpenForms.Cast<Form>().OfType<TranslationResultForm>().Single();
                    try
                    {
                        var details = Descendants(form).OfType<TextBox>().Single();
                        Assert.IsFalse(details.Text.Contains("PRIVATE-SOURCE-SECRET"));
                        StringAssert.Contains(details.Text, string.Format(UiText.Get(language, "TextBlockFailure"), 2,
                            UiText.Get(language, "TextTranslateStage"), "InvalidOperationException"));
                        var timer = (System.Windows.Forms.Timer)typeof(TranslationResultForm).GetField("_timer", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(form)!;
                        Assert.IsFalse(timer.Enabled); Assert.IsTrue(details.ReadOnly);
                        Assert.AreEqual(UiText.Get(language, "Close"), Descendants(form).OfType<Button>().Single().Text);
                    }
                    finally { form.Close(); form.Dispose(); }
                }
            });
        }

        private static IEnumerable<Control> Descendants(Control parent)
        {
            foreach (Control child in parent.Controls)
            { yield return child; foreach (var descendant in Descendants(child)) yield return descendant; }
        }
    }
}
