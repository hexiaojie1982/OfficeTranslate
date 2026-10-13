using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ExcelParallelCellsTests
    {
        [TestMethod]
        public void Settings_DefaultAndJsonRoundTrip()
        {
            var json = new JavaScriptSerializer();
            Assert.AreEqual(4, new TranslationSettings().MaxParallelCells);
            Assert.AreEqual(4, json.Deserialize<TranslationSettings>("{}").MaxParallelCells);
            var saved = json.Serialize(new TranslationSettings { MaxParallelCells = 8 });
            Assert.AreEqual(8, json.Deserialize<TranslationSettings>(saved).MaxParallelCells);
        }

        [DataTestMethod]
        [DataRow(0)]
        [DataRow(-1)]
        [DataRow(33)]
        public void Settings_InvalidConcurrencyRejected(int value)
        {
            Assert.ThrowsException<InvalidOperationException>(() =>
                new TranslationSettings { MaxParallelCells = value }.Validate());
        }

        [DataTestMethod]
        [DataRow(1, false)]
        [DataRow(4, false)]
        [DataRow(4, true)]
        [Timeout(20000)]
        public async Task BoundedRequests_ReverseCompletion_WritesCorrectTargetAndProgress(int limit, bool bilingual)
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            using (var started = new SemaphoreSlim(0))
            using (var written = new SemaphoreSlim(0))
            {
                var count = limit * 2;
                var gates = Enumerable.Range(0, count).Select(_ => NewGate<TranslationResult>()).ToArray();
                var entered = new ConcurrentQueue<int>();
                var writes = new List<int>();
                var reports = new List<string>();
                var active = 0;
                var peak = 0;
                var targets = Enumerable.Range(0, count).Select(i => new ExcelTextTarget(i.ToString(), (text, mode) =>
                {
                    Assert.IsTrue(pump.Dispatcher.IsUiThread);
                    Assert.AreEqual(bilingual, mode);
                    Assert.AreEqual((bilingual ? i + Environment.NewLine : "") + "译文" + i, text);
                    writes.Add(i);
                    written.Release();
                })).ToArray();
                var summary = new TranslationTaskSummary(count + 2); // image items remain for the serial OCR loop
                var run = ExcelTextTranslation.RunAsync(targets,
                    new TranslationSettings { MaxParallelCells = limit, BilingualMode = bilingual }, summary,
                    async (text, token) =>
                    {
                        var i = int.Parse(text);
                        var now = Interlocked.Increment(ref active);
                        int old;
                        do { old = Volatile.Read(ref peak); }
                        while (now > old && Interlocked.CompareExchange(ref peak, now, old) != old);
                        entered.Enqueue(i);
                        started.Release();
                        try { return await gates[i].Task.ConfigureAwait(false); }
                        finally { Interlocked.Decrement(ref active); }
                    }, CancellationToken.None, message =>
                    {
                        Assert.IsTrue(pump.Dispatcher.IsUiThread);
                        reports.Add(message);
                    }, pump.Dispatcher);
                try
                {
                    for (var i = 0; i < limit; i++) Assert.IsTrue(await started.WaitAsync(5000));
                    Assert.AreEqual(limit, entered.Count, "The producer must stop at the configured limit.");
                    for (var i = limit - 1; i >= 0; i--)
                    {
                        gates[i].SetResult(TranslationResult.Translated("译文" + i + "\r\n"));
                        Assert.IsTrue(await written.WaitAsync(5000));
                    }
                    for (var i = 0; i < limit; i++) Assert.IsTrue(await started.WaitAsync(5000));
                    for (var i = count - 1; i >= limit; i--)
                    {
                        gates[i].SetResult(TranslationResult.Translated("译文" + i));
                        Assert.IsTrue(await written.WaitAsync(5000));
                    }
                    await run;
                    Assert.AreEqual(limit, peak);
                    Assert.AreEqual(count, summary.Translated);
                    CollectionAssert.AreEqual(Enumerable.Range(0, limit).Reverse()
                        .Concat(Enumerable.Range(limit, limit).Reverse()).ToArray(), writes.ToArray());
                    CollectionAssert.AreEqual(Enumerable.Range(1, count)
                        .Select(i => $"OfficeTranslate：已完成 {i}/{count + 2}").ToArray(), reports.Skip(1).ToArray());
                }
                finally
                {
                    foreach (var gate in gates) gate.TrySetResult(TranslationResult.Skipped(""));
                    await run;
                }
            }
        }

        [TestMethod]
        [Timeout(20000)]
        public async Task Cancel_StopsQueuedWritesAndWaitingCells_DrainsBeforeReturning()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            using (var cancel = new CancellationTokenSource())
            using (var entered = new SemaphoreSlim(0))
            using (var blockEntered = new ManualResetEventSlim())
            using (var releaseUi = new ManualResetEventSlim())
            {
                var result = NewGate<TranslationResult>();
                var calls = 0;
                var writes = 0;
                var summary = new TranslationTaskSummary(20);
                var targets = Enumerable.Range(0, 20).Select(i => new ExcelTextTarget("source", (_, __) => writes++)).ToArray();
                var run = ExcelTextTranslation.RunAsync(targets, new TranslationSettings(), summary,
                    async (_, token) => { Interlocked.Increment(ref calls); entered.Release(); return await result.Task; },
                    cancel.Token, _ => { }, pump.Dispatcher);
                for (var i = 0; i < 4; i++) Assert.IsTrue(await entered.WaitAsync(5000));
                var blocker = pump.Dispatcher.InvokeAsync(() => { blockEntered.Set(); releaseUi.Wait(10000); });
                try
                {
                    Assert.IsTrue(blockEntered.Wait(5000));
                    // Queue a real dispatcher callback behind the blocker, then cancel.
                    var lateWrite = pump.Dispatcher.InvokeAsync(() => { cancel.Token.ThrowIfCancellationRequested(); writes++; });
                    cancel.Cancel();
                    Assert.IsFalse(run.IsCompleted, "Must drain translations that have not observed cancellation yet.");
                    result.SetResult(TranslationResult.Translated("译文"));
                    releaseUi.Set();
                    await Assert.ThrowsExceptionAsync<OperationCanceledException>(async () => await lateWrite);
                    await Assert.ThrowsExceptionAsync<OperationCanceledException>(async () => await run);
                    Assert.AreEqual(4, calls);
                    Assert.AreEqual(0, writes);
                    Assert.AreEqual(0, summary.Translated);
                }
                finally { releaseUi.Set(); result.TrySetResult(TranslationResult.Skipped("")); await blocker; }
            }
        }

        [TestMethod]
        [Timeout(20000)]
        public async Task Failure_CancelsSiblings_PreservesWrittenSummaryAndOriginalError()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            using (var started = new SemaphoreSlim(0))
            using (var written = new SemaphoreSlim(0))
            {
                var gates = Enumerable.Range(0, 10).Select(_ => NewGate<TranslationResult>()).ToArray();
                var summary = new TranslationTaskSummary(10);
                var calls = 0;
                var writes = 0;
                var targets = Enumerable.Range(0, 10).Select(i => new ExcelTextTarget(i.ToString(), (_, __) =>
                { writes++; written.Release(); })).ToArray();
                var run = ExcelTextTranslation.RunAsync(targets, new TranslationSettings(), summary,
                    async (text, token) =>
                    {
                        var gate = gates[int.Parse(text)];
                        using (token.Register(() => gate.TrySetCanceled()))
                        {
                            Interlocked.Increment(ref calls); started.Release();
                            return await gate.Task;
                        }
                    }, CancellationToken.None, _ => { }, pump.Dispatcher);
                for (var i = 0; i < 4; i++) Assert.IsTrue(await started.WaitAsync(5000));
                gates[0].SetResult(TranslationResult.Translated("译文"));
                Assert.IsTrue(await written.WaitAsync(5000));
                Assert.IsTrue(await started.WaitAsync(5000));
                var original = new InvalidOperationException("model failed");
                gates[1].SetException(original);
                var caught = await Assert.ThrowsExceptionAsync<InvalidOperationException>(async () => await run);
                Assert.AreSame(original, caught);
                Assert.AreEqual(5, calls);
                Assert.AreEqual(1, writes);
                Assert.AreEqual(1, summary.Translated);
            }
        }

        [TestMethod]
        public void Summary_ConcurrentMixedResults_ExactCounts()
        {
            const int count = 10000;
            var summary = new TranslationTaskSummary(count * 5);
            Parallel.For(0, count, _ =>
            {
                summary.Record(TranslationResult.Translated("ok"));
                summary.Record(TranslationResult.Cached("cache"));
                summary.Record(TranslationResult.Skipped("skip"));
                summary.Record(TranslationResult.Translated("partial", true, 2, 1));
                summary.Record(TranslationResult.Unchanged("review", 3));
            });
            Assert.AreEqual(count * 2, summary.Translated);
            Assert.AreEqual(count, summary.CacheHits);
            Assert.AreEqual(count, summary.SkippedNoSource);
            Assert.AreEqual(count * 2, summary.NeedsReview);
            Assert.AreEqual(count * 2, summary.TextNeedsReview);
            Assert.AreEqual(count, summary.TextPartiallyTranslated);
            Assert.AreEqual(count, summary.TextUnchangedNeedsReview);
            Assert.AreEqual(count * 2, summary.FallbackTranslatedParagraphs);
            Assert.AreEqual(count * 4, summary.FallbackReviewParagraphs);
        }

        [TestMethod]
        [Timeout(15000)]
        public async Task Cancel_BetweenNetworkCompletionAndUiWrite_DoesNotWrite()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            using (var cancel = new CancellationTokenSource())
            using (var entered = new ManualResetEventSlim())
            using (var release = new ManualResetEventSlim())
            {
                // Inline continuations let SetResult return only after the worker has
                // queued its write on the blocked UI thread.
                var network = new TaskCompletionSource<TranslationResult>();
                var started = NewGate<bool>();
                var summary = new TranslationTaskSummary(1);
                var writes = 0;
                var run = ExcelTextTranslation.RunAsync(new[] { new ExcelTextTarget("source", (_, __) => writes++) },
                    new TranslationSettings(), summary, (_, __) => { started.SetResult(true); return network.Task; },
                    cancel.Token, _ => { }, pump.Dispatcher);
                await started.Task;
                var blocker = pump.Dispatcher.InvokeAsync(() => { entered.Set(); release.Wait(10000); });
                try
                {
                    Assert.IsTrue(entered.Wait(5000));
                    await Task.Run(() => network.SetResult(TranslationResult.Translated("译文")));
                    Assert.IsFalse(run.IsCompleted);
                    cancel.Cancel();
                }
                finally { release.Set(); await blocker; }
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(async () => await run);
                Assert.AreEqual(0, writes);
                Assert.AreEqual(0, summary.Translated);
            }
        }

        [TestMethod]
        public async Task AlreadyCanceled_StartsNoRequests()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            {
                var summary = new TranslationTaskSummary(1);
                var run = ExcelTextTranslation.RunAsync(new[] { new ExcelTextTarget("source", (_, __) => Assert.Fail()) },
                    new TranslationSettings(), summary, (_, __) => throw new AssertFailedException("Unexpected request"),
                    new CancellationToken(true), _ => Assert.Fail(), pump.Dispatcher);
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(async () => await run);
                Assert.AreEqual(0, summary.Translated);
            }
        }

        [TestMethod]
        public async Task UnchangedAndSkipped_DoNotWriteButAdvanceProgress()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            {
                var summary = new TranslationTaskSummary(2);
                var reports = new List<string>();
                await ExcelTextTranslation.RunAsync(new[] {
                    new ExcelTextTarget("skip", (_, __) => Assert.Fail()),
                    new ExcelTextTarget("review", (_, __) => Assert.Fail()) }, new TranslationSettings(), summary,
                    (text, _) => Task.FromResult(text == "skip" ? TranslationResult.Skipped(text) : TranslationResult.Unchanged(text)),
                    CancellationToken.None, reports.Add, pump.Dispatcher);
                Assert.AreEqual(1, summary.SkippedNoSource);
                Assert.AreEqual(1, summary.TextUnchangedNeedsReview);
                Assert.AreEqual("OfficeTranslate：已完成 2/2", reports.Last());
            }
        }

        [TestMethod]
        public async Task WriteFailure_IsNotCountedAsTranslated()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            {
                var summary = new TranslationTaskSummary(1);
                var error = new InvalidOperationException("write failed");
                var run = ExcelTextTranslation.RunAsync(new[] { new ExcelTextTarget("source", (_, __) => throw error) },
                    new TranslationSettings(), summary, (_, __) => Task.FromResult(TranslationResult.Translated("译文")),
                    CancellationToken.None, _ => { }, pump.Dispatcher);
                Assert.AreSame(error, await Assert.ThrowsExceptionAsync<InvalidOperationException>(async () => await run));
                Assert.AreEqual(0, summary.Translated);
            }
        }

        [TestMethod]
        [Timeout(20000)]
        public async Task Dispatcher_MultiplePoolCallers_AllCallbacksUseCapturedSta()
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            {
                var ids = new HashSet<int>();
                var tasks = Enumerable.Range(0, 100).Select(i => Task.Run(() => pump.Dispatcher.InvokeAsync(() =>
                {
                    Assert.AreEqual(pump.StaThreadId, Thread.CurrentThread.ManagedThreadId);
                    Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                    Assert.IsTrue(ids.Add(i));
                    return i;
                }))).ToArray();
                var results = await Task.WhenAll(tasks);
                CollectionAssert.AreEqual(Enumerable.Range(0, 100).ToArray(), results);
                Assert.AreEqual(100, ids.Count);
            }
        }

        private static TaskCompletionSource<T> NewGate<T>() =>
            new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
    }
}
