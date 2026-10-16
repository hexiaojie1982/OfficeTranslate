using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeTranslate.Core
{
    // Text is a snapshot; the write delegate is invoked only on the Office STA.
    internal sealed class WordTextTarget
    {
        public WordTextTarget(string text, Action<string> write) { Text = text; Write = write; }
        public string Text { get; }
        public Action<string> Write { get; }
    }

    internal static class WordTextTranslation
    {
        internal static async Task RunAsync(IReadOnlyList<WordTextTarget> targets,
            TranslationSettings settings, TranslationTaskSummary summary, int imageCount,
            Func<string, CancellationToken, Task<TranslationResult>> translate,
            CancellationToken token, Action<string> progress, OfficeUiDispatcher ui,
            Func<string, int, int, string> formatProgress)
        {
            settings.ValidateParallelTextBlocks();
            // A bounded rolling window limits both requests and buffered results.
            // Only this consumer commits, in source order. No worker touches COM.
            var pending = new Queue<Task<Outcome>>();
            var next = 0;
            using (var stop = CancellationTokenSource.CreateLinkedTokenSource(token))
            {
                var workToken = stop.Token;
                async Task<Outcome> StartAsync(int index)
                {
                    workToken.ThrowIfCancellationRequested();
                    await ui.InvokeAsync(() =>
                    {
                        workToken.ThrowIfCancellationRequested();
                        progress(formatProgress("TextTranslating", imageCount + index + 1, summary.Total));
                    }).ConfigureAwait(false);
                    workToken.ThrowIfCancellationRequested();
                    try
                    {
                        var text = targets[index].Text;
                        var hasText = false;
                        foreach (var c in text) if (char.IsLetterOrDigit(c)) { hasText = true; break; }
                        var result = hasText
                            ? await translate(text, workToken).ConfigureAwait(false)
                            : TranslationResult.Skipped(text);
                        workToken.ThrowIfCancellationRequested();
                        return new Outcome(result, null);
                    }
                    catch (OperationCanceledException) when (workToken.IsCancellationRequested) { throw; }
                    catch (Exception ex) { return new Outcome(null, ex); }
                }

                try
                {
                    for (var index = 0; index < targets.Count; index++)
                    {
                        token.ThrowIfCancellationRequested();
                        while (next < targets.Count && pending.Count < settings.MaxParallelTextBlocks)
                            pending.Enqueue(StartAsync(next++));
                        // Keep the head in the queue until consumed so finally drains it too.
                        var outcome = await pending.Peek().ConfigureAwait(false);
                        var currentIndex = index;
                        await ui.InvokeAsync(() =>
                        {
                            token.ThrowIfCancellationRequested();
                            if (outcome.Error != null)
                                summary.RecordTextFailure(currentIndex + 1, false, outcome.Error);
                            else
                            {
                                var result = outcome.Result!;
                                if (result.Changed)
                                {
                                    progress(formatProgress("TextWriting", currentIndex + 1, targets.Count));
                                    // A progress subscriber can cancel synchronously too.
                                    token.ThrowIfCancellationRequested();
                                    try { targets[currentIndex].Write(result.Text.TrimEnd('\r', '\a')); }
                                    catch (Exception ex) when (!(ex is OperationCanceledException && token.IsCancellationRequested))
                                    {
                                        summary.RecordTextFailure(currentIndex + 1, true, ex);
                                        progress(formatProgress("TextCompleted", imageCount + currentIndex + 1, summary.Total));
                                        return;
                                    }
                                }
                                // Record only after the entire write has committed, even if
                                // cancellation arrives during that synchronous commit.
                                summary.Record(result);
                            }
                            progress(formatProgress("TextCompleted", imageCount + currentIndex + 1, summary.Total));
                        }).ConfigureAwait(false);
                        _ = pending.Dequeue(); // Already awaited and committed above.
                    }
                    token.ThrowIfCancellationRequested();
                }
                finally
                {
                    stop.Cancel();
                    // Observe every worker before disposing the client / exposing Summary.
                    // Worker faults here are induced cancellation or a callback failure;
                    // the original exception from the consumer is preserved.
                    try { await Task.WhenAll(pending).ConfigureAwait(false); }
                    // Every task is consumed on success. If any remain here, the
                    // consumer already failed/canceled; secondary worker/callback
                    // errors must not replace that original exception.
                    catch (Exception) { }
                }
            }
        }

        private sealed class Outcome
        {
            public Outcome(TranslationResult? result, Exception? error) { Result = result; Error = error; }
            public TranslationResult? Result { get; }
            public Exception? Error { get; }
        }
    }
}
