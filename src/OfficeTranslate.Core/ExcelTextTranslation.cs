using System;
using System.Collections.Generic;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeTranslate.Core
{
    // The source string is captured on the UI thread; only Write touches COM.
    internal sealed class ExcelTextTarget
    {
        private readonly Action<string, bool> _write;
        public ExcelTextTarget(string text, Action<string, bool> write) { Text = text; _write = write; }
        public string Text { get; }
        public void Write(string value, bool bilingual) => _write(value, bilingual);
    }

    internal static class ExcelTextTranslation
    {
        internal static async Task RunAsync(IReadOnlyList<ExcelTextTarget> targets,
            TranslationSettings settings, TranslationTaskSummary summary,
            Func<string, CancellationToken, Task<TranslationResult>> translate,
            CancellationToken token, Action<string> progress, OfficeUiDispatcher ui)
        {
            settings.ValidateParallelCells();
            var completed = 0;
            ExceptionDispatchInfo? failure = null;
            using (var stop = CancellationTokenSource.CreateLinkedTokenSource(token))
            using (var slots = new SemaphoreSlim(settings.MaxParallelCells, settings.MaxParallelCells))
            {
                var pending = new List<Task>();
                var workToken = stop.Token;
                async Task TranslateOneAsync(ExcelTextTarget target)
                {
                    try
                    {
                        workToken.ThrowIfCancellationRequested();
                        var result = await translate(target.Text, workToken).ConfigureAwait(false);
                        workToken.ThrowIfCancellationRequested();
                        await ui.InvokeAsync(() =>
                        {
                            // Recheck after queuing: cancellation must prevent late writes.
                            workToken.ThrowIfCancellationRequested();
                            if (result.Changed)
                            {
                                var translated = result.Text.TrimEnd('\r', '\n');
                                target.Write(settings.BilingualMode ? target.Text + Environment.NewLine + translated : translated,
                                    settings.BilingualMode);
                            }
                            summary.Record(result);
                            // Increment and report on the same UI queue so progress never regresses.
                            var current = Interlocked.Increment(ref completed);
                            progress($"OfficeTranslate：已完成 {current}/{summary.Total}");
                        }).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        // Retain the original failure, not a sibling's induced cancellation.
                        Interlocked.CompareExchange(ref failure, ExceptionDispatchInfo.Capture(ex), null);
                        stop.Cancel();
                    }
                    finally { slots.Release(); }
                }

                try
                {
                    if (targets.Count > 0)
                        await ui.InvokeAsync(() =>
                        {
                            workToken.ThrowIfCancellationRequested();
                            progress($"OfficeTranslate：正在翻译（已完成 0/{summary.Total}）");
                        }).ConfigureAwait(false);
                    foreach (var target in targets)
                    {
                        await slots.WaitAsync(workToken).ConfigureAwait(false);
                        // Retain only in-flight tasks, even for very large worksheets.
                        pending.RemoveAll(task => task.IsCompleted);
                        pending.Add(TranslateOneAsync(target));
                    }
                }
                catch (OperationCanceledException) when (workToken.IsCancellationRequested) { }
                finally
                {
                    // Do not dispose the semaphore/client or expose Summary while workers can write.
                    await Task.WhenAll(pending).ConfigureAwait(false);
                }
                failure?.Throw();
                token.ThrowIfCancellationRequested();
            }
        }
    }
}
