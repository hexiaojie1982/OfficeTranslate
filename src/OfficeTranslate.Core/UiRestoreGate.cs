using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeTranslate.Core
{
    // A one-shot lease for best-effort UI teardown. No COM objects here: the
    // real host supplies its restore action. Test helpers can use the SAME
    // implementation on an actual Office UI queue without modifying the host.
    public static class UiRestoreGate
    {
        public const int DefaultWaitMs = 2000;

        public static async Task RunAsync(OfficeUiDispatcher ui, Func<bool> isCurrent,
            Action restore, Action<string>? skipped = null, int waitMs = DefaultWaitMs)
        {
            if (ui == null) throw new ArgumentNullException(nameof(ui));
            if (isCurrent == null) throw new ArgumentNullException(nameof(isCurrent));
            if (restore == null) throw new ArgumentNullException(nameof(restore));
            if (waitMs <= 0) throw new ArgumentOutOfRangeException(nameof(waitMs));
            var ticket = new Ticket();
            try
            {
                var queued = ui.InvokeAsync(() =>
                {
                    if (Interlocked.CompareExchange(ref ticket.State, 1, 0) != 0) return;
                    if (!isCurrent()) { skipped?.Invoke("superseded"); return; }
                    if (ticket.ElapsedMs > waitMs) { skipped?.Invoke("stale"); return; }
                    restore();
                });
                // Do not depend on a blocked Office synchronization context to
                // expire the ticket. A started COM call cannot be interrupted.
                var finished = await Task.WhenAny(queued, Task.Delay(waitMs)).ConfigureAwait(false);
                if (ReferenceEquals(finished, queued)) await queued.ConfigureAwait(false);
                else if (Interlocked.CompareExchange(ref ticket.State, 2, 0) == 0) skipped?.Invoke("expired");
            }
            catch { /* best-effort: never mask the task's original outcome */ }
        }

        private sealed class Ticket
        {
            public int State; // pending=0, running=1, expired=2
            private readonly long _posted = Stopwatch.GetTimestamp();
            public double ElapsedMs => (Stopwatch.GetTimestamp() - _posted) * 1000.0 / Stopwatch.Frequency;
        }
    }
}
