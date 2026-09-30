using System;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeTranslate.Core
{
    // Marshals COM / clipboard / Office-writeback work back to the Office UI
    // (STA) thread that owns the host application.
    //
    // M2: an add-in's async continuations are NOT guaranteed to resume on the
    // Office STA thread. Observed in review: the first menu translation in an
    // Excel session works, the second fails inside clipboard capture with
    // "Current thread must be set to single thread apartment (STA) mode
    // before OLE calls can be made". The ambient SynchronizationContext on
    // the Office thread is fragile (it depends on which WinForms controls
    // happened to be created on it and when), so every operation that touches
    // COM or the clipboard goes through this dispatcher explicitly instead
    // of relying on await context capture. Network I/O stays off the UI
    // thread; only the COM/clipboard/writeback sections are dispatched.
    public sealed class OfficeUiDispatcher
    {
        private readonly Thread _uiThread;
        private readonly SynchronizationContext _context;
        private readonly string _host;

        private OfficeUiDispatcher(Thread uiThread, SynchronizationContext context, string host)
        {
            _uiThread = uiThread;
            _context = context;
            _host = host;
        }

        // Captures the calling thread as the Office UI thread. Must be called
        // on the Office main (STA) thread -- in practice, at the top of the
        // ribbon button handler, before the first await. Office does not
        // install a SynchronizationContext on its main thread, so when none
        // is present one is installed explicitly; Post/Send then have a
        // deterministic pump target on this thread.
        public static OfficeUiDispatcher Capture(string host)
        {
            var thread = Thread.CurrentThread;
            var context = SynchronizationContext.Current;
            if (context == null)
            {
                context = new WindowsFormsSynchronizationContext();
                SynchronizationContext.SetSynchronizationContext(context);
            }
            var dispatcher = new OfficeUiDispatcher(thread, context, host);
            dispatcher.LogProbe("menu_entry");
            return dispatcher;
        }

        public bool IsUiThread => Thread.CurrentThread == _uiThread;

        // Runs func on the Office UI thread and blocks the caller until it
        // finishes. Safe to call from the UI thread itself (runs inline, so
        // it can never deadlock against a Send issued on the same thread).
        public T Invoke<T>(Func<T> func)
        {
            if (func == null) throw new ArgumentNullException(nameof(func));
            if (IsUiThread) return func();
            var holder = new ResultHolder<T>();
            _context.Send(_ =>
            {
                try { holder.Value = func(); }
                catch (Exception ex) { holder.Error = ExceptionDispatchInfo.Capture(ex); }
            }, null);
            holder.Error?.Throw();
            return holder.Value;
        }

        public void Invoke(Action action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            Invoke(() => { action(); return 0; });
        }

        // Posts func to the Office UI thread without blocking it. The
        // returned task completes with func's result (or exception) once the
        // UI thread has run it.
        public Task<T> InvokeAsync<T>(Func<T> func)
        {
            if (func == null) throw new ArgumentNullException(nameof(func));
            if (IsUiThread)
            {
                try { return Task.FromResult(func()); }
                catch (Exception ex) { return Task.FromException<T>(ex); }
            }
            var tcs = new TaskCompletionSource<T>();
            _context.Post(_ =>
            {
                try { tcs.SetResult(func()); }
                catch (Exception ex) { tcs.SetException(ex); }
            }, null);
            return tcs.Task;
        }

        public Task InvokeAsync(Action action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            return InvokeAsync(() => { action(); return 0; });
        }

        // M2 diagnostics: thread id, apartment state and sync-context type at
        // each stage, written to the diagnostics log so a lost STA context
        // can be located (menu entry, around awaits, around each capture).
        // Records no document content, only thread facts.
        public void LogProbe(string stage)
        {
            var thread = Thread.CurrentThread;
            var context = SynchronizationContext.Current;
            ImageOverlayDiagnostics.LogThreadProbe(_host, stage,
                thread.ManagedThreadId,
                thread.GetApartmentState().ToString(),
                context == null ? "null" : (context.GetType().FullName ?? "unknown"),
                IsUiThread);
        }

        private sealed class ResultHolder<T>
        {
            public T Value = default(T);
            public ExceptionDispatchInfo? Error;
        }
    }
}
