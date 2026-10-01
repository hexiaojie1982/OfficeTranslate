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
    // before OLE calls can be made". Every operation that touches COM or the
    // clipboard goes through this dispatcher explicitly instead of relying
    // on await context capture. Network I/O stays off the UI thread; only
    // the COM/clipboard/writeback sections are dispatched.
    //
    // N1: the ambient SynchronizationContext is NEVER trusted. The real
    // Word ribbon entry carries the base System.Threading.
    // SynchronizationContext, whose Send/Post run inline on the CALLING
    // thread -- keeping it made cross-thread Invoke/InvokeAsync silently
    // run on MTA pool threads (measured on the release DLL). The dispatcher
    // therefore owns a private WindowsFormsSynchronizationContext created
    // on the captured STA thread, which pins that thread and its message
    // pump as the only dispatch target. The ambient context is left
    // untouched (never read, never replaced).
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
        // ribbon button handler, before the first await. Fails fast when the
        // calling thread is not STA: silently dispatching to a thread-pool
        // thread is worse than an explicit error.
        public static OfficeUiDispatcher Capture(string host)
        {
            var thread = Thread.CurrentThread;
            if (thread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException(
                    "OfficeTranslate 必须在 Office UI (STA) 线程上启动；当前线程不是 STA，无法安全调度 COM 操作。");
            // Created on this thread: per the .NET reference source the
            // constructor captures Thread.CurrentThread as the destination
            // thread together with the thread's WinForms marshaling
            // control, so Post/Send pump back here. Private to this
            // dispatcher; the ambient SynchronizationContext is ignored.
            var context = new WindowsFormsSynchronizationContext();
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

        // M2 diagnostics: thread id, apartment state and dispatch-target
        // context type at each stage, written to the diagnostics log so a
        // lost STA context can be located (menu entry, around awaits, around
        // each capture). Records no document content, only thread facts.
        // N1: reports the dispatcher's own context (the actual Post/Send
        // target), not the ambient one, which this dispatcher ignores.
        public void LogProbe(string stage)
        {
            var thread = Thread.CurrentThread;
            ImageOverlayDiagnostics.LogThreadProbe(_host, stage,
                thread.ManagedThreadId,
                thread.GetApartmentState().ToString(),
                _context.GetType().FullName ?? "unknown",
                IsUiThread);
        }

        private sealed class ResultHolder<T>
        {
            public T Value = default(T);
            public ExceptionDispatchInfo? Error;
        }
    }
}
