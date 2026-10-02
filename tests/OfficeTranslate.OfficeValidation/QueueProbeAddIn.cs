using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using OfficeTranslate.Core;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeTranslate.OfficeValidation
{
    // NOT in the solution/MSI and never shipped to users. Register this
    // temporary COM helper only for a controlled Office-UI acceptance run.
    // Only used under an exported temporary HKCU review override, never shipped.
    [ComVisible(true), Guid("7898D80A-8CB7-4E18-9D52-3DC5901786D7")]
    [ProgId("OfficeTranslate.OfficeValidation"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class QueueProbeAddIn : Extensibility.IDTExtensibility2
    {
        private Word.Application _word;
        private OfficeUiDispatcher _ui;
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint process);
        public void OnConnection(object app, Extensibility.ext_ConnectMode mode, object instance, ref Array custom)
        {
            _word = (Word.Application)app;
            _ui = OfficeUiDispatcher.Capture("Word-acceptance");
            ((Microsoft.Office.Core.COMAddIn)instance).Object = this;
        }

        public void RunQueueProbe(string mode, string output)
        {
            // Automation calls can arrive on an RPC MTA. Use the dispatcher
            // captured by the real Office OnConnection UI callback.
            _ui.Invoke(() => RunQueueProbeOnUi(mode, output));
        }

        private void RunQueueProbeOnUi(string mode, string output)
        {
            if (mode != "expire" && mode != "supersede" && mode != "normal") throw new ArgumentException("mode");
            if (File.Exists(output)) throw new InvalidOperationException("Preserve existing evidence");
            var ui = _ui;
            uint process;
            var nativeUiThread = GetWindowThreadProcessId(new IntPtr(_word.ActiveWindow.Hwnd), out process);
            if (nativeUiThread != GetCurrentThreadId()) throw new InvalidOperationException("Not on Word's actual window thread");
            var thread = Thread.CurrentThread.ManagedThreadId;
            var doc = _word.Documents.Add();
            doc.Content.Text = "0123456789abcdef";
            doc.Range(0, 5).Select();
            var saved = _word.Selection.Range.Duplicate;
            var generation = 1;
            var calls = 0;
            var reason = "none";
            var posted = new ManualResetEventSlim();
            var done = Task.Run(async () =>
            {
                try
                {
                    var gate = UiRestoreGate.RunAsync(ui, () => Volatile.Read(ref generation) == 1,
                        () => { Interlocked.Increment(ref calls); saved.Select(); }, r => reason = r);
                    posted.Set();
                    await gate;
                    // Barrier enters the REAL Word UI queue after the late
                    // callback, proving it has been drained before assertion.
                    var actual = await ui.InvokeAsync(() =>
                        "start=" + _word.Selection.Start + " end=" + _word.Selection.End
                        + " story=" + (int)_word.Selection.StoryType
                        + " is_ui=" + ui.IsUiThread
                        + " ui_thread=" + Thread.CurrentThread.ManagedThreadId);
                    File.WriteAllText(output, "mode=" + mode + " calls=" + calls + " reason=" + reason
                        + " " + actual + " entry_thread=" + thread
                        + " native_ui_thread=" + nativeUiThread
                        + " core_mvid=" + typeof(UiRestoreGate).Module.ModuleVersionId);
                }
                catch (Exception ex) { File.WriteAllText(output, "FAIL " + ex); }
                finally { posted.Dispose(); }
            });
            if (!posted.Wait(5000)) throw new TimeoutException("Probe worker did not post");
            // Intentionally block the ACTUAL Office STA with no message pump.
            // Select the user's newer range before releasing the queued callback.
            if (mode == "expire") Thread.Sleep(2400);
            if (mode == "supersede") { Interlocked.Increment(ref generation); Thread.Sleep(200); }
            if (mode != "normal") doc.Range(7, 12).Select();
        }
        public void OnDisconnection(Extensibility.ext_DisconnectMode mode, ref Array custom) { _word = null; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }
    }
}
