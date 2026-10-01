using Microsoft.Office.Interop.Word;
using OfficeTranslate.Core;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;
using WordApplication = Microsoft.Office.Interop.Word.Application;

namespace OfficeTranslate.WordAddIn
{
    [ComVisible(true)]
    [Guid("7898D80A-8CB7-4E18-9D52-3DC5901786D7")]
    [ProgId("OfficeTranslate.WordAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class AddIn : Extensibility.IDTExtensibility2, Microsoft.Office.Core.IRibbonExtensibility
    {
        private WordApplication? _word;
        private Microsoft.Office.Core.IRibbonUI? _ribbon;
        private CancellationTokenSource? _cancellation;
        private TranslationProgressForm? _progressForm;
        private readonly SettingsStore _settingsStore = new SettingsStore();
        private TranslationSettings _settings = new TranslationSettings();

        public string GetCustomUI(string ribbonId)
        {
            var name = typeof(AddIn).Namespace + ".Ribbon.xml";
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            using (var reader = new StreamReader(stream ?? throw new InvalidOperationException("找不到 Ribbon.xml"))) return reader.ReadToEnd();
        }

        public void OnRibbonLoad(Microsoft.Office.Core.IRibbonUI ribbon) { _ribbon = ribbon; }
        public object GetRibbonImage(Microsoft.Office.Core.IRibbonControl control) => RibbonImageFactory.Get(control.Id);
        public object GetLanguageItemImage(Microsoft.Office.Core.IRibbonControl control) => RibbonImageFactory.GetLanguage(control.Tag);
        public string GetLanguageItemLabel(Microsoft.Office.Core.IRibbonControl control) => UiText.Language(_settings.UiLanguage, control.Tag);
        public string GetRibbonLabel(Microsoft.Office.Core.IRibbonControl control)
        {
            var id = control.Id;
            var key = id.EndsWith("LanguageGroup") ? "LanguageGroup" : id.EndsWith("TranslateGroup") ? "TranslateGroup" :
                id.EndsWith("ToolsGroup") ? "ToolsGroup" : id.EndsWith("SwapLanguages") ? "Swap" :
                id.EndsWith("Selection") ? "Selection" : id.EndsWith("Document") ? "Document" :
                id.EndsWith("Bilingual") ? "Bilingual" : id.EndsWith("ImageOcr") ? "ImageOcr" :
                id.EndsWith("Cancel") ? "Cancel" : "Settings";
            var value = UiText.Get(_settings.UiLanguage, key);
            return _settings.UiLanguage == "en" ? value : string.Join("\u2060", value.ToCharArray());
        }
        public bool GetBilingualMode(Microsoft.Office.Core.IRibbonControl control) => _settings.BilingualMode;
        public bool GetImageOcrMode(Microsoft.Office.Core.IRibbonControl control) => _settings.ImageOcrEnabled;
        public bool GetImageOcrVisible(Microsoft.Office.Core.IRibbonControl control) => true;
        public void ImageOcrModeChanged(Microsoft.Office.Core.IRibbonControl control, bool pressed) { _settings.ImageOcrEnabled = pressed; }
        public string GetSourceLanguageMenuLabel(Microsoft.Office.Core.IRibbonControl control) => UiText.Get(_settings.UiLanguage, "Source");
        public string GetTargetLanguageMenuLabel(Microsoft.Office.Core.IRibbonControl control) => UiText.Get(_settings.UiLanguage, "Target");
        public string GetSourceLanguageTip(Microsoft.Office.Core.IRibbonControl control) => string.Format(UiText.Get(_settings.UiLanguage, "SourceTip"), UiText.Language(_settings.UiLanguage, _settings.SourceLanguage));
        public string GetTargetLanguageTip(Microsoft.Office.Core.IRibbonControl control) => string.Format(UiText.Get(_settings.UiLanguage, "TargetTip"), UiText.Language(_settings.UiLanguage, _settings.TargetLanguage));
        public object GetSourceLanguageImage(Microsoft.Office.Core.IRibbonControl control) => RibbonImageFactory.GetLanguage(_settings.SourceLanguage);
        public object GetTargetLanguageImage(Microsoft.Office.Core.IRibbonControl control) => RibbonImageFactory.GetLanguage(_settings.TargetLanguage);
        public void SelectSourceLanguage(Microsoft.Office.Core.IRibbonControl control)
        {
            _settings.SourceLanguage = control.Tag;
            _ribbon?.InvalidateControl("OfficeTranslate.SourceLanguageMenu");
        }
        public void SelectTargetLanguage(Microsoft.Office.Core.IRibbonControl control)
        {
            _settings.TargetLanguage = control.Tag;
            _ribbon?.InvalidateControl("OfficeTranslate.TargetLanguageMenu");
        }
        public void SwapLanguages(Microsoft.Office.Core.IRibbonControl control)
        {
            var oldTarget = _settings.TargetLanguage;
            if (_settings.SourceLanguage == "自动检测")
            {
                _settings.SourceLanguage = oldTarget;
                _settings.TargetLanguage = oldTarget == "简体中文" ? "英语" : "简体中文";
            }
            else
            {
                _settings.TargetLanguage = _settings.SourceLanguage;
                _settings.SourceLanguage = oldTarget;
            }
            _ribbon?.InvalidateControl("OfficeTranslate.SourceLanguageMenu");
            _ribbon?.InvalidateControl("OfficeTranslate.TargetLanguageMenu");
        }

        public void BilingualModeChanged(Microsoft.Office.Core.IRibbonControl control, bool pressed)
        {
            _settings.BilingualMode = pressed;
        }
        public async void TranslateSelection(Microsoft.Office.Core.IRibbonControl control) => await RunAsync(false);
        public async void TranslateDocument(Microsoft.Office.Core.IRibbonControl control) => await RunAsync(true);
        public void OpenSettings(Microsoft.Office.Core.IRibbonControl control)
        {
            using (var form = new SettingsForm(_settingsStore))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    ReloadPersistentSettings();
                    _ribbon?.Invalidate();
                }
            }
        }

        private async Task RunAsync(bool wholeDocument)
        {
            if (_word == null || _cancellation != null) return;
            _cancellation = new CancellationTokenSource();
            // M2: capture the Office UI (STA) thread at the ribbon entry
            // point, before the first await. All COM, clipboard, and
            // writeback operations are dispatched back to this thread by
            // the translation service.
            var ui = OfficeUiDispatcher.Capture("Word");
            // O1: capture the task's own document window BEFORE the first
            // await. Process.MainWindowHandle is not necessarily the window
            // the ribbon was invoked from when several document windows are
            // open; progress/result/error windows must all be owned by the
            // task window so dismissing them returns activation to the
            // right document instead of switching to another one.
            var taskWindow = GetTaskWindow();
            // S1: hoisted so the finally can run the final selection
            // restore after the UI teardown.
            WordTranslationService? service = null;
            try
            {
                var settings = LoadForTranslation();
                settings.Validate();
                _progressForm = new TranslationProgressForm(settings.UiLanguage, () => _cancellation?.Cancel());
                _progressForm.ShowFor(taskWindow);
                var bilingual = settings.BilingualMode;
                service = new WordTranslationService(_word);
                _progressForm.SetStatus("OfficeTranslate：正在读取文档…");
                var summary = await service.TranslateAsync(wholeDocument, bilingual, settings, _cancellation.Token, _progressForm.SetStatus, ui);
                _progressForm.CloseAndShowResult(summary, taskWindow);
                _progressForm = null;
            }
            catch (OperationCanceledException) { }
            // N1: error presentation is UI work and must enter the same UI
            // dispatch boundary; the ambient SynchronizationContext cannot
            // be trusted to bring the post-await continuation back to STA.
            // O1: close the progress window BEFORE showing the error, and
            // give the error dialog the task window as its owner. Showing
            // an ownerless MessageBox while the progress form is still open
            // lets Windows reactivate a different document window on dismiss.
            catch (Exception ex)
            {
                // S1: boundary probe BEFORE the progress close. The service
                // finally already logged selection_restore_diag (matched=1);
                // this shows whether the selection is still on the target
                // when the AddIn teardown starts.
                WordSelectionProbe.Log(_word, "before_progress_close", taskWindow);
                // P1 (afa3812 review): the progress form was owned by the
                // task window; if that window died mid-task, Close must not
                // throw and mask the real error below.
                try { _progressForm?.CloseSafely(); } catch { }
                _progressForm = null;
                WordSelectionProbe.Log(_word, "after_progress_close", taskWindow);
                ui.Invoke(() =>
                {
                    // P1 (c77bf85 review): the cached task window may be gone
                    // while the task waited on the network (user closed the
                    // source document) -- and worse, its HWND value can be
                    // recycled by a new window so IsWindow alone still
                    // passes. The review proved that owning the error dialog
                    // with such a handle reproduces the stuck blank modal
                    // blocking Word's exit (error_dialog_shown without
                    // error_dialog_dismissed / after_final_restore).
                    // ResolveErrorDialogOwner therefore tests membership in
                    // the current Word window set: cached window while it is
                    // still one of this instance's windows, else the live
                    // ActiveWindow, else ownerless.
                    var window = ResolveErrorDialogOwner(taskWindow);
                    // S1: the old bare markers could not answer what the
                    // selection was when the dialog appeared or was
                    // dismissed; the probe snapshots it at both points.
                    WordSelectionProbe.Log(_word, "error_dialog_shown", window);
                    try
                    {
                        if (window != IntPtr.Zero) MessageBox.Show(new TranslationProgressForm.WindowHandle(window), ex.Message, "OfficeTranslate", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        else MessageBox.Show(ex.Message, "OfficeTranslate", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        // P1: the dismiss probe is written once Show returns
                        // or throws, so the next review can tell "dialog
                        // hung" (dismissed missing) apart from "dialog
                        // teardown threw" (dismissed without
                        // after_final_restore). If Show itself hangs, this
                        // finally never runs either -- the only real fix is
                        // never giving it a bad owner in the first place
                        // (see ResolveErrorDialogOwner above).
                        WordSelectionProbe.Log(_word, "error_dialog_dismissed", window);
                    }
                });
            }
            finally
            {
                _progressForm?.CloseSafely(); _progressForm = null;
                // S1: final restore AFTER the host UI teardown. The
                // boundary probes showed closing the owned progress window
                // reactivates Word and moves Application.Selection back to
                // the pre-restore state, so the old service-finally restore
                // did not stick. Same one-shot ticket, expiry and
                // generation rules (N1); best-effort, never throws.
                if (service != null) await service.RestoreTaskSelectionAsync(ui);
                WordSelectionProbe.Log(_word, "after_final_restore", taskWindow);
                // S1: start the read-only post-restore drift watch (never
                // awaited: the task is over, the watch only observes).
                StartPostRestoreWatch(ui, _word, taskWindow);
                _cancellation.Dispose(); _cancellation = null;
            }
        }

        // S1: read-only drift watch after the final restore. The error
        // path showed the selection moving back to the failed image about
        // 7.6s after after_final_restore while success/cancel stay stable,
        // so this samples (never restores) at +2/+5/+10s to bracket the
        // drift in time. Fire-and-forget: never throws, stops itself after
        // the last sample, holds no ticket and never touches the selection
        // (N1 untouched). By construction of the finally above, the
        // progress form is already null and any error dialog already
        // dismissed when the watch runs.
        private static async void StartPostRestoreWatch(OfficeUiDispatcher ui, WordApplication? word, IntPtr taskWindow)
        {
            try
            {
                foreach (var delayMs in new[] { 2000, 3000, 5000 })
                {
                    await Task.Delay(delayMs).ConfigureAwait(false);
                    var w = word;
                    var hwnd = taskWindow;
                    try { await ui.InvokeAsync(() => WordSelectionProbe.Log(w, "post_restore_watch", hwnd)).ConfigureAwait(false); }
                    catch { }
                }
            }
            catch { }
        }

        // O1: the document window this task was invoked from. Read on the UI
        // thread at ribbon entry, before any await; IntPtr.Zero when there
        // is no usable window (falls back to the previous unowned behavior).
        private IntPtr GetTaskWindow()
        {
            try
            {
                var window = _word?.ActiveWindow;
                if (window == null) return IntPtr.Zero;
                return new IntPtr(window.Hwnd);
            }
            catch { return IntPtr.Zero; }
        }

        // P1 (c77bf85 review): re-validate the cached task window right
        // before it is used as the error dialog's owner. Must be called on
        // the UI thread (reads _word.Windows / ActiveWindow via COM).
        // Never throws; the resolution is logged metadata-only so the next
        // review can see whether the cached HWND was still one of this
        // instance's windows and which window was chosen.
        //
        // IsWindow alone is NOT sufficient: the review proved the cached
        // HWND of the closed document A can still pass IsWindow (the handle
        // value gets recycled by a new window) while no longer belonging to
        // this Word instance's window set -- owning the error dialog with it
        // reproduced the stuck blank modal that blocked Word's exit (no
        // error_dialog_dismissed / after_final_restore). Membership in the
        // CURRENT _word.Windows collection is the validity test.
        private IntPtr ResolveErrorDialogOwner(IntPtr cachedTaskWindow)
        {
            var wordHwnds = new System.Collections.Generic.HashSet<long>();
            int windowCount = -1;
            try
            {
                var windows = _word?.Windows;
                if (windows != null)
                {
                    windowCount = windows.Count;
                    foreach (Microsoft.Office.Interop.Word.Window w in windows)
                    {
                        try { wordHwnds.Add(new IntPtr(w.Hwnd).ToInt64()); }
                        catch { }
                    }
                }
            }
            catch { }
            bool cachedValid = cachedTaskWindow != IntPtr.Zero && wordHwnds.Contains(cachedTaskWindow.ToInt64());
            IntPtr activeHwnd = IntPtr.Zero;
            try
            {
                var active = _word?.ActiveWindow;
                if (active != null) activeHwnd = new IntPtr(active.Hwnd);
                // Fresh from the live instance, but still require a live
                // handle: never own a dialog with a dead HWND.
                if (activeHwnd != IntPtr.Zero && !IsWindow(activeHwnd)) activeHwnd = IntPtr.Zero;
            }
            catch { activeHwnd = IntPtr.Zero; }
            var chosen = cachedValid ? cachedTaskWindow : activeHwnd;
            ImageOverlayDiagnostics.LogCaptureFailure("Word",
                "error_dialog_owner stage=resolve"
                + " cachedHwnd=" + cachedTaskWindow.ToInt64().ToString("X")
                + " cachedInWordWindows=" + (cachedValid ? "1" : "0")
                + " activeHwnd=" + activeHwnd.ToInt64().ToString("X")
                + " wordWindows=" + windowCount
                + " chosenHwnd=" + chosen.ToInt64().ToString("X"));
            return chosen;
        }

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            _word = (WordApplication)application;
            _settings = _settingsStore.Load();
        }
        private TranslationSettings LoadForTranslation()
        {
            var settings = _settingsStore.Load();
            settings.SourceLanguage = _settings.SourceLanguage;
            settings.TargetLanguage = _settings.TargetLanguage;
            settings.BilingualMode = _settings.BilingualMode;
            settings.ImageOcrEnabled = _settings.ImageOcrEnabled;
            return settings;
        }
        private void ReloadPersistentSettings()
        {
            var source = _settings.SourceLanguage; var target = _settings.TargetLanguage; var bilingual = _settings.BilingualMode; var imageOcr = _settings.ImageOcrEnabled;
            _settings = _settingsStore.Load();
            _settings.SourceLanguage = source; _settings.TargetLanguage = target; _settings.BilingualMode = bilingual;
            _settings.ImageOcrEnabled = imageOcr;
        }
        public void OnDisconnection(Extensibility.ext_DisconnectMode removeMode, ref Array custom) { _cancellation?.Cancel(); _progressForm?.CloseSafely(); _word = null; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { _cancellation?.Cancel(); _progressForm?.CloseSafely(); }
    }
}
