using OfficeTranslate.Core;
using OfficeTranslate.WordAddIn;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
#if EXCEL
using HostApplication = Microsoft.Office.Interop.Excel.Application;
using HostService = OfficeTranslate.ExcelAddIn.ExcelTranslationService;
namespace OfficeTranslate.ExcelAddIn
#else
using HostApplication = Microsoft.Office.Interop.PowerPoint.Application;
using HostService = OfficeTranslate.PowerPointAddIn.PowerPointTranslationService;
namespace OfficeTranslate.PowerPointAddIn
#endif
{
    [ComVisible(true)]
#if EXCEL
    [Guid("9F8BE921-0DAD-45A4-9B64-81AA2F8DE101")]
    [ProgId("OfficeTranslate.ExcelAddIn")]
#else
    [Guid("A56AA843-3192-4EE2-88AD-205E1B2D1102")]
    [ProgId("OfficeTranslate.PowerPointAddIn")]
#endif
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class AddIn : Extensibility.IDTExtensibility2, Microsoft.Office.Core.IRibbonExtensibility
    {
        private HostApplication? _host;
        private Microsoft.Office.Core.IRibbonUI? _ribbon;
        private CancellationTokenSource? _cancellation;
        private readonly SettingsStore _store = new SettingsStore();
        private TranslationSettings _settings = new TranslationSettings();
        private TranslationProgressForm? _progressForm;

        public string GetCustomUI(string ribbonId)
        {
            var name = typeof(AddIn).Namespace + ".Ribbon.xml";
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            using (var reader = new StreamReader(stream ?? throw new InvalidOperationException("Ribbon.xml not found"))) return reader.ReadToEnd();
        }
        public void OnRibbonLoad(Microsoft.Office.Core.IRibbonUI ribbon) => _ribbon = ribbon;
        public object GetRibbonImage(Microsoft.Office.Core.IRibbonControl c) => RibbonImageFactory.Get(c.Id);
        public object GetLanguageItemImage(Microsoft.Office.Core.IRibbonControl c) => RibbonImageFactory.GetLanguage(c.Tag);
        public object GetSourceLanguageImage(Microsoft.Office.Core.IRibbonControl c) => RibbonImageFactory.GetLanguage(_settings.SourceLanguage);
        public object GetTargetLanguageImage(Microsoft.Office.Core.IRibbonControl c) => RibbonImageFactory.GetLanguage(_settings.TargetLanguage);
        public string GetLanguageItemLabel(Microsoft.Office.Core.IRibbonControl c) => UiText.Language(_settings.UiLanguage, c.Tag);
        public string GetSourceLanguageMenuLabel(Microsoft.Office.Core.IRibbonControl c) => UiText.Get(_settings.UiLanguage, "Source");
        public string GetTargetLanguageMenuLabel(Microsoft.Office.Core.IRibbonControl c) => UiText.Get(_settings.UiLanguage, "Target");
        public string GetSourceLanguageTip(Microsoft.Office.Core.IRibbonControl c) => string.Format(UiText.Get(_settings.UiLanguage, "SourceTip"), UiText.Language(_settings.UiLanguage, _settings.SourceLanguage));
        public string GetTargetLanguageTip(Microsoft.Office.Core.IRibbonControl c) => string.Format(UiText.Get(_settings.UiLanguage, "TargetTip"), UiText.Language(_settings.UiLanguage, _settings.TargetLanguage));
        public bool GetBilingualMode(Microsoft.Office.Core.IRibbonControl c) => _settings.BilingualMode;
        public bool GetImageOcrMode(Microsoft.Office.Core.IRibbonControl c) => _settings.ImageOcrEnabled;
        public bool GetImageOcrVisible(Microsoft.Office.Core.IRibbonControl c) => true;
        public string GetRibbonLabel(Microsoft.Office.Core.IRibbonControl c)
        {
            var id = c.Id; var key = id.EndsWith("LanguageGroup") ? "LanguageGroup" : id.EndsWith("TranslateGroup") ? "TranslateGroup" : id.EndsWith("ToolsGroup") ? "ToolsGroup" : id.EndsWith("SwapLanguages") ? "Swap" : id.EndsWith("Selection") ? "Selection" : id.EndsWith("Document") ? "Document" : id.EndsWith("Bilingual") ? "Bilingual" : id.EndsWith("ImageOcr") ? "ImageOcr" : id.EndsWith("Cancel") ? "Cancel" : "Settings";
            var value = UiText.Get(_settings.UiLanguage, key); return _settings.UiLanguage == "en" ? value : string.Join("\u2060", value.ToCharArray());
        }
        public void SelectSourceLanguage(Microsoft.Office.Core.IRibbonControl c) { _settings.SourceLanguage = c.Tag; _ribbon?.InvalidateControl("OfficeTranslate.SourceLanguageMenu"); }
        public void SelectTargetLanguage(Microsoft.Office.Core.IRibbonControl c) { _settings.TargetLanguage = c.Tag; _ribbon?.InvalidateControl("OfficeTranslate.TargetLanguageMenu"); }
        public void SwapLanguages(Microsoft.Office.Core.IRibbonControl c)
        {
            var target = _settings.TargetLanguage;
            if (_settings.SourceLanguage == "自动检测") { _settings.SourceLanguage = target; _settings.TargetLanguage = target == "简体中文" ? "英语" : "简体中文"; }
            else { _settings.TargetLanguage = _settings.SourceLanguage; _settings.SourceLanguage = target; }
            _ribbon?.Invalidate();
        }
        public void BilingualModeChanged(Microsoft.Office.Core.IRibbonControl c, bool pressed) { _settings.BilingualMode = pressed; }
        public void ImageOcrModeChanged(Microsoft.Office.Core.IRibbonControl c, bool pressed) { _settings.ImageOcrEnabled = pressed; }
        public async void TranslateSelection(Microsoft.Office.Core.IRibbonControl c) => await RunAsync(false);
        public async void TranslateDocument(Microsoft.Office.Core.IRibbonControl c) => await RunAsync(true);
        public void OpenSettings(Microsoft.Office.Core.IRibbonControl c) { using (var form = new SettingsForm(_store)) if (form.ShowDialog() == DialogResult.OK) { ReloadPersistentSettings(); _ribbon?.Invalidate(); } }
        private async Task RunAsync(bool whole)
        {
            if (_host == null || _cancellation != null) return; _cancellation = new CancellationTokenSource();
            // M2: capture the Office UI (STA) thread at the ribbon entry
            // point, before the first await. All COM, clipboard, and
            // writeback operations are dispatched back to this thread by
            // the translation service.
            var ui = OfficeUiDispatcher.Capture(
#if EXCEL
                "Excel"
#else
                "PowerPoint"
#endif
                );
            // The host window handle is a COM property; read it here on the
            // UI thread instead of after the network awaits.
            var hostWindow = GetHostWindow();
            try { var settings = LoadForTranslation(); settings.Validate(); _progressForm = new TranslationProgressForm(settings.UiLanguage, () => _cancellation?.Cancel()); _progressForm.ShowFor(hostWindow); var service = new HostService(_host); SetStatus("OfficeTranslate：正在准备翻译…"); var summary = await service.TranslateAsync(whole, settings, _cancellation.Token, SetStatus, ui); _progressForm.CloseAndShowResult(summary, hostWindow); _progressForm = null; }
            catch (OperationCanceledException) { SetStatus("OfficeTranslate：已取消"); }
            catch (Exception ex)
            {
                CloseProgress();
                // P1 (afa3812 review, Word parity): the entry-cached host
                // window is re-validated before owning the error dialog; a
                // dead owner HWND left a stuck modal blocking the host in
                // the Word review. Resolved on the UI thread: cached while
                // alive, else a fresh read of the host window, else
                // ownerless. Never throws.
                // N1: error presentation must run inside the UI dispatch
                // boundary, not on whatever thread the await resumed on.
                ui.Invoke(() =>
                {
                    var window = ResolveErrorDialogOwner(hostWindow);
                    if (window != IntPtr.Zero) MessageBox.Show(new HostWindow(window), ex.Message, "OfficeTranslate", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else MessageBox.Show(ex.Message, "OfficeTranslate", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
            finally { CloseProgress(); _cancellation.Dispose(); _cancellation = null; }
        }
        // P1 (afa3812 review, Word parity): re-validate the cached host
        // window before it owns the error dialog. Called on the UI thread
        // (GetHostWindow reads a COM property). Never throws: cached while
        // alive, else a fresh read of the host window, else ownerless.
        private IntPtr ResolveErrorDialogOwner(IntPtr cachedHostWindow)
        {
            try { if (cachedHostWindow != IntPtr.Zero && IsWindow(cachedHostWindow)) return cachedHostWindow; } catch { }
            try
            {
                var fresh = GetHostWindow();
                if (fresh != IntPtr.Zero && IsWindow(fresh)) return fresh;
            }
            catch { }
            return IntPtr.Zero;
        }
        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);
        private IntPtr GetHostWindow() {
#if EXCEL
            return _host == null ? IntPtr.Zero : new IntPtr(_host.Hwnd);
#else
            return _host == null ? IntPtr.Zero : new IntPtr(_host.HWND);
#endif
        }
        private void SetStatus(string text) => _progressForm?.SetStatus(text);
        private TranslationSettings LoadForTranslation()
        {
            var settings = _store.Load();
            settings.SourceLanguage = _settings.SourceLanguage;
            settings.TargetLanguage = _settings.TargetLanguage;
            settings.BilingualMode = _settings.BilingualMode;
            settings.ImageOcrEnabled = _settings.ImageOcrEnabled;
            return settings;
        }
        private void ReloadPersistentSettings()
        {
            var source = _settings.SourceLanguage; var target = _settings.TargetLanguage; var bilingual = _settings.BilingualMode; var imageOcr = _settings.ImageOcrEnabled;
            _settings = _store.Load();
            _settings.SourceLanguage = source; _settings.TargetLanguage = target; _settings.BilingualMode = bilingual;
            _settings.ImageOcrEnabled = imageOcr;
        }
        private void CloseProgress()
        {
            _progressForm?.CloseSafely(); _progressForm = null;
        }
        private sealed class HostWindow : IWin32Window
        {
            public HostWindow(IntPtr handle) { Handle = handle; }
            public IntPtr Handle { get; }
        }
        public void OnConnection(object application, Extensibility.ext_ConnectMode mode, object instance, ref Array custom) { System.Windows.Forms.Application.EnableVisualStyles(); _host = (HostApplication)application; _settings = _store.Load(); }
        public void OnDisconnection(Extensibility.ext_DisconnectMode mode, ref Array custom) { _cancellation?.Cancel(); CloseProgress(); _host = null; }
        public void OnAddInsUpdate(ref Array custom) { } public void OnStartupComplete(ref Array custom) { } public void OnBeginShutdown(ref Array custom) { _cancellation?.Cancel(); CloseProgress(); }

    }
}
