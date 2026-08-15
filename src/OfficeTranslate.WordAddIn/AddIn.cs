using Microsoft.Office.Interop.Word;
using OfficeTranslate.Core;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Linq;
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
        private readonly SettingsStore _settingsStore = new SettingsStore();
        private TranslationSettings _settings = new TranslationSettings();
        private static readonly string[] SourceLanguages = { "自动检测", "简体中文", "繁體中文", "英语", "日语", "韩语", "法语", "德语", "西班牙语", "俄语", "葡萄牙语", "意大利语", "阿拉伯语" };
        private static readonly string[] TargetLanguages = SourceLanguages.Skip(1).ToArray();

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
                id.EndsWith("Bilingual") ? "Bilingual" : id.EndsWith("Cancel") ? "Cancel" : "Settings";
            var value = UiText.Get(_settings.UiLanguage, key);
            return _settings.UiLanguage == "en" ? value : string.Join("\u2060", value.ToCharArray());
        }
        public int GetSourceLanguageCount(Microsoft.Office.Core.IRibbonControl control) => SourceLanguages.Length;
        public string GetSourceLanguageLabel(Microsoft.Office.Core.IRibbonControl control, int index) => SourceLanguages[index];
        public int GetSourceLanguageIndex(Microsoft.Office.Core.IRibbonControl control) => Math.Max(0, Array.IndexOf(SourceLanguages, _settings.SourceLanguage));
        public int GetTargetLanguageCount(Microsoft.Office.Core.IRibbonControl control) => TargetLanguages.Length;
        public string GetTargetLanguageLabel(Microsoft.Office.Core.IRibbonControl control, int index) => TargetLanguages[index];
        public int GetTargetLanguageIndex(Microsoft.Office.Core.IRibbonControl control) => Math.Max(0, Array.IndexOf(TargetLanguages, _settings.TargetLanguage));
        public bool GetBilingualMode(Microsoft.Office.Core.IRibbonControl control) => _settings.BilingualMode;
        public string GetSourceLanguageMenuLabel(Microsoft.Office.Core.IRibbonControl control) => UiText.Get(_settings.UiLanguage, "Source");
        public string GetTargetLanguageMenuLabel(Microsoft.Office.Core.IRibbonControl control) => UiText.Get(_settings.UiLanguage, "Target");
        public string GetSourceLanguageTip(Microsoft.Office.Core.IRibbonControl control) => string.Format(UiText.Get(_settings.UiLanguage, "SourceTip"), UiText.Language(_settings.UiLanguage, _settings.SourceLanguage));
        public string GetTargetLanguageTip(Microsoft.Office.Core.IRibbonControl control) => string.Format(UiText.Get(_settings.UiLanguage, "TargetTip"), UiText.Language(_settings.UiLanguage, _settings.TargetLanguage));
        public object GetSourceLanguageImage(Microsoft.Office.Core.IRibbonControl control) => RibbonImageFactory.GetLanguage(_settings.SourceLanguage);
        public object GetTargetLanguageImage(Microsoft.Office.Core.IRibbonControl control) => RibbonImageFactory.GetLanguage(_settings.TargetLanguage);
        public void SelectSourceLanguage(Microsoft.Office.Core.IRibbonControl control)
        {
            _settings.SourceLanguage = control.Tag;
            SaveLanguageSettings();
            _ribbon?.InvalidateControl("OfficeTranslate.SourceLanguageMenu");
        }
        public void SelectTargetLanguage(Microsoft.Office.Core.IRibbonControl control)
        {
            _settings.TargetLanguage = control.Tag;
            SaveLanguageSettings();
            _ribbon?.InvalidateControl("OfficeTranslate.TargetLanguageMenu");
        }
        public void SourceLanguageChanged(Microsoft.Office.Core.IRibbonControl control, string selectedId, int selectedIndex)
        {
            _settings.SourceLanguage = SourceLanguages[selectedIndex]; SaveLanguageSettings();
        }
        public void TargetLanguageChanged(Microsoft.Office.Core.IRibbonControl control, string selectedId, int selectedIndex)
        {
            _settings.TargetLanguage = TargetLanguages[selectedIndex]; SaveLanguageSettings();
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
            SaveLanguageSettings();
            _ribbon?.InvalidateControl("OfficeTranslate.SourceLanguageMenu");
            _ribbon?.InvalidateControl("OfficeTranslate.TargetLanguageMenu");
        }

        public void BilingualModeChanged(Microsoft.Office.Core.IRibbonControl control, bool pressed)
        {
            _settings.BilingualMode = pressed;
            SaveLanguageSettings();
        }

        private void SaveLanguageSettings()
        {
            try { _settingsStore.Save(_settings); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "OfficeTranslate", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
        public async void TranslateSelection(Microsoft.Office.Core.IRibbonControl control) => await RunAsync(false);
        public async void TranslateDocument(Microsoft.Office.Core.IRibbonControl control) => await RunAsync(true);
        public void CancelTranslation(Microsoft.Office.Core.IRibbonControl control) => _cancellation?.Cancel();
        public void OpenSettings(Microsoft.Office.Core.IRibbonControl control)
        {
            using (var form = new SettingsForm(_settingsStore))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    _settings = _settingsStore.Load();
                    _ribbon?.Invalidate();
                }
            }
        }

        private async Task RunAsync(bool wholeDocument)
        {
            if (_word == null || _cancellation != null) return;
            _cancellation = new CancellationTokenSource();
            try
            {
                var settings = _settingsStore.Load();
                settings.Validate();
                var bilingual = settings.BilingualMode;
                var service = new WordTranslationService(_word);
                _word.StatusBar = "OfficeTranslate：正在读取文档…";
                await service.TranslateAsync(wholeDocument, bilingual, settings, _cancellation.Token, status => _word.StatusBar = status);
                _word.StatusBar = bilingual ? "OfficeTranslate：双语排版完成" : "OfficeTranslate：翻译完成";
            }
            catch (OperationCanceledException) { if (_word != null) _word.StatusBar = "OfficeTranslate：已取消"; }
            catch (Exception ex) { MessageBox.Show(ex.Message, "OfficeTranslate", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { _cancellation.Dispose(); _cancellation = null; }
        }

        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            _word = (WordApplication)application;
            _settings = _settingsStore.Load();
        }
        public void OnDisconnection(Extensibility.ext_DisconnectMode removeMode, ref Array custom) { _cancellation?.Cancel(); _word = null; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) => _cancellation?.Cancel();
    }
}
