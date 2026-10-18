using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using OfficeTranslate.Core;
using OfficeTranslate.WordAddIn;
#if WORD
namespace OfficeTranslate.WordAddIn
#elif EXCEL
namespace OfficeTranslate.ExcelAddIn
#else
namespace OfficeTranslate.PowerPointAddIn
#endif
{
    public sealed partial class AddIn : Microsoft.Office.Core.ICustomTaskPaneConsumer
    {
        private Microsoft.Office.Core.ICTPFactory? _chatFactory;
        private readonly Dictionary<int, Microsoft.Office.Core.CustomTaskPane> _chatPanes = new Dictionary<int, Microsoft.Office.Core.CustomTaskPane>();
        private System.Windows.Forms.Timer? _chatCleanup;
        public void CTPFactoryAvailable(Microsoft.Office.Core.ICTPFactory factory) { _chatFactory = factory; }
        public void OpenChat(Microsoft.Office.Core.IRibbonControl control)
        {
            ReloadPersistentSettings();
            var stage = "active-window";
            try { OpenChatCore(ref stage); }
            catch (Exception error)
            {
                System.Windows.Forms.MessageBox.Show(UiText.Get(_settings.UiLanguage, "ChatPaneError") + "\r\n\r\n"
                    + UiText.Get(_settings.UiLanguage, "ChatErrorDetails") + " (Ctrl+C):\r\n"
                    + ChatPaneDetails(error, _settings, stage), UiText.Get(_settings.UiLanguage, "Chat"));
            }
        }
        private void OpenChatCore(ref string stage)
        {
#if WORD
            var window = _word!.ActiveWindow; const string progId = "OfficeTranslate.WordChatPane";
#elif EXCEL
            var window = _host!.ActiveWindow; const string progId = "OfficeTranslate.ExcelChatPane";
#else
            var window = _host!.ActiveWindow; const string progId = "OfficeTranslate.PowerPointChatPane";
#endif
            stage = "window-handle";
            int id = GetChatWindowHandle(window);
            if (id == 0) throw new InvalidOperationException("The document window has no handle.");
            if (_chatPanes.TryGetValue(id, out var existing))
            {
                try
                {
                    object owner = existing.Window;
                    if (ReferenceEquals(owner, (object)window) || owner.Equals((object)window)) { existing.Visible = true; return; }
                }
                catch { }
                DisposeChat(existing); _chatPanes.Remove(id);
            }
            stage = "create-pane";
            if (_chatFactory == null) throw new InvalidOperationException("Office has not supplied the Chat task-pane factory.");
            // Office exposes Title as read-only. Keep its frame title as the product
            // name; the localized heading inside the pane updates with the UI language.
            var pane = _chatFactory.CreateCTP(progId, UiText.Get(_settings.UiLanguage, "ChatPaneTitle"), window);
            try
            {
                stage = "initialize-control";
                var content = (DocumentChatPane)pane.ContentControl;
                // TargetLanguage is session-only. Load persisted service settings and
                // overlay the current ribbon languages, just like document translation.
                content.Initialize(() => DocumentChatHost.Capture(window), LoadForTranslation);
                stage = "show-pane";
                pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight; pane.Width = 440;
                pane.Visible = true; _chatPanes.Add(id, pane);
            }
            catch { DisposeChat(pane); throw; }
            if (_chatCleanup == null)
            {
                _chatCleanup = new System.Windows.Forms.Timer { Interval = 3000 };
                _chatCleanup.Tick += (s, e) =>
                {
                    foreach (var key in new List<int>(_chatPanes.Keys)) if (!ChatWindowAlive(new IntPtr(key)))
                    { DisposeChat(_chatPanes[key]); _chatPanes.Remove(key); }
                };
                _chatCleanup.Start();
            }
        }
        // PowerPoint exposes DocumentWindow.HWND through its typed COM interface,
        // but IDispatch/dynamic invocation fails with MissingMemberException.
        // Keep this read early-bound in all hosts; do not substitute Application.HWND
        // (panes and their cleanup must belong to the originating document window).
        internal static int GetChatWindowHandle(object window)
        {
#if WORD
            return ((Microsoft.Office.Interop.Word.Window)window).Hwnd;
#elif EXCEL
            return ((Microsoft.Office.Interop.Excel.Window)window).Hwnd;
#else
            return ((Microsoft.Office.Interop.PowerPoint.DocumentWindow)window).HWND;
#endif
        }
        internal static string ChatPaneDetails(Exception error, TranslationSettings settings, string stage)
        {
            var details = UiText.Get(settings.UiLanguage, "ChatStage") + ": " + stage + "\r\n" + ChatErrors.Describe(error, settings);
            for (Exception? current = error; current != null; current = current.InnerException)
            {
                details += "\r\n" + current.GetType().Name + " HRESULT=0x" + current.HResult.ToString("X8");
                // Method names only: no file paths, arguments, locals, document text or prompts.
                foreach (var frame in (new StackTrace(current, false).GetFrames() ?? new StackFrame[0]).Take(6))
                {
                    var method = frame.GetMethod();
                    if (method != null) details += "\r\n  at " + method.DeclaringType?.FullName + "." + method.Name;
                }
            }
            return details;
        }
        [DllImport("user32.dll", EntryPoint = "IsWindow")] private static extern bool ChatWindowAlive(IntPtr hwnd);
        public string GetChatTip(Microsoft.Office.Core.IRibbonControl control) => UiText.Get(_settings.UiLanguage, "ChatRibbonTip");
        private void RefreshChatLanguage()
        {
            foreach (var pane in _chatPanes.Values)
                try { if (pane.ContentControl is DocumentChatPane content) content.RefreshLanguage(); } catch { }
        }
        private static void DisposeChat(Microsoft.Office.Core.CustomTaskPane pane)
        {
            try { if (pane.ContentControl is DocumentChatPane content) { content.Stop(); content.Dispose(); } } catch { }
            try { pane.Delete(); } catch { }
        }
        private void CloseChat()
        { _chatCleanup?.Stop(); _chatCleanup?.Dispose(); _chatCleanup = null; foreach (var pane in _chatPanes.Values) DisposeChat(pane); _chatPanes.Clear(); _chatFactory = null; }
    }
}
