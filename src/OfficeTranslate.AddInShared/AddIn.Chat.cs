using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
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
            try
            {
#if WORD
                dynamic window = _word!.ActiveWindow; const string progId = "OfficeTranslate.WordChatPane";
#elif EXCEL
                dynamic window = _host!.ActiveWindow; const string progId = "OfficeTranslate.ExcelChatPane";
#else
                dynamic window = _host!.ActiveWindow; const string progId = "OfficeTranslate.PowerPointChatPane";
#endif
                int id = window.HWND;
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
                if (_chatFactory == null) throw new InvalidOperationException();
                var pane = _chatFactory.CreateCTP(progId, UiText.Get(_settings.UiLanguage, "Chat"), window);
                try
                {
                    var content = (DocumentChatPane)pane.ContentControl;
#if WORD
                    content.Initialize(() => DocumentChatHost.Capture(window), () => _settingsStore.Load());
#else
                    content.Initialize(() => DocumentChatHost.Capture(window), () => _store.Load());
#endif
                    pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight; pane.Width = 440;
                    _chatPanes.Add(id, pane); pane.Visible = true;
                }
                catch { pane.Delete(); throw; }
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
            catch { System.Windows.Forms.MessageBox.Show(UiText.Get(_settings.UiLanguage, "ChatPaneError"), UiText.Get(_settings.UiLanguage, "Chat")); }
        }
        [DllImport("user32.dll", EntryPoint = "IsWindow")] private static extern bool ChatWindowAlive(IntPtr hwnd);
        private static void DisposeChat(Microsoft.Office.Core.CustomTaskPane pane)
        {
            try { if (pane.ContentControl is DocumentChatPane content) { content.Stop(); content.Dispose(); } } catch { }
            try { pane.Delete(); } catch { }
        }
        private void CloseChat()
        { _chatCleanup?.Stop(); _chatCleanup?.Dispose(); _chatCleanup = null; foreach (var pane in _chatPanes.Values) DisposeChat(pane); _chatPanes.Clear(); _chatFactory = null; }
    }
}
