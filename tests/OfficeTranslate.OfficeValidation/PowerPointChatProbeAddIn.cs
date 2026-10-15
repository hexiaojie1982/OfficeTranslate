using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using OfficeTranslate.Core;
using Office = Microsoft.Office.Core;
using Ppt = Microsoft.Office.Interop.PowerPoint;

namespace OfficeTranslate.OfficeValidation
{
    // Test-only Office entry: forwards the REAL Office factory to the unchanged
    // product AddIn. This assembly is not in the solution/MSI and is never shipped.
    [ComVisible(true), Guid("1CFF20FA-E1F4-4CE9-88AC-092303010001")]
    [ProgId("OfficeTranslate.PowerPointChatProbe"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class PowerPointChatProbeAddIn : Extensibility.IDTExtensibility2, Office.ICustomTaskPaneConsumer
    {
        private readonly PowerPointAddIn.AddIn _addin = new PowerPointAddIn.AddIn();
        private Ppt.Application _app;
        private OfficeUiDispatcher _ui;
        private const BindingFlags PrivateInstance = BindingFlags.Instance | BindingFlags.NonPublic;
        public void OnConnection(object application, Extensibility.ext_ConnectMode mode, object instance, ref Array custom)
        {
            // Use missing fixture settings; never read or change the user's model/key.
            var store = typeof(PowerPointAddIn.AddIn).GetField("_store", PrivateInstance).GetValue(_addin);
            typeof(SettingsStore).GetField("_path", PrivateInstance).SetValue(store,
                Path.Combine(Path.GetTempPath(), "OfficeTranslate-ChatProbe-" + Guid.NewGuid() + ".json"));
            _app = (Ppt.Application)application;
            _addin.OnConnection(application, mode, instance, ref custom);
            _ui = OfficeUiDispatcher.Capture("PowerPoint-Chat-acceptance");
            ((Office.COMAddIn)instance).Object = this;
        }
        public void CTPFactoryAvailable(Office.ICTPFactory factory) => _addin.CTPFactoryAvailable(factory);
        public void Run(string output) => _ui.Invoke(() => RunOnUi(output));
        private void RunOnUi(string output)
        {
            var stage = "create-document";
            try
            {
                Check(_ui.IsUiThread, "Office UI thread");
                var first = _app.Presentations.Add(Office.MsoTriState.msoTrue);
                var slide = first.Slides.Add(1, Ppt.PpSlideLayout.ppLayoutBlank);
                slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 20, 20, 300, 100).TextFrame.TextRange.Text = "Chat probe fixture";
                var window = first.Windows[1];
                dynamic late = window;
                var failedLate = false;
                try { int ignored = late.HWND; }
                catch (MissingMemberException) { failedLate = true; }
                Check(failedLate, "Baseline dynamic HWND reproduces MissingMemberException");
                var id = window.HWND;
                Check(id != 0, "Typed DocumentWindow.HWND succeeds");
                stage = "create-native-pane";
                var core = typeof(PowerPointAddIn.AddIn).GetMethod("OpenChatCore", PrivateInstance);
                Check(core != null, "Candidate method missing; loaded " + typeof(PowerPointAddIn.AddIn).Assembly.Location);
                core.Invoke(_addin, new object[] { "active-window" });
                stage = "inspect-native-pane";
                var panes = (IDictionary)typeof(PowerPointAddIn.AddIn).GetField("_chatPanes", PrivateInstance).GetValue(_addin);
                Check(panes.Count == 1, "One native pane created");
                var pane = (Office.CustomTaskPane)panes[id];
                Check(pane.Visible && pane.Width == 440 && pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionRight, "Native pane visible/right/440");
                var control = (UserControl)pane.ContentControl;
                Check(control.Controls.Count > 0, "COM activated actual product UserControl and initialized UI");
                stage = "capture-context";
                var capture = (Func<ChatContext>)control.GetType().GetField("_capture", PrivateInstance).GetValue(control);
                var context = capture();
                Check(context.Text.Contains("Chat probe fixture"), "Real presentation context captured");
                stage = "ribbon-reopen";
                pane.Visible = false;
                _addin.OpenChat(null); // Exact Ribbon entry, after the throwing core was verified.
                Check(pane.Visible && panes.Count == 1, "Ribbon entry reopens same pane");
                stage = "second-document";
                var second = _app.Presentations.Add(Office.MsoTriState.msoTrue);
                second.Slides.Add(1, Ppt.PpSlideLayout.ppLayoutBlank);
                core.Invoke(_addin, new object[] { "active-window" });
                Check(panes.Count == 2, "Two document windows have independent panes");
                stage = "origin-writeback";
                context.Apply(ChatMode.Summarize, ChatWriteAction.InsertAtStart, "Probe summary");
                Check(first.Slides.Count == 2 && second.Slides.Count == 1, "Writeback stays with originating presentation");
                stage = "close-cleanup";
                first.Saved = Office.MsoTriState.msoTrue; first.Close();
                var until = Stopwatch.StartNew();
                while (panes.Count != 1 && until.ElapsedMilliseconds < 6000) { Application.DoEvents(); Thread.Sleep(20); }
                Check(panes.Count == 1, "Closed document pane cleaned up");
                stage = "new-document-reopen";
                second.Saved = Office.MsoTriState.msoTrue; second.Close();
                var third = _app.Presentations.Add(Office.MsoTriState.msoTrue);
                third.Slides.Add(1, Ppt.PpSlideLayout.ppLayoutBlank);
                core.Invoke(_addin, new object[] { "active-window" });
                var reopened = (Office.CustomTaskPane)panes[third.Windows[1].HWND];
                Check(reopened.Visible, "Pane opens after documents closed/reopened");
                third.Saved = Office.MsoTriState.msoTrue; third.Close();
                File.WriteAllText(output, "PASS: real PowerPoint UI STA; baseline late HWND fails; typed HWND succeeds; native COM control activation, Ribbon reopen, multi-window isolation, origin writeback, cleanup and new-document reopen.\r\nVersion="
                    + typeof(PowerPointAddIn.AddIn).Assembly.GetName().Version);
            }
            catch (Exception error) { File.WriteAllText(output, "FAIL Stage=" + stage + ": " + error); }
        }
        private static void Check(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
        public void OnDisconnection(Extensibility.ext_DisconnectMode mode, ref Array custom) => _addin.OnDisconnection(mode, ref custom);
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) => _addin.OnBeginShutdown(ref custom);
    }
}
