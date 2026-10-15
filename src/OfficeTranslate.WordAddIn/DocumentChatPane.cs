using OfficeTranslate.Core;
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeTranslate.WordAddIn
{
    [ComVisible(true)]
#if WORD
    [Guid("EFD0B5E1-8082-4C8C-9429-3011D0230001"), ProgId("OfficeTranslate.WordChatPane")]
#elif EXCEL
    [Guid("EFD0B5E1-8082-4C8C-9429-3011D0230002"), ProgId("OfficeTranslate.ExcelChatPane")]
#else
    [Guid("EFD0B5E1-8082-4C8C-9429-3011D0230003"), ProgId("OfficeTranslate.PowerPointChatPane")]
#endif
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class DocumentChatPane : UserControl
    {
        private readonly DocumentChatSession _session = new DocumentChatSession();
        private readonly TextBox _transcript = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, MaxLength = 0, BackColor = Color.White };
        private readonly TextBox _input = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, MaxLength = 0 };
        private readonly Label _context = new Label { AutoSize = true, ForeColor = Color.FromArgb(90, 104, 122), MaximumSize = new Size(440, 0) };
        private readonly Label _status = new Label { AutoSize = true, MaximumSize = new Size(440, 0) };
        private readonly FlowLayoutPanel _commands = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        private readonly FlowLayoutPanel _writes = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        private readonly Button _cancel = new Button();
        private CancellationTokenSource? _pending;
        internal ChatTransport? Transport { get; set; }
        internal Func<ChatWriteAction, bool>? ConfirmWrite { get; set; }
        private Func<ChatContext>? _capture;
        private Func<TranslationSettings>? _settings;
        private string _language = "zh-CN", _prefix = "", _result = "";
        private ChatContext? _resultContext;
        private ChatMode _resultMode;
        private bool _written;
        public DocumentChatPane()
        {
            Font = new Font("Microsoft YaHei UI", 9F); BackColor = Color.FromArgb(247, 249, 252); Size = new Size(420, 650);
        }
        internal void Initialize(Func<ChatContext> capture, Func<TranslationSettings> settings)
        {
            _capture = capture; _settings = settings; _language = settings().UiLanguage;
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 8 };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (var i = 0; i < 8; i++) root.RowStyles.Add(new RowStyle(i == 3 ? SizeType.Percent : i == 6 ? SizeType.Absolute : SizeType.AutoSize, i == 3 ? 100 : i == 6 ? 90 : 0));
            root.Controls.Add(new Label { Text = T("Chat"), AutoSize = true, ForeColor = Color.FromArgb(32, 55, 88), Font = new Font(Font.FontFamily, 14F, FontStyle.Bold) }, 0, 0);
            root.Controls.Add(_context, 0, 1);
            AddButton(_commands, "ChatSummary", () => Start(ChatMode.Summarize)); AddButton(_commands, "ChatRewrite", () => Start(ChatMode.Rewrite));
            AddButton(_commands, "ChatClear", Clear); AddButton(_commands, "ChatRefresh", RefreshContext);
            AddButton(_commands, "ChatSettings", () => { using (var form = new ChatSettingsForm(new SettingsStore())) form.ShowDialog(this); });
            root.Controls.Add(_commands, 0, 2); root.Controls.Add(_transcript, 0, 3); root.Controls.Add(_writes, 0, 4); root.Controls.Add(_status, 0, 5);
            root.Controls.Add(_input, 0, 6);
            var sendBar = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
            AddButton(sendBar, "ChatSend", () => Start(ChatMode.Question));
            Style(_cancel, T("Cancel")); _cancel.Click += (s, e) => _pending?.Cancel(); _cancel.Enabled = false; sendBar.Controls.Add(_cancel);
            root.Controls.Add(sendBar, 0, 7); Controls.Add(root);
            _input.KeyDown += (s, e) => { if (e.Control && e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Start(ChatMode.Question); } };
            Enter += (s, e) => { if (_pending == null) RefreshContext(); };
            VisibleChanged += (s, e) => { if (Visible && _pending == null) RefreshContext(); };
            _transcript.Text = T("ChatEmpty"); RefreshContext();
        }
        private string T(string key) => UiText.Get(_language, key);
        private static void Style(Button b, string text)
        { b.Text = text; b.AutoSize = true; b.FlatStyle = FlatStyle.Flat; b.FlatAppearance.BorderSize = 0; b.BackColor = Color.FromArgb(45, 108, 223); b.ForeColor = Color.White; b.UseVisualStyleBackColor = false; b.Padding = new Padding(8, 4, 8, 4); }
        private Button AddButton(FlowLayoutPanel panel, string key, Action action)
        { var b = new Button(); Style(b, T(key)); b.Click += (s, e) => action(); panel.Controls.Add(b); return b; }
        private void SetStatus(string key, bool error = false)
        { _status.Text = T(key); _status.ForeColor = error ? Color.FromArgb(190, 45, 45) : Color.FromArgb(90, 104, 122); }
        private void RefreshContext()
        {
            try { var c = _capture!(); _context.Text = string.Format(T("ChatContext"), T(c.IsSelection ? "ChatSelected" : "ChatFull"), c.Text.Trim().Length); }
            catch { _context.Text = T("ChatNoDocument"); }
        }
        private void ClearWrites()
        { while (_writes.Controls.Count > 0) { var control = _writes.Controls[0]; _writes.Controls.RemoveAt(0); control.Dispose(); } }
        private void Clear()
        { if (_pending != null) return; _session.Clear(); _transcript.Text = T("ChatEmpty"); _prefix = _result = ""; _resultContext = null; ClearWrites(); SetStatus("ChatCleared"); }
        private async void Start(ChatMode mode)
        {
            if (_pending != null || _capture == null || _settings == null) return;
            var question = mode == ChatMode.Question ? _input.Text.Trim() : T(mode == ChatMode.Summarize ? "ChatSummaryRequest" : "ChatRewriteRequest");
            if (string.IsNullOrWhiteSpace(question)) { SetStatus("ChatNeedInput", true); return; }
            var ui = OfficeUiDispatcher.Capture("Chat");
            var cancellation = new CancellationTokenSource(); _pending = cancellation;
            _commands.Enabled = false; _cancel.Enabled = true; ClearWrites(); _resultContext = null; _result = "";
            try
            {
                var context = _capture(); var settings = _settings();
                _context.Text = string.Format(T("ChatContext"), T(context.IsSelection ? "ChatSelected" : "ChatFull"), context.Text.Trim().Length);
                if (string.IsNullOrWhiteSpace(context.Text)) { SetStatus("ChatNoText", true); return; }
                _prefix = "";
                foreach (var m in _session.History) _prefix += T(m.Role == "user" ? "ChatYou" : "ChatAssistant") + "\r\n" + m.Content + "\r\n\r\n";
                _prefix += T("ChatYou") + "\r\n" + question + "\r\n\r\n" + T("ChatAssistant") + "\r\n";
                _transcript.Text = _prefix; SetStatus("ChatLoading");
                using (var client = new TranslationClient())
                {
                    // Office's ambient SynchronizationContext may dispatch to an MTA pool thread.
                    // Use the same explicit STA dispatcher as translation and OCR.
                    ChatTransport transport = Transport ?? client.ChatAsync;
                    var reply = await _session.SendAsync(context, question, mode, settings, transport,
                        text => ui.Invoke(() => { if (!IsDisposed && _pending == cancellation) { if (text.Length == 0) SetStatus("ChatRecovering"); _transcript.Text = _prefix + text; _transcript.SelectionStart = _transcript.TextLength; _transcript.ScrollToCaret(); } }),
                        (i, n) => ui.Invoke(() => { if (!IsDisposed && _pending == cancellation) { if (n == 0) SetStatus("ChatLoading"); else _status.Text = string.Format(T("ChatChunkProgress"), i, n); } }), cancellation.Token).ConfigureAwait(false);
                    ui.Invoke(() =>
                    {
                        if (IsDisposed || cancellation.IsCancellationRequested) return;
                        _result = reply.Text; _transcript.Clear();
                        foreach (var m in _session.History) _transcript.AppendText(T(m.Role == "user" ? "ChatYou" : "ChatAssistant") + "\r\n" + m.Content + "\r\n\r\n");
                        _prefix = ""; _resultMode = mode; _resultContext = context; _written = false;
                        if (mode == ChatMode.Question) _input.Clear();
                        SetStatus(reply.UsedFallback ? "ChatFallback" : "ChatDone");
                        if (reply.UsedSummaries) _status.Text += " " + T("ChatReduced");
                        BuildWrites();
                    });
                }
            }
            catch (Exception) when (cancellation.IsCancellationRequested) { ui.Invoke(() => { if (!IsDisposed) SetStatus("ChatCancelled"); }); }
            // Server bodies may echo private input or keys. Never display/log raw exception messages.
            catch { ui.Invoke(() => { if (!IsDisposed) SetStatus("ChatError", true); }); }
            finally { ui.Invoke(() => { if (!IsDisposed) { _commands.Enabled = true; _cancel.Enabled = false; } _pending = null; }); cancellation.Dispose(); }
        }
        private void BuildWrites()
        {
            foreach (ChatWriteAction action in Enum.GetValues(typeof(ChatWriteAction)))
            {
                if (_resultMode == ChatMode.Question || (_resultMode == ChatMode.Summarize && action != ChatWriteAction.InsertAtStart) ||
                    (_resultMode == ChatMode.Rewrite && action == ChatWriteAction.InsertAtStart)) continue;
                var allowed = _resultContext?.Allows(_resultMode, action, _result) == true;
                var key = action == ChatWriteAction.ReplaceSelection ? "ChatReplace" : action == ChatWriteAction.InsertBelow ? "ChatInsertBelow" : "ChatInsertStart";
                var button = AddButton(_writes, key, () =>
                {
                    if (_written || _resultContext == null) return;
                    var confirmed = ConfirmWrite != null ? ConfirmWrite(action) : MessageBox.Show(this, T("ChatWriteConfirm"), T(key), MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
                    if (!confirmed) return;
                    try { _resultContext.Apply(_resultMode, action, _result); _written = true; _writes.Enabled = false; SetStatus("ChatWritten"); }
                    catch { SetStatus("ChatWriteError", true); }
                });
                button.Enabled = allowed;
            }
            _writes.Enabled = true;
            if (_resultMode == ChatMode.Rewrite && _resultContext?.CanReplace != true) _status.Text += " " + T("ChatReplaceUnavailable");
        }
        internal void Stop() { _pending?.Cancel(); _session.Clear(); _prefix = _result = ""; _resultContext = null; }
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            var width = Math.Max(100, ClientSize.Width - 32);
            _context.MaximumSize = new Size(width, 0); _status.MaximumSize = new Size(width, 0);
        }
        protected override void Dispose(bool disposing)
        { if (disposing) { Stop(); _resultContext = null; _capture = null; _settings = null; } base.Dispose(disposing); }
    }
    internal sealed class ChatSettingsForm : Form
    {
        public ChatSettingsForm(SettingsStore store)
        {
            var settings = store.Load(); string T(string key) => UiText.Get(settings.UiLanguage, key);
            Text = T("ChatSettings"); Font = new Font("Microsoft YaHei UI", 9F); BackColor = Color.FromArgb(247, 249, 252);
            Size = new Size(620, 600); MinimumSize = new Size(520, 500); StartPosition = FormStartPosition.CenterParent; ShowInTaskbar = false;
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 11 };
            TextBox Add(string label, string value) { root.Controls.Add(new Label { Text = T(label), AutoSize = true }); var box = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Text = value, Dock = DockStyle.Fill, Height = 80, MaxLength = 0 }; root.Controls.Add(box); return box; }
            var q = Add("ChatQuestionPrompt", settings.ChatQuestionSystemPrompt); var s = Add("ChatSummaryPrompt", settings.ChatSummarySystemPrompt); var r = Add("ChatRewritePrompt", settings.ChatRewriteSystemPrompt);
            root.Controls.Add(new Label { Text = T("ChatBudget"), AutoSize = true });
            var budget = new NumericUpDown { Minimum = 2048, Maximum = 1048576, Value = Math.Max(2048, Math.Min(1048576, settings.ChatContextTokens)), Dock = DockStyle.Fill }; root.Controls.Add(budget);
            root.Controls.Add(new Label { Text = T("ChatOutputBudget"), AutoSize = true });
            var output = new NumericUpDown { Minimum = 256, Maximum = 524288, Value = Math.Max(256, Math.Min(524288, settings.ChatOutputTokens)), Dock = DockStyle.Fill }; root.Controls.Add(output);
            var save = new Button { Text = T("Save"), AutoSize = true, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(45,108,223), ForeColor = Color.White }; save.FlatAppearance.BorderSize = 0;
            save.Click += (a, b) => { if (output.Value > budget.Value / 2 || string.IsNullOrWhiteSpace(q.Text) || string.IsNullOrWhiteSpace(s.Text) || string.IsNullOrWhiteSpace(r.Text)) { MessageBox.Show(this, T("ChatInvalidSettings"), Text); return; }
                try { settings.ChatQuestionSystemPrompt = q.Text; settings.ChatSummarySystemPrompt = s.Text; settings.ChatRewriteSystemPrompt = r.Text; settings.ChatContextTokens = (int)budget.Value; settings.ChatOutputTokens = (int)output.Value; store.Save(settings); Close(); } catch { MessageBox.Show(this, T("ChatSettingsError"), Text); } };
            root.Controls.Add(save); Controls.Add(root); AcceptButton = save;
        }
    }
}
