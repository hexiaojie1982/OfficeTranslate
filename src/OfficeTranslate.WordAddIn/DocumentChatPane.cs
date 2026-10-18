using OfficeTranslate.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
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
        private DocumentChatSession _session = new DocumentChatSession();
        internal ChatHistoryStore HistoryStore { get; set; } = new ChatHistoryStore();
        internal SettingsStore ChatSettingsStore { get; set; } = new SettingsStore();
        private readonly ChatConversationView _transcript = new ChatConversationView { Dock = DockStyle.Fill };
        private readonly TextBox _input = new TextBox { Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, MaxLength = 0 };
        private readonly Label _context = new Label { AutoSize = true, ForeColor = Color.FromArgb(90, 104, 122), MaximumSize = new Size(440, 0) };
        private readonly Label _status = new Label { AutoSize = true, MaximumSize = new Size(440, 0) };
        private readonly TableLayoutPanel _commands = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Margin = Padding.Empty };
        private readonly CheckBox _useTranslationSettings = new CheckBox { Tag = "ChatUseTranslationSettings", AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 8) };
        private bool _refreshingSettings;
        private readonly FlowLayoutPanel _writes = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        private readonly Button _cancel = new Button();
        private readonly ToolTip _tips = new ToolTip();
        private readonly System.Windows.Forms.Timer _streamTimer = new System.Windows.Forms.Timer { Interval = 50 };
        private readonly object _streamSync = new object();
        private string? _queuedOutput;
        private string _renderedOutput = "", _statusKey = "";
        private bool _recovering, _acceptingStream, _hasTurn;
        private readonly List<string> _statusSuffixes = new List<string>();
        private object[] _statusArgs = new object[0];
        private ChatContext? _displayContext;
        private CancellationTokenSource? _pending;
        internal ChatTransport? Transport { get; set; }
        internal Func<ChatWriteAction, bool>? ConfirmWrite { get; set; }
        internal Func<SaveFileDialog, DialogResult>? ExportDialog { get; set; }
        internal Action<Form>? HistoryDialog { get; set; }
        private Func<ChatContext>? _capture;
        private Func<TranslationSettings>? _settings;
        private string _language = "zh-CN", _result = "";
        private string _turnQuestion = "", _errorDetails = "";
        private DateTimeOffset _turnTime;
        private IReadOnlyList<ChatMessage> _turnHistory = new ChatMessage[0];
        private ChatContext? _resultContext;
        private ChatMode _resultMode;
        private bool _written;
        public DocumentChatPane()
        {
            Font = new Font("Microsoft YaHei UI", 9F); BackColor = Color.FromArgb(240, 244, 250); ForeColor = Color.FromArgb(32, 55, 88); Size = new Size(420, 650);
            DoubleBuffered = true;
            _streamTimer.Tick += (s, e) => FlushOutput();
        }
        internal void Initialize(Func<ChatContext> capture, Func<TranslationSettings> settings)
        {
            _session = new DocumentChatSession(HistoryStore);
            _capture = capture; _settings = settings; _language = settings().UiLanguage;
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 9 };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (var i = 0; i < 9; i++) root.RowStyles.Add(new RowStyle(i == 3 ? SizeType.Percent : i == 7 ? SizeType.Absolute : SizeType.AutoSize, i == 3 ? 100 : i == 7 ? 90 : 0));
            root.Controls.Add(new Label { Tag = "Chat", Text = T("Chat"), AutoSize = true, ForeColor = Color.FromArgb(32, 55, 88), Font = new Font(Font.FontFamily, 14F, FontStyle.Bold) }, 0, 0);
            root.Controls.Add(_context, 0, 1);
            var functions = ButtonRow(); var operations = ButtonRow();
            _commands.Controls.Add(functions, 0, 0); _commands.Controls.Add(_useTranslationSettings, 0, 1); _commands.Controls.Add(operations, 0, 2);
            AddButton(functions, "ChatTranslate", () => Start(ChatMode.Translate), true);
            AddButton(functions, "ChatSummary", () => Start(ChatMode.Summarize), true); AddButton(functions, "ChatRewrite", () => Start(ChatMode.Rewrite), true);
            AddButton(functions, "ChatSettings", () => { using (var form = new ChatSettingsForm(ChatSettingsStore)) form.ShowDialog(this); RefreshLanguage(); }, true);
            _useTranslationSettings.CheckedChanged += (s, e) => SaveTranslationSettingsChoice();
            AddButton(operations, "ChatClear", Clear); AddButton(operations, "ChatRefresh", RefreshContext);
            AddButton(operations, "ChatExport", Export);
            AddButton(operations, "ChatHistory", () =>
            {
                RefreshLanguage();
                using (var form = new ChatHistoryForm(HistoryStore, _language))
                { if (HistoryDialog != null) HistoryDialog(form); else form.ShowDialog(this); }
            });
            root.Controls.Add(_commands, 0, 2); root.Controls.Add(_transcript, 0, 3); root.Controls.Add(_writes, 0, 4); root.Controls.Add(_status, 0, 5);
            root.Controls.Add(new Label { Tag = "ChatInputHint", Text = T("ChatInputHint"), AutoSize = true }, 0, 6);
            root.Controls.Add(_input, 0, 7);
            var sendBar = ButtonRow();
            AddButton(sendBar, "ChatSend", () => Start(ChatMode.Question), true);
            _cancel.Tag = "Cancel"; Style(_cancel, T("Cancel")); _cancel.Click += (s, e) => _pending?.Cancel(); _cancel.Enabled = false; sendBar.Controls.Add(_cancel);
            root.Controls.Add(sendBar, 0, 8); Controls.Add(root);
            _input.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter && !e.Shift) { e.SuppressKeyPress = true; Start(ChatMode.Question); } };
            Enter += (s, e) => { RefreshLanguage(); if (_pending == null) RefreshContext(); };
            VisibleChanged += (s, e) => { if (Visible) { RefreshLanguage(); if (_pending == null) RefreshContext(); } };
            RenderTranscript(); RefreshLanguage(); RefreshContext();
        }
        private string T(string key) => UiText.Get(_language, key);
        private static FlowLayoutPanel ButtonRow() => new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill, Margin = Padding.Empty, Padding = Padding.Empty };
        private static void Style(Button b, string text, bool primary = false)
        {
            b.Text = text; b.AutoSize = false; b.Height = 32;
            b.Width = Math.Max(72, TextRenderer.MeasureText(text, b.Font).Width + 24);
            b.Margin = new Padding(0, 0, 8, 8); b.FlatStyle = FlatStyle.Flat; b.FlatAppearance.BorderSize = 0;
            b.BackColor = primary ? Color.FromArgb(45, 108, 223) : Color.FromArgb(231, 236, 243);
            b.ForeColor = primary ? Color.White : Color.FromArgb(45, 60, 80);
            b.UseVisualStyleBackColor = false; b.Cursor = Cursors.Hand;
        }
        private Button AddButton(FlowLayoutPanel panel, string key, Action action, bool primary = false)
        { var b = new Button { Tag = key, Font = Font }; Style(b, T(key), primary); b.Click += (s, e) => action(); panel.Controls.Add(b); return b; }
        internal void RefreshLanguage()
        {
            if (_settings == null) return;
            var settings = _settings();
            var language = settings.UiLanguage; var changed = language != _language; _language = language;
            SuspendLayout();
            try
            {
                Localize(Controls);
                _refreshingSettings = true;
                try { _useTranslationSettings.Checked = settings.ChatUseTranslationSettings; }
                finally { _refreshingSettings = false; }
                _input.AccessibleName = T("ChatInputHint"); _transcript.AccessibleName = T("ChatTranscript");
                _tips.SetToolTip(_context, T("ChatContextTip")); _tips.SetToolTip(_input, T("ChatInputHint"));
                ShowContext();
                if (_statusKey.Length > 0) RenderStatus();
                if (changed)
                {
                    if (_hasTurn || _pending == null) RenderTranscript();
                }
            }
            finally { ResumeLayout(true); }
        }
        private void SaveTranslationSettingsChoice()
        {
            if (_refreshingSettings || _settings == null) return;
            var settings = _settings();
            var previous = settings.ChatUseTranslationSettings;
            settings.ChatUseTranslationSettings = _useTranslationSettings.Checked;
            try { ChatSettingsStore.Save(settings); }
            catch
            {
                settings.ChatUseTranslationSettings = previous;
                _refreshingSettings = true;
                try { _useTranslationSettings.Checked = previous; }
                finally { _refreshingSettings = false; }
                SetStatus("ChatSettingsError", true);
            }
        }
        private void Localize(Control.ControlCollection controls)
        {
            foreach (Control control in controls)
            {
                if (control.Tag is string key)
                {
                    if (control is Button button) Style(button, T(key), key == "ChatTranslate" || key == "ChatSummary" || key == "ChatRewrite" || key == "ChatSettings" || key == "ChatSend" || key == "ChatReplace" || key == "ChatInsertBelow" || key == "ChatInsertStart");
                    else control.Text = T(key);
                    if (key == "ChatRefresh") _tips.SetToolTip(control, T("ChatRefreshTip"));
                }
                Localize(control.Controls);
            }
        }
        private void RenderTranscript() => _transcript.ShowConversation(_hasTurn ? _turnHistory : _session.History,
            T("ChatYou"), T("ChatAssistant"), T("ChatEmpty"), _hasTurn ? _turnQuestion : null,
            _hasTurn ? _renderedOutput : "", ErrorText(), T("ChatUserAvatar"), T("ChatAiAvatar"), _turnTime, _language, _acceptingStream);
        private string ErrorText() => _errorDetails.Length == 0 ? "" : "\r\n\r\n" + T("ChatErrorDetails") + "\r\n" + _errorDetails;
        private void SetStatus(string key, bool error = false)
        { _statusKey = key; _statusArgs = new object[0]; _statusSuffixes.Clear(); RenderStatus(); _status.ForeColor = error ? Color.FromArgb(190, 45, 45) : Color.FromArgb(90, 104, 122); }
        private void RenderStatus()
        { _status.Text = string.Format(T(_statusKey), _statusArgs); foreach (var key in _statusSuffixes) _status.Text += " " + T(key); }
        private void AppendStatus(string key) { _statusSuffixes.Add(key); RenderStatus(); }
        private void RefreshContext()
        {
            RefreshLanguage();
            try { _displayContext = _capture!(); }
            catch { _displayContext = null; }
            ShowContext();
        }
        private void ShowContext() => _context.Text = _displayContext == null ? T("ChatNoDocument") : string.Format(T("ChatContext"), T(_displayContext.IsSelection ? "ChatSelected" : "ChatFull"), _displayContext.Text.Trim().Length);
        // Network callbacks only replace a snapshot. The UI timer batches tokens;
        // no per-token dispatcher calls, layout or sibling invalidation occurs.
        private void QueueOutput(string text, CancellationTokenSource request)
        {
            lock (_streamSync)
            {
                if (!_acceptingStream || _pending != request || request.IsCancellationRequested) return;
                _queuedOutput = text; if (text.Length == 0) _recovering = true;
            }
        }
        private void FlushOutput()
        {
            string? text; bool recovering;
            lock (_streamSync) { text = _queuedOutput; _queuedOutput = null; recovering = _recovering; _recovering = false; }
            if (IsDisposed || text == null) return;
            if (recovering) SetStatus("ChatRecovering");
            _transcript.ShowOutput(text); _renderedOutput = text;
        }
        private void EndStream(bool flush)
        {
            _streamTimer.Stop();
            lock (_streamSync) { _acceptingStream = false; }
            if (flush) FlushOutput();
            lock (_streamSync) { _queuedOutput = null; _recovering = false; }
        }
        private void ClearWrites()
        { while (_writes.Controls.Count > 0) { var control = _writes.Controls[0]; _writes.Controls.RemoveAt(0); control.Dispose(); } }
        private void Clear()
        { if (_pending != null) return; if (!_session.Clear()) { SetStatus("ChatHistorySaveFailed", true); return; } _hasTurn = false; _renderedOutput = _result = _turnQuestion = _errorDetails = ""; _turnHistory = new ChatMessage[0]; _resultContext = null; RenderTranscript(); ClearWrites(); SetStatus("ChatCleared"); }
        private void Export()
        {
            RefreshLanguage();
            var history = _session.History;
            if (history.Count == 0) { SetStatus("ChatExportEmpty"); return; }
            try
            {
                using (var dialog = new SaveFileDialog
                {
                    Title = T("ChatExportTitle"), Filter = T("ChatExportFilter"), FilterIndex = 1,
                    DefaultExt = "md", AddExtension = true, OverwritePrompt = true, RestoreDirectory = true,
                    FileName = string.Format(CultureInfo.InvariantCulture, T("ChatExportFileName"), DateTime.Now)
                })
                {
                    if ((ExportDialog != null ? ExportDialog(dialog) : dialog.ShowDialog(this)) != DialogResult.OK) return;
                    var markdown = dialog.FilterIndex == 1;
                    File.WriteAllText(dialog.FileName, ChatTranscriptExport.Format(history, markdown, T), new UTF8Encoding(true));
                    SetStatus("ChatExportSuccess");
                }
            }
            catch (Exception ex)
            {
                // Never display/log raw filesystem errors: they may include paths, account data or content.
                SetStatus("ChatExportFailed", true);
                AppendStatus(ex is UnauthorizedAccessException || ex is System.Security.SecurityException ? "ChatExportAccessError" :
                    ex is IOException ? "ChatExportIoError" : ex is ArgumentException || ex is NotSupportedException ? "ChatExportPathError" : "ChatExportUnknownError");
            }
        }
        private async void Start(ChatMode mode)
        {
            if (_pending != null || _capture == null || _settings == null) return;
            RefreshLanguage();
            var settings = _settings();
            string question;
            switch (mode)
            {
                case ChatMode.Translate: question = string.Format(T("ChatTranslateRequest"), settings.TargetLanguage); break;
                case ChatMode.Summarize: question = T("ChatSummaryRequest"); break;
                case ChatMode.Rewrite: question = T("ChatRewriteRequest"); break;
                default: question = _input.Text.Trim(); break;
            }
            if (string.IsNullOrWhiteSpace(question)) { SetStatus("ChatNeedInput", true); return; }
            var ui = OfficeUiDispatcher.Capture("Chat");
            var cancellation = new CancellationTokenSource();
            lock (_streamSync) { _pending = cancellation; _queuedOutput = null; _recovering = false; _acceptingStream = true; }
            TranslationSettings? diagnosticSettings = null; ChatContext? diagnosticContext = null;
            _commands.Enabled = false; _cancel.Enabled = true; ClearWrites(); _resultContext = null; _result = "";
            try
            {
                diagnosticSettings = settings; var context = diagnosticContext = _capture();
                _displayContext = context; ShowContext();
                if (string.IsNullOrWhiteSpace(context.Text)) { SetStatus("ChatNoText", true); return; }
                _turnHistory = _session.History; _turnQuestion = question; _turnTime = DateTimeOffset.Now; _errorDetails = ""; _hasTurn = true;
                _renderedOutput = ""; RenderTranscript(); SetStatus("ChatLoading"); _streamTimer.Start();
                using (var client = new TranslationClient())
                {
                    // Office's ambient SynchronizationContext may dispatch to an MTA pool thread.
                    // Use the same explicit STA dispatcher as translation and OCR.
                    ChatTransport transport = Transport ?? client.ChatAsync;
                    var reply = await _session.SendAsync(context, question, mode, settings, transport,
                        text => QueueOutput(text, cancellation),
                        (i, n) => ui.Invoke(() => { if (!IsDisposed && _pending == cancellation) { if (n == 0) SetStatus("ChatLoading"); else { _statusKey = "ChatChunkProgress"; _statusArgs = new object[] { i, n }; _statusSuffixes.Clear(); RenderStatus(); } } }), cancellation.Token).ConfigureAwait(false);
                    ui.Invoke(() =>
                    {
                        if (IsDisposed || cancellation.IsCancellationRequested) return;
                        EndStream(false); _result = reply.Text; _hasTurn = false; RenderTranscript();
                        _turnQuestion = ""; _turnHistory = new ChatMessage[0]; _resultMode = mode; _resultContext = context; _written = false;
                        if (mode == ChatMode.Question) _input.Clear();
                        SetStatus(reply.UsedFallback ? "ChatFallback" : "ChatDone");
                        if (reply.UsedSummaries) AppendStatus("ChatReduced");
                        BuildWrites();
                        if (_session.HistorySaveFailed) SetStatus("ChatHistorySaveFailed", true);
                    });
                }
            }
            catch (Exception) when (cancellation.IsCancellationRequested) { ui.Invoke(() => { if (!IsDisposed) { EndStream(true); SetStatus("ChatCancelled"); } }); }
            catch (Exception ex)
            {
                // Chat transport normalizes provider/parser errors before they reach the pane.
                ui.Invoke(() =>
                {
                    if (!IsDisposed)
                    {
                        EndStream(true); SetStatus(ChatErrors.UiKey(ex), true);
                        _errorDetails = ChatErrors.Describe(ex, diagnosticSettings, diagnosticContext?.Text ?? "", question);
                        RenderTranscript();
                    }
                });
            }
            finally { ui.Invoke(() => { if (!IsDisposed) { EndStream(false); if (_hasTurn) RenderTranscript(); _commands.Enabled = true; _cancel.Enabled = false; } lock (_streamSync) { _pending = null; } }); cancellation.Dispose(); }
        }
        private void BuildWrites()
        {
            // Translation and rewriting correspond to the source position: replace the
            // captured selection or insert below it; never insert at the document start.
            var sourceResult = _resultMode == ChatMode.Rewrite || _resultMode == ChatMode.Translate;
            foreach (ChatWriteAction action in Enum.GetValues(typeof(ChatWriteAction)))
            {
                if (_resultMode == ChatMode.Question || (_resultMode == ChatMode.Summarize && action != ChatWriteAction.InsertAtStart) ||
                    (sourceResult && action == ChatWriteAction.InsertAtStart)) continue;
                var allowed = _resultContext?.Allows(_resultMode, action, _result) == true;
                var key = action == ChatWriteAction.ReplaceSelection ? "ChatReplace" : action == ChatWriteAction.InsertBelow ? "ChatInsertBelow" : "ChatInsertStart";
                var button = AddButton(_writes, key, () =>
                {
                    if (_written || _resultContext == null) return;
                    var confirmed = ConfirmWrite != null ? ConfirmWrite(action) : MessageBox.Show(this, T("ChatWriteConfirm"), T(key), MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
                    if (!confirmed) return;
                    try { _resultContext.Apply(_resultMode, action, _result); _written = true; _writes.Enabled = false; SetStatus("ChatWritten"); }
                    catch { SetStatus("ChatWriteError", true); }
                }, true);
                button.Enabled = allowed;
            }
            _writes.Enabled = true;
            if (sourceResult && _resultContext?.CanReplace != true) AppendStatus("ChatReplaceUnavailable");
        }
        internal void Stop() { _pending?.Cancel(); EndStream(false); _session.Clear(); _hasTurn = false; _renderedOutput = _result = _turnQuestion = _errorDetails = ""; _turnHistory = new ChatMessage[0]; _resultContext = null; }
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            var width = Math.Max(100, ClientSize.Width - 32);
            _context.MaximumSize = new Size(width, 0); _status.MaximumSize = new Size(width, 0);
            _useTranslationSettings.MaximumSize = new Size(width, 0);
        }
        protected override void Dispose(bool disposing)
        { if (disposing) { Stop(); _streamTimer.Dispose(); _tips.Dispose(); _resultContext = _displayContext = null; _capture = null; _settings = null; } base.Dispose(disposing); }
    }
    internal sealed class ChatSettingsForm : Form
    {
        public ChatSettingsForm(SettingsStore store)
        {
            var settings = store.Load(); string T(string key) => UiText.Get(settings.UiLanguage, key);
            Text = T("ChatSettings"); Font = new Font("Microsoft YaHei UI", 9F); BackColor = Color.FromArgb(247, 249, 252);
            Size = new Size(620, 720); MinimumSize = new Size(520, 500); StartPosition = FormStartPosition.CenterParent; ShowInTaskbar = false;
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), AutoScroll = true, ColumnCount = 1, RowCount = 13 };
            TextBox Add(string label, string value) { root.Controls.Add(new Label { Text = T(label), AutoSize = true }); var box = new TextBox { Tag = label, Multiline = true, ScrollBars = ScrollBars.Vertical, Text = value, Dock = DockStyle.Fill, Height = 80, MaxLength = 0 }; root.Controls.Add(box); return box; }
            var q = Add("ChatQuestionPrompt", settings.ChatQuestionSystemPrompt); var t = Add("ChatTranslatePrompt", settings.ChatTranslateSystemPrompt); var s = Add("ChatSummaryPrompt", settings.ChatSummarySystemPrompt); var r = Add("ChatRewritePrompt", settings.ChatRewriteSystemPrompt);
            root.Controls.Add(new Label { Text = T("ChatBudget"), AutoSize = true });
            var budget = new NumericUpDown { Minimum = 2048, Maximum = 1048576, Value = Math.Max(2048, Math.Min(1048576, settings.ChatContextTokens)), Dock = DockStyle.Fill }; root.Controls.Add(budget);
            root.Controls.Add(new Label { Text = T("ChatOutputBudget"), AutoSize = true });
            var output = new NumericUpDown { Minimum = 256, Maximum = 524288, Value = Math.Max(256, Math.Min(524288, settings.ChatOutputTokens)), Dock = DockStyle.Fill }; root.Controls.Add(output);
            var save = new Button { Text = T("Save"), AutoSize = true, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(45,108,223), ForeColor = Color.White }; save.FlatAppearance.BorderSize = 0;
            save.Click += (a, b) => { if (output.Value > budget.Value / 2 || string.IsNullOrWhiteSpace(t.Text) || string.IsNullOrWhiteSpace(q.Text) || string.IsNullOrWhiteSpace(s.Text) || string.IsNullOrWhiteSpace(r.Text)) { MessageBox.Show(this, T("ChatInvalidSettings"), Text); return; }
                try { settings.ChatTranslateSystemPrompt = t.Text; settings.ChatQuestionSystemPrompt = q.Text; settings.ChatSummarySystemPrompt = s.Text; settings.ChatRewriteSystemPrompt = r.Text; settings.ChatContextTokens = (int)budget.Value; settings.ChatOutputTokens = (int)output.Value; store.Save(settings); Close(); } catch { MessageBox.Show(this, T("ChatSettingsError"), Text); } };
            root.Controls.Add(save); Controls.Add(root); AcceptButton = save;
        }
    }
}
