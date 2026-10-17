using System;
using System.Drawing;
using System.Globalization;
using System.Collections.Generic;
using System.Windows.Forms;
using OfficeTranslate.Core;

namespace OfficeTranslate.WordAddIn
{
    internal sealed class ChatHistoryForm : Form
    {
        private readonly ChatHistoryStore _store;
        private readonly string _language;
        private readonly ListView _list = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, MultiSelect = false, HideSelection = false };
        private readonly ChatConversationView _details = new ChatConversationView { Dock = DockStyle.Fill };
        private readonly Label _metadata = new Label { Dock = DockStyle.Fill, AutoSize = true };
        private readonly Label _status = new Label { Dock = DockStyle.Fill, AutoSize = true };
        private readonly Button _delete = new Button { AutoSize = true };
        private readonly Button _deleteAll = new Button { AutoSize = true };
        internal Func<string, bool>? ConfirmDelete { get; set; }
        private string T(string key) => UiText.Get(_language, key);
        public ChatHistoryForm(ChatHistoryStore store, string language)
        {
            _store = store; _language = language;
            Text = T("ChatHistoryTitle"); Font = new Font("Microsoft YaHei UI", 9F);
            Size = new Size(920, 680); MinimumSize = new Size(650, 480); StartPosition = FormStartPosition.CenterParent; ShowInTaskbar = false;
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 5 };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 38));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 62)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _list.Columns.Add(T("ChatHistoryHost"), 100); _list.Columns.Add(T("ChatHistoryTime"), 170);
            _list.Columns.Add(T("ChatHistoryDocument"), 190); _list.Columns.Add(T("ChatHistoryQuestion"), 360);
            _list.AccessibleName = T("ChatHistoryTitle"); _details.AccessibleName = T("ChatHistoryDetails");
            _list.SelectedIndexChanged += (s, e) => ShowSelected();
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            _delete.Text = T("ChatHistoryDelete"); _delete.Click += (s, e) => DeleteSelected();
            _deleteAll.Text = T("ChatHistoryDeleteAll"); _deleteAll.Click += (s, e) => DeleteAll();
            var close = new Button { Text = T("Close"), AutoSize = true, DialogResult = DialogResult.Cancel };
            buttons.Controls.AddRange(new Control[] { _delete, _deleteAll, close }); CancelButton = close;
            root.Controls.Add(_status, 0, 0); root.Controls.Add(_list, 0, 1); root.Controls.Add(_metadata, 0, 2); root.Controls.Add(_details, 0, 3); root.Controls.Add(buttons, 0, 4);
            Controls.Add(root); LoadHistory();
        }
        private string? SelectedId => _list.SelectedItems.Count == 1 ? _list.SelectedItems[0].Tag as string : null;
        private void LoadHistory()
        {
            _list.Items.Clear(); ClearDetails(); _delete.Enabled = _deleteAll.Enabled = false;
            try
            {
                var records = _store.List(out var unreadable);
                foreach (var record in records)
                {
                    var row = new ListViewItem(record.Host) { Tag = record.SessionId, UseItemStyleForSubItems = false };
                    row.ForeColor = record.Host == "Word" ? Color.FromArgb(24, 90, 189) :
                        record.Host == "Excel" ? Color.FromArgb(16, 124, 65) :
                        record.Host == "PowerPoint" ? Color.FromArgb(180, 71, 35) : _list.ForeColor;
                    row.SubItems.Add(record.CreatedAt.ToLocalTime().ToString("g", CultureInfo.CurrentCulture));
                    row.SubItems.Add(record.DocumentName); row.SubItems.Add(record.Preview); _list.Items.Add(row);
                }
                _status.Text = unreadable > 0 ? T("ChatHistoryUnreadable") : records.Count == 0 ? T("ChatHistoryEmpty") : T("ChatHistoryHint");
                _deleteAll.Enabled = records.Count > 0 || unreadable > 0;
            }
            catch { _status.Text = T("ChatHistoryLoadFailed"); }
        }
        private void ShowSelected()
        {
            ClearDetails(); _delete.Enabled = SelectedId != null;
            if (SelectedId == null) return;
            try
            {
                var record = _store.Read(SelectedId);
                if (record == null) { _status.Text = T("ChatHistoryLoadFailed"); return; }
                _metadata.Text = T("ChatHistoryDocument") + ": " + record.DocumentName + "\r\n" + T("ChatHistoryTime") + ": " + record.CreatedAt;
                var messages = new List<ChatMessage>();
                foreach (var turn in record.Turns)
                {
                    DateTimeOffset.TryParse(turn.Timestamp, CultureInfo.InvariantCulture, DateTimeStyles.None, out var timestamp);
                    messages.Add(new ChatMessage("user", turn.Question, timestamp));
                    messages.Add(new ChatMessage("assistant", turn.Answer, timestamp));
                }
                ShowDetails(messages);
                _details.ScrollToStart();
            }
            catch { _status.Text = T("ChatHistoryLoadFailed"); }
        }
        private void ShowDetails(IReadOnlyList<ChatMessage> messages) => _details.ShowConversation(messages,
            T("ChatExportQuestion"), T("ChatExportAnswer"), "", null, "", "", T("ChatUserAvatar"), T("ChatAiAvatar"), default, _language, false);
        private void ClearDetails() { _metadata.Text = ""; ShowDetails(new ChatMessage[0]); }
        private bool Confirm(string key) => ConfirmDelete != null ? ConfirmDelete(key) :
            MessageBox.Show(this, T(key), Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes;
        private void DeleteSelected()
        {
            var id = SelectedId;
            if (id == null || !Confirm("ChatHistoryDeleteConfirm")) return;
            try { _store.Delete(id); LoadHistory(); }
            catch { _status.Text = T("ChatHistoryDeleteFailed"); }
        }
        private void DeleteAll()
        {
            if (!Confirm("ChatHistoryDeleteAllConfirm")) return;
            try { _store.DeleteAll(); LoadHistory(); }
            catch { LoadHistory(); _status.Text = T("ChatHistoryDeleteFailed"); }
        }
    }
}
