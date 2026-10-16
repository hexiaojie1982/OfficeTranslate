using System;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using OfficeTranslate.Core;

namespace OfficeTranslate.WordAddIn
{
    internal sealed class ChatHistoryForm : Form
    {
        private readonly ChatHistoryStore _store;
        private readonly string _language;
        private readonly ListView _list = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, MultiSelect = false, HideSelection = false };
        private readonly TextBox _details = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = true, MaxLength = 0, BackColor = Color.White };
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
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 4 };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 38));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 62)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _list.Columns.Add(T("ChatHistoryTime"), 180); _list.Columns.Add(T("ChatHistoryDocument"), 200); _list.Columns.Add(T("ChatHistoryQuestion"), 440);
            _list.AccessibleName = T("ChatHistoryTitle"); _details.AccessibleName = T("ChatHistoryDetails");
            _list.SelectedIndexChanged += (s, e) => ShowSelected();
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            _delete.Text = T("ChatHistoryDelete"); _delete.Click += (s, e) => DeleteSelected();
            _deleteAll.Text = T("ChatHistoryDeleteAll"); _deleteAll.Click += (s, e) => DeleteAll();
            var close = new Button { Text = T("Close"), AutoSize = true, DialogResult = DialogResult.Cancel };
            buttons.Controls.AddRange(new Control[] { _delete, _deleteAll, close }); CancelButton = close;
            root.Controls.Add(_status, 0, 0); root.Controls.Add(_list, 0, 1); root.Controls.Add(_details, 0, 2); root.Controls.Add(buttons, 0, 3);
            Controls.Add(root); LoadHistory();
        }
        private string? SelectedId => _list.SelectedItems.Count == 1 ? _list.SelectedItems[0].Tag as string : null;
        private void LoadHistory()
        {
            _list.Items.Clear(); _details.Clear(); _delete.Enabled = _deleteAll.Enabled = false;
            try
            {
                var records = _store.List(out var unreadable);
                foreach (var record in records)
                {
                    var row = new ListViewItem(record.CreatedAt.ToLocalTime().ToString("g", CultureInfo.CurrentCulture)) { Tag = record.SessionId };
                    row.SubItems.Add(record.DocumentName); row.SubItems.Add(record.Preview); _list.Items.Add(row);
                }
                _status.Text = unreadable > 0 ? T("ChatHistoryUnreadable") : records.Count == 0 ? T("ChatHistoryEmpty") : T("ChatHistoryHint");
                _deleteAll.Enabled = records.Count > 0 || unreadable > 0;
            }
            catch { _status.Text = T("ChatHistoryLoadFailed"); }
        }
        private void ShowSelected()
        {
            _details.Clear(); _delete.Enabled = SelectedId != null;
            if (SelectedId == null) return;
            try
            {
                var record = _store.Read(SelectedId);
                if (record == null) { _status.Text = T("ChatHistoryLoadFailed"); return; }
                var text = new StringBuilder();
                text.Append(T("ChatHistoryDocument")).Append(": ").AppendLine(record.DocumentName);
                text.Append(T("ChatHistoryTime")).Append(": ").AppendLine(record.CreatedAt).AppendLine();
                foreach (var turn in record.Turns)
                {
                    text.Append(T("ChatExportTimestamp")).Append(": ").AppendLine(turn.Timestamp);
                    text.AppendLine(T("ChatExportQuestion")).AppendLine(turn.Question).AppendLine();
                    text.AppendLine(T("ChatExportAnswer")).AppendLine(turn.Answer).AppendLine();
                }
                // WinForms EDIT controls require CRLF for multiline display. Only
                // normalize this view; JSON keeps every original source character.
                _details.Text = text.ToString().Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\r\n");
                _details.SelectionStart = 0; _details.SelectionLength = 0;
            }
            catch { _status.Text = T("ChatHistoryLoadFailed"); }
        }
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
