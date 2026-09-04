using OfficeTranslate.Core;
using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace OfficeTranslate.WordAddIn
{
    internal sealed class SettingsForm : Form
    {
        private readonly SettingsStore _store;
        private readonly ComboBox _uiLanguage = new ComboBox();
        private readonly ComboBox _provider = new ComboBox();
        private readonly TextBox _baseUrl = new TextBox();
        private readonly TextBox _apiKey = new TextBox();
        private readonly ComboBox _model = new ComboBox();
        private readonly ComboBox _imageModel = new ComboBox();
        private readonly Button _refreshModels = new Button();
        private readonly NumericUpDown _chunkSize = new NumericUpDown();
        private readonly ComboBox _translationStyle = new ComboBox();
        private readonly TextBox _instructions = new TextBox();
        private readonly TextBox _glossary = new TextBox();
        private readonly Label _status = new Label();
        private readonly CancellationTokenSource _cancellation = new CancellationTokenSource();
        private TranslationSettings _existing = new TranslationSettings();
        private bool _loading;
        private ProviderKind _activeProvider;

        public SettingsForm(SettingsStore store)
        {
            _store = store;
            _existing = _store.Load();
            Text = T("SettingsTitle");
            Width = 680; Height = 760; MinimumSize = new Size(600, 680);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.FromArgb(247, 249, 252);
            BuildUi(); LoadValues();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _cancellation.Cancel(); _cancellation.Dispose(); base.OnFormClosed(e);
        }

        private void BuildUi()
        {
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), ColumnCount = 1, RowCount = 3 };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(new Label { Text = T("Heading"), Font = new Font(Font.FontFamily, 15F, FontStyle.Bold), ForeColor = Color.FromArgb(32, 55, 88), AutoSize = true, Margin = new Padding(0, 0, 0, 14) }, 0, 0);

            var card = new TableLayoutPanel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(18), ColumnCount = 2, RowCount = 11 };
            card.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110)); card.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (var i = 0; i < 8; i++) card.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            card.RowStyles.Add(new RowStyle(SizeType.Percent, 38)); card.RowStyles.Add(new RowStyle(SizeType.Percent, 62)); card.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            _uiLanguage.DropDownStyle = ComboBoxStyle.DropDownList; _uiLanguage.Items.AddRange(UiText.LanguageNames);
            _provider.DropDownStyle = ComboBoxStyle.DropDownList; _provider.Items.AddRange(new object[] { "Ollama API", "OpenAI-compatible API" }); _provider.SelectedIndexChanged += ProviderChanged;
            _apiKey.UseSystemPasswordChar = true; _model.DropDownStyle = ComboBoxStyle.DropDownList; _imageModel.DropDownStyle = ComboBoxStyle.DropDownList;
            _chunkSize.Minimum = 500; _chunkSize.Maximum = 30000; _chunkSize.Increment = 500;
            _translationStyle.DropDownStyle = ComboBoxStyle.DropDownList; _translationStyle.SelectedIndexChanged += TranslationStyleChanged;
            _instructions.Multiline = true; _instructions.ScrollBars = ScrollBars.Vertical;
            _glossary.Multiline = true; _glossary.ScrollBars = ScrollBars.Both; _glossary.AcceptsReturn = true; _glossary.WordWrap = false;

            AddRow(card, 0, T("UiLanguage"), _uiLanguage); AddRow(card, 1, T("Provider"), _provider); AddRow(card, 2, "Base URL", _baseUrl); AddRow(card, 3, "API Key", _apiKey);
            var modelPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true, Margin = Padding.Empty };
            modelPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); modelPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _model.Dock = DockStyle.Fill; StyleButton(_refreshModels, T("Refresh"), false); _refreshModels.Margin = new Padding(8, 0, 0, 0); _refreshModels.Click += RefreshModelsClicked;
            modelPanel.Controls.Add(_model, 0, 0); modelPanel.Controls.Add(_refreshModels, 1, 0); AddRow(card, 4, T("Model"), modelPanel);
            AddRow(card, 5, T("ImageModel"), _imageModel);
            AddRow(card, 6, T("Chunk"), _chunkSize); AddRow(card, 7, T("TranslationStyle"), _translationStyle); AddRow(card, 8, T("Instructions"), _instructions); AddRow(card, 9, T("Glossary"), _glossary);
            _status.Text = T("InitialStatus"); _status.AutoSize = true; _status.ForeColor = Color.FromArgb(90, 104, 122); _status.Margin = new Padding(0, 10, 0, 0);
            card.Controls.Add(_status, 0, 10); card.SetColumnSpan(_status, 2); root.Controls.Add(card, 0, 1);

            var buttonBar = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 3, Padding = new Padding(0, 14, 0, 0) };
            buttonBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); buttonBar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); buttonBar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var about = new Button(); StyleButton(about, T("About"), false); about.Click += (s, e) => { var code = _uiLanguage.SelectedIndex >= 0 ? UiText.LanguageCodes[_uiLanguage.SelectedIndex] : _existing.UiLanguage; using (var form = new AboutForm(code)) form.ShowDialog(this); };
            buttonBar.Controls.Add(about, 0, 0);
            var save = new Button(); StyleButton(save, T("Save"), true); save.Click += SaveClicked;
            var cancel = new Button { DialogResult = DialogResult.Cancel }; StyleButton(cancel, T("CancelButton"), false);
            cancel.Margin = new Padding(0, 0, 8, 0); save.Margin = Padding.Empty;
            buttonBar.Controls.Add(cancel, 1, 0); buttonBar.Controls.Add(save, 2, 0); root.Controls.Add(buttonBar, 0, 2);
            Controls.Add(root); AcceptButton = save; CancelButton = cancel;
        }

        private static void StyleButton(Button button, string text, bool primary)
        {
            button.Text = text; button.AutoSize = true; button.Height = 34; button.Padding = new Padding(12, 4, 12, 4); button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = primary ? 0 : 1; button.BackColor = primary ? Color.FromArgb(45, 108, 223) : Color.White;
            button.ForeColor = primary ? Color.White : Color.FromArgb(42, 58, 78); button.FlatAppearance.BorderColor = Color.FromArgb(190, 200, 214);
        }

        private static void AddRow(TableLayoutPanel layout, int row, string label, Control control)
        {
            control.Dock = DockStyle.Fill; control.Margin = new Padding(0, 4, 0, 8);
            layout.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = Color.FromArgb(55, 70, 90) }, 0, row); layout.Controls.Add(control, 1, row);
        }

        private void LoadValues()
        {
            _loading = true;
            var uiIndex = Array.IndexOf(UiText.LanguageCodes, _existing.UiLanguage); _uiLanguage.SelectedIndex = uiIndex >= 0 ? uiIndex : 0;
            _activeProvider = _existing.Provider; _provider.SelectedIndex = _activeProvider == ProviderKind.Ollama ? 0 : 1; ApplyProviderProfile(_activeProvider);
            SetImageModels(new[] { _existing.ImageModel }, _existing.ImageModel);
            _chunkSize.Value = Math.Max(_chunkSize.Minimum, Math.Min(_chunkSize.Maximum, _existing.MaxCharactersPerChunk));
            var styleId = _existing.TranslationStyle;
            if (!string.IsNullOrWhiteSpace(_existing.CustomInstructions) && string.IsNullOrWhiteSpace(styleId)) styleId = "Custom";
            if (Array.IndexOf(TranslationStyleCatalog.Ids, styleId) < 0) styleId = "Custom";
            _translationStyle.Items.Clear();
            foreach (var id in TranslationStyleCatalog.Ids) _translationStyle.Items.Add(UiText.TranslationStyle(_existing.UiLanguage, id));
            _translationStyle.SelectedIndex = Math.Max(0, Array.IndexOf(TranslationStyleCatalog.Ids, styleId));
            _instructions.Text = string.IsNullOrWhiteSpace(_existing.CustomInstructions) && styleId != "Custom" ? TranslationStyleCatalog.GetPrompt(styleId) : _existing.CustomInstructions;
            _glossary.Text = _existing.Glossary; _loading = false;
        }

        private void TranslationStyleChanged(object sender, EventArgs e)
        {
            if (_loading || _translationStyle.SelectedIndex < 0) return;
            var styleId = TranslationStyleCatalog.Ids[_translationStyle.SelectedIndex];
            if (styleId != "Custom") _instructions.Text = TranslationStyleCatalog.GetPrompt(styleId);
        }

        private void ProviderChanged(object sender, EventArgs e)
        {
            if (_loading) return;
            CaptureProviderProfile(_activeProvider);
            _activeProvider = _provider.SelectedIndex == 0 ? ProviderKind.Ollama : ProviderKind.OpenAiCompatible;
            ApplyProviderProfile(_activeProvider); _status.Text = T("ProviderChanged");
        }

        private void CaptureProviderProfile(ProviderKind provider)
        {
            var model = _model.SelectedItem?.ToString() ?? string.Empty;
            if (provider == ProviderKind.Ollama)
            {
                _existing.OllamaBaseUrl = _baseUrl.Text.Trim(); _existing.OllamaApiKey = _apiKey.Text;
                if (!string.IsNullOrWhiteSpace(model)) _existing.OllamaModel = model;
            }
            else
            {
                _existing.OpenAiBaseUrl = _baseUrl.Text.Trim(); _existing.OpenAiApiKey = _apiKey.Text;
                if (!string.IsNullOrWhiteSpace(model)) _existing.OpenAiModel = model;
            }
        }

        private void ApplyProviderProfile(ProviderKind provider)
        {
            var ollama = provider == ProviderKind.Ollama;
            _baseUrl.Text = ollama ? _existing.OllamaBaseUrl : _existing.OpenAiBaseUrl;
            _apiKey.Text = ollama ? _existing.OllamaApiKey : _existing.OpenAiApiKey;
            var model = ollama ? _existing.OllamaModel : _existing.OpenAiModel;
            SetModels(new[] { model }, model);
        }

        private async void RefreshModelsClicked(object sender, EventArgs e)
        {
            _refreshModels.Enabled = false; _status.ForeColor = Color.FromArgb(45, 108, 223); _status.Text = T("LoadingModels");
            try
            {
                using (var client = new TranslationClient())
                {
                    var models = await client.GetModelsAsync(BuildSettings(false), _cancellation.Token);
                    if (models.Count == 0) throw new InvalidOperationException(T("NoModels"));
                    var selected = _activeProvider == ProviderKind.Ollama ? _existing.OllamaModel : _existing.OpenAiModel;
                    SetModels(models, selected); _status.ForeColor = Color.FromArgb(30, 130, 76); _status.Text = string.Format(T("ModelsLoaded"), models.Count);
                    SetImageModels(models, _existing.ImageModel);
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) { _status.ForeColor = Color.FromArgb(190, 45, 45); _status.Text = "× " + ex.Message; }
            finally { if (!IsDisposed) _refreshModels.Enabled = true; }
        }

        private void SetModels(System.Collections.Generic.IEnumerable<string> models, string selected)
        {
            _model.BeginUpdate(); _model.Items.Clear(); foreach (var model in models) if (!string.IsNullOrWhiteSpace(model)) _model.Items.Add(model); _model.EndUpdate();
            var index = _model.Items.IndexOf(selected); _model.SelectedIndex = index >= 0 ? index : (_model.Items.Count > 0 ? 0 : -1);
        }

        private void SetImageModels(System.Collections.Generic.IEnumerable<string> models, string selected)
        {
            _imageModel.BeginUpdate(); _imageModel.Items.Clear();
            _imageModel.Items.Add("（使用文本模型）");
            foreach (var model in models) if (!string.IsNullOrWhiteSpace(model)) _imageModel.Items.Add(model);
            _imageModel.EndUpdate();
            var index = string.IsNullOrWhiteSpace(selected) ? 0 : _imageModel.Items.IndexOf(selected);
            _imageModel.SelectedIndex = index >= 0 ? index : 0;
        }

        private TranslationSettings BuildSettings(bool requireModel)
        {
            CaptureProviderProfile(_activeProvider);
            var settings = new TranslationSettings {
                Provider = _provider.SelectedIndex == 0 ? ProviderKind.Ollama : ProviderKind.OpenAiCompatible,
                BaseUrl = _baseUrl.Text.Trim(), ApiKey = _apiKey.Text, Model = _model.SelectedItem?.ToString() ?? (requireModel ? string.Empty : "model-discovery"),
                SourceLanguage = _existing.SourceLanguage, TargetLanguage = _existing.TargetLanguage,
                UiLanguage = _uiLanguage.SelectedIndex >= 0 ? UiText.LanguageCodes[_uiLanguage.SelectedIndex] : _existing.UiLanguage,
                BilingualMode = _existing.BilingualMode, MaxCharactersPerChunk = (int)_chunkSize.Value, TimeoutSeconds = _existing.TimeoutSeconds,
                TranslationStyle = _translationStyle.SelectedIndex >= 0 ? TranslationStyleCatalog.Ids[_translationStyle.SelectedIndex] : "Custom",
                CustomInstructions = _instructions.Text.Trim(), Glossary = _glossary.Text.Trim(),
                OllamaBaseUrl = _existing.OllamaBaseUrl, OllamaApiKey = _existing.OllamaApiKey, OllamaModel = _existing.OllamaModel,
                OpenAiBaseUrl = _existing.OpenAiBaseUrl, OpenAiApiKey = _existing.OpenAiApiKey, OpenAiModel = _existing.OpenAiModel
                , ImageModel = _imageModel.SelectedIndex <= 0 ? string.Empty : _imageModel.SelectedItem?.ToString() ?? string.Empty
            };
            if (requireModel) settings.Validate(); else settings.ValidateEndpoint(); return settings;
        }

        private void SaveClicked(object sender, EventArgs e)
        {
            try { _store.Save(BuildSettings(true)); DialogResult = DialogResult.OK; Close(); }
            catch (Exception ex) { MessageBox.Show(ex.Message, T("SaveError"), MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        private string T(string key) => UiText.Get(_existing.UiLanguage, key);
    }
}
