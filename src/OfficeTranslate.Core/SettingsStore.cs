using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace OfficeTranslate.Core
{
    public sealed class SettingsStore
    {
        private readonly string _path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OfficeTranslate", "settings.json");
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer();

        public TranslationSettings Load()
        {
            try
            {
                if (!File.Exists(_path)) return new TranslationSettings();
                var raw = File.ReadAllText(_path, Encoding.UTF8);
                var serialized = raw ?? string.Empty;
                if (string.IsNullOrWhiteSpace(serialized)) return new TranslationSettings();
                var settings = _json.Deserialize<TranslationSettings>(serialized) ?? new TranslationSettings();
                // Ribbon language choices and bilingual mode are session-only. Ignore
                // values left behind by older versions of settings.json.
                settings.SourceLanguage = "自动检测";
                settings.TargetLanguage = "简体中文";
                settings.BilingualMode = false;
                if (serialized.IndexOf("\"TranslationStyle\"", StringComparison.OrdinalIgnoreCase) < 0)
                    settings.TranslationStyle = string.IsNullOrWhiteSpace(settings.CustomInstructions) ? "ProfessionalReport" : "Custom";
                if (!string.IsNullOrEmpty(settings.ApiKey)) settings.ApiKey = Unprotect(settings.ApiKey);
                var hasProviderProfiles = serialized.IndexOf("\"OllamaBaseUrl\"", StringComparison.OrdinalIgnoreCase) >= 0;
                if (hasProviderProfiles)
                {
                    if (!string.IsNullOrEmpty(settings.OllamaApiKey)) settings.OllamaApiKey = Unprotect(settings.OllamaApiKey);
                    if (!string.IsNullOrEmpty(settings.OpenAiApiKey)) settings.OpenAiApiKey = Unprotect(settings.OpenAiApiKey);
                }
                else if (settings.Provider == ProviderKind.Ollama)
                {
                    settings.OllamaBaseUrl = settings.BaseUrl; settings.OllamaApiKey = settings.ApiKey; settings.OllamaModel = settings.Model;
                }
                else
                {
                    settings.OpenAiBaseUrl = settings.BaseUrl; settings.OpenAiApiKey = settings.ApiKey; settings.OpenAiModel = settings.Model;
                }
                return settings;
            }
            catch { return new TranslationSettings(); }
        }

        public void Save(TranslationSettings settings)
        {
            settings.Validate();
            var copy = _json.Deserialize<TranslationSettings>(_json.Serialize(settings))!;
            if (!string.IsNullOrEmpty(copy.ApiKey)) copy.ApiKey = Protect(copy.ApiKey);
            if (!string.IsNullOrEmpty(copy.OllamaApiKey)) copy.OllamaApiKey = Protect(copy.OllamaApiKey);
            if (!string.IsNullOrEmpty(copy.OpenAiApiKey)) copy.OpenAiApiKey = Protect(copy.OpenAiApiKey);
            var raw = _json.Serialize(copy);
            Directory.CreateDirectory(Path.GetDirectoryName(_path));
            File.WriteAllText(_path, raw, Encoding.UTF8);
        }

        private static string Protect(string value) => Convert.ToBase64String(
            ProtectedData.Protect(Encoding.UTF8.GetBytes(value), null, DataProtectionScope.CurrentUser));
        private static string Unprotect(string value) => Encoding.UTF8.GetString(
            ProtectedData.Unprotect(Convert.FromBase64String(value), null, DataProtectionScope.CurrentUser));
    }
}
