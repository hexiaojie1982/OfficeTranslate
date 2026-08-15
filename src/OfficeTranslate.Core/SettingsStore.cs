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
                var settings = _json.Deserialize<TranslationSettings>(raw) ?? new TranslationSettings();
                if (raw.IndexOf("\"TranslationStyle\"", StringComparison.OrdinalIgnoreCase) < 0)
                    settings.TranslationStyle = string.IsNullOrWhiteSpace(settings.CustomInstructions) ? "ProfessionalReport" : "Custom";
                if (!string.IsNullOrEmpty(settings.ApiKey)) settings.ApiKey = Unprotect(settings.ApiKey);
                return settings;
            }
            catch { return new TranslationSettings(); }
        }

        public void Save(TranslationSettings settings)
        {
            settings.Validate();
            var copy = _json.Deserialize<TranslationSettings>(_json.Serialize(settings))!;
            if (!string.IsNullOrEmpty(copy.ApiKey)) copy.ApiKey = Protect(copy.ApiKey);
            Directory.CreateDirectory(Path.GetDirectoryName(_path));
            File.WriteAllText(_path, _json.Serialize(copy), Encoding.UTF8);
        }

        private static string Protect(string value) => Convert.ToBase64String(
            ProtectedData.Protect(Encoding.UTF8.GetBytes(value), null, DataProtectionScope.CurrentUser));
        private static string Unprotect(string value) => Encoding.UTF8.GetString(
            ProtectedData.Unprotect(Convert.FromBase64String(value), null, DataProtectionScope.CurrentUser));
    }
}
