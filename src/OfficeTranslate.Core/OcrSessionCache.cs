using System;
using System.Collections.Concurrent;
using System.Security.Cryptography;

namespace OfficeTranslate.Core
{
    // Process-local OCR content only; parsing and host layout run again on every hit.
    internal static class OcrSessionCache
    {
        private const int MaximumEntries = 2000;
        private static readonly ConcurrentDictionary<string, string> Entries =
            new ConcurrentDictionary<string, string>(StringComparer.Ordinal);
        private static readonly ConcurrentQueue<string> InsertionOrder = new ConcurrentQueue<string>();

        public static string CreateKey(TranslationSettings settings, byte[] imageBytes)
        {
            string imageHash;
            using (var sha = SHA256.Create())
                imageHash = Convert.ToBase64String(sha.ComputeHash(imageBytes));

            // Pure OCR detects all languages; translation language/style settings
            // do not affect its prompt. Keep URL path casing (endpoints can differ).
            var material = string.Join("\u001F", new[]
            {
                settings.Provider.ToString(),
                (settings.BaseUrl ?? string.Empty).TrimEnd('/'),
                string.IsNullOrWhiteSpace(settings.ImageModel) ? settings.Model : settings.ImageModel,
                imageHash
            });
            return TranslationSessionCache.Hash(material);
        }

        public static bool TryGet(string key, out string recognized) => Entries.TryGetValue(key, out recognized);

        public static void Store(string key, string recognized)
        {
            if (!Entries.TryAdd(key, recognized))
            {
                Entries[key] = recognized;
                return;
            }

            InsertionOrder.Enqueue(key);
            while (Entries.Count > MaximumEntries && InsertionOrder.TryDequeue(out var oldest))
                Entries.TryRemove(oldest, out _);
        }
    }
}
