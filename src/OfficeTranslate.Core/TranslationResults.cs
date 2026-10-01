using System;
using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text;

namespace OfficeTranslate.Core
{
    public sealed class TranslationResult
    {
        internal TranslationResult(string text, bool changed, bool fromCache, bool skippedNoSource, bool needsReview)
        {
            Text = text;
            Changed = changed;
            FromCache = fromCache;
            SkippedNoSource = skippedNoSource;
            NeedsReview = needsReview;
        }

        public string Text { get; }
        public bool Changed { get; }
        public bool FromCache { get; }
        public bool SkippedNoSource { get; }
        public bool NeedsReview { get; }

        internal static TranslationResult Translated(string text, bool needsReview = false) =>
            new TranslationResult(text, true, false, false, needsReview);

        internal static TranslationResult Cached(string text) =>
            new TranslationResult(text, true, true, false, false);

        internal static TranslationResult Skipped(string text) =>
            new TranslationResult(text, false, false, true, false);

        internal static TranslationResult Unchanged(string text) =>
            new TranslationResult(text, false, false, false, true);
    }

    public sealed class TranslationTaskSummary
    {
        public TranslationTaskSummary(int total) => Total = total;

        public int Total { get; }
        public int Translated { get; private set; }
        public int CacheHits { get; private set; }
        public int SkippedNoSource { get; private set; }
        public int NeedsReview { get; private set; }
        public int OcrNoText { get; private set; }

        public void Record(TranslationResult result)
        {
            if (result.FromCache) CacheHits++;
            else if (result.SkippedNoSource) SkippedNoSource++;
            else if (result.NeedsReview) NeedsReview++;
            else if (result.Changed) Translated++;
            else NeedsReview++;
        }

        public void RecordSkipped() => SkippedNoSource++;

        public void RecordImage(bool recognized)
        {
            if (recognized) Translated++;
            else OcrNoText++;
        }

        public void RecordImageNeedsReview() => NeedsReview++;
    }

    internal static class TranslationSessionCache
    {
        private const int MaximumEntries = 2000;
        private static readonly ConcurrentDictionary<string, string> Entries =
            new ConcurrentDictionary<string, string>(StringComparer.Ordinal);
        private static readonly ConcurrentQueue<string> InsertionOrder = new ConcurrentQueue<string>();

        public static string CreateKey(TranslationSettings settings, string text)
        {
            var material = string.Join("\u001F", new[]
            {
                settings.Provider.ToString(),
                (settings.BaseUrl ?? string.Empty).TrimEnd('/').ToLowerInvariant(),
                settings.Model ?? string.Empty,
                settings.SourceLanguage ?? string.Empty,
                settings.TargetLanguage ?? string.Empty,
                settings.TranslationStyle ?? string.Empty,
                settings.CustomInstructions ?? string.Empty,
                settings.Glossary ?? string.Empty,
                settings.MaxCharactersPerChunk.ToString(),
                text ?? string.Empty
            });
            return Hash(material);
        }

        public static bool TryGet(string key, out string translated) => Entries.TryGetValue(key, out translated);

        public static void Store(string key, string translated)
        {
            if (!Entries.TryAdd(key, translated))
            {
                Entries[key] = translated;
                return;
            }

            InsertionOrder.Enqueue(key);
            while (Entries.Count > MaximumEntries && InsertionOrder.TryDequeue(out var oldest))
                Entries.TryRemove(oldest, out _);
        }

        public static string Hash(string value)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty));
                return Convert.ToBase64String(bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_');
            }
        }
    }
}
