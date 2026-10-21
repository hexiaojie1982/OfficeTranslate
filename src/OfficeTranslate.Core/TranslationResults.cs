using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace OfficeTranslate.Core
{
    public sealed class TranslationResult
    {
        internal TranslationResult(string text, bool changed, bool fromCache, bool skippedNoSource, bool needsReview,
            int fallbackTranslatedParagraphs = 0, int fallbackReviewParagraphs = 0)
        {
            Text = text;
            Changed = changed;
            FromCache = fromCache;
            SkippedNoSource = skippedNoSource;
            NeedsReview = needsReview;
            FallbackTranslatedParagraphs = fallbackTranslatedParagraphs;
            FallbackReviewParagraphs = fallbackReviewParagraphs;
        }

        public string Text { get; }
        public bool Changed { get; }
        public bool FromCache { get; }
        public bool SkippedNoSource { get; }
        public bool NeedsReview { get; }
        public int FallbackTranslatedParagraphs { get; }
        public int FallbackReviewParagraphs { get; }

        internal static TranslationResult Translated(string text, bool needsReview = false,
            int fallbackTranslatedParagraphs = 0, int fallbackReviewParagraphs = 0) =>
            new TranslationResult(text, true, false, false, needsReview, fallbackTranslatedParagraphs, fallbackReviewParagraphs);

        internal static TranslationResult Cached(string text) =>
            new TranslationResult(text, true, true, false, false);

        internal static TranslationResult Skipped(string text) =>
            new TranslationResult(text, false, false, true, false);

        internal static TranslationResult Unchanged(string text, int fallbackReviewParagraphs = 0) =>
            new TranslationResult(text, false, false, false, true, 0, fallbackReviewParagraphs);
    }

    public sealed class TextBlockFailure
    {
        internal TextBlockFailure(int blockNumber, bool duringWrite, Exception error)
        { BlockNumber = blockNumber; DuringWrite = duringWrite; ErrorType = error.GetType().Name; }
        public int BlockNumber { get; }
        public bool DuringWrite { get; }
        // Exception messages can contain provider echoes of document text or secrets.
        public string ErrorType { get; }
    }

    public sealed class TranslationTaskSummary
    {
        private readonly object _recordLock = new object();
        private readonly List<string> _unplacedImageNotes = new List<string>();
        private readonly List<TextBlockFailure> _textBlockFailures = new List<TextBlockFailure>();
        public IReadOnlyList<TextBlockFailure> TextBlockFailures => _textBlockFailures.AsReadOnly();
        public void RecordTextFailure(int blockNumber, bool duringWrite, Exception error)
        {
            lock (_recordLock)
            {
                _textBlockFailures.Add(new TextBlockFailure(blockNumber, duringWrite, error));
                NeedsReview++; TextNeedsReview++;
            }
        }
        public TranslationTaskSummary(int total) => Total = total;

        public int Total { get; }
        public int Translated { get; private set; }
        public int CacheHits { get; private set; }
        public int SkippedNoSource { get; private set; }
        public int NeedsReview { get; private set; }
        public int OcrNoText { get; private set; }
        public int TextNeedsReview { get; private set; }
        public int TextUnchangedNeedsReview { get; private set; }
        public int TextPartiallyTranslated { get; private set; }
        public int FallbackTranslatedParagraphs { get; private set; }
        public int FallbackReviewParagraphs { get; private set; }
        public int ImageNeedsReview { get; private set; }
        public int OcrImages { get; private set; }
        public int OcrRawRegions { get; private set; }
        public int OcrKeptRegions { get; private set; }
        public int OcrDroppedRegions { get; private set; }
        public int OcrRegionLimitDropped { get; private set; }
        public int OcrBboxNotes { get; private set; }
        public int OcrEmptyTranslations { get; private set; }
        public IReadOnlyList<string> UnplacedImageNotes => _unplacedImageNotes.AsReadOnly();

        public void Record(TranslationResult result)
        {
            // Text workers may complete concurrently. Image recording stays
            // sequential and starts only after all text workers have drained.
            lock (_recordLock)
            {
                FallbackTranslatedParagraphs += result.FallbackTranslatedParagraphs;
                FallbackReviewParagraphs += result.FallbackReviewParagraphs;
                if (result.FromCache) CacheHits++;
                else if (result.SkippedNoSource) SkippedNoSource++;
                else if (result.NeedsReview)
                {
                    NeedsReview++; TextNeedsReview++;
                    if (result.Changed) { Translated++; TextPartiallyTranslated++; }
                    else TextUnchangedNeedsReview++;
                }
                else if (result.Changed) Translated++;
                else { NeedsReview++; TextNeedsReview++; TextUnchangedNeedsReview++; }
            }
        }

        public void RecordSkipped() => SkippedNoSource++;

        public void RecordImage(bool recognized)
        {
            if (recognized) Translated++;
            else OcrNoText++;
        }

        public void RecordImageNeedsReview() { NeedsReview++; ImageNeedsReview++; }

        public void RecordImageRecognition(ImageRecognitionResult result)
        {
            OcrImages++;
            OcrRawRegions += result.Statistics.RawCount;
            OcrKeptRegions += result.Statistics.RetainedCount;
            OcrDroppedRegions += result.Statistics.DroppedCount;
            OcrRegionLimitDropped += result.Statistics.RegionLimitDroppedCount;
            OcrBboxNotes += result.Statistics.InvalidBboxDowngradedCount;
            OcrEmptyTranslations += result.Statistics.EmptyTranslationCount;
            if (result.RegionsTruncated) RecordImageNeedsReview();
            if (result.Regions.Count == 0) OcrNoText++;
            else if (result.Regions.Any(r => !string.IsNullOrWhiteSpace(r.Translation))) Translated++;
            // A source-only result is NOT a translated image or empty OCR.
            // Its forced side-note is counted by the normal host review path.
        }

        // Complete, recoverable text in the result window, never OCR text in
        // metadata-only diagnostic logs or silently truncated content.
        public void RecordImageNoteFailure(string ownerId, string noteText, string reason) =>
            _unplacedImageNotes.Add("图片 " + ownerId + "\r\n" + reason + "\r\n" + noteText);
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
