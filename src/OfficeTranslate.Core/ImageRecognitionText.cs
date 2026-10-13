using System;
using System.Collections.Generic;

namespace OfficeTranslate.Core
{
    // Adapts the existing image parser's transport result for source-only OCR.
    // Never trusts an unsolicited translation. The translation parser stays unchanged.
    public static class ImageRecognitionText
    {
        public static string NoteText(ImageOverlayNoteChunk chunk) =>
            "【图片原文待检查】" + chunk.Text.Substring(ImageOverlayNotes.Header.Length);

        public static ImageRecognitionResult FromImageResponse(ImageRecognitionResult response)
        {
            var regions = new List<ImageTranslationRegion>();
            var stats = new ImageRecognitionStatistics
            {
                RawCount = response.Statistics.RawCount,
                NonDictionaryCount = response.Statistics.NonDictionaryCount,
                EmptyTextCount = response.Statistics.EmptyTextCount,
                InvalidBboxCount = response.Statistics.InvalidBboxCount
            };
            foreach (var item in response.Regions)
            {
                if (string.IsNullOrWhiteSpace(item.Source)) { stats.EmptyTextCount++; continue; }
                var valid = !float.IsNaN(item.X1) && !float.IsNaN(item.Y1) &&
                    !float.IsNaN(item.X2) && !float.IsNaN(item.Y2) &&
                    item.X1 >= 0 && item.Y1 >= 0 && item.X2 <= 1000 && item.Y2 <= 1000 &&
                    item.X2 > item.X1 && item.Y2 > item.Y1;
                var region = new ImageTranslationRegion
                {
                    Source = item.Source,
                    // Existing region/result types are reused; no translation is generated.
                    Translation = string.Empty,
                    X1 = item.X1, Y1 = item.Y1, X2 = item.X2, Y2 = item.Y2,
                    ReviewReason = valid ? string.Empty : "模型坐标损坏或无效；原文保留为待检查旁注，不作定位覆盖。"
                };
                if (!valid) { stats.InvalidBboxDowngradedCount++; stats.ReviewCount++; }
                regions.Add(region);
            }
            stats.RetainedCount = regions.Count;
            return new ImageRecognitionResult(regions.AsReadOnly(), stats)
            { FinishReason = response.FinishReason, Truncated = response.Truncated };
        }
    }

    public static class ImageRecognitionTags
    {
        // Sub-prefixes keep the frozen legacy IsOwnMarker/text collectors excluding
        // these artifacts, while exact owner matching isolates cleanup in both directions.
        public const string OverlayPrefix = "OfficeTranslateOCR:Recognition:";
        public const string NotePrefix = "OfficeTranslateOCR-Note:Recognition:";
        public static string OverlayFor(string ownerId) => OverlayPrefix + ownerId;
        public static string NoteFor(string ownerId) => NotePrefix + ownerId;
        public static bool WordExcelMarkerBelongsTo(string marker, string ownerId) =>
            marker == OverlayFor(ownerId) || marker == NoteFor(ownerId);
        public static string PowerPointOverlayFor(string ownerId) => "Recognition:" + ownerId;
        public static string PowerPointNoteFor(string ownerId) => "Recognition-Note:" + ownerId;
    }
}
