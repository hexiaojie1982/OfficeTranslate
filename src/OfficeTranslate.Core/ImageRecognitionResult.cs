using System;
using System.Collections.Generic;
using System.Globalization;
using System.Web.Script.Serialization;

namespace OfficeTranslate.Core
{
    public sealed class ImageRecognitionStatistics
    {
        public int RawCount { get; internal set; }
        public int RetainedCount { get; internal set; }
        public int InvalidBboxCount { get; internal set; }
        public int InvalidBboxDowngradedCount { get; internal set; }
        public int EmptyTranslationCount { get; internal set; }
        public int EmptyTranslationDroppedCount { get; internal set; }
        public int NonDictionaryCount { get; internal set; }
        public int EmptyTextCount { get; internal set; }
        public int RegionLimitDroppedCount { get; internal set; }
        public int DroppedCount => NonDictionaryCount + EmptyTextCount + RegionLimitDroppedCount;
        public int ReviewCount { get; internal set; }
    }

    public sealed class ImageRecognitionResult
    {
        internal ImageRecognitionResult(IReadOnlyList<ImageTranslationRegion> regions, ImageRecognitionStatistics statistics)
        { Regions = regions; Statistics = statistics; }
        public IReadOnlyList<ImageTranslationRegion> Regions { get; }
        public ImageRecognitionStatistics Statistics { get; }
        public string FinishReason { get; internal set; } = "unknown";
        public bool Truncated { get; internal set; }
        public bool RegionsTruncated => Statistics.RegionLimitDroppedCount > 0;
    }

    // Host-independent parsing; counts are returned, including empty results.
    public static class ImageRegionParser
    {
        public const int MaximumRegions = 512;
        public static ImageRecognitionResult Parse(string content)
        {
            if (content.Length > ResponseLimits.Image) throw new InvalidOperationException("图片区域 JSON 超过 16 MiB 字符上限。");
            var start = content.IndexOf('{'); var end = content.LastIndexOf('}');
            if (start < 0 || end <= start) throw new InvalidOperationException("图片模型没有返回有效的 JSON 坐标。");
            var root = new JavaScriptSerializer { MaxJsonLength = ResponseLimits.Image }
                .DeserializeObject(content.Substring(start, end - start + 1)) as Dictionary<string, object>;
            if (root == null || !root.TryGetValue("regions", out var value) || !(value is object[] items))
                throw new InvalidOperationException("图片模型返回的 regions 格式无效。");
            var stats = new ImageRecognitionStatistics { RawCount = items.Length,
                RegionLimitDroppedCount = Math.Max(0, items.Length - MaximumRegions) };
            var regions = new List<ImageTranslationRegion>();
            for (var i = 0; i < Math.Min(items.Length, MaximumRegions); i++)
            {
                var item = items[i];
                if (!(item is Dictionary<string, object> dict)) { stats.NonDictionaryCount++; continue; }
                var region = new ImageTranslationRegion { Source = Text(dict, "source"), Translation = Text(dict, "translation") };
                var boxOk = dict.TryGetValue("bbox", out var b) && b is object[] box && box.Length == 4;
                var coords = boxOk ? (object[])b! : Array.Empty<object>();
                region.X1 = boxOk ? Coordinate(coords[0]) : float.NaN;
                region.Y1 = boxOk ? Coordinate(coords[1]) : float.NaN;
                region.X2 = boxOk ? Coordinate(coords[2]) : float.NaN;
                region.Y2 = boxOk ? Coordinate(coords[3]) : float.NaN;
                boxOk = boxOk && region.X1 >= 0 && region.Y1 >= 0 && region.X2 <= 1000 && region.Y2 <= 1000
                    && region.X2 > region.X1 && region.Y2 > region.Y1;
                if (!boxOk) stats.InvalidBboxCount++;
                var emptyTranslation = string.IsNullOrWhiteSpace(region.Translation);
                if (emptyTranslation) stats.EmptyTranslationCount++;
                if (emptyTranslation && string.IsNullOrWhiteSpace(region.Source))
                { stats.EmptyTranslationDroppedCount++; stats.EmptyTextCount++; continue; }
                if (!boxOk)
                {
                    stats.InvalidBboxDowngradedCount++;
                    region.ReviewReason = "模型坐标损坏或无效；文字保留为待检查旁注，不作定位覆盖。";
                }
                if (emptyTranslation)
                    region.ReviewReason += (region.ReviewReason.Length > 0 ? " " : "") + "模型已识别原文，但未提供译文。";
                if (region.RequiresReview) stats.ReviewCount++;
                regions.Add(region);
            }
            stats.RetainedCount = regions.Count;
            return new ImageRecognitionResult(regions.AsReadOnly(), stats);
        }

        private static string Text(Dictionary<string, object> dict, string key) =>
            dict.TryGetValue(key, out var value) && value is string text ? text : string.Empty;

        private static float Coordinate(object value)
        {
            if (value == null || value is bool) return float.NaN;
            try
            {
                var f = Convert.ToSingle(value, CultureInfo.InvariantCulture);
                return float.IsNaN(f) || float.IsInfinity(f) ? float.NaN : f;
            }
            catch { return float.NaN; }
        }
    }
}
