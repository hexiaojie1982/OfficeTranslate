using System;

namespace OfficeTranslate.Core
{
    public sealed class ImageOverlayMetrics
    {
        public ImageOverlayMetrics(float fontSize, bool fits)
        {
            FontSize = fontSize;
            Fits = fits;
        }

        public float FontSize { get; }
        public bool Fits { get; }
    }

    public static class ImageOverlayLayout
    {
        public const float MinFontSize = 8F;
        public const float MaxFontSize = 18F;

        private const float LineHeightFactor = 1.3F;
        private const float HorizontalMargin = 4F;
        private const float VerticalMargin = 2F;

        // Finds the largest font size in [MinFontSize, MaxFontSize] such that the
        // text wraps inside the given region instead of growing the box downward.
        // Returns Fits=false when even the minimum size overflows; callers must
        // then degrade to an explicit side-note instead of a fake overlay.
        public static ImageOverlayMetrics FitFontSize(float width, float height, string text)
        {
            width = Normalize(width, 24F);
            height = Normalize(height, 12F);
            // R6: no fictional minimum usable sizes. A region narrower than the
            // margins genuinely cannot hold text and must degrade to SideNote
            // instead of pretending a 12pt-wide area exists.
            var usableWidth = width - HorizontalMargin;
            var usableHeight = height - VerticalMargin;
            if (usableWidth <= 0F || usableHeight <= 0F)
                return new ImageOverlayMetrics(MinFontSize, false);

            var low = MinFontSize;
            var high = Math.Min(MaxFontSize, usableHeight / LineHeightFactor);
            if (high < low || !FitsAt(low, usableWidth, usableHeight, text))
                return new ImageOverlayMetrics(MinFontSize, false);

            // Binary search the largest fitting size; 0.5pt resolution is plenty.
            for (var i = 0; i < 8; i++)
            {
                var mid = (low + high) / 2F;
                if (FitsAt(mid, usableWidth, usableHeight, text)) low = mid;
                else high = mid;
            }
            return new ImageOverlayMetrics(low, true);
        }

        // Height estimator for side-notes, where vertical growth is acceptable
        // because the note explicitly does NOT claim positional coverage.
        // R2: no height cap. A side-note must preserve every translation
        // completely; clipping it would silently lose content.
        public static float EstimateNoteHeight(float width, string text, float fontSize)
        {
            width = Normalize(width, 160F);
            fontSize = Normalize(fontSize, 9F);
            var usableWidth = Math.Max(12F, width - HorizontalMargin);
            var lines = CountVisualLines(text, usableWidth, fontSize);
            return lines * fontSize * LineHeightFactor + 8F;
        }

        private static bool FitsAt(float fontSize, float usableWidth, float usableHeight, string text)
        {
            return CountVisualLines(text, usableWidth, fontSize) * fontSize * LineHeightFactor <= usableHeight;
        }

        private static int CountVisualLines(string text, float usableWidth, float fontSize)
        {
            var logicalLines = (text ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace('\r', '\n')
                .Split(new[] { '\n' }, StringSplitOptions.None);
            var visualLineCount = 0;
            foreach (var line in logicalLines)
                visualLineCount += Math.Max(1, (int)Math.Ceiling(MeasureUnits(line) * fontSize / usableWidth));
            return Math.Max(1, visualLineCount);
        }

        private static float MeasureUnits(string text)
        {
            if (string.IsNullOrEmpty(text)) return 1F;
            var units = 0F;
            foreach (var character in text)
            {
                if (char.IsWhiteSpace(character)) units += 0.35F;
                else if (character <= 0x7F) units += 0.58F;
                else units += 1F;
            }
            return Math.Max(1F, units);
        }

        private static float Normalize(float value, float fallback)
        {
            return float.IsNaN(value) || float.IsInfinity(value) || value <= 0 ? fallback : value;
        }
    }
}
