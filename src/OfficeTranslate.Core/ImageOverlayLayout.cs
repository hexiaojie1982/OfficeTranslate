using System;

namespace OfficeTranslate.Core
{
    public sealed class ImageOverlayMetrics
    {
        public ImageOverlayMetrics(float fontSize, float height)
        {
            FontSize = fontSize;
            Height = height;
        }

        public float FontSize { get; }
        public float Height { get; }
    }

    public static class ImageOverlayLayout
    {
        public static ImageOverlayMetrics Calculate(float width, float originalHeight, string text)
        {
            width = Normalize(width, 24F);
            originalHeight = Normalize(originalHeight, 12F);

            // Keep OCR text readable and expand the box vertically instead of
            // shrinking long bilingual content into the original one-line box.
            var fontSize = Math.Max(8F, Math.Min(18F, originalHeight * 0.55F));
            var usableWidth = Math.Max(12F, width - 4F);
            var logicalLines = (text ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace('\r', '\n')
                .Split(new[] { '\n' }, StringSplitOptions.None);

            var visualLineCount = 0;
            foreach (var line in logicalLines)
            {
                var units = MeasureUnits(line);
                visualLineCount += Math.Max(1, (int)Math.Ceiling(units * fontSize / usableWidth));
            }

            var requiredHeight = visualLineCount * fontSize * 1.4F + 4F;
            // A defensive ceiling avoids invalid Office shape dimensions if a
            // vision model unexpectedly returns a very large block of text.
            var height = Math.Min(720F, Math.Max(originalHeight, requiredHeight));
            return new ImageOverlayMetrics(fontSize, height);
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
