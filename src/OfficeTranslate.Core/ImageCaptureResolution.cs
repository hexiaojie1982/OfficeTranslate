using System;

namespace OfficeTranslate.Core
{
    public sealed class ImageCaptureResolution
    {
        private ImageCaptureResolution(int width, int height) { Width = width; Height = height; }
        public int Width { get; }
        public int Height { get; }
        public const int TargetLongEdge = 2560;
        public const double MaxDisplayMultiplier = 12;

        // Use available original pixels, never invent detail by upsampling.
        // The display-relative multiplier avoids sending huge data for icons.
        public static ImageCaptureResolution Plan(int sourceWidth, int sourceHeight, float widthPt, float heightPt,
            int targetLongEdge = TargetLongEdge)
        {
            if (sourceWidth <= 0 || sourceHeight <= 0 || targetLongEdge <= 0 ||
                !FinitePositive(widthPt) || !FinitePositive(heightPt)) throw new ArgumentOutOfRangeException();
            var sourceLong = Math.Max(sourceWidth, sourceHeight);
            var displayLong = Math.Max(widthPt, heightPt) * 96.0 / 72.0;
            var scale = Math.Min(1.0, Math.Min(targetLongEdge, displayLong * MaxDisplayMultiplier) / sourceLong);
            scale = Math.Min(scale, 8000.0 / sourceLong);
            scale = Math.Min(scale, Math.Sqrt(32_000_000.0 / ((double)sourceWidth * sourceHeight)));
            return new ImageCaptureResolution(Math.Max(1, (int)Math.Floor(sourceWidth * scale)),
                Math.Max(1, (int)Math.Floor(sourceHeight * scale)));
        }

        public static bool SameAspect(int width, int height, float widthPt, float heightPt)
        {
            if (width <= 0 || height <= 0 || !FinitePositive(widthPt) || !FinitePositive(heightPt)) return false;
            var pixelAspect = (double)width / height;
            var shapeAspect = (double)widthPt / heightPt;
            // Stricter than the planner's 5%: direct original-image extraction
            // must prove the untransformed bitmap maps to the Office rectangle.
            return Math.Abs(pixelAspect - shapeAspect) / Math.Max(pixelAspect, shapeAspect) <= 0.01;
        }

        // Export uses a whole-slide scale, NOT a desired shape pixel width.
        // Keep a uniform scale (including cropped/rotated fallback appearances).
        public static Tuple<int, int> PowerPointExportScale(float widthPt, float heightPt, float slideWidthPt, float slideHeightPt)
        {
            if (!FinitePositive(slideWidthPt) || !FinitePositive(slideHeightPt) ||
                !FinitePositive(widthPt) || !FinitePositive(heightPt)) throw new ArgumentOutOfRangeException();
            var displayLong = Math.Max(widthPt, heightPt) * 96.0 / 72.0;
            var factor = Math.Min(TargetLongEdge / displayLong, MaxDisplayMultiplier);
            // A rotated shape's bounding diagonal is a conservative upper
            // bound on the export extent, whatever the rotation angle.
            factor = Math.Min(factor, 8000 / (Math.Sqrt((double)widthPt * widthPt + (double)heightPt * heightPt) * 96.0 / 72.0));
            return Tuple.Create(Math.Max(1, checked((int)Math.Floor(slideWidthPt * factor))),
                Math.Max(1, checked((int)Math.Floor(slideHeightPt * factor))));
        }
        private static bool FinitePositive(float value) => value > 0 && !float.IsInfinity(value) && !float.IsNaN(value);
    }
}
