using System;

namespace OfficeTranslate.Core
{
    public enum ImageOverlayVerdict
    {
        // The region geometry is trustworthy: draw the planned overlay here.
        Place,
        // Geometry is not trustworthy (bad bbox, aspect mismatch, rotated image,
        // or text that cannot fit). Do NOT draw a positioned overlay; the caller
        // must surface an explicit side-note instead.
        SideNote
    }

    public sealed class PlannedOverlay
    {
        public PlannedOverlay(ImageOverlayVerdict verdict, string reason,
            float left, float top, float width, float height, float fontSize,
            float widthGrowthFactor = 1F, float heightGrowthFactor = 1F)
        {
            Verdict = verdict;
            Reason = reason ?? string.Empty;
            Left = left;
            Top = top;
            Width = width;
            Height = height;
            FontSize = fontSize;
            WidthGrowthFactor = verdict == ImageOverlayVerdict.SideNote ? 0F : widthGrowthFactor;
            HeightGrowthFactor = verdict == ImageOverlayVerdict.SideNote ? 0F : heightGrowthFactor;
        }

        public ImageOverlayVerdict Verdict { get; }
        public string Reason { get; }

        // Offsets/sizes in Office points, relative to the source image's own
        // origin. The caller adds the image's Left/Top in whichever reference
        // frame the host uses, so no cross-frame conversion happens here.
        // Starts at the bbox-mapped size. If needed, grows around its center
        // independently up to 5x per dimension, inside the image, minimizing area.
        public float Left { get; }
        public float Top { get; }
        public float Width { get; }
        public float Height { get; }
        public float FontSize { get; }
        public float WidthGrowthFactor { get; }
        public float HeightGrowthFactor { get; }
        public float GrowthFactor => Math.Max(WidthGrowthFactor, HeightGrowthFactor);
    }

    // Pure, host-independent geometry: validates the model's 0-1000 bbox,
    // checks that the captured PNG actually maps onto the Office shape rect,
    // and fits the text inside the region. No COM, no network, fully unit
    // testable. Callers must still log the returned geometry via
    // ImageOverlayDiagnostics so model error can be told apart from
    // coordinate-conversion error on real machines.
    //
    // Coordinate contract (R3, review 2026-09-30): the bbox MUST be in
    // "final rendered PNG" coordinates, i.e. where the text appears in the
    // PNG the model actually saw. All three hosts capture the rendered
    // appearance (PowerPoint Export, Word CopyAsPicture, Excel CopyPicture):
    // a flipped picture exports with its text already flipped, so the
    // model's bbox already matches the display coordinates and the planner
    // performs NO flip compensation. If a future caller feeds an
    // untransformed bitmap, it must convert the bbox itself before calling.
    public static class ImageOverlayPlanner
    {
        private const float BboxScale = 1000F;
        private const float MinBboxUnits = 3F;
        private const float MaxGrowthFactor = 5F;
        private const float MaxAspectDeviation = 0.05F;
        private const float MaxRotationDegrees = 0.5F;
        // Tolerance for float rounding when validating the final rect against
        // the image rect. Anything beyond this is a real out-of-bounds error.
        private const float BoundsEpsilon = 0.05F;

        public static PlannedOverlay Plan(
            float x1, float y1, float x2, float y2,
            int pixelWidth, int pixelHeight,
            float shapeWidth, float shapeHeight,
            string text,
            float rotationDegrees)
        {
            if (float.IsNaN(x1) || float.IsNaN(y1) || float.IsNaN(x2) || float.IsNaN(y2) ||
                float.IsInfinity(x1) || float.IsInfinity(y1) || float.IsInfinity(x2) || float.IsInfinity(y2) ||
                !(x1 >= 0F && x2 <= BboxScale && y1 >= 0F && y2 <= BboxScale))
                return SideNote("模型返回的坐标超出 0-1000 范围或为非有限值。");
            if (!(x2 > x1 && y2 > y1))
                return SideNote("模型返回的坐标框无效（右下角不在右上角的右下方）。");
            if (x2 - x1 < MinBboxUnits || y2 - y1 < MinBboxUnits)
                return SideNote("模型返回的坐标框过小，不可信。");
            // R7: explicit finite checks. NaN fails every comparison, so the
            // old `shapeWidth <= 0` style checks silently let NaN through and
            // produced Place with NaN geometry.
            if (pixelWidth <= 0 || pixelHeight <= 0 ||
                !IsValidDimension(shapeWidth) || !IsValidDimension(shapeHeight))
                return SideNote("图片尺寸无效，无法换算坐标。");
            if (float.IsNaN(rotationDegrees) || float.IsInfinity(rotationDegrees) ||
                Math.Abs(rotationDegrees) > MaxRotationDegrees)
                return SideNote("图片存在旋转或旋转角度无效，坐标无法可靠换算。");

            // The captured PNG must be the same picture as the Office shape.
            // A skewed aspect ratio means DPI/padding anomalies or a stale
            // capture, and any linear mapping would misplace the overlay.
            var pixelAspect = (float)pixelWidth / pixelHeight;
            var shapeAspect = shapeWidth / shapeHeight;
            var deviation = Math.Abs(pixelAspect - shapeAspect) / Math.Max(pixelAspect, shapeAspect);
            if (deviation > MaxAspectDeviation)
                return SideNote("捕获图像长宽比与 Office 图片不一致（偏差 " +
                    (deviation * 100F).ToString("0.#") + "%），坐标无法可靠换算。");

            // Validate the exact starting rectangle before any growth. Invalid
            // geometry must never be rescued by the expansion path.
            var left = shapeWidth * x1 / BboxScale;
            var top = shapeHeight * y1 / BboxScale;
            var width = shapeWidth * (x2 - x1) / BboxScale;
            var height = shapeHeight * (y2 - y1) / BboxScale;

            if (float.IsNaN(left) || float.IsNaN(top) || float.IsNaN(width) || float.IsNaN(height) ||
                float.IsInfinity(left) || float.IsInfinity(top) || float.IsInfinity(width) || float.IsInfinity(height))
                return SideNote("坐标换算结果为非有限值，无法定位。");
            if (left < -BoundsEpsilon || top < -BoundsEpsilon ||
                left + width > shapeWidth + BoundsEpsilon ||
                top + height > shapeHeight + BoundsEpsilon)
                return SideNote("换算后的文本框超出图片边界，无法定位。");

            var fit = ImageOverlayLayout.FitFontSize(width, height, text);
            if (!fit.Fits)
            {
                var maxWidth = Math.Min(shapeWidth, width * MaxGrowthFactor);
                var maxHeight = Math.Min(shapeHeight, height * MaxGrowthFactor);
                if (!ImageOverlayLayout.FitFontSize(maxWidth, maxHeight, text).Fits)
                    return SideNote("译文在放大至上限的区域内仍放不下（已尝试最小字号），改用旁注。");

                FindMinimumArea(width, height, maxWidth, maxHeight, text, out var grownWidth, out var grownHeight);
                // Translate at the edges instead of clipping away the fitting
                // dimensions. The whole frame must stay inside the image.
                left = Math.Max(0F, Math.Min(shapeWidth - grownWidth, left + width / 2F - grownWidth / 2F));
                top = Math.Max(0F, Math.Min(shapeHeight - grownHeight, top + height / 2F - grownHeight / 2F));
                var widthGrowth = grownWidth / width;
                var heightGrowth = grownHeight / height;
                width = grownWidth;
                height = grownHeight;
                fit = ImageOverlayLayout.FitFontSize(width, height, text);
                if (!fit.Fits || left < -BoundsEpsilon || top < -BoundsEpsilon ||
                    left + width > shapeWidth + BoundsEpsilon || top + height > shapeHeight + BoundsEpsilon)
                    return SideNote("译文在放大并钳制后的区域内仍放不下，改用旁注。");
                return new PlannedOverlay(ImageOverlayVerdict.Place, string.Empty,
                    left, top, width, height, fit.FontSize, widthGrowth, heightGrowth);
            }

            return new PlannedOverlay(ImageOverlayVerdict.Place, string.Empty,
                left, top, width, height, fit.FontSize);
        }

        private static bool IsValidDimension(float value)
        {
            return !float.IsNaN(value) && !float.IsInfinity(value) && value > 0F;
        }

        // FitFontSize is unchanged and is the only layout oracle. At the
        // minimum font its required height is a staircase of wrapped-line
        // counts. The least-width/least-height corner of EVERY fitting stair
        // is an area candidate; any interior point has no smaller area.
        private static void FindMinimumArea(float minWidth, float minHeight, float maxWidth, float maxHeight,
            string text, out float bestWidth, out float bestHeight)
        {
            bool Fits(float w, float h) => ImageOverlayLayout.FitFontSize(w, h, text).Fits;
            var heightFloor = FirstFittingFloat(minHeight, maxHeight, h => Fits(maxWidth, h));
            var width = FirstFittingFloat(minWidth, maxWidth, w => Fits(w, maxHeight));
            bestWidth = width;
            bestHeight = maxHeight;
            var bestArea = double.PositiveInfinity;
            while (true)
            {
                var height = FirstFittingFloat(minHeight, maxHeight, h => Fits(width, h));
                var area = (double)width * height;
                if (area < bestArea) { bestArea = area; bestWidth = width; bestHeight = height; }
                if (height == heightFloor || width == maxWidth) break;
                // One float below the exact fitting height forces the next
                // wrapping stair. No guessed font metrics or fixed grid.
                var below = FromBits(ToBits(height) - 1);
                var nextWidth = FirstFittingFloat(width, maxWidth, w => Fits(w, below));
                // All later widths are larger and cannot beat this area bound.
                if ((double)nextWidth * heightFloor >= bestArea) break;
                width = nextWidth;
            }
        }

        // Positive finite floats have the same ordering as their integer bits.
        // Exact representable endpoints avoid rounding that would keep us on
        // the same wrapping stair indefinitely. Caller proves high fits.
        private static float FirstFittingFloat(float low, float high, Func<float, bool> fits)
        {
            if (fits(low)) return low;
            var lo = ToBits(low);
            var hi = ToBits(high);
            while (hi - lo > 1)
            {
                var mid = lo + (hi - lo) / 2;
                if (fits(FromBits(mid))) hi = mid;
                else lo = mid;
            }
            return FromBits(hi);
        }
        private static int ToBits(float value) => BitConverter.ToInt32(BitConverter.GetBytes(value), 0);
        private static float FromBits(int value) => BitConverter.ToSingle(BitConverter.GetBytes(value), 0);

        private static PlannedOverlay SideNote(string reason)
        {
            return new PlannedOverlay(ImageOverlayVerdict.SideNote, reason, 0F, 0F, 0F, 0F, 0F, 0F);
        }
    }
}
