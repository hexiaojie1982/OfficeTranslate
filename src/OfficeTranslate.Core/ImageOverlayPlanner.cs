using System;

namespace OfficeTranslate.Core
{
    public enum ImageOverlayVerdict
    {
        // The region geometry is trustworthy: draw the overlay exactly here.
        Place,
        // Geometry is not trustworthy (bad bbox, aspect mismatch, rotated image,
        // or text that cannot fit). Do NOT draw a positioned overlay; the caller
        // must surface an explicit side-note instead.
        SideNote
    }

    public sealed class PlannedOverlay
    {
        public PlannedOverlay(ImageOverlayVerdict verdict, string reason,
            float left, float top, float width, float height, float fontSize)
        {
            Verdict = verdict;
            Reason = reason ?? string.Empty;
            Left = left;
            Top = top;
            Width = width;
            Height = height;
            FontSize = fontSize;
        }

        public ImageOverlayVerdict Verdict { get; }
        public string Reason { get; }

        // Offsets/sizes in Office points, relative to the source image's own
        // origin. The caller adds the image's Left/Top in whichever reference
        // frame the host uses, so no cross-frame conversion happens here.
        // Width/Height are the EXACT bbox-mapped sizes: the planner never
        // expands a small region (R6). If the text cannot fit even at the
        // minimum font size, the verdict is SideNote instead.
        public float Left { get; }
        public float Top { get; }
        public float Width { get; }
        public float Height { get; }
        public float FontSize { get; }
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
        private const float MinBboxUnits = 10F;
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

            // R6: exact bbox mapping, never expanded. A region smaller than the
            // text is handled by FitFontSize below, which degrades to SideNote.
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
                return SideNote("译文在原文区域内放不下（已尝试最小字号），改用旁注。");

            return new PlannedOverlay(ImageOverlayVerdict.Place, string.Empty,
                left, top, width, height, fit.FontSize);
        }

        private static bool IsValidDimension(float value)
        {
            return !float.IsNaN(value) && !float.IsInfinity(value) && value > 0F;
        }

        private static PlannedOverlay SideNote(string reason)
        {
            return new PlannedOverlay(ImageOverlayVerdict.SideNote, reason, 0F, 0F, 0F, 0F, 0F);
        }
    }
}
