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
    public static class ImageOverlayPlanner
    {
        private const float BboxScale = 1000F;
        private const float MinBboxUnits = 10F;
        private const float MaxAspectDeviation = 0.05F;
        private const float MaxRotationDegrees = 0.5F;

        public static PlannedOverlay Plan(
            float x1, float y1, float x2, float y2,
            int pixelWidth, int pixelHeight,
            float shapeWidth, float shapeHeight,
            string text,
            float rotationDegrees,
            bool flipHorizontal,
            bool flipVertical)
        {
            if (!(x1 >= 0F && x2 <= BboxScale && y1 >= 0F && y2 <= BboxScale))
                return SideNote("模型返回的坐标超出 0-1000 范围。");
            if (!(x2 > x1 && y2 > y1))
                return SideNote("模型返回的坐标框无效（右下角不在右上角的右下方）。");
            if (x2 - x1 < MinBboxUnits || y2 - y1 < MinBboxUnits)
                return SideNote("模型返回的坐标框过小，不可信。");
            if (pixelWidth <= 0 || pixelHeight <= 0 || shapeWidth <= 0 || shapeHeight <= 0)
                return SideNote("图片尺寸无效，无法换算坐标。");
            if (Math.Abs(rotationDegrees) > MaxRotationDegrees)
                return SideNote("图片已旋转 " + rotationDegrees.ToString("0.#") + "°，坐标无法可靠换算。");

            // The captured PNG must be the same picture as the Office shape.
            // A skewed aspect ratio means DPI/padding anomalies or a stale
            // capture, and any linear mapping would misplace the overlay.
            var pixelAspect = (float)pixelWidth / pixelHeight;
            var shapeAspect = shapeWidth / shapeHeight;
            var deviation = Math.Abs(pixelAspect - shapeAspect) / Math.Max(pixelAspect, shapeAspect);
            if (deviation > MaxAspectDeviation)
                return SideNote("捕获图像长宽比与 Office 图片不一致（偏差 " +
                    (deviation * 100F).ToString("0.#") + "%），坐标无法可靠换算。");

            // Mirror the bbox when the displayed image is flipped: the model
            // sees the flipped rendering, the shape frame is unflipped.
            float bx1 = x1, bx2 = x2, by1 = y1, by2 = y2;
            if (flipHorizontal) { bx1 = BboxScale - x2; bx2 = BboxScale - x1; }
            if (flipVertical) { by1 = BboxScale - y2; by2 = BboxScale - y1; }

            var left = shapeWidth * bx1 / BboxScale;
            var top = shapeHeight * by1 / BboxScale;
            var width = Math.Max(8F, shapeWidth * (bx2 - bx1) / BboxScale);
            var height = Math.Max(8F, shapeHeight * (by2 - by1) / BboxScale);

            var fit = ImageOverlayLayout.FitFontSize(width, height, text);
            if (!fit.Fits)
                return SideNote("译文在原文区域内放不下（已尝试最小字号），改用旁注。");

            return new PlannedOverlay(ImageOverlayVerdict.Place, string.Empty,
                left, top, width, height, fit.FontSize);
        }

        private static PlannedOverlay SideNote(string reason)
        {
            return new PlannedOverlay(ImageOverlayVerdict.SideNote, reason, 0F, 0F, 0F, 0F, 0F);
        }
    }
}
