namespace OfficeTranslate.Core
{
    // C3: validation for coordinates read back from the Office object model
    // (e.g. Range.Information[wdHorizontalPositionRelativeToPage]). Word
    // reports -1 when the position cannot be determined (anchor not visible),
    // and -1 is not a coordinate. NaN/Infinity are never coordinates either.
    public static class ImageOverlayGeometry
    {
        public static bool IsUsablePageCoordinate(float value)
        {
            return !float.IsNaN(value) && !float.IsInfinity(value) && value != -1F;
        }
    }
}
