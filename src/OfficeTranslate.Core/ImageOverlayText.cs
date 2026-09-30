namespace OfficeTranslate.Core
{
    public static class ImageOverlayText
    {
        // Truncates text for exception messages and logs so a failure report
        // never dumps an unbounded translation into the UI or a log file.
        public static string Truncate(string text, int maxLength = 500)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            if (maxLength <= 0) return string.Empty;
            return text.Length <= maxLength ? text : text.Substring(0, maxLength) + "…";
        }
    }
}
