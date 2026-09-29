using System;
using System.Security.Cryptography;
using System.Text;

namespace OfficeTranslate.Core
{
    // R5: stable, persistable owner identity for "which source image does this
    // overlay/note belong to". Cleanup matches on the owner id only, never on
    // region centers or absolute character anchors:
    // - overlapping images have different ids, so one image's cleanup cannot
    //   delete another image's translations;
    // - Word anchor Start/End shifts when shapes are added, but the id does
    //   not depend on anchors at all, so re-translation still finds old boxes.
    //
    // The id is the host shape's Name ("wd:Picture 1"), which Office keeps
    // unique per document/sheet/slide and persists across save/reopen/move/
    // resize. Word inline shapes have no Name, so they fall back to a content
    // hash of the captured PNG (stable across sessions as long as the image
    // itself is unchanged).
    public static class ImageOverlayIdentity
    {
        public static string ForNamedShape(string hostPrefix, string shapeName, byte[] pngBytes)
        {
            var name = string.IsNullOrWhiteSpace(shapeName)
                ? "noname-" + ContentHash(pngBytes)
                : shapeName.Trim();
            return hostPrefix + ":" + name;
        }

        public static string ForContent(string hostPrefix, byte[] pngBytes)
        {
            return hostPrefix + ":img-" + ContentHash(pngBytes);
        }

        public static string ContentHash(byte[] pngBytes)
        {
            if (pngBytes == null || pngBytes.Length == 0) return "empty";
            using (var sha = SHA256.Create())
            {
                var hash = sha.ComputeHash(pngBytes);
                var builder = new StringBuilder(32);
                // First 16 bytes (32 hex chars) is plenty for per-document identity.
                for (var i = 0; i < 16; i++) builder.Append(hash[i].ToString("x2"));
                return builder.ToString();
            }
        }
    }

    // Marker scheme for the shapes this add-in creates. Word/Excel store the
    // marker in AlternativeText, PowerPoint in Tags["OfficeTranslateOCR"].
    // Only shapes created by us are ever touched; the source image's own
    // AlternativeText/description is never modified.
    public static class ImageOverlayTags
    {
        public const string OverlayPrefix = "OfficeTranslateOCR:";
        public const string NotePrefix = "OfficeTranslateOCR-Note:";

        public static string OverlayFor(string ownerId) => OverlayPrefix + ownerId;
        public static string NoteFor(string ownerId) => NotePrefix + ownerId;

        // Pre-owner-id markers from the first 2.1.24 build. They can never be
        // attributed to an image again, so cleanup reaps them instead of
        // leaving them orphaned (they are our own artifacts, not user content).
        public static bool IsLegacyOverlayMarker(string marker) => marker == "OfficeTranslateOCR";
        public static bool IsLegacyNoteMarker(string marker) => marker == "OfficeTranslateOCR-Note";

        // Word/Excel: marker is the full AlternativeText.
        public static bool WordExcelMarkerBelongsTo(string marker, string ownerId)
        {
            return marker == OverlayFor(ownerId) ||
                   marker == NoteFor(ownerId) ||
                   IsLegacyOverlayMarker(marker) ||
                   IsLegacyNoteMarker(marker);
        }

        // PowerPoint: marker is the Tags["OfficeTranslateOCR"] value.
        public static bool PowerPointTagBelongsTo(string tagValue, string ownerId)
        {
            return tagValue == ownerId ||
                   tagValue == "Note:" + ownerId ||
                   tagValue == "1" ||
                   tagValue == "Note";
        }

        // Excludes our own shapes from ordinary text/image collection.
        public static bool IsOwnMarker(string marker)
        {
            if (string.IsNullOrEmpty(marker)) return false;
            return marker.StartsWith(OverlayPrefix, StringComparison.Ordinal) ||
                   marker.StartsWith(NotePrefix, StringComparison.Ordinal) ||
                   IsLegacyOverlayMarker(marker) ||
                   IsLegacyNoteMarker(marker);
        }
    }
}
