using System;
using System.Security.Cryptography;
using System.Text;

namespace OfficeTranslate.Core
{
    // R5/C2/C4: stable, persistable owner identity for "which source image
    // does this overlay/note belong to". Cleanup matches on the owner id
    // only, never on region centers or absolute character anchors:
    // - overlapping images have different ids, so one image's cleanup cannot
    //   delete another image's translations;
    // - Word anchor Start/End shifts when shapes are added, but the id does
    //   not depend on anchors at all, so re-translation still finds old boxes;
    // - C4: markers that cannot be attributed to an image are retained, never
    //   treated as belonging to whatever image happens to be translated.
    //
    // The id is the host shape's Name ("wd:Picture 1"), which Office keeps
    // unique per document/sheet/slide and persists across save/reopen/move/
    // resize. Word inline shapes have no Name, so each inline instance gets a
    // persistent GUID stored in a document bookmark (C2): two identical
    // pictures are different instances and must never share an owner. The
    // source image's own AlternativeText/description is never modified; the
    // content hash is only an auxiliary fallback, never the primary owner.
    public static class ImageOverlayIdentity
    {
        public static string ForNamedShape(string hostPrefix, string shapeName, byte[] pngBytes)
        {
            var name = string.IsNullOrWhiteSpace(shapeName)
                ? "noname-" + ContentHash(pngBytes)
                : shapeName.Trim();
            return hostPrefix + ":" + name;
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

        // Pre-owner-id markers (first 2.1.24 review build, and the 2.1.23 and
        // earlier OCR markers that share the same values). They can never be
        // attributed to a specific source image again, so cleanup RETAINS
        // them by default (C4): deleting them while translating an unrelated
        // image would destroy other images' translations. The user removes
        // them manually if desired. They are still recognized as our own
        // artifacts so ordinary text/image collection keeps excluding them.
        public static bool IsLegacyOverlayMarker(string marker) => marker == "OfficeTranslateOCR";
        public static bool IsLegacyNoteMarker(string marker) => marker == "OfficeTranslateOCR-Note";

        // Word/Excel: marker is the full AlternativeText. Only markers that
        // provably belong to THIS image (by owner id) are cleaned up. Legacy
        // id-less markers are never treated as belonging to any image.
        public static bool WordExcelMarkerBelongsTo(string marker, string ownerId)
        {
            return marker == OverlayFor(ownerId) ||
                   marker == NoteFor(ownerId);
        }

        // PowerPoint: marker is the Tags["OfficeTranslateOCR"] value. Same
        // rule: only provably-owned tags are cleaned up.
        public static bool PowerPointTagBelongsTo(string tagValue, string ownerId)
        {
            return tagValue == ownerId ||
                   tagValue == "Note:" + ownerId;
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
