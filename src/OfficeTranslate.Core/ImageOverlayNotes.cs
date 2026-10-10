using System.Collections.Generic;
using System.Text;

namespace OfficeTranslate.Core
{
    // Hosts collect degraded regions first, then place owner-tagged chunks.
    // Cleanup happens once per image, never once per chunk.
    public static partial class ImageOverlayNotes
    {
        // C1: the header must be non-empty and recognizable. An empty header
        // both mislabels the note and breaks naive counting helpers that
        // advance by needle.Length (an empty needle never advances).
        public const string Header = "【图片译文待检查】";

        public static string Combine(IReadOnlyList<string> entries)
        {
            var builder = new StringBuilder(Header);
            if (entries != null)
            {
                foreach (var entry in entries)
                {
                    builder.Append("\r\r");
                    builder.Append(entry);
                }
            }
            return builder.ToString();
        }
    }
}
