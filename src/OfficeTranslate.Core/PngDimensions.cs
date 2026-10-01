using System;

namespace OfficeTranslate.Core
{
    // Reads PNG pixel dimensions straight from the IHDR chunk so the planner
    // can verify the captured bitmap before any coordinate math. No decoder,
    // no GDI+, no Office interop: pure and unit testable.
    public static class PngDimensions
    {
        private static readonly byte[] Signature = { 137, 80, 78, 71, 13, 10, 26, 10 };

        public static bool TryRead(byte[] png, out int width, out int height)
        {
            width = 0;
            height = 0;
            try
            {
                if (png == null || png.Length < 33) return false;
                for (var i = 0; i < Signature.Length; i++)
                    if (png[i] != Signature[i]) return false;
                // After the 8-byte signature: 4-byte length, 4-byte "IHDR",
                // then width/height as big-endian uint32.
                if (png[12] != 'I' || png[13] != 'H' || png[14] != 'D' || png[15] != 'R') return false;
                width = ReadBigEndianInt32(png, 16);
                height = ReadBigEndianInt32(png, 20);
                return width > 0 && height > 0 && width <= 32768 && height <= 32768;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static int ReadBigEndianInt32(byte[] data, int offset)
        {
            return (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];
        }
    }
}
