using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ExportPngBudgetTests
    {
        private static byte[] Png()
        {
            using (var image = new Bitmap(120, 80))
            using (var stream = new MemoryStream()) { image.Save(stream, ImageFormat.Png); return stream.ToArray(); }
        }
        private static void WithFile(byte[] data, Action<string> test)
        {
            var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".png");
            try { File.WriteAllBytes(path, data); test(path); }
            finally { File.Delete(path); }
        }
        [TestMethod]
        public void ExportPng_CompleteIhdrIsAcceptedWithoutChangingBytes()
        {
            var png = Png(); WithFile(png, path => CollectionAssert.AreEqual(png, ImageCaptureBudget.ReadPngFile(path)));
        }
        [TestMethod]
        public void ExportPng_ByteLimitAndZeroBudgetAreExplicitFailures()
        {
            WithFile(Png(), path =>
            {
                Assert.ThrowsException<ImageCaptureLimitException>(() => ImageCaptureBudget.ReadPngFile(path, 8));
                Assert.ThrowsException<ImageCaptureLimitException>(() => ImageCaptureBudget.ReadPngFile(path, 0));
            });
        }
        [TestMethod]
        public void ExportPng_TruncatedHeaderAndInvalidSignatureAreRejected()
        {
            WithFile(new byte[24], path => Assert.ThrowsException<InvalidDataException>(() => ImageCaptureBudget.ReadPngFile(path)));
            WithFile(new byte[33], path => Assert.ThrowsException<InvalidDataException>(() => ImageCaptureBudget.ReadPngFile(path)));
        }
        [DataTestMethod]
        [DataRow(8001, 1)]
        [DataRow(1, 8001)]
        [DataRow(6000, 6000)]
        public void ExportPng_PixelBudgetIsCheckedBeforeReadingPayload(int width, int height)
        {
            var png = Png();
            for (var i = 0; i < 4; i++) { png[16 + i] = (byte)(width >> (24 - i * 8)); png[20 + i] = (byte)(height >> (24 - i * 8)); }
            WithFile(png, path => Assert.ThrowsException<ImageCaptureLimitException>(() => ImageCaptureBudget.ReadPngFile(path)));
        }
    }
}
