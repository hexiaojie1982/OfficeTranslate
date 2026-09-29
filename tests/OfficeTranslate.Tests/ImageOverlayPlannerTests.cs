using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public sealed class ImageOverlayPlannerTests
    {
        [TestMethod]
        public void Plan_ValidBbox_MapsFractionsToShapeFrame()
        {
            // 1000x800 PNG on a 500x400pt shape: same aspect, bbox [100,200,600,400].
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 600F, 400F,
                1000, 800, 500F, 400F, "Hi", 0F, false, false);

            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            Assert.AreEqual(50F, plan.Left, 0.001);
            Assert.AreEqual(80F, plan.Top, 0.001);
            Assert.AreEqual(250F, plan.Width, 0.001);
            Assert.AreEqual(80F, plan.Height, 0.001);
            Assert.IsTrue(plan.FontSize >= ImageOverlayLayout.MinFontSize);
        }

        [TestMethod]
        public void Plan_InvertedBbox_GoesToSideNote()
        {
            var plan = ImageOverlayPlanner.Plan(600F, 200F, 100F, 400F,
                1000, 800, 500F, 400F, "Hi", 0F, false, false);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
            Assert.IsFalse(string.IsNullOrWhiteSpace(plan.Reason));
        }

        [TestMethod]
        public void Plan_OutOfRangeBbox_GoesToSideNote()
        {
            var plan = ImageOverlayPlanner.Plan(-5F, 200F, 600F, 400F,
                1000, 800, 500F, 400F, "Hi", 0F, false, false);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
        }

        [TestMethod]
        public void Plan_NaNBbox_GoesToSideNote()
        {
            // Corrupt model output (NaN) must not become a fake overlay.
            var plan = ImageOverlayPlanner.Plan(float.NaN, 200F, 600F, 400F,
                1000, 800, 500F, 400F, "Hi", 0F, false, false);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
        }

        [TestMethod]
        public void Plan_TinyBbox_GoesToSideNote()
        {
            // 5x5 units on the 0-1000 scale is noise, not a text region.
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 105F, 205F,
                1000, 800, 500F, 400F, "Hi", 0F, false, false);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
        }

        [TestMethod]
        public void Plan_AspectMismatch_GoesToSideNote()
        {
            // PNG is 2:1 but the shape is 1:1: the capture cannot be trusted.
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 600F, 400F,
                1000, 500, 400F, 400F, "Hi", 0F, false, false);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
            Assert.IsTrue(plan.Reason.Contains("长宽比"));
        }

        [TestMethod]
        public void Plan_RotatedImage_GoesToSideNote()
        {
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 600F, 400F,
                1000, 800, 500F, 400F, "Hi", 90F, false, false);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
            Assert.IsTrue(plan.Reason.Contains("旋转"));
        }

        [TestMethod]
        public void Plan_HorizontalFlip_MirrorsBbox()
        {
            // Displayed (flipped) fraction 0.1-0.3 maps to shape fraction 0.7-0.9.
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 300F, 400F,
                1000, 800, 500F, 400F, "Hi", 0F, true, false);

            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            Assert.AreEqual(350F, plan.Left, 0.001);
            Assert.AreEqual(100F, plan.Width, 0.001);
        }

        [TestMethod]
        public void Plan_TextTooLong_GoesToSideNote()
        {
            // A wall of text can never fit the small region: no fake overlay.
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 300F, 260F,
                1000, 800, 500F, 400F, new string('中', 500), 0F, false, false);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
            Assert.IsTrue(plan.Reason.Contains("放不下"));
        }

        [TestMethod]
        public void Plan_ZeroPixelSize_GoesToSideNote()
        {
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 600F, 400F,
                0, 0, 500F, 400F, "Hi", 0F, false, false);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
        }
    }

    [TestClass]
    public sealed class PngDimensionsTests
    {
        [TestMethod]
        public void TryRead_ValidPng_ReturnsDimensions()
        {
            var png = BuildMinimalPng(640, 480);

            Assert.IsTrue(PngDimensions.TryRead(png, out var width, out var height));
            Assert.AreEqual(640, width);
            Assert.AreEqual(480, height);
        }

        [TestMethod]
        public void TryRead_Garbage_ReturnsFalse()
        {
            Assert.IsFalse(PngDimensions.TryRead(new byte[] { 1, 2, 3 }, out _, out _));
            Assert.IsFalse(PngDimensions.TryRead(new byte[0], out _, out _));
        }

        private static byte[] BuildMinimalPng(int width, int height)
        {
            // 8-byte signature + 4-byte length + "IHDR" + width/height BE + CRC.
            var png = new byte[33];
            byte[] signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
            signature.CopyTo(png, 0);
            png[8] = 0; png[9] = 0; png[10] = 0; png[11] = 13;
            png[12] = (byte)'I'; png[13] = (byte)'H'; png[14] = (byte)'D'; png[15] = (byte)'R';
            WriteBigEndian(png, 16, width);
            WriteBigEndian(png, 20, height);
            return png;
        }

        private static void WriteBigEndian(byte[] data, int offset, int value)
        {
            data[offset] = (byte)(value >> 24);
            data[offset + 1] = (byte)(value >> 16);
            data[offset + 2] = (byte)(value >> 8);
            data[offset + 3] = (byte)value;
        }
    }
}
