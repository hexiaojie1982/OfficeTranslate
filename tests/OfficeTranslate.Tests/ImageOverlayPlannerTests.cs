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
                1000, 800, 500F, 400F, "Hi", 0F);

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
                1000, 800, 500F, 400F, "Hi", 0F);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
            Assert.IsFalse(string.IsNullOrWhiteSpace(plan.Reason));
        }

        [TestMethod]
        public void Plan_OutOfRangeBbox_GoesToSideNote()
        {
            var plan = ImageOverlayPlanner.Plan(-5F, 200F, 600F, 400F,
                1000, 800, 500F, 400F, "Hi", 0F);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
        }

        [TestMethod]
        public void Plan_NaNBbox_GoesToSideNote()
        {
            // Corrupt model output (NaN) must not become a fake overlay.
            var plan = ImageOverlayPlanner.Plan(float.NaN, 200F, 600F, 400F,
                1000, 800, 500F, 400F, "Hi", 0F);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
        }

        [TestMethod]
        public void Plan_TinyBbox_GoesToSideNote()
        {
            // 5x5 units on the 0-1000 scale is noise, not a text region.
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 105F, 205F,
                1000, 800, 500F, 400F, "Hi", 0F);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
        }

        [TestMethod]
        public void Plan_AspectMismatch_GoesToSideNote()
        {
            // PNG is 2:1 but the shape is 1:1: the capture cannot be trusted.
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 600F, 400F,
                1000, 500, 400F, 400F, "Hi", 0F);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
            Assert.IsTrue(plan.Reason.Contains("长宽比"));
        }

        [TestMethod]
        public void Plan_RotatedImage_GoesToSideNote()
        {
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 600F, 400F,
                1000, 800, 500F, 400F, "Hi", 90F);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
            Assert.IsTrue(plan.Reason.Contains("旋转"));
        }

        [TestMethod]
        public void Plan_RenderedPngCoords_MapDirectlyWithoutMirroring()
        {
            // R3 (review 2026-09-30): the bbox is in FINAL RENDERED PNG
            // coordinates. An asymmetric bbox on the right side must map
            // directly to the right side of the shape: no flip compensation.
            // (The old Plan_HorizontalFlip_MirrorsBbox validated the wrong
            // premise: hosts capture the already-flipped rendering.)
            var plan = ImageOverlayPlanner.Plan(700F, 200F, 900F, 400F,
                1000, 800, 500F, 400F, "Hi", 0F);

            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            Assert.AreEqual(350F, plan.Left, 0.001);
            Assert.AreEqual(100F, plan.Width, 0.001);
        }

        [TestMethod]
        public void Plan_EdgeTinyRegion_DoesNotExpandBeyondImage()
        {
            // R6 (review 2026-09-30): bbox=[980,800,1000,1000] on a 100x100pt
            // image maps to a 2pt-wide strip at the right edge. The planner
            // must NOT expand it to 8pt (right=106 > image right=100), and
            // the text cannot fit anyway, so the verdict is SideNote.
            var plan = ImageOverlayPlanner.Plan(980F, 800F, 1000F, 1000F,
                200, 200, 100F, 100F, "A", 0F);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
        }

        [TestMethod]
        public void Plan_SmallRegion_KeepsExactSizeWithinBounds()
        {
            // bbox=[700,700,1000,1000] on 100x100pt: exact 30x30pt box, text
            // fits, and the box stays inside the image rect.
            var plan = ImageOverlayPlanner.Plan(700F, 700F, 1000F, 1000F,
                200, 200, 100F, 100F, "A", 0F);

            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            Assert.AreEqual(70F, plan.Left, 0.001);
            Assert.AreEqual(30F, plan.Width, 0.001);
            Assert.IsTrue(plan.Left + plan.Width <= 100.001F);
            Assert.IsTrue(plan.Top + plan.Height <= 100.001F);
        }

        [TestMethod]
        public void Plan_NaNShapeWidth_GoesToSideNote()
        {
            // R7: NaN shape dimensions used to slip through `<= 0` checks and
            // produce Place with NaN geometry.
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 600F, 400F,
                1000, 800, float.NaN, 400F, "Hi", 0F);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
        }

        [TestMethod]
        public void Plan_InfiniteShapeHeight_GoesToSideNote()
        {
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 600F, 400F,
                1000, 800, 500F, float.PositiveInfinity, "Hi", 0F);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
        }

        [TestMethod]
        public void Plan_NaNRotation_GoesToSideNote()
        {
            // NaN rotation used to pass `Math.Abs(rot) > 0.5` (false for NaN).
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 600F, 400F,
                1000, 800, 500F, 400F, "Hi", float.NaN);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
        }

        [TestMethod]
        public void Plan_TextTooLong_GoesToSideNote()
        {
            // A wall of text can never fit the small region: no fake overlay.
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 300F, 260F,
                1000, 800, 500F, 400F, new string('中', 500), 0F);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
            Assert.IsTrue(plan.Reason.Contains("放不下"));
        }

        [TestMethod]
        public void Plan_ZeroPixelSize_GoesToSideNote()
        {
            var plan = ImageOverlayPlanner.Plan(100F, 200F, 600F, 400F,
                0, 0, 500F, 400F, "Hi", 0F);

            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
        }
    }

    [TestClass]
    public sealed class ImageOverlayNotesTests
    {
        [TestMethod]
        public void Combine_MultipleEntries_KeepsEveryTranslation()
        {
            // R2: one image gets ONE combined side-note; every degraded
            // region's translation must survive in it.
            var combined = ImageOverlayNotes.Combine(new[]
            {
                "坐标不可信。\rREGION_ONE",
                "放不下。\rREGION_TWO",
            });

            Assert.IsTrue(combined.Contains("REGION_ONE"));
            Assert.IsTrue(combined.Contains("REGION_TWO"));
            Assert.AreEqual(1, CountOccurrences(combined, ImageOverlayNotes.Header));
        }

        [TestMethod]
        public void Combine_SingleEntry_StillHasHeader()
        {
            var combined = ImageOverlayNotes.Combine(new[] { "原因。\r正文" });

            Assert.IsTrue(combined.StartsWith(ImageOverlayNotes.Header));
            Assert.IsTrue(combined.Contains("正文"));
        }

        [TestMethod]
        public void Header_IsNonEmpty_AndRecognizable()
        {
            // C1: the header must never be empty; an empty header both
            // mislabels the note and hangs naive counting helpers.
            Assert.IsFalse(string.IsNullOrEmpty(ImageOverlayNotes.Header));
            Assert.IsTrue(CountOccurrences(
                ImageOverlayNotes.Combine(new[] { "a\r一", "b\r二" }),
                ImageOverlayNotes.Header) == 1);
        }

        [TestMethod]
        public void CountOccurrences_EmptyNeedle_ReturnsZeroInsteadOfHanging()
        {
            // C1 regression: IndexOf("", i) always returns i, so without the
            // guard this call would never terminate.
            Assert.AreEqual(0, CountOccurrences("anything", ""));
            Assert.AreEqual(0, CountOccurrences("anything", null!));
        }

        private static int CountOccurrences(string text, string needle)
        {
            // C1: an empty needle would match at every position without
            // advancing (IndexOf("", i) == i, i += 0) and loop forever.
            if (string.IsNullOrEmpty(needle)) return 0;
            var count = 0;
            var index = 0;
            while ((index = text.IndexOf(needle, index, System.StringComparison.Ordinal)) >= 0)
            {
                count++;
                index += needle.Length;
            }
            return count;
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
