using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public sealed class ImageOverlayLayoutTests
    {
        [TestMethod]
        public void Calculate_ShortText_KeepsOriginalHeight()
        {
            // fontSize = 20 * 0.55 = 11; one visual line needs 11 * 1.4 + 4 = 19.4 < 20.
            var metrics = ImageOverlayLayout.Calculate(200F, 20F, "Hello");

            Assert.AreEqual(11F, metrics.FontSize, 0.001);
            Assert.AreEqual(20F, metrics.Height, 0.001);
        }

        [TestMethod]
        public void Calculate_LongText_ExpandsHeight()
        {
            // fontSize = max(8, 12 * 0.55) = 8; 100 CJK units on 96pt width wrap to
            // 9 visual lines: 9 * 8 * 1.4 + 4 = 104.8.
            var metrics = ImageOverlayLayout.Calculate(100F, 12F, new string('中', 100));

            Assert.AreEqual(8F, metrics.FontSize, 0.001);
            Assert.AreEqual(104.8F, metrics.Height, 0.01);
            Assert.IsTrue(metrics.Height > 12F);
        }

        [TestMethod]
        public void Calculate_InvalidDimensions_FallsBackToSafeDefaults()
        {
            var metrics = ImageOverlayLayout.Calculate(0F, -5F, "x");

            Assert.AreEqual(8F, metrics.FontSize, 0.001);
            Assert.IsTrue(metrics.Height >= 12F);
            Assert.IsTrue(metrics.Height <= 720F);
        }

        [TestMethod]
        public void Calculate_HugeText_HeightIsCapped()
        {
            var metrics = ImageOverlayLayout.Calculate(200F, 20F, new string('中', 100000));

            Assert.AreEqual(720F, metrics.Height, 0.001);
        }
    }

    [TestClass]
    public sealed class TranslationResultTests
    {
        [TestMethod]
        public void Unchanged_KeepsOriginalAndNeedsReview()
        {
            var result = TranslationResult.Unchanged("原文");

            Assert.AreEqual("原文", result.Text);
            Assert.IsFalse(result.Changed);
            Assert.IsTrue(result.NeedsReview);
            Assert.IsFalse(result.SkippedNoSource);
        }

        [TestMethod]
        public void Translated_MarksChangedWithoutReview()
        {
            var result = TranslationResult.Translated("译文");

            Assert.IsTrue(result.Changed);
            Assert.IsFalse(result.NeedsReview);
        }

        [TestMethod]
        public void TaskSummary_RecordsUnchangedAsNeedsReview()
        {
            var summary = new TranslationTaskSummary(2);
            summary.Record(TranslationResult.Unchanged("a"));
            summary.Record(TranslationResult.Translated("b"));

            Assert.AreEqual(1, summary.NeedsReview);
            Assert.AreEqual(1, summary.Translated);
        }
    }
}
