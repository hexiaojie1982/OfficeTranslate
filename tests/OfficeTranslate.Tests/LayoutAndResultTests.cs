using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public sealed class ImageOverlayLayoutTests
    {
        [TestMethod]
        public void FitFontSize_ShortText_UsesLargeSizeAndFits()
        {
            // Single short line in a 200x20 region: the largest fitting size wins.
            var metrics = ImageOverlayLayout.FitFontSize(200F, 20F, "Hello");

            Assert.IsTrue(metrics.Fits);
            Assert.IsTrue(metrics.FontSize >= 8F && metrics.FontSize <= 18F);
        }

        [TestMethod]
        public void FitFontSize_LongText_ShrinksInsteadOfGrowing()
        {
            // 100 CJK chars on a 100x12 region cannot fit even at 8pt:
            // no downward growth is allowed anymore, so this must NOT fit.
            var metrics = ImageOverlayLayout.FitFontSize(100F, 12F, new string('中', 100));

            Assert.IsFalse(metrics.Fits);
        }

        [TestMethod]
        public void FitFontSize_MediumText_ShrinksToFit()
        {
            // 20 CJK chars on a 200x30 region: must fit by shrinking, not growing.
            var metrics = ImageOverlayLayout.FitFontSize(200F, 30F, new string('中', 20));

            Assert.IsTrue(metrics.Fits);
            Assert.IsTrue(metrics.FontSize >= 8F);
        }

        [TestMethod]
        public void FitFontSize_InvalidDimensions_FallsBackToSafeDefaults()
        {
            // Defaults are 24x12: usable height 10pt cannot host even 8pt text
            // (8 * 1.3 = 10.4 > 10), so the honest answer is "does not fit".
            var metrics = ImageOverlayLayout.FitFontSize(0F, -5F, "x");

            Assert.AreEqual(8F, metrics.FontSize, 0.001);
            Assert.IsFalse(metrics.Fits);
        }

        [TestMethod]
        public void EstimateNoteHeight_GrowsWithText()
        {
            var shortNote = ImageOverlayLayout.EstimateNoteHeight(200F, "short", 9F);
            var longNote = ImageOverlayLayout.EstimateNoteHeight(200F, new string('中', 200), 9F);

            Assert.IsTrue(shortNote > 0F && shortNote < longNote);
            Assert.IsTrue(longNote <= 400F);
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
