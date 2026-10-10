using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ImageRecognitionRecallTests
    {
        [TestMethod]
        public void SmallOriginal_IsNeverUpsampled()
        {
            var p = ImageCaptureResolution.Plan(400, 300, 180, 135);
            Assert.AreEqual(400, p.Width); Assert.AreEqual(300, p.Height);
        }
        [TestMethod]
        public void DenseOriginalAtSmallDisplay_Reaches2560WithSameAspect()
        {
            var p = ImageCaptureResolution.Plan(3200, 2400, 180, 135);
            Assert.AreEqual(2560, p.Width); Assert.AreEqual(1920, p.Height);
            Assert.IsTrue(ImageCaptureResolution.SameAspect(p.Width, p.Height, 180, 135));
        }
        [TestMethod]
        public void TinyIcon_IsBoundedByDisplayMultiplier()
        {
            var p = ImageCaptureResolution.Plan(6000, 6000, 12, 12);
            Assert.AreEqual(192, p.Width); Assert.AreEqual(192, p.Height);
        }
        [TestMethod]
        public void HugeInput_IsBoundedBySideAndTotalPixels()
        {
            foreach (var source in new[] { Tuple.Create(50000, 30000), Tuple.Create(50000, 1000), Tuple.Create(30000, 30000) })
            {
                var p = ImageCaptureResolution.Plan(source.Item1, source.Item2, 20000, 20000, 50000);
                Assert.IsTrue(p.Width <= 8000 && p.Height <= 8000);
                Assert.IsTrue((long)p.Width * p.Height <= 32_000_000);
            }
        }
        [TestMethod]
        public void PortraitAndLandscape_PreserveAspectWithinPixelRounding()
        {
            foreach (var source in new[] { Tuple.Create(6000, 4000), Tuple.Create(4000, 6000) })
            {
                var p = ImageCaptureResolution.Plan(source.Item1, source.Item2, source.Item1 / 10F, source.Item2 / 10F);
                Assert.IsTrue(ImageCaptureResolution.SameAspect(p.Width, p.Height, source.Item1, source.Item2));
                Assert.IsTrue(Math.Max(p.Width, p.Height) <= 2560);
            }
        }
        [TestMethod]
        public void InvalidResolutionOrShapeDimensions_AreRejected()
        {
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => ImageCaptureResolution.Plan(0, 10, 10, 10));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => ImageCaptureResolution.Plan(10, 10, float.NaN, 10));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => ImageCaptureResolution.Plan(10, 10, 10, float.PositiveInfinity));
            Assert.IsFalse(ImageCaptureResolution.SameAspect(100, 50, 100, 100));
        }
        [TestMethod]
        public void NormalJson_PreservesFullTextsAndCoordinates()
        {
            var r = ImageRegionParser.Parse("```json\n{\"regions\":[{\"source\":\"first\\nsecond\",\"translation\":\"一\\n二\",\"bbox\":[10,20,990,980]}]}\n```");
            Assert.AreEqual(1, r.Statistics.RawCount); Assert.AreEqual(1, r.Statistics.RetainedCount);
            Assert.AreEqual(0, r.Statistics.DroppedCount); Assert.AreEqual(0, r.Statistics.ReviewCount);
            Assert.AreEqual("first\nsecond", r.Regions[0].Source); Assert.AreEqual("一\n二", r.Regions[0].Translation);
            Assert.AreEqual(990F, r.Regions[0].X2);
            Assert.AreEqual("first\nsecond\r一\n二", r.Regions[0].DisplayText(true));
        }
        [TestMethod]
        public void SourceOnly_IsKeptForReviewNeverAsTranslatedOverlay()
        {
            var r = ImageRegionParser.Parse("{\"regions\":[{\"source\":\"recognized\",\"translation\":\"\",\"bbox\":[0,0,100,100]}]}");
            Assert.AreEqual(1, r.Statistics.EmptyTranslationCount); Assert.AreEqual(0, r.Statistics.DroppedCount);
            Assert.IsTrue(r.Regions[0].RequiresReview); Assert.AreEqual("", r.Regions[0].Translation);
            StringAssert.Contains(r.Regions[0].DisplayText(false), "未提供译文");
            StringAssert.Contains(r.Regions[0].DisplayText(true), "recognized");
            var summary = new TranslationTaskSummary(1); summary.RecordImageRecognition(r); summary.RecordImageNeedsReview();
            Assert.AreEqual(0, summary.Translated); Assert.AreEqual(0, summary.OcrNoText);
            Assert.AreEqual(1, summary.ImageNeedsReview); Assert.AreEqual(0, summary.TextNeedsReview);
        }
        [TestMethod]
        public void BrokenBbox_WithUsableTextIsDowngradedNotDropped()
        {
            var r = ImageRegionParser.Parse("{\"regions\":[{\"source\":\"x\",\"translation\":\"译文\",\"bbox\":[1,2,3]}]}");
            Assert.AreEqual(1, r.Statistics.InvalidBboxCount); Assert.AreEqual(1, r.Statistics.InvalidBboxDowngradedCount);
            Assert.AreEqual(1, r.Statistics.RetainedCount); Assert.AreEqual(0, r.Statistics.DroppedCount);
            Assert.IsTrue(r.Regions[0].RequiresReview); Assert.AreEqual("译文", r.Regions[0].Translation);
        }
        [TestMethod]
        public void NonFiniteOutOfRangeReversedAndBooleanBoxes_AreDowngraded()
        {
            foreach (var box in new[] { "[0,0,1001,100]", "[10,10,1,1]", "[true,0,10,10]", "[\"NaN\",0,10,10]", "[null,0,10,10]" })
            {
                var r = ImageRegionParser.Parse("{\"regions\":[{\"source\":\"x\",\"translation\":\"y\",\"bbox\":" + box + "}]}");
                Assert.IsTrue(r.Regions[0].RequiresReview, box); Assert.AreEqual(1, r.Statistics.InvalidBboxDowngradedCount);
            }
        }
        [TestMethod]
        public void NonDictionaryAndEmptyText_AreCountedWithoutLosingUsableRegions()
        {
            var r = ImageRegionParser.Parse("{\"regions\":[null,1,[],{}, {\"source\":\"x\"},{\"translation\":\"y\",\"bbox\":[0,0,10,10]}]}");
            Assert.AreEqual(6, r.Statistics.RawCount); Assert.AreEqual(3, r.Statistics.NonDictionaryCount);
            Assert.AreEqual(1, r.Statistics.EmptyTranslationDroppedCount); Assert.AreEqual(4, r.Statistics.DroppedCount);
            Assert.AreEqual(2, r.Statistics.RetainedCount);
            Assert.AreEqual(r.Statistics.RawCount, r.Statistics.RetainedCount + r.Statistics.DroppedCount);
        }
        [TestMethod]
        public void ZeroRegions_StillReturnsStatisticsAndSummary()
        {
            var r = ImageRegionParser.Parse("{\"regions\":[]}");
            Assert.AreEqual(0, r.Statistics.RawCount); Assert.AreEqual(0, r.Statistics.RetainedCount);
            var s = new TranslationTaskSummary(1); s.RecordImageRecognition(r);
            Assert.AreEqual(1, s.OcrImages); Assert.AreEqual(1, s.OcrNoText);
            StringAssert.Contains(ImageOverlayDiagnostics.FormatOcrSummary("Word", 2560, 1920, r.Statistics,
                "stop", false, "parsed", 12), "raw=0 kept=0 dropped=0");
        }
        [TestMethod]
        public void MalformedRegions_IsExplicitFailure()
        {
            Assert.ThrowsException<InvalidOperationException>(() => ImageRegionParser.Parse("{\"regions\":{}}"));
            Assert.ThrowsException<InvalidOperationException>(() => ImageRegionParser.Parse("No JSON"));
        }
        [TestMethod]
        public void OcrMetadata_DoesNotLogServerSuppliedContent()
        {
            var line = ImageOverlayDiagnostics.FormatOcrSummary("secret-host", 10, 20, null,
                "secret-prompt-key-text", true, "failed", 2);
            StringAssert.Contains(line, "host=Image"); StringAssert.Contains(line, "finish_reason=other");
            StringAssert.Contains(line, "parsed=False"); Assert.IsFalse(line.Contains("secret"));
        }
    }
}
