using System;
using System.Globalization;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public sealed class ImageOverlayExpansionTests
    {
        private static void InBounds(PlannedOverlay plan, float imageWidth, float imageHeight)
        {
            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            Assert.IsTrue(plan.Left >= 0 && plan.Top >= 0);
            Assert.IsTrue(plan.Left + plan.Width <= imageWidth + 0.05F);
            Assert.IsTrue(plan.Top + plan.Height <= imageHeight + 0.05F);
            Assert.IsTrue(plan.WidthGrowthFactor >= 1 && plan.WidthGrowthFactor <= 5);
            Assert.IsTrue(plan.HeightGrowthFactor >= 1 && plan.HeightGrowthFactor <= 5);
            Assert.IsTrue(plan.FontSize >= ImageOverlayLayout.MinFontSize);
        }

        [TestMethod]
        public void Plan_SmallTrustedBox_GrowsMinimallyAroundCenter()
        {
            var plan = ImageOverlayPlanner.Plan(500, 500, 506, 506, 1000, 1000, 1000, 1000, "A", 0);
            InBounds(plan, 1000, 1000);
            Assert.IsTrue(plan.WidthGrowthFactor > 1 && plan.HeightGrowthFactor < 5);
            Assert.AreEqual(503F, plan.Left + plan.Width / 2, 0.001);
            Assert.AreEqual(503F, plan.Top + plan.Height / 2, 0.001);
            Assert.AreEqual(plan.Width / 6, plan.WidthGrowthFactor, 0.00001);
            Assert.AreEqual(plan.Height / 6, plan.HeightGrowthFactor, 0.00001);
            // Existing MeasureUnits has a one-unit floor even for one ASCII
            // letter: 4pt margins + 6pt = 10pt minimum width.
            Assert.AreEqual(10F, plan.Width, 0.0001);
            Assert.IsTrue(ImageOverlayLayout.FitFontSize(plan.Width, plan.Height, "A").Fits);
            // 6pt needs 6*1.3 + 2 = 9.8pt height, not the full 30pt cap.
            Assert.AreEqual(9.8F, plan.Height, 0.0001);
            Assert.IsFalse(ImageOverlayLayout.FitFontSize(plan.Width - 0.0001F, plan.Height, "A").Fits);
            Assert.IsFalse(ImageOverlayLayout.FitFontSize(plan.Width, plan.Height - 0.0001F, "A").Fits);
        }

        [DataTestMethod]
        [DataRow(0F, 0F)]
        [DataRow(994F, 0F)]
        [DataRow(0F, 994F)]
        [DataRow(994F, 994F)]
        public void Plan_EdgeBox_TranslatesGrownFrameInsideImage(float x, float y)
        {
            var plan = ImageOverlayPlanner.Plan(x, y, x + 6, y + 6, 1000, 1000, 1000, 1000, "A", 0);
            InBounds(plan, 1000, 1000);
            Assert.IsTrue(plan.Width > 6 && plan.Height > 6);
            Assert.AreEqual(x == 0 ? 0 : 1000 - plan.Width, plan.Left, 0.001);
            Assert.AreEqual(y == 0 ? 0 : 1000 - plan.Height, plan.Top, 0.001);
            Assert.AreEqual(9.8F, plan.Height, 0.0001);
        }

        [TestMethod]
        public void Plan_ImageWidthCap_PreservesFitAndPerDimensionLimit()
        {
            var text = new string('中', 400);
            var plan = ImageOverlayPlanner.Plan(0, 490, 1000, 510, 1000, 1000, 1000, 1000, text, 0);
            InBounds(plan, 1000, 1000);
            Assert.AreEqual(0F, plan.Left); Assert.AreEqual(1000F, plan.Width);
            Assert.AreEqual(500F, plan.Top + plan.Height / 2, 0.001);
            Assert.IsTrue(plan.Height > 20 && plan.Height < 60);
            Assert.AreEqual(plan.Height / 20, plan.GrowthFactor, 0.00001);
            Assert.IsTrue(ImageOverlayLayout.FitFontSize(plan.Width, plan.Height, text).Fits);
        }

        [TestMethod]
        public void Plan_FiveTimesStillTooSmall_UsesExistingSideNoteChannel()
        {
            var plan = ImageOverlayPlanner.Plan(500, 500, 506, 506, 1000, 1000, 1000, 1000, new string('中', 20), 0);
            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
            StringAssert.Contains(plan.Reason, "放不下"); Assert.AreEqual(0F, plan.GrowthFactor);
            Assert.AreEqual(0F, plan.Width); Assert.AreEqual(0F, plan.Height);
        }

        [DataTestMethod]
        [DataRow(2.9F, 6F, ImageOverlayVerdict.SideNote)]
        [DataRow(6F, 2.9F, ImageOverlayVerdict.SideNote)]
        [DataRow(3F, 3F, ImageOverlayVerdict.Place)]
        [DataRow(9.999F, 9.999F, ImageOverlayVerdict.Place)]
        public void Plan_TrustGateIsThreeUnits(float width, float height, ImageOverlayVerdict verdict)
        {
            var plan = ImageOverlayPlanner.Plan(0, 0, width, height, 1000, 1000, 1000, 1000, "A", 0);
            Assert.AreEqual(verdict, plan.Verdict);
            if (verdict == ImageOverlayVerdict.Place) InBounds(plan, 1000, 1000);
            else { StringAssert.Contains(plan.Reason, "过小，不可信"); Assert.AreEqual(0F, plan.GrowthFactor); }
        }

        [TestMethod]
        public void Plan_AlreadyFits_KeepsExactRectAndGrowthOne()
        {
            var plan = ImageOverlayPlanner.Plan(100, 100, 500, 500, 1000, 1000, 100, 100, "A", 0);
            InBounds(plan, 100, 100);
            Assert.AreEqual(10F, plan.Left); Assert.AreEqual(10F, plan.Top);
            Assert.AreEqual(40F, plan.Width); Assert.AreEqual(40F, plan.Height); Assert.AreEqual(1F, plan.GrowthFactor);
        }

        [DataTestMethod]
        [DataRow(-1F, 0F, 6F, 6F, 1000, 1000, 1000F, 1000F, 0F, "0-1000")]
        [DataRow(0F, 0F, 1001F, 6F, 1000, 1000, 1000F, 1000F, 0F, "0-1000")]
        [DataRow(float.NaN, 0F, 6F, 6F, 1000, 1000, 1000F, 1000F, 0F, "非有限")]
        [DataRow(float.PositiveInfinity, 0F, 6F, 6F, 1000, 1000, 1000F, 1000F, 0F, "非有限")]
        [DataRow(10F, 0F, 6F, 6F, 1000, 1000, 1000F, 1000F, 0F, "坐标框无效")]
        [DataRow(0F, 0F, 6F, 6F, 0, 1000, 1000F, 1000F, 0F, "尺寸无效")]
        [DataRow(0F, 0F, 6F, 6F, 1000, 1000, float.NaN, 1000F, 0F, "尺寸无效")]
        [DataRow(0F, 0F, 6F, 6F, 1000, 1000, 1000F, float.PositiveInfinity, 0F, "尺寸无效")]
        [DataRow(0F, 0F, 6F, 6F, 1000, 1000, 1000F, 1000F, 0.51F, "旋转")]
        [DataRow(0F, 0F, 6F, 6F, 1000, 1000, 1000F, 1000F, float.NaN, "旋转")]
        [DataRow(0F, 0F, 6F, 6F, 1000, 500, 1000F, 1000F, 0F, "长宽比")]
        public void Plan_InvalidInput_IsRejectedBeforeGrowth(float x1, float y1, float x2, float y2,
            int pw, int ph, float sw, float sh, float rotation, string reason)
        {
            var plan = ImageOverlayPlanner.Plan(x1, y1, x2, y2, pw, ph, sw, sh, "A", rotation);
            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
            StringAssert.Contains(plan.Reason, reason); Assert.AreEqual(0F, plan.GrowthFactor);
        }

        [TestMethod]
        public void Plan_ValidationOrder_RemainsRangeThenBoxThenTrustThenDimensionsThenRotationThenAspect()
        {
            AssertReason(-1, 0, 1, 1, "0-1000");
            AssertReason(1, 0, 0, 1, "坐标框无效");
            AssertReason(0, 0, 2.9F, 2.9F, "过小");
            // All later guards are also invalid; dimensions must win.
            AssertReason(0, 0, 6, 6, "尺寸无效");
            var rotation = ImageOverlayPlanner.Plan(0, 0, 6, 6, 1000, 500, 1000, 1000, "A", 90);
            StringAssert.Contains(rotation.Reason, "旋转");
        }
        private static void AssertReason(float x1, float y1, float x2, float y2, string reason)
        {
            var plan = ImageOverlayPlanner.Plan(x1, y1, x2, y2, 0, 500, float.NaN, 1000, "A", 90);
            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict); StringAssert.Contains(plan.Reason, reason);
        }

        [DataTestMethod]
        [DataRow("Word")]
        [DataRow("Excel")]
        [DataRow("PowerPoint")]
        public void Diagnostics_GrowFieldIsConsistentInvariantAndMetadataOnly(string host)
        {
            var previous = Thread.CurrentThread.CurrentCulture;
            try
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
                var plan = ImageOverlayPlanner.Plan(500, 500, 506, 506, 1000, 1000, 1000, 1000, "A", 0);
                var entry = new ImageOverlayDiagnosticEntry { Host = host, Verdict = plan.Verdict.ToString(),
                    WidthGrowthFactor = plan.WidthGrowthFactor, HeightGrowthFactor = plan.HeightGrowthFactor };
                var line = ImageOverlayDiagnostics.Format(entry);
                StringAssert.Contains(line, "host=" + host); StringAssert.Contains(line, "verdict=Place");
                StringAssert.Contains(line, "grow=1.7x1.6"); Assert.IsFalse(line.Contains("PRIVATE_OCR_TEXT"));
                entry.WidthGrowthFactor = 1; entry.HeightGrowthFactor = 1;
                StringAssert.Contains(ImageOverlayDiagnostics.Format(entry), "grow=1.0x1.0");
                var note = new PlannedOverlay(ImageOverlayVerdict.SideNote, "reason", 0, 0, 0, 0, 0);
                entry.Verdict = note.Verdict.ToString(); entry.WidthGrowthFactor = note.WidthGrowthFactor;
                entry.HeightGrowthFactor = note.HeightGrowthFactor;
                StringAssert.Contains(ImageOverlayDiagnostics.Format(entry), "grow=0.0");
            }
            finally { Thread.CurrentThread.CurrentCulture = previous; }
        }
    }
}
