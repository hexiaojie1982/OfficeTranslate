using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public sealed class IndependentOverlayExpansionTests
    {
        [TestMethod]
        public void Plan_WideShortText_GrowsOnlyWidthWhenThatIsMinimumArea()
        {
            // At 1pt, three CJK units fit in one line after width-only growth.
            var plan = ImageOverlayPlanner.Plan(400, 400, 405, 403.3F, 1000, 1000, 1000, 1000, "甲乙丙", 0);
            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            Assert.AreEqual(7F, plan.Width, 0.0001); Assert.AreEqual(3.3F, plan.Height, 0.0001);
            Assert.AreEqual(1.4F, plan.WidthGrowthFactor, 0.0001); Assert.AreEqual(1F, plan.HeightGrowthFactor, 0.0001);
            Assert.AreEqual(402.5F, plan.Left + plan.Width / 2, 0.001);
            Assert.AreEqual(401.65F, plan.Top + plan.Height / 2, 0.001);
        }

        [TestMethod]
        public void Plan_PreviousWideFixture_NowFitsWithoutGrowthBelowSixPoints()
        {
            // Keep the old 10x24 fixture as a regression: lowering the floor
            // now allows fitting without growth.
            var plan = ImageOverlayPlanner.Plan(400, 400, 410, 424, 1000, 1000, 1000, 1000, "ABCDE", 0);
            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            Assert.AreEqual(10F, plan.Width); Assert.AreEqual(24F, plan.Height);
            Assert.AreEqual(1F, plan.WidthGrowthFactor);
            Assert.AreEqual(1F, plan.HeightGrowthFactor);
            Assert.IsTrue(plan.FontSize >= 1 && plan.FontSize < 6);
            Assert.AreEqual(405F, plan.Left + plan.Width / 2, 0.001);
            Assert.AreEqual(412F, plan.Top + plan.Height / 2, 0.001);
            Assert.IsTrue(ImageOverlayLayout.FitFontSize(plan.Width, plan.Height, "ABCDE").Fits);
        }

        [TestMethod]
        public void Plan_IndependentCombination_MinimizesAreaAcrossAllWrapCounts()
        {
            var text = new string('中', 10);
            var plan = ImageOverlayPlanner.Plan(400, 400, 405, 403, 1000, 1000, 1000, 1000, text, 0);
            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            // Ten CJK units at 1pt: optimum is two lines, 9 x 4.6pt.
            Assert.AreEqual(9F, plan.Width, 0.0001); Assert.AreEqual(4.6F, plan.Height, 0.0001);
            var area = (double)plan.Width * plan.Height;
            for (var lines = 1; lines <= 9; lines++)
            {
                var w = Math.Max(5, 4 + 10.0 / lines);
                var h = Math.Max(3, 2 + lines * 1.3);
                if (w <= 25 && h <= 15) Assert.IsTrue(area <= w * h + 0.001, "Area at wrap count " + lines);
            }
            Assert.IsTrue(ImageOverlayLayout.FitFontSize(plan.Width, plan.Height, text).Fits);
        }

        [TestMethod]
        public void Plan_MultilineMinimumArea_MatchesExhaustiveGridOracle()
        {
            const string text = "ABCDEFGH\nABC\n中文";
            var plan = ImageOverlayPlanner.Plan(400, 400, 405, 403, 1000, 1000, 1000, 1000, text, 0);
            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            var area = (double)plan.Width * plan.Height;
            // Independent broad oracle using unchanged FitFontSize, not the
            // production stair-walk. Every fitting grid pair must be >= optimum.
            for (var w = 5F; w <= 25; w += 0.5F)
                for (var h = 3F; h <= 15; h += 0.5F)
                    if (ImageOverlayLayout.FitFontSize(w, h, text).Fits)
                        Assert.IsTrue(area <= (double)w * h + 0.001);
            Assert.IsTrue(ImageOverlayLayout.FitFontSize(plan.Width, plan.Height, text).Fits);
        }

        [TestMethod]
        public void Plan_WidthAtImageCap_StillGrowsHeightIndependently()
        {
            var plan = ImageOverlayPlanner.Plan(0, 490, 1000, 510, 1000, 1000, 1000, 1000, new string('中', 15000), 0);
            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            Assert.AreEqual(1000F, plan.Width); Assert.AreEqual(1F, plan.WidthGrowthFactor);
            Assert.IsTrue(plan.HeightGrowthFactor > 1 && plan.HeightGrowthFactor <= 5);
            Assert.AreEqual(500F, plan.Top + plan.Height / 2, 0.001);
        }

        [TestMethod]
        public void Plan_SaturatedHeightCap_CannotFit_ReturnsSideNote()
        {
            var plan = ImageOverlayPlanner.Plan(400, 400, 500, 403, 1000, 1000, 1000, 1000, "A\nB\nC\nD\nE\nF\nG\nH\nI\nJ\nK", 0);
            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
            StringAssert.Contains(plan.Reason, "放不下");
            Assert.AreEqual(0F, plan.WidthGrowthFactor); Assert.AreEqual(0F, plan.HeightGrowthFactor);
        }

        [TestMethod]
        public void Plan_ShortHeight_NowPlacesWithMinimalGrowth()
        {
            var plan = ImageOverlayPlanner.Plan(400, 400, 500, 403, 1000, 1000, 1000, 1000, "A", 0);
            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            Assert.AreEqual(1F, plan.WidthGrowthFactor);
            Assert.AreEqual(3.3F / 3F, plan.HeightGrowthFactor, 0.0001);
        }
    }
}
