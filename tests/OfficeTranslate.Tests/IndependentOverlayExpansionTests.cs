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
            // At the 6pt floor, an 18pt original height makes width-only growth
            // the minimum-area choice (a 24pt height would favor height-only).
            var plan = ImageOverlayPlanner.Plan(400, 400, 410, 418, 1000, 1000, 1000, 1000, "ABCDE", 0);
            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            Assert.AreEqual(12.7F, plan.Width, 0.0001); Assert.AreEqual(18F, plan.Height);
            Assert.AreEqual(1.27F, plan.WidthGrowthFactor, 0.0001); Assert.AreEqual(1F, plan.HeightGrowthFactor);
            Assert.AreEqual(405F, plan.Left + plan.Width / 2, 0.001);
            Assert.AreEqual(409F, plan.Top + plan.Height / 2, 0.001);
        }

        [TestMethod]
        public void Plan_PreviousWideFixture_NowGrowsOnlyHeightAtSixPoints()
        {
            // Keep the old 10x24 fixture as a regression: lowering the floor
            // legitimately changes its minimum-area direction, not just size.
            var plan = ImageOverlayPlanner.Plan(400, 400, 410, 424, 1000, 1000, 1000, 1000, "ABCDE", 0);
            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            Assert.AreEqual(10F, plan.Width); Assert.AreEqual(25.4F, plan.Height, 0.0001);
            Assert.AreEqual(1F, plan.WidthGrowthFactor);
            Assert.AreEqual(25.4F / 24F, plan.HeightGrowthFactor, 0.0001);
            Assert.AreEqual(405F, plan.Left + plan.Width / 2, 0.001);
            Assert.AreEqual(412F, plan.Top + plan.Height / 2, 0.001);
            Assert.IsTrue(ImageOverlayLayout.FitFontSize(plan.Width, plan.Height, "ABCDE").Fits);
        }

        [TestMethod]
        public void Plan_IndependentCombination_MinimizesAreaAcrossAllWrapCounts()
        {
            const string text = "ABCDEFGHIJ";
            var plan = ImageOverlayPlanner.Plan(400, 400, 410, 420, 1000, 1000, 1000, 1000, text, 0);
            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            // At 6pt: ten ASCII letters measure 34.8pt, margins 4/2pt.
            // The optimum is three lines, 15.6 x 25.4, not width-first or
            // height-first 5x (two lines require 21.4 x 20).
            Assert.AreEqual(15.6F, plan.Width, 0.0001); Assert.AreEqual(25.4F, plan.Height, 0.0001);
            var area = (double)plan.Width * plan.Height;
            for (var lines = 1; lines <= 9; lines++)
            {
                var w = Math.Max(10, 4 + 34.8 / lines);
                var h = Math.Max(20, 2 + lines * 7.8);
                if (w <= 50 && h <= 100) Assert.IsTrue(area <= w * h + 0.001, "Area at wrap count " + lines);
            }
            Assert.IsTrue(ImageOverlayLayout.FitFontSize(plan.Width, plan.Height, text).Fits);
        }

        [TestMethod]
        public void Plan_MultilineMinimumArea_MatchesExhaustiveGridOracle()
        {
            const string text = "ABCDEFGH\nABC\n中文";
            var plan = ImageOverlayPlanner.Plan(400, 400, 410, 420, 1000, 1000, 1000, 1000, text, 0);
            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            var area = (double)plan.Width * plan.Height;
            // Independent broad oracle using unchanged FitFontSize, not the
            // production stair-walk. Every fitting grid pair must be >= optimum.
            for (var w = 10F; w <= 50; w += 0.5F)
                for (var h = 20F; h <= 100; h += 0.5F)
                    if (ImageOverlayLayout.FitFontSize(w, h, text).Fits)
                        Assert.IsTrue(area <= (double)w * h + 0.001);
            Assert.IsTrue(ImageOverlayLayout.FitFontSize(plan.Width, plan.Height, text).Fits);
        }

        [TestMethod]
        public void Plan_WidthAtImageCap_StillGrowsHeightIndependently()
        {
            var plan = ImageOverlayPlanner.Plan(0, 490, 1000, 510, 1000, 1000, 1000, 1000, new string('中', 400), 0);
            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            Assert.AreEqual(1000F, plan.Width); Assert.AreEqual(1F, plan.WidthGrowthFactor);
            Assert.IsTrue(plan.HeightGrowthFactor > 1 && plan.HeightGrowthFactor <= 5);
            Assert.AreEqual(500F, plan.Top + plan.Height / 2, 0.001);
        }

        [TestMethod]
        public void Plan_SaturatedHeightCap_CannotFit_ReturnsSideNote()
        {
            var plan = ImageOverlayPlanner.Plan(400, 400, 500, 403, 1000, 1000, 1000, 1000, "A\nB", 0);
            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
            StringAssert.Contains(plan.Reason, "放不下");
            Assert.AreEqual(0F, plan.WidthGrowthFactor); Assert.AreEqual(0F, plan.HeightGrowthFactor);
        }

        [TestMethod]
        public void Plan_BetweenOldAndNewHeightCaps_NowPlaces()
        {
            var plan = ImageOverlayPlanner.Plan(400, 400, 500, 403, 1000, 1000, 1000, 1000, "A", 0);
            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            Assert.AreEqual(1F, plan.WidthGrowthFactor);
            Assert.IsTrue(plan.HeightGrowthFactor > 3 && plan.HeightGrowthFactor <= 5);
        }
    }
}
