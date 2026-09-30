using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    // C4: cleanup must only delete shapes that provably belong to the image
    // being translated. Legacy id-less markers can never be attributed, so
    // they are retained by default instead of being reaped globally.
    [TestClass]
    public class ImageOverlayIdentityTests
    {
        [TestMethod]
        public void WordExcel_OwnOverlayMarker_BelongsToOwner()
        {
            var owner = "wd:Picture 1";
            Assert.IsTrue(ImageOverlayTags.WordExcelMarkerBelongsTo(
                ImageOverlayTags.OverlayFor(owner), owner));
            Assert.IsTrue(ImageOverlayTags.WordExcelMarkerBelongsTo(
                ImageOverlayTags.NoteFor(owner), owner));
        }

        [TestMethod]
        public void WordExcel_OtherOwnersMarker_DoesNotBelong()
        {
            Assert.IsFalse(ImageOverlayTags.WordExcelMarkerBelongsTo(
                ImageOverlayTags.OverlayFor("wd:Picture 2"), "wd:Picture 1"));
            Assert.IsFalse(ImageOverlayTags.WordExcelMarkerBelongsTo(
                ImageOverlayTags.NoteFor("wd:Picture 2"), "wd:Picture 1"));
        }

        [TestMethod]
        public void WordExcel_LegacyMarkers_AreRetained_NotAttributedToAnyone()
        {
            // C4: translating ANY image must not delete these.
            Assert.IsFalse(ImageOverlayTags.WordExcelMarkerBelongsTo("OfficeTranslateOCR", "wd:Picture 1"));
            Assert.IsFalse(ImageOverlayTags.WordExcelMarkerBelongsTo("OfficeTranslateOCR-Note", "wd:Picture 1"));
            Assert.IsFalse(ImageOverlayTags.WordExcelMarkerBelongsTo("OfficeTranslateOCR", "xl:Picture 7"));
        }

        [TestMethod]
        public void WordExcel_LegacyMarkers_StillExcludedFromCollection()
        {
            // Retained != collected: they are our artifacts, not user content.
            Assert.IsTrue(ImageOverlayTags.IsOwnMarker("OfficeTranslateOCR"));
            Assert.IsTrue(ImageOverlayTags.IsOwnMarker("OfficeTranslateOCR-Note"));
            Assert.IsFalse(ImageOverlayTags.IsOwnMarker("Figure 1: 用户自己的图片描述"));
            Assert.IsFalse(ImageOverlayTags.IsOwnMarker(""));
        }

        [TestMethod]
        public void PowerPoint_OwnTag_BelongsToOwner()
        {
            var owner = "pp:Picture 3";
            Assert.IsTrue(ImageOverlayTags.PowerPointTagBelongsTo(owner, owner));
            Assert.IsTrue(ImageOverlayTags.PowerPointTagBelongsTo("Note:" + owner, owner));
        }

        [TestMethod]
        public void PowerPoint_LegacyTags_AreRetained_NotAttributedToAnyone()
        {
            // C4: the pre-owner-id tag values "1" and "Note" must not be
            // treated as belonging to whatever image is being translated.
            Assert.IsFalse(ImageOverlayTags.PowerPointTagBelongsTo("1", "pp:Picture 3"));
            Assert.IsFalse(ImageOverlayTags.PowerPointTagBelongsTo("Note", "pp:Picture 3"));
            Assert.IsFalse(ImageOverlayTags.PowerPointTagBelongsTo("1", "pp:Picture 9"));
        }

        [TestMethod]
        public void ForNamedShape_BlankName_FallsBackToContentHash()
        {
            var a = ImageOverlayIdentity.ForNamedShape("wd", "   ", new byte[] { 1, 2, 3 });
            var b = ImageOverlayIdentity.ForNamedShape("wd", "", new byte[] { 1, 2, 3 });
            Assert.AreEqual(a, b);
            Assert.IsTrue(a.StartsWith("wd:noname-"));
        }

        [TestMethod]
        public void ContentHash_IsStable_AndDistinct()
        {
            var h1 = ImageOverlayIdentity.ContentHash(new byte[] { 1, 2, 3 });
            var h2 = ImageOverlayIdentity.ContentHash(new byte[] { 1, 2, 3 });
            var h3 = ImageOverlayIdentity.ContentHash(new byte[] { 1, 2, 4 });
            Assert.AreEqual(h1, h2);
            Assert.AreNotEqual(h1, h3);
            Assert.AreEqual("empty", ImageOverlayIdentity.ContentHash(null!));
            Assert.AreEqual("empty", ImageOverlayIdentity.ContentHash(new byte[0]));
        }
    }

    // C3: Word's Range.Information[] returns -1 when a page position cannot
    // be determined; -1, NaN and Infinity must never be used as coordinates.
    [TestClass]
    public class ImageOverlayGeometryTests
    {
        [TestMethod]
        public void IsUsablePageCoordinate_AcceptsOrdinaryPoints()
        {
            Assert.IsTrue(ImageOverlayGeometry.IsUsablePageCoordinate(0F));
            Assert.IsTrue(ImageOverlayGeometry.IsUsablePageCoordinate(90.35F));
            Assert.IsTrue(ImageOverlayGeometry.IsUsablePageCoordinate(1000F));
        }

        [TestMethod]
        public void IsUsablePageCoordinate_RejectsMinusOne_NaN_Infinity()
        {
            Assert.IsFalse(ImageOverlayGeometry.IsUsablePageCoordinate(-1F));
            Assert.IsFalse(ImageOverlayGeometry.IsUsablePageCoordinate(float.NaN));
            Assert.IsFalse(ImageOverlayGeometry.IsUsablePageCoordinate(float.PositiveInfinity));
            Assert.IsFalse(ImageOverlayGeometry.IsUsablePageCoordinate(float.NegativeInfinity));
        }
    }

    // M1: canonical bookmark choice among bookmarks that provably name the
    // same inline-image instance. Pure logic (no COM): order is by bookmark
    // start, then end, then name. Bookmark enumeration order is unspecified
    // and must not decide.
    [TestClass]
    public class ChooseCanonicalBookmarkTests
    {
        private static System.Tuple<string, int, int> B(string name, int start, int end)
            => System.Tuple.Create(name, start, end);

        [TestMethod]
        public void SingleCandidate_IsCanonical_NoDuplicates()
        {
            var chosen = ImageOverlayIdentity.ChooseCanonicalBookmark(
                new System.Collections.Generic.List<System.Tuple<string, int, int>>
                {
                    B("OTImg_aaa", 10, 12),
                });
            Assert.AreEqual("OTImg_aaa", chosen.Item1);
            Assert.AreEqual(0, chosen.Item2.Count);
        }

        [TestMethod]
        public void EarliestStart_Wins()
        {
            var chosen = ImageOverlayIdentity.ChooseCanonicalBookmark(
                new System.Collections.Generic.List<System.Tuple<string, int, int>>
                {
                    B("OTImg_bbb", 12, 14),
                    B("OTImg_aaa", 9, 12),
                });
            Assert.AreEqual("OTImg_aaa", chosen.Item1);
            CollectionAssert.AreEqual(
                new System.Collections.Generic.List<string> { "OTImg_bbb" }, chosen.Item2);
        }

        [TestMethod]
        public void SameStart_EarliestEnd_Wins()
        {
            var chosen = ImageOverlayIdentity.ChooseCanonicalBookmark(
                new System.Collections.Generic.List<System.Tuple<string, int, int>>
                {
                    B("OTImg_bbb", 10, 15),
                    B("OTImg_aaa", 10, 12),
                });
            Assert.AreEqual("OTImg_aaa", chosen.Item1);
            CollectionAssert.AreEqual(
                new System.Collections.Generic.List<string> { "OTImg_bbb" }, chosen.Item2);
        }

        [TestMethod]
        public void SameRange_NameOrdinal_Wins_RegardlessOfInputOrder()
        {
            // The M1 menu repro: an expanded bookmark (9-12) and a fresh one
            // (10-12) both containing the same inline shape (11-12).
            var first = new System.Collections.Generic.List<System.Tuple<string, int, int>>
            {
                B("OTImg_zzz", 9, 12),
                B("OTImg_aaa", 10, 12),
            };
            var second = new System.Collections.Generic.List<System.Tuple<string, int, int>>
            {
                B("OTImg_aaa", 10, 12),
                B("OTImg_zzz", 9, 12),
            };
            var c1 = ImageOverlayIdentity.ChooseCanonicalBookmark(first);
            var c2 = ImageOverlayIdentity.ChooseCanonicalBookmark(second);
            // 9 < 10, so the expanded bookmark is canonical in both orders.
            Assert.AreEqual("OTImg_zzz", c1.Item1);
            Assert.AreEqual("OTImg_zzz", c2.Item1);
            CollectionAssert.AreEqual(
                new System.Collections.Generic.List<string> { "OTImg_aaa" }, c1.Item2);
            CollectionAssert.AreEqual(
                new System.Collections.Generic.List<string> { "OTImg_aaa" }, c2.Item2);
        }

        [TestMethod]
        public void IdenticalRanges_NameOrdinal_BreaksTie()
        {
            var chosen = ImageOverlayIdentity.ChooseCanonicalBookmark(
                new System.Collections.Generic.List<System.Tuple<string, int, int>>
                {
                    B("OTImg_zzz", 10, 12),
                    B("OTImg_aaa", 10, 12),
                    B("OTImg_mmm", 10, 12),
                });
            Assert.AreEqual("OTImg_aaa", chosen.Item1);
            CollectionAssert.AreEqual(
                new System.Collections.Generic.List<string> { "OTImg_mmm", "OTImg_zzz" }, chosen.Item2);
        }
    }
}
