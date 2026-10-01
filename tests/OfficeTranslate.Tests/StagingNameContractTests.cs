using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    // R1: the staging bookmark name contract. A tighten staging is
    // "OTTmp_" + the 32 hex chars of its canonical "OTImg_" name, so an
    // interrupted tighten can be recovered -- and its identity re-adopted
    // -- from the name alone. Only the exact 38-char hex shape is a
    // verifiable identity record; everything else is ignored.
    [TestClass]
    public class StagingNameContractTests
    {
        private const string Hex32 = "0123456789abcdef0123456789abcdef";

        [TestMethod]
        public void VerifiableStaging_Valid38CharHex_ReturnsTrue()
        {
            Assert.IsTrue(ImageOverlayIdentity.IsVerifiableStagingName("OTTmp_" + Hex32));
        }

        [TestMethod]
        public void VerifiableStaging_UppercaseHex_ReturnsTrue()
        {
            Assert.IsTrue(ImageOverlayIdentity.IsVerifiableStagingName(
                "OTTmp_" + Hex32.ToUpperInvariant()));
        }

        [TestMethod]
        public void VerifiableStaging_Old14CharRandom_ReturnsFalse()
        {
            Assert.IsFalse(ImageOverlayIdentity.IsVerifiableStagingName("OTTmp_ab12cd34"));
        }

        [TestMethod]
        public void VerifiableStaging_CanonicalPrefix_ReturnsFalse()
        {
            Assert.IsFalse(ImageOverlayIdentity.IsVerifiableStagingName("OTImg_" + Hex32));
        }

        [TestMethod]
        public void VerifiableStaging_NonHexChar_ReturnsFalse()
        {
            Assert.IsFalse(ImageOverlayIdentity.IsVerifiableStagingName(
                "OTTmp_" + Hex32.Substring(0, 31) + "g"));
        }

        [TestMethod]
        public void VerifiableStaging_NullOrEmpty_ReturnsFalse()
        {
            Assert.IsFalse(ImageOverlayIdentity.IsVerifiableStagingName(null));
            Assert.IsFalse(ImageOverlayIdentity.IsVerifiableStagingName(string.Empty));
        }

        [TestMethod]
        public void VerifiableStaging_UserBookmarkWithPrefix_ReturnsFalse()
        {
            Assert.IsFalse(ImageOverlayIdentity.IsVerifiableStagingName("OTTmp_myBookmark"));
        }

        [TestMethod]
        public void CanonicalNameForStaging_MapsBackToCanonical()
        {
            Assert.AreEqual(
                "OTImg_" + Hex32,
                ImageOverlayIdentity.CanonicalNameForStaging("OTTmp_" + Hex32));
        }

        [TestMethod]
        public void CanonicalNameForStaging_Unverifiable_Throws()
        {
            Assert.ThrowsException<ArgumentException>(() =>
                ImageOverlayIdentity.CanonicalNameForStaging("OTTmp_ab12cd34"));
        }

        [TestMethod]
        public void StagingNameForCanonical_RoundTrip()
        {
            string canonical = "OTImg_" + Hex32;
            string staging = ImageOverlayIdentity.StagingNameForCanonical(canonical);
            Assert.AreEqual("OTTmp_" + Hex32, staging);
            Assert.AreEqual(canonical, ImageOverlayIdentity.CanonicalNameForStaging(staging));
        }

        [TestMethod]
        public void StagingNameForCanonical_ForeignShape_ReturnsEmpty()
        {
            Assert.AreEqual(string.Empty,
                ImageOverlayIdentity.StagingNameForCanonical("OTImg_short"));
            Assert.AreEqual(string.Empty,
                ImageOverlayIdentity.StagingNameForCanonical("SomeOtherBookmark"));
            Assert.AreEqual(string.Empty,
                ImageOverlayIdentity.StagingNameForCanonical(null));
        }

        [TestMethod]
        public void BookmarkNames_StayWithinWord40CharLimit()
        {
            Assert.AreEqual(38, ("OTImg_" + Hex32).Length);
            Assert.AreEqual(38, ("OTTmp_" + Hex32).Length);
        }
    }
}
