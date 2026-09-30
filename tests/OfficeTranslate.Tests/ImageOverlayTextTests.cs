using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    // W1/W2: failure reports embed translations, so truncation must be
    // correct: never throw, never dump unbounded text.
    [TestClass]
    public class ImageOverlayTextTests
    {
        [TestMethod]
        public void Truncate_NullOrEmpty_ReturnsEmpty()
        {
            Assert.AreEqual(string.Empty, ImageOverlayText.Truncate(null));
            Assert.AreEqual(string.Empty, ImageOverlayText.Truncate(string.Empty));
        }

        [TestMethod]
        public void Truncate_ShortText_ReturnsUnchanged()
        {
            Assert.AreEqual("abc", ImageOverlayText.Truncate("abc"));
        }

        [TestMethod]
        public void Truncate_LongText_CutsAtDefaultLimitWithEllipsis()
        {
            var text = new string('x', 600);
            var result = ImageOverlayText.Truncate(text);
            Assert.AreEqual(501, result.Length);
            Assert.IsTrue(result.EndsWith("…"));
        }

        [TestMethod]
        public void Truncate_CustomLimit_Respected()
        {
            Assert.AreEqual("ab…", ImageOverlayText.Truncate("abcdef", 2));
        }

        [TestMethod]
        public void Truncate_NonPositiveLimit_ReturnsEmpty()
        {
            Assert.AreEqual(string.Empty, ImageOverlayText.Truncate("abc", 0));
            Assert.AreEqual(string.Empty, ImageOverlayText.Truncate("abc", -3));
        }
    }
}
