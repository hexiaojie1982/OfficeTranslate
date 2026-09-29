using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public sealed class SourceLanguageProtectorTests
    {
        [TestMethod]
        public void Protect_MixedChineseEnglish_ProtectsEnglishSegment()
        {
            var text = "今天天气不错，API运行正常。";
            var protectedText = SourceLanguageProtector.Protect(text, "简体中文");

            Assert.IsTrue(protectedText.HasProtectedSegments);
            Assert.IsFalse(protectedText.IsFullyProtected);
            Assert.AreEqual(text, protectedText.OriginalText);
            // The foreign segment is replaced by a placeholder token.
            StringAssert.Contains(protectedText.Text, "⟦OT_KEEP_");
            Assert.IsFalse(protectedText.Text.Contains("API"));
        }

        [TestMethod]
        public void Restore_RoundTrip_PutsProtectedSegmentBack()
        {
            var protectedText = SourceLanguageProtector.Protect("今天天气不错，API运行正常。", "简体中文");

            var restored = protectedText.Restore("今天天气很好，⟦OT_KEEP_0001⟧运行正常。");

            Assert.AreEqual("今天天气很好，API运行正常。", restored);
        }

        [TestMethod]
        public void Protect_PureEnglish_WithChineseSource_IsFullyProtected()
        {
            var protectedText = SourceLanguageProtector.Protect("hello world", "简体中文");

            Assert.IsTrue(protectedText.IsFullyProtected);
            Assert.IsFalse(protectedText.HasProtectedSegments && !protectedText.IsFullyProtected);
        }

        [TestMethod]
        public void Restore_TamperedToken_ThrowsProtectedContentException()
        {
            var protectedText = SourceLanguageProtector.Protect("调用API接口", "简体中文");

            Assert.ThrowsException<ProtectedContentException>(() =>
                protectedText.Restore("调用⟦OT_KEEP_9999⟧接口"));
        }

        [TestMethod]
        public void Restore_DroppedToken_ThrowsProtectedContentException()
        {
            var protectedText = SourceLanguageProtector.Protect("调用API接口", "简体中文");

            // The model translated the sentence but silently dropped the placeholder.
            Assert.ThrowsException<ProtectedContentException>(() =>
                protectedText.Restore("调用接口"));
        }

        [TestMethod]
        public void PreservesProtectedSegmentsInOrder_DetectsReorderAndLoss()
        {
            var protectedText = SourceLanguageProtector.Protect("使用API和SDK开发", "简体中文");

            Assert.IsTrue(protectedText.PreservesProtectedSegmentsInOrder("使用API和SDK开发"));
            Assert.IsFalse(protectedText.PreservesProtectedSegmentsInOrder("使用SDK和API开发"));
            Assert.IsFalse(protectedText.PreservesProtectedSegmentsInOrder("使用API开发"));
        }

        [TestMethod]
        public void Protect_LayoutSeparators_RoundTripsNewlines()
        {
            var protectedText = SourceLanguageProtector.Protect("第一行\r\n第二行", "简体中文");

            Assert.IsTrue(protectedText.HasProtectedSegments);
            var restored = protectedText.Restore("Line1⟦OT_KEEP_0001⟧Line2");

            Assert.AreEqual("Line1\r\nLine2", restored);
        }

        [TestMethod]
        public void IsAutomatic_RecognizesAutoDetect()
        {
            Assert.IsTrue(SourceLanguageProtector.IsAutomatic("自动检测"));
            Assert.IsTrue(SourceLanguageProtector.IsAutomatic(""));
            Assert.IsTrue(SourceLanguageProtector.IsAutomatic(null!));
            Assert.IsFalse(SourceLanguageProtector.IsAutomatic("简体中文"));
        }
    }
}
