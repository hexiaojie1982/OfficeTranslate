using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class WordOriginalImageCaptureTests
    {
        internal static string Package(string change = "", int width = 320, int height = 240)
        {
            string image;
            using (var bitmap = new Bitmap(width, height))
            using (var stream = new MemoryStream())
            {
                using (var graphics = Graphics.FromImage(bitmap)) graphics.Clear(change == "transparent" ? Color.Transparent : Color.White);
                bitmap.Save(stream, ImageFormat.Png); image = Convert.ToBase64String(stream.ToArray());
            }
            var xml = "<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage' xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing' xmlns:pic='http://schemas.openxmlformats.org/drawingml/2006/picture' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>" +
                "<pkg:part pkg:name='/word/document.xml'><pkg:xmlData><wp:inline><pic:pic><pic:blipFill><a:blip r:embed='rId1'/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm/><a:prstGeom prst='rect'/></pic:spPr></pic:pic></wp:inline></pkg:xmlData></pkg:part>" +
                "<pkg:part pkg:name='/word/_rels/document.xml.rels'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' Target='media/image1.png'/></Relationships></pkg:xmlData></pkg:part>" +
                "<pkg:part pkg:name='/word/media/image1.png' pkg:contentType='image/png'><pkg:binaryData>" + image + "</pkg:binaryData></pkg:part></pkg:package>";
            switch (change)
            {
                case "crop": return xml.Replace("<a:stretch>", "<a:srcRect l='100'/><a:stretch>");
                case "flip": return xml.Replace("<a:xfrm/>", "<a:xfrm flipH='1'/>");
                case "rotate": return xml.Replace("<a:xfrm/>", "<a:xfrm rot='60000'/>");
                case "effect": return xml.Replace("<a:xfrm/>", "<a:xfrm/><a:effectDag/>");
                case "nonrect": return xml.Replace("prst='rect'", "prst='ellipse'");
                case "linked": return xml.Replace("Target='media", "TargetMode='External' Target='media");
                case "svg": return xml.Replace("<a:blip r:embed='rId1'/>", "<a:blip r:embed='rId1'><a:extLst><a:ext><replacement xmlns='urn:unknown'/></a:ext></a:extLst></a:blip>");
                case "float": return xml.Replace("wp:inline", "wp:anchor");
                case "group": return xml.Replace("<pic:pic>", "<a:grpSp><pic:pic>").Replace("</pic:pic>", "</pic:pic></a:grpSp>");
                case "multiple-inline": return xml.Replace("</wp:inline>", "</wp:inline><wp:inline/>");
                default: return xml;
            }
        }
        [TestMethod]
        public void OriginalRasterExtraction_IsReadOnlyAndDoesNotInventPixels()
        {
            var xml = Package(); var before = xml;
            Assert.IsTrue(WordOriginalImageCapture.TryCapture(xml, 180, 135, 102400, out var bytes, out var reason, out var w, out var h), reason);
            Assert.AreEqual(before, xml); Assert.AreEqual(320, w); Assert.AreEqual(240, h);
            Assert.IsTrue(PngDimensions.TryRead(bytes, out var pw, out var ph)); Assert.AreEqual(w, pw); Assert.AreEqual(h, ph);
        }
        [TestMethod]
        public void TransformsEffectsExternalAndUnknownExtensions_RequireRenderedFallback()
        {
            foreach (var kind in new[] { "crop", "flip", "rotate", "effect", "nonrect", "linked", "svg", "float", "group", "multiple-inline", "transparent" })
            {
                Assert.IsFalse(WordOriginalImageCapture.TryCapture(Package(kind), 180, 135, 102400,
                    out var bytes, out var reason, out _, out _), kind);
                Assert.AreEqual(0, bytes.Length); Assert.IsFalse(string.IsNullOrWhiteSpace(reason), kind);
            }
        }
        [TestMethod]
        public void AspectMismatch_RequiresRenderedFallback()
        {
            Assert.IsFalse(WordOriginalImageCapture.TryCapture(Package(), 100, 100, 102400, out _, out var reason, out _, out _));
            Assert.AreEqual("source-aspect-mismatch", reason);
        }
        [TestMethod]
        public void ExtractionPngBudget_IsNotSilentlyBypassedByFallback()
        {
            Assert.ThrowsException<ImageCaptureLimitException>(() => WordOriginalImageCapture.TryCapture(Package(), 180, 135, 8, out _, out _, out _, out _));
        }
        [TestMethod]
        public void DtdAndInvalidPackage_AreRejectedSafely()
        {
            Assert.IsFalse(WordOriginalImageCapture.TryCapture("<!DOCTYPE x [<!ENTITY y SYSTEM 'file:///bad'>]><x>&y;</x>", 180, 135, 102400, out _, out _, out _, out _));
            Assert.IsFalse(WordOriginalImageCapture.TryCapture("garbage", 180, 135, 102400, out _, out _, out _, out _));
        }
        [TestMethod]
        public void Original3200x2400_InSyntheticFlatOpc_Produces2560x1920()
        {
            Assert.IsTrue(WordOriginalImageCapture.TryCapture(Package(width: 3200, height: 2400), 180, 135,
                ImageCaptureBudget.MaxSinglePngBytes, out var bytes, out var reason, out var sourceW, out var sourceH), reason);
            Assert.AreEqual(3200, sourceW); Assert.AreEqual(2400, sourceH);
            Assert.IsTrue(PngDimensions.TryRead(bytes, out var w, out var h));
            Assert.AreEqual(2560, w); Assert.AreEqual(1920, h);
        }
        [TestMethod]
        public void OriginalPixelsOverSafetyLimit_RequireRenderedFallback()
        {
            Assert.IsFalse(WordOriginalImageCapture.TryCapture(Package(width: 8001, height: 20), 8001, 20,
                ImageCaptureBudget.MaxSinglePngBytes, out _, out var reason, out _, out _));
            Assert.AreEqual("source-pixels-limit", reason);
        }
    }
}
