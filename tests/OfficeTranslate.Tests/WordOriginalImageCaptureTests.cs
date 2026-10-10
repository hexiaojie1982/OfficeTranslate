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
        internal static string Package(string change = "", int width = 320, int height = 240, byte[]? jpeg = null)
        {
            string image;
            if (jpeg != null) image = Convert.ToBase64String(jpeg);
            else
            {
                using (var bitmap = new Bitmap(width, height))
                using (var stream = new MemoryStream())
                {
                    using (var graphics = Graphics.FromImage(bitmap)) graphics.Clear(change == "transparent" ? Color.Transparent : Color.White);
                    bitmap.Save(stream, ImageFormat.Png); image = Convert.ToBase64String(stream.ToArray());
                }
            }
            var xml = "<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage' xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing' xmlns:pic='http://schemas.openxmlformats.org/drawingml/2006/picture' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>" +
                "<pkg:part pkg:name='/word/document.xml'><pkg:xmlData><wp:inline><pic:pic><pic:blipFill><a:blip r:embed='rId1'/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm/><a:prstGeom prst='rect'/></pic:spPr></pic:pic></wp:inline></pkg:xmlData></pkg:part>" +
                "<pkg:part pkg:name='/word/_rels/document.xml.rels'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' Target='media/image1.png'/></Relationships></pkg:xmlData></pkg:part>" +
                "<pkg:part pkg:name='/word/media/image1.png' pkg:contentType='image/png'><pkg:binaryData>" + image + "</pkg:binaryData></pkg:part></pkg:package>";
            if (jpeg != null) xml = xml.Replace("image/png", "image/jpeg").Replace("image1.png", "image1.jpg");
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

        // A real GDI+ JPEG with an explicit APP1 Exif/TIFF IFD. Avoid
        // SetPropertyItem/Save silently dropping the metadata under test.
        private static byte[] PatternJpeg(int orientation, bool shortValue = false, bool bigEndian = false)
        {
            byte[] jpeg;
            using (var bitmap = new Bitmap(120, 80, PixelFormat.Format24bppRgb))
            using (var graphics = Graphics.FromImage(bitmap))
            using (var stream = new MemoryStream())
            {
                for (var i = 0; i < PatternColors.Length; i++)
                    using (var brush = new SolidBrush(PatternColors[i]))
                        graphics.FillRectangle(brush, (i % 3) * 40, (i / 3) * 40, 40, 40);
                bitmap.Save(stream, ImageFormat.Jpeg); jpeg = stream.ToArray();
            }
            // APP1 length=34 (2-byte length + Exif header 6 + TIFF 26).
            // IFD0: one Orientation entry, SHORT count=1. BYTE count=1
            // intentionally exposes an unparseable one-byte property.
            var exif = bigEndian
                ? new byte[] { 0xff,0xe1,0,34, 69,120,105,102,0,0, 77,77,0,42,0,0,0,8,
                    0,1, 1,18,0,3,0,0,0,1,0,(byte)orientation,0,0, 0,0,0,0 }
                : new byte[] { 0xff,0xe1,0,34, 69,120,105,102,0,0, 73,73,42,0,8,0,0,0,
                    1,0, 18,1,(byte)(shortValue ? 1 : 3),0,1,0,0,0,(byte)orientation,0,0,0, 0,0,0,0 };
            using (var result = new MemoryStream())
            {
                result.Write(jpeg, 0, 2); result.Write(exif, 0, exif.Length);
                result.Write(jpeg, 2, jpeg.Length - 2); return result.ToArray();
            }
        }

        private static readonly Color[] PatternColors = { Color.Red, Color.Lime, Color.Blue, Color.Yellow, Color.Magenta, Color.Cyan };

        // Independent coordinate oracle, not derived from RotateFlipType:
        // EXIF 2 reverses columns; 6 moves stored (x,y) to (H-1-y,x).
        private static Point ViewedPoint(int orientation, int x, int y)
        {
            switch (orientation)
            {
                case 1: return new Point(x, y);
                case 2: return new Point(119 - x, y);
                case 3: return new Point(119 - x, 79 - y);
                case 4: return new Point(x, 79 - y);
                case 5: return new Point(y, x);
                case 6: return new Point(79 - y, x);
                case 7: return new Point(79 - y, 119 - x);
                case 8: return new Point(y, 119 - x);
                default: throw new ArgumentOutOfRangeException(nameof(orientation));
            }
        }

        private static void AssertPattern(Bitmap bitmap, int orientation)
        {
            for (var i = 0; i < PatternColors.Length; i++)
            {
                var point = ViewedPoint(orientation, (i % 3) * 40 + 20, (i / 3) * 40 + 20);
                var actual = bitmap.GetPixel(point.X, point.Y); var expected = PatternColors[i];
                // JPEG chroma compression can change channels slightly.
                Assert.IsTrue(Math.Abs(actual.R - expected.R) <= 12 && Math.Abs(actual.G - expected.G) <= 12 &&
                    Math.Abs(actual.B - expected.B) <= 12, "orientation=" + orientation + " cell=" + i + " actual=" + actual);
            }
        }

        [DataTestMethod]
        [DataRow(1)] [DataRow(2)] [DataRow(3)] [DataRow(4)]
        [DataRow(5)] [DataRow(6)] [DataRow(7)] [DataRow(8)]
        public void GdiRotateFlip_AllExifMappingsMatchIndependentPixelCoordinates(int orientation)
        {
            using (var stream = new MemoryStream(PatternJpeg(orientation)))
            using (var image = (Bitmap)Image.FromStream(stream))
            {
                Assert.AreEqual(120, image.Width); Assert.AreEqual(80, image.Height);
                Assert.AreEqual(orientation, (int)BitConverter.ToUInt16(image.GetPropertyItem(0x112).Value, 0));
                // FromStream has not applied EXIF automatically.
                AssertPattern(image, 1);
                var transforms = new[] { RotateFlipType.RotateNoneFlipNone, RotateFlipType.RotateNoneFlipX,
                    RotateFlipType.Rotate180FlipNone, RotateFlipType.Rotate180FlipX, RotateFlipType.Rotate90FlipX,
                    RotateFlipType.Rotate90FlipNone, RotateFlipType.Rotate270FlipX, RotateFlipType.Rotate270FlipNone };
                image.RotateFlip(transforms[orientation - 1]);
                Assert.AreEqual(orientation >= 5 ? 80 : 120, image.Width);
                Assert.AreEqual(orientation >= 5 ? 120 : 80, image.Height);
                AssertPattern(image, orientation);
            }
        }

        private static void AssertExifCapture(int orientation, bool bigEndian = false)
        {
            var swapped = orientation >= 5;
            Assert.IsTrue(WordOriginalImageCapture.TryCapture(Package(jpeg: PatternJpeg(orientation, bigEndian: bigEndian)),
                swapped ? 80 : 120, swapped ? 120 : 80, 102400, out var bytes, out var reason, out var w, out var h), reason);
            Assert.AreEqual(swapped ? 80 : 120, w); Assert.AreEqual(swapped ? 120 : 80, h);
            using (var stream = new MemoryStream(bytes))
            using (var image = (Bitmap)Image.FromStream(stream))
            {
                Assert.AreEqual(w, image.Width); Assert.AreEqual(h, image.Height);
                AssertPattern(image, orientation);
                Assert.IsFalse(Array.Exists(image.PropertyIdList, id => id == 0x112), "Output must not carry an orientation requiring a second rotation.");
            }
        }

        [TestMethod]
        public void Exif6_CorrectsStoredPixelsAndSwapsCaptureDimensions() => AssertExifCapture(6);

        [TestMethod]
        public void Exif2_CorrectsLeftRightMirrorPixels() => AssertExifCapture(2);

        [DataTestMethod]
        [DataRow(1)] [DataRow(3)] [DataRow(4)] [DataRow(5)] [DataRow(7)] [DataRow(8)]
        public void OtherValidExifOrientations_PreserveCorrectViewedPixels(int orientation) => AssertExifCapture(orientation);

        [TestMethod]
        public void ExifBigEndian6_IsNormalizedByGdiAndCorrectedOnce() => AssertExifCapture(6, bigEndian: true);

        [DataTestMethod]
        [DataRow(0)] [DataRow(9)]
        public void InvalidExifOrientation_HasDistinctFallbackReason(int orientation)
        {
            Assert.IsFalse(WordOriginalImageCapture.TryCapture(Package(jpeg: PatternJpeg(orientation)), 120, 80,
                102400, out var bytes, out var reason, out _, out _));
            Assert.AreEqual("exif-orientation-invalid", reason); Assert.AreEqual(0, bytes.Length);
        }

        [TestMethod]
        public void OneByteExifOrientation_HasDistinctFallbackWithoutThrowing()
        {
            var jpeg = PatternJpeg(6, shortValue: true);
            using (var stream = new MemoryStream(jpeg))
            using (var image = Image.FromStream(stream)) Assert.AreEqual(1, image.GetPropertyItem(0x112).Value.Length, "Fixture must expose the malformed property.");
            Assert.IsFalse(WordOriginalImageCapture.TryCapture(Package(jpeg: jpeg), 120, 80,
                102400, out var bytes, out var reason, out _, out _));
            Assert.AreEqual("exif-orientation-invalid", reason); Assert.AreEqual(0, bytes.Length);
        }
    }
}
