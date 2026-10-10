using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class OriginalImageAllHostsTests
    {
        private const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        private const string P = "http://schemas.openxmlformats.org/presentationml/2006/main";

        private static void Entry(ZipArchive zip, string path, byte[] bytes)
        {
            using (var stream = zip.CreateEntry(path).Open()) stream.Write(bytes, 0, bytes.Length);
        }
        private static void Xml(ZipArchive zip, string path, string xml) => Entry(zip, path, Encoding.UTF8.GetBytes(xml));
        private static string Relationship(string id, string kind, string target, bool external = false) =>
            "<Relationship Id='" + id + "' Type='" + R + "/" + kind + "' Target='" + target + "'" + (external ? " TargetMode='External'" : "") + "/>";
        private static void Rels(ZipArchive zip, string path, params string[] rels) =>
            Xml(zip, path, "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" + string.Join("", rels) + "</Relationships>");

        private static MemoryStream Package(string host, string change = "", byte[]? jpeg = null)
        {
            var output = new MemoryStream();
            using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            {
                var ns = host == "Excel" ? Xdr : P;
                var picture = "<h:pic><h:nvPicPr><h:cNvPr id='7' name='Photo'/></h:nvPicPr><h:blipFill><a:blip r:embed='img'/><a:stretch><a:fillRect/></a:stretch></h:blipFill><h:spPr><a:xfrm/><a:prstGeom prst='rect'/></h:spPr></h:pic>";
                switch (change)
                {
                    case "crop": picture = picture.Replace("<a:stretch>", "<a:srcRect l='10000'/><a:stretch>"); break;
                    case "fill": picture = picture.Replace("<a:fillRect/>", "<a:fillRect t='100'/>"); break;
                    case "flip": picture = picture.Replace("<a:xfrm/>", "<a:xfrm flipV='true'/>"); break;
                    case "rotate": picture = picture.Replace("<a:xfrm/>", "<a:xfrm rot='5400000'/>"); break;
                    case "effect": picture = picture.Replace("<a:xfrm/>", "<a:xfrm/><a:effectLst><a:outerShdw/></a:effectLst>"); break;
                    case "nonrect": picture = picture.Replace("prst='rect'", "prst='ellipse'"); break;
                    case "placeholder": picture = picture.Replace("</h:nvPicPr>", "<h:nvPr><h:ph/></h:nvPr></h:nvPicPr>"); break;
                    case "extension": picture = picture.Replace("<a:blip r:embed='img'/>", "<a:blip r:embed='img'><a:extLst><a:ext><unknown/></a:ext></a:extLst></a:blip>"); break;
                    case "shape-extension": picture = picture.Replace("</h:spPr>", "<a:extLst><a:ext/></a:extLst></h:spPr>"); break;
                    case "dpi": picture = picture.Replace("<a:blip r:embed='img'/>", "<a:blip r:embed='img'><a:extLst><a:ext><useLocalDpi xmlns='http://schemas.microsoft.com/office/drawing/2010/main' val='0'/></a:ext></a:extLst></a:blip>"); break;
                    case "group": picture = "<h:grpSp>" + picture + "</h:grpSp>"; break;
                    case "alternate": picture = "<AlternateContent>" + picture + "</AlternateContent>"; break;
                    case "duplicate-name": picture += "<h:sp><h:cNvPr name='Photo'/></h:sp>"; break;
                    case "missing-name": picture = picture.Replace("name='Photo'", "name='Other'"); break;
                }
                var root = "<root xmlns:h='" + ns + "' xmlns:a='" + A + "' xmlns:r='" + R + "'>" + picture + "</root>";
                if (change == "dtd") root = "<!DOCTYPE root [<!ENTITY e SYSTEM 'file:///secret'>]>" + root;
                var prefix = host == "Excel" ? "xl" : "ppt";
                var imagePart = prefix + "/media/original.jpg";
                if (host == "Excel")
                {
                    Xml(zip, "xl/workbook.xml", "<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='" + R + "'><sheets><sheet name='Data' sheetId='41' r:id='sheetRel'/><sheet name='Other' sheetId='1' r:id='otherRel'/></sheets></workbook>");
                    Rels(zip, "xl/_rels/workbook.xml.rels", Relationship("sheetRel", "worksheet", "worksheets/sheet17.xml"), Relationship("otherRel", "worksheet", "worksheets/sheet1.xml"));
                    var drawing = "<drawing r:id='draw'/>";
                    if (change == "duplicate-drawing") drawing += drawing;
                    Xml(zip, "xl/worksheets/sheet17.xml", "<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='" + R + "'>" + drawing + "</worksheet>");
                    Rels(zip, "xl/worksheets/_rels/sheet17.xml.rels", Relationship("draw", "drawing", "../drawings/drawing23.xml"));
                    Xml(zip, "xl/drawings/drawing23.xml", root);
                    Rels(zip, "xl/drawings/_rels/drawing23.xml.rels", Relationship("img", "image", "../media/original.jpg", change == "external"));
                }
                else
                {
                    Xml(zip, "ppt/presentation.xml", "<p:presentation xmlns:p='" + P + "' xmlns:r='" + R + "'><p:sldIdLst><p:sldId id='701' r:id='slideRel'/><p:sldId id='256' r:id='otherRel'/></p:sldIdLst></p:presentation>");
                    Rels(zip, "ppt/_rels/presentation.xml.rels", Relationship("slideRel", "slide", "slides/slide9.xml"), Relationship("otherRel", "slide", "slides/slide1.xml"));
                    Xml(zip, "ppt/slides/slide9.xml", root);
                    Rels(zip, "ppt/slides/_rels/slide9.xml.rels", Relationship("img", "image", "../media/original.jpg", change == "external"));
                }
                Entry(zip, imagePart, jpeg ?? WordOriginalImageCaptureTests.PatternJpeg(1));
                if (change == "duplicate-entry") Entry(zip, imagePart, WordOriginalImageCaptureTests.PatternJpeg(2));
            }
            output.Position = 0; return output;
        }

        private static bool Capture(Stream zip, string host, out byte[] png, out string reason, float width = 120, float height = 80,
            string? identity = null, long budget = ImageCaptureBudget.MaxSinglePngBytes) =>
            OriginalImageCapture.TryCapturePackage(zip, host, identity ?? (host == "Excel" ? "Data" : "701"), "Photo",
                width, height, budget, out png, out reason, out _, out _);

        [DataTestMethod]
        [DataRow("Excel")]
        [DataRow("PowerPoint")]
        public void PackageIdentityAndPictureName_SelectRelationshipNotPartIndex(string host)
        {
            using (var zip = Package(host))
            {
                var before = zip.ToArray();
                Assert.IsTrue(Capture(zip, host, out var png, out var reason), reason);
                Assert.IsTrue(PngDimensions.TryRead(png, out var w, out var h));
                Assert.AreEqual(120, w); Assert.AreEqual(80, h);
                CollectionAssert.AreEqual(before, zip.ToArray()); Assert.IsTrue(zip.CanRead);
            }
        }

        [DataTestMethod]
        [DataRow("Excel")]
        [DataRow("PowerPoint")]
        public void UnsafeOrAmbiguousPackagePicture_RequiresRenderedFallback(string host)
        {
            foreach (var change in new[] { "crop", "fill", "flip", "rotate", "effect", "nonrect", "placeholder", "extension", "shape-extension", "group", "alternate", "external", "duplicate-name", "missing-name", "duplicate-entry" })
                using (var zip = Package(host, change))
                {
                    Assert.IsFalse(Capture(zip, host, out var png, out var reason), change);
                    Assert.AreEqual(0, png.Length); Assert.IsFalse(string.IsNullOrWhiteSpace(reason), change);
                }
        }

        [DataTestMethod]
        [DataRow("Excel", 1)] [DataRow("Excel", 2)] [DataRow("Excel", 3)] [DataRow("Excel", 4)]
        [DataRow("Excel", 5)] [DataRow("Excel", 6)] [DataRow("Excel", 7)] [DataRow("Excel", 8)]
        [DataRow("PowerPoint", 1)] [DataRow("PowerPoint", 2)] [DataRow("PowerPoint", 3)] [DataRow("PowerPoint", 4)]
        [DataRow("PowerPoint", 5)] [DataRow("PowerPoint", 6)] [DataRow("PowerPoint", 7)] [DataRow("PowerPoint", 8)]
        public void PackageExif_AllOrientationsReuseWordCorrectedPixels(string host, int orientation)
        {
            var jpeg = WordOriginalImageCaptureTests.PatternJpeg(orientation);
            var width = orientation >= 5 ? 80 : 120; var height = orientation >= 5 ? 120 : 80;
            using (var zip = Package(host, jpeg: jpeg))
            {
                Assert.IsTrue(Capture(zip, host, out var png, out var reason, width, height), reason);
                Assert.IsTrue(WordOriginalImageCapture.TryCapture(WordOriginalImageCaptureTests.Package(jpeg: jpeg), width, height,
                    ImageCaptureBudget.MaxSinglePngBytes, out var expected, out reason, out _, out _), reason);
                CollectionAssert.AreEqual(expected, png);
            }
        }

        [DataTestMethod]
        [DataRow("Excel")]
        [DataRow("PowerPoint")]
        public void PackageInvalidExifAndAspectMismatch_AreExplicitFallbacks(string host)
        {
            using (var zip = Package(host, jpeg: WordOriginalImageCaptureTests.PatternJpeg(9)))
            {
                Assert.IsFalse(Capture(zip, host, out _, out var reason)); Assert.AreEqual("exif-orientation-invalid", reason);
            }
            using (var zip = Package(host))
            {
                Assert.IsFalse(Capture(zip, host, out _, out var reason, 100, 100)); Assert.AreEqual("source-aspect-mismatch", reason);
            }
        }

        [DataTestMethod]
        [DataRow("Excel")]
        [DataRow("PowerPoint")]
        public void PackageMissingSheetOrSlide_DoesNotGuessAnotherPart(string host)
        {
            using (var zip = Package(host)) Assert.IsFalse(Capture(zip, host, out _, out _, identity: "unknown"));
            using (var zip = Package(host)) Assert.IsFalse(Capture(zip, host, out _, out _, identity: host == "Excel" ? "Other" : "256"));
        }

        [TestMethod]
        public void MultipleWorksheetDrawings_AreNotGuessed()
        {
            using (var zip = Package("Excel", "duplicate-drawing"))
            {
                Assert.IsFalse(Capture(zip, "Excel", out _, out var reason)); Assert.AreEqual("drawing-not-unique", reason);
            }
        }

        [DataTestMethod]
        [DataRow("Excel")]
        [DataRow("PowerPoint")]
        public void PackagePngBudgetFailure_IsNotBypassedByFallback(string host)
        {
            using (var zip = Package(host)) Assert.ThrowsException<ImageCaptureLimitException>(() => Capture(zip, host, out _, out _, budget: 8));
        }

        [DataTestMethod]
        [DataRow("Excel")]
        [DataRow("PowerPoint")]
        public void DirtyUnsavedAndLegacyDocuments_DoNotReadDisk(string host)
        {
            foreach (var saved in new[] { false, true })
            {
                var path = saved ? "absent." + (host == "Excel" ? "xls" : "ppt") : "absent." + (host == "Excel" ? "xlsx" : "pptx");
                Assert.IsFalse(OriginalImageCapture.TryCaptureFile(path, saved, host, "Data", "Photo", 120, 80, 102400,
                    out _, out var reason, out _, out _));
                Assert.AreEqual(saved ? "unsupported-package-format" : "unsaved-document", reason);
            }
            Assert.IsFalse(OriginalImageCapture.TryCaptureFile("", true, host, "Data", "Photo", 120, 80, 102400, out _, out var empty, out _, out _));
            Assert.AreEqual("unsaved-document", empty);
        }

        [DataTestMethod]
        [DataRow("Excel", ".xlsx")]
        [DataRow("Excel", ".xlsm")]
        [DataRow("PowerPoint", ".pptx")]
        [DataRow("PowerPoint", ".pptm")]
        public void SavedPackageFile_IsReadableWhileHostKeepsWriteSharing(string host, string extension)
        {
            var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
            try
            {
                using (var zip = Package(host)) File.WriteAllBytes(path, zip.ToArray());
                var before = File.ReadAllBytes(path);
                using (var occupied = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    Assert.IsTrue(OriginalImageCapture.TryCaptureFile(path, true, host, host == "Excel" ? "Data" : "701", "Photo", 120, 80, 102400,
                        out _, out var reason, out _, out _), reason);
                }
                CollectionAssert.AreEqual(before, File.ReadAllBytes(path));
            }
            finally { File.Delete(path); }
        }

        [TestMethod]
        public void FloatingWordAnchor_UniqueNamedPictureUsesOriginalWithoutMutation()
        {
            var xml = WordOriginalImageCaptureTests.Package("float", jpeg: WordOriginalImageCaptureTests.PatternJpeg(6))
                .Replace("<wp:anchor>", "<wp:anchor><wp:docPr name='FloatingPhoto'/><wp:effectExtent l='0' t='0' r='0' b='0'/>");
            var before = xml;
            Assert.IsTrue(WordOriginalImageCapture.TryCapture(xml, 80, 120, 102400, "FloatingPhoto",
                out var png, out var reason, out var sourceW, out var sourceH), reason);
            Assert.AreEqual("anchor-original-raster", reason); Assert.AreEqual(80, sourceW); Assert.AreEqual(120, sourceH);
            Assert.IsTrue(PngDimensions.TryRead(png, out var w, out var h)); Assert.AreEqual(80, w); Assert.AreEqual(120, h);
            Assert.AreEqual(before, xml);
        }

        [TestMethod]
        public void FloatingWordAnchor_MismatchAmbiguityAndTransformsFallBack()
        {
            var xml = WordOriginalImageCaptureTests.Package("float").Replace("<wp:anchor>", "<wp:anchor><wp:docPr name='Photo'/>");
            foreach (var candidate in new[] { xml.Replace("name='Photo'", "name='Other'"),
                xml.Replace("<wp:docPr name='Photo'/>", "<wp:docPr name='Photo'/><wp:effectExtent l='2'/>") ,
                xml.Replace("<a:xfrm/>", "<a:xfrm rot='60000'/>") ,
                xml.Replace("<a:xfrm/>", "<a:xfrm flipH='true'/>") ,
                xml.Replace("<a:stretch>", "<a:srcRect t='100'/><a:stretch>") ,
                xml.Replace("</wp:anchor>", "</wp:anchor><wp:anchor><wp:docPr name='Photo'/></wp:anchor>") })
            {
                Assert.IsFalse(WordOriginalImageCapture.TryCapture(candidate, 180, 135, 102400, "Photo", out _, out var reason, out _, out _));
                Assert.IsFalse(string.IsNullOrWhiteSpace(reason));
            }
        }

        [TestMethod]
        public void PowerPointExportScale_UsesSlideRelativePointDimensions()
        {
            var scale = ImageCaptureResolution.PowerPointExportScale(180, 135, 960, 540);
            Assert.AreEqual(10240, scale.Item1); Assert.AreEqual(5760, scale.Item2);
            var icon = ImageCaptureResolution.PowerPointExportScale(12, 12, 960, 540);
            Assert.AreEqual(11520, icon.Item1); Assert.AreEqual(6480, icon.Item2);
            var large = ImageCaptureResolution.PowerPointExportScale(3840, 2160, 960, 540);
            Assert.AreEqual(480, large.Item1); Assert.AreEqual(270, large.Item2);
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => ImageCaptureResolution.PowerPointExportScale(0, 12, 960, 540));
        }

        [TestMethod]
        public void OcrSummary_CaptureMetadataIsExplicitAndRejectsContent()
        {
            StringAssert.Contains(ImageOverlayDiagnostics.FormatOcrSummary("Excel", 120, 80, null, "stop", false, "parsed", 2,
                "original", "package-original-raster"), "capture=original capture_reason=package-original-raster");
            StringAssert.Contains(ImageOverlayDiagnostics.FormatOcrSummary("PowerPoint", 120, 80, null, "stop", false, "parsed", 2,
                "fallback", "crop-or-fill-transform-scaled-export"), "capture=fallback capture_reason=crop-or-fill-transform-scaled-export");
            var line = ImageOverlayDiagnostics.FormatOcrSummary("Word", 1, 1, null, "stop", false, "parsed", 0, "secret prompt", "path/key\r\ntext");
            StringAssert.Contains(line, "capture=unknown capture_reason=unknown"); Assert.IsFalse(line.Contains("secret")); Assert.IsFalse(line.Contains("path/key"));
        }

        [DataTestMethod]
        [DataRow("Excel")]
        [DataRow("PowerPoint")]
        public void PackageLargeOriginal_UsesShared2560ResolutionBudget(string host)
        {
            byte[] jpeg;
            using (var bitmap = new Bitmap(3200, 2400))
            using (var graphics = Graphics.FromImage(bitmap))
            using (var stream = new MemoryStream())
            {
                graphics.Clear(Color.White); bitmap.Save(stream, ImageFormat.Jpeg); jpeg = stream.ToArray();
            }
            using (var zip = Package(host, jpeg: jpeg))
            {
                Assert.IsTrue(Capture(zip, host, out var png, out var reason, 180, 135), reason);
                Assert.IsTrue(PngDimensions.TryRead(png, out var w, out var h)); Assert.AreEqual(2560, w); Assert.AreEqual(1920, h);
            }
        }

        [DataTestMethod]
        [DataRow("Excel")]
        [DataRow("PowerPoint")]
        public void PackageXmlDtdCorruptRasterAndTransparency_AreExplicitFallbacks(string host)
        {
            using (var zip = Package(host, "dtd")) Assert.IsFalse(Capture(zip, host, out _, out _));
            using (var zip = Package(host, jpeg: new byte[] { 1, 2, 3 })) Assert.IsFalse(Capture(zip, host, out _, out _));
            byte[] transparent;
            using (var bitmap = new Bitmap(120, 80))
            using (var stream = new MemoryStream()) { bitmap.Save(stream, ImageFormat.Png); transparent = stream.ToArray(); }
            using (var zip = Package(host, jpeg: transparent))
            {
                Assert.IsFalse(Capture(zip, host, out _, out var reason)); Assert.AreEqual("transparent-raster", reason);
            }
        }

        [DataTestMethod]
        [DataRow("Excel")]
        [DataRow("PowerPoint")]
        public void KnownDpiHint_StillUsesOriginalPixels(string host)
        {
            using (var zip = Package(host, "dpi")) Assert.IsTrue(Capture(zip, host, out _, out var reason), reason);
        }
    }
}
