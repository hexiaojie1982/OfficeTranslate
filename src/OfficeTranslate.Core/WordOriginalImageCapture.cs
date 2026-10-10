using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;

namespace OfficeTranslate.Core
{
    // Read-only Flat OPC extraction from an INLINE Range.WordOpenXML.
    // No document resize, temporary document, clipboard or view change.
    public static class WordOriginalImageCapture
    {
        private static readonly XNamespace Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace Pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

        public static bool TryCapture(string flatOpc, float widthPt, float heightPt, long maxPngBytes,
            out byte[] bytes, out string reason, out int sourceWidth, out int sourceHeight)
        {
            bytes = Array.Empty<byte>(); sourceWidth = sourceHeight = 0; reason = "unsupported-package";
            try
            {
                XDocument xml;
                using (var reader = XmlReader.Create(new StringReader(flatOpc), new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null })) xml = XDocument.Load(reader);
                var document = xml.Descendants(Pkg + "part").SingleOrDefault(p => (string?)p.Attribute(Pkg + "name") == "/word/document.xml");
                if (document == null) return false;
                var pictures = document.Descendants(Pic + "pic").ToArray();
                if (pictures.Length != 1 || document.Descendants(Wp + "anchor").Any() || document.Descendants(Wp + "inline").Count() != 1)
                { reason = "not-single-inline-picture"; return false; }
                var picture = pictures[0];
                if (picture.Ancestors(Wp + "inline").Count() != 1 || picture.Ancestors().Any(e =>
                    e.Name.LocalName == "grpSp" || e.Name.LocalName == "wgp" || e.Name.LocalName == "group" || e.Name.LocalName == "AlternateContent"))
                { reason = "group-or-alternate-rendering"; return false; }
                if (picture.Descendants(A + "custGeom").Any() || picture.Descendants(A + "prstGeom")
                    .Any(g => (string?)g.Attribute("prst") != "rect"))
                { reason = "non-rectangular-picture"; return false; }
                foreach (var node in picture.Descendants())
                {
                    if (node.Name == A + "xfrm" && node.Attributes().Any(a =>
                        (a.Name.LocalName == "rot" || a.Name.LocalName == "flipH" || a.Name.LocalName == "flipV") &&
                        a.Value != "0" && a.Value != "false"))
                    { reason = "rotation-or-flip"; return false; }
                    if ((node.Name == A + "srcRect" || node.Name == A + "fillRect") && node.Attributes().Any(a => a.Value != "0"))
                    { reason = "crop-or-fill-transform"; return false; }
                    if (node.Name == A + "tile" || node.Name == A + "effectDag" || node.Name == A + "effectLst" && node.HasElements)
                    { reason = "image-effect-or-tile"; return false; }
                }
                var blips = picture.Descendants(A + "blip").ToArray();
                if (blips.Length != 1 || blips[0].Attribute(R + "link") != null)
                { reason = "linked-or-ambiguous-image"; return false; }
                if (blips[0].Elements().Any(e => e.Name != A + "extLst"))
                { reason = "image-effect"; return false; }
                // SVG replacements and unknown extensions can render differently
                // from the embedded raster. Only the harmless DPI hint is allowed.
                XNamespace a14 = "http://schemas.microsoft.com/office/drawing/2010/main";
                if (blips[0].Descendants().Any(e => e.Name != A + "extLst" && e.Name != A + "ext" && e.Name != a14 + "useLocalDpi"))
                { reason = "unknown-image-extension"; return false; }
                var rid = (string?)blips[0].Attribute(R + "embed");
                var relPart = xml.Descendants(Pkg + "part").SingleOrDefault(p =>
                    (string?)p.Attribute(Pkg + "name") == "/word/_rels/document.xml.rels");
                var relationship = relPart?.Descendants(Rel + "Relationship").SingleOrDefault(r => (string?)r.Attribute("Id") == rid);
                if (relationship == null || (string?)relationship.Attribute("TargetMode") == "External" ||
                    !((string?)relationship.Attribute("Type") ?? "").EndsWith("/image", StringComparison.Ordinal)) return false;
                var target = (string?)relationship.Attribute("Target");
                if (string.IsNullOrEmpty(target)) return false;
                var partPath = new Uri(new Uri("http://package/word/document.xml"), target).AbsolutePath;
                var part = xml.Descendants(Pkg + "part").SingleOrDefault(p => (string?)p.Attribute(Pkg + "name") == partPath);
                var contentType = (string?)part?.Attribute(Pkg + "contentType");
                if (contentType != "image/png" && contentType != "image/jpeg" && contentType != "image/bmp")
                { reason = "unsupported-raster-format"; return false; }
                var encoded = part?.Element(Pkg + "binaryData")?.Value;
                if (string.IsNullOrWhiteSpace(encoded) || encoded!.Length > ImageCaptureBudget.MaxSinglePngBytes * 4 / 3)
                { reason = "source-bytes-limit"; return false; }
                using (var stream = new MemoryStream(Convert.FromBase64String(encoded)))
                using (var original = Image.FromStream(stream))
                {
                    sourceWidth = original.Width; sourceHeight = original.Height;
                    if (sourceWidth > 8000 || sourceHeight > 8000 || (long)sourceWidth * sourceHeight > 32_000_000)
                    { reason = "source-pixels-limit"; return false; }
                    if (!TryNormalizeExifOrientation(original, out reason)) return false;
                    // RotateFlip affects the decoded pixels, not merely these
                    // numbers. All remaining checks and encoding use that same
                    // corrected image; 5-8 exchange its width and height.
                    sourceWidth = original.Width; sourceHeight = original.Height;
                    // Transparent text can depend on the document background.
                    // Do not replace its rendered appearance with a white PNG.
                    if (HasTransparency(original))
                    { reason = "transparent-raster"; return false; }
                    if (!ImageCaptureResolution.SameAspect(sourceWidth, sourceHeight, widthPt, heightPt))
                    { reason = "source-aspect-mismatch"; return false; }
                    var plan = ImageCaptureResolution.Plan(sourceWidth, sourceHeight, widthPt, heightPt);
                    using (var bitmap = new Bitmap(plan.Width, plan.Height, PixelFormat.Format32bppArgb))
                    {
                        using (var graphics = Graphics.FromImage(bitmap))
                        {
                            graphics.Clear(Color.White);
                            graphics.CompositingMode = CompositingMode.SourceOver;
                            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                            graphics.DrawImage(original, new Rectangle(0, 0, plan.Width, plan.Height));
                        }
                        bytes = ImageCaptureBudget.EncodePng(bitmap, maxPngBytes);
                    }
                }
                reason = "inline-original-raster"; return true;
            }
            catch (ImageCaptureLimitException) { throw; }
            catch (Exception ex) { reason = "extraction-" + ex.GetType().Name; return false; }
        }

        private static bool TryNormalizeExifOrientation(Image image, out string reason)
        {
            reason = string.Empty;
            ushort orientation;
            try
            {
                if (!image.PropertyIdList.Contains(0x112)) return true;
                var item = image.GetPropertyItem(0x112);
                // EXIF Orientation is precisely one SHORT. GDI+ exposes its
                // value in native byte order, including big-endian JPEG TIFFs.
                if (item.Type != 3 || item.Len != 2 || item.Value == null || item.Value.Length != 2)
                { reason = "exif-orientation-invalid"; return false; }
                orientation = BitConverter.ToUInt16(item.Value, 0);
                if (orientation < 1 || orientation > 8)
                { reason = "exif-orientation-invalid"; return false; }
            }
            catch (Exception)
            {
                // Unreadable/short EXIF is an explicit rendered fallback, not
                // an outer extraction exception or a guessed orientation.
                reason = "exif-orientation-invalid"; return false;
            }
            if (orientation == 1) return true;
            RotateFlipType transform;
            switch (orientation)
            {
                case 2: transform = RotateFlipType.RotateNoneFlipX; break;
                case 3: transform = RotateFlipType.Rotate180FlipNone; break;
                case 4: transform = RotateFlipType.Rotate180FlipX; break;
                case 5: transform = RotateFlipType.Rotate90FlipX; break;
                case 6: transform = RotateFlipType.Rotate90FlipNone; break;
                case 7: transform = RotateFlipType.Rotate270FlipX; break;
                case 8: transform = RotateFlipType.Rotate270FlipNone; break;
                default: reason = "exif-orientation-invalid"; return false;
            }
            // FromStream leaves JPEG stored pixels unrotated (verified with
            // asymmetric pixel tests). Normalize once; the new PNG below is
            // drawn into a fresh bitmap and does not inherit the EXIF tag.
            try { image.RotateFlip(transform); return true; }
            catch (Exception) { reason = "exif-orientation-transform-failed"; return false; }
        }

        private static bool HasTransparency(Image image)
        {
            if (!(image is Bitmap bitmap) || !Image.IsAlphaPixelFormat(image.PixelFormat)) return false;
            var data = bitmap.LockBits(new Rectangle(0, 0, image.Width, image.Height),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                var row = new byte[image.Width * 4];
                for (var y = 0; y < image.Height; y++)
                {
                    Marshal.Copy(IntPtr.Add(data.Scan0, y * data.Stride), row, 0, row.Length);
                    for (var x = 3; x < row.Length; x += 4) if (row[x] != 255) return true;
                }
                return false;
            }
            finally { bitmap.UnlockBits(data); }
        }
    }
}
