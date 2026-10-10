using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OfficeTranslate.Core
{
    // Read-only Flat OPC extraction. Floating anchors require an exact name.
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
            => TryCapture(flatOpc, widthPt, heightPt, maxPngBytes, null,
                out bytes, out reason, out sourceWidth, out sourceHeight);

        public static bool TryCapture(string flatOpc, float widthPt, float heightPt, long maxPngBytes, string? floatingName,
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
                XElement picture;
                if (floatingName == null)
                {
                    var pictures = document.Descendants(Pic + "pic").ToArray();
                    if (pictures.Length != 1 || document.Descendants(Wp + "anchor").Any() || document.Descendants(Wp + "inline").Count() != 1)
                    { reason = "not-single-inline-picture"; return false; }
                    picture = pictures[0];
                    if (picture.Ancestors(Wp + "inline").Count() != 1) return false;
                }
                else
                {
                    var anchors = document.Descendants(Wp + "anchor").Where(e =>
                        (string?)e.Element(Wp + "docPr")?.Attribute("name") == floatingName && floatingName.Length > 0).ToArray();
                    if (anchors.Length != 1 || anchors[0].Descendants(Pic + "pic").Count() != 1)
                    { reason = "anchor-name-not-unique"; return false; }
                    var anchor = anchors[0];
                    // Position/wrapping do not transform the pixels. Effects and
                    // relative sizing do: without proving them, use rendering.
                    if (anchor.Descendants().Any(e => e.Name.LocalName == "sizeRelH" || e.Name.LocalName == "sizeRelV") ||
                        anchor.Elements(Wp + "effectExtent").Any(e => e.Attributes().Any(a => a.Value != "0")))
                    { reason = "anchor-effect-or-relative-size"; return false; }
                    picture = anchor.Descendants(Pic + "pic").Single();
                }
                if (!OriginalImageCapture.ValidatePicture(picture, out var rid, out reason)) return false;
                var relPart = xml.Descendants(Pkg + "part").SingleOrDefault(p =>
                    (string?)p.Attribute(Pkg + "name") == "/word/_rels/document.xml.rels");
                var relationship = relPart?.Descendants(Rel + "Relationship").SingleOrDefault(r => (string?)r.Attribute("Id") == rid);
                if (relationship == null || (string?)relationship.Attribute("TargetMode") == "External" ||
                    !((string?)relationship.Attribute("Type") ?? "").EndsWith("/image", StringComparison.Ordinal))
                { reason = "missing-or-external-image-relationship"; return false; }
                var target = (string?)relationship.Attribute("Target");
                if (string.IsNullOrEmpty(target)) { reason = "missing-image-target"; return false; }
                var partUri = new Uri(new Uri("http://package/word/document.xml"), target);
                if (partUri.Host != "package" || partUri.Query.Length != 0 || partUri.Fragment.Length != 0)
                { reason = "external-image-target"; return false; }
                var partPath = Uri.UnescapeDataString(partUri.AbsolutePath);
                var part = xml.Descendants(Pkg + "part").SingleOrDefault(p => (string?)p.Attribute(Pkg + "name") == partPath);
                var contentType = (string?)part?.Attribute(Pkg + "contentType");
                if (contentType != "image/png" && contentType != "image/jpeg" && contentType != "image/bmp")
                { reason = "unsupported-raster-format"; return false; }
                var encoded = part?.Element(Pkg + "binaryData")?.Value;
                if (string.IsNullOrWhiteSpace(encoded) || encoded!.Length > ImageCaptureBudget.MaxSinglePngBytes * 4 / 3)
                { reason = "source-bytes-limit"; return false; }
                if (!OriginalImageCapture.TryEncode(Convert.FromBase64String(encoded), widthPt, heightPt, maxPngBytes,
                    out bytes, out reason, out sourceWidth, out sourceHeight)) return false;
                reason = floatingName == null ? "inline-original-raster" : "anchor-original-raster"; return true;
            }
            catch (ImageCaptureLimitException) { throw; }
            catch (Exception ex) { reason = "extraction-" + ex.GetType().Name; return false; }
        }

    }
}
