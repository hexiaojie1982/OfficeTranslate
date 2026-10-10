using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;

namespace OfficeTranslate.Core
{
    // No Office references: package identity, rendering guards and encoding
    // are shared by all hosts. Never guess a part from a sheet/slide index.
    public static class OriginalImageCapture
    {
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const long MaxXmlBytes = 8 * 1024 * 1024;

        internal static bool ValidatePicture(XElement picture, out string rid, out string reason)
        {
            rid = ""; reason = "unsupported-picture";
            if (picture.Ancestors().Any(e => e.Name.LocalName == "grpSp" || e.Name.LocalName == "wgp" ||
                e.Name.LocalName == "group" || e.Name.LocalName == "AlternateContent"))
            { reason = "group-or-alternate-rendering"; return false; }
            // A slide placeholder can inherit geometry/effects from layout or
            // master parts that this pic alone cannot prove. Render it instead.
            if (picture.Descendants().Any(e => e.Name.LocalName == "ph"))
            { reason = "inherited-placeholder-rendering"; return false; }
            if (picture.Descendants(A + "custGeom").Any() || picture.Descendants(A + "prstGeom").Any(g => (string?)g.Attribute("prst") != "rect"))
            { reason = "non-rectangular-picture"; return false; }
            foreach (var node in picture.Descendants())
            {
                if (node.Name == A + "xfrm" && node.Attributes().Any(a =>
                    (a.Name.LocalName == "rot" || a.Name.LocalName == "flipH" || a.Name.LocalName == "flipV") && a.Value != "0" && a.Value != "false"))
                { reason = "rotation-or-flip"; return false; }
                if ((node.Name == A + "srcRect" || node.Name == A + "fillRect") && node.Attributes().Any(a => a.Value != "0"))
                { reason = "crop-or-fill-transform"; return false; }
                if (node.Name == A + "tile" || node.Name == A + "effectDag" || node.Name == A + "effectLst" && node.HasElements ||
                    node.Name == A + "scene3d" || node.Name == A + "sp3d" || node.Name.LocalName == "style" || node.Name == A + "ln")
                { reason = "image-effect-or-tile"; return false; }
                // Extensions outside blip may also affect rendering; only the
                // known harmless blip useLocalDpi hint is accepted below.
                if (node.Name.LocalName == "extLst" && !node.Ancestors(A + "blip").Any())
                { reason = "unknown-image-extension"; return false; }
            }
            var blips = picture.Descendants(A + "blip").ToArray();
            if (blips.Length != 1 || blips[0].Attribute(R + "link") != null || string.IsNullOrWhiteSpace((string?)blips[0].Attribute(R + "embed")))
            { reason = "linked-or-ambiguous-image"; return false; }
            if (blips[0].Elements().Any(e => e.Name != A + "extLst"))
            { reason = "image-effect"; return false; }
            XNamespace a14 = "http://schemas.microsoft.com/office/drawing/2010/main";
            if (blips[0].Descendants().Any(e => e.Name != A + "extLst" && e.Name != A + "ext" && e.Name != a14 + "useLocalDpi"))
            { reason = "unknown-image-extension"; return false; }
            rid = (string)blips[0].Attribute(R + "embed")!; reason = "safe-picture"; return true;
        }

        internal static bool TryEncode(byte[] data, float widthPt, float heightPt, long maxBytes,
            out byte[] png, out string reason, out int sourceWidth, out int sourceHeight)
        {
            png = Array.Empty<byte>(); sourceWidth = sourceHeight = 0; reason = "source-bytes-limit";
            if (data.LongLength > ImageCaptureBudget.MaxSinglePngBytes) return false;
            using (var stream = new MemoryStream(data))
            using (var image = Image.FromStream(stream))
            {
                // MIME/extension alone is not proof of a supported raster.
                if (image.RawFormat.Guid != ImageFormat.Png.Guid && image.RawFormat.Guid != ImageFormat.Jpeg.Guid && image.RawFormat.Guid != ImageFormat.Bmp.Guid)
                { reason = "unsupported-raster-format"; return false; }
                sourceWidth = image.Width; sourceHeight = image.Height;
                if (sourceWidth > 8000 || sourceHeight > 8000 || (long)sourceWidth * sourceHeight > 32_000_000)
                { reason = "source-pixels-limit"; return false; }
                if (!TryNormalizeExifOrientation(image, out reason)) return false;
                sourceWidth = image.Width; sourceHeight = image.Height;
                if (HasTransparency(image)) { reason = "transparent-raster"; return false; }
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
                        graphics.DrawImage(image, new Rectangle(0, 0, plan.Width, plan.Height));
                    }
                    png = ImageCaptureBudget.EncodePng(bitmap, maxBytes);
                }
            }
            reason = "package-original-raster"; return true;
        }

        // Host must prove Saved == true. Opening a saved package for a dirty
        // document could return old pixels even when its name/aspect still match.
        public static bool TryCaptureFile(string path, bool saved, string host, string partIdentity, string shapeName,
            float widthPt, float heightPt, long maxBytes, out byte[] png, out string reason, out int sourceWidth, out int sourceHeight)
        {
            png = Array.Empty<byte>(); sourceWidth = sourceHeight = 0;
            reason = "unsaved-document";
            if (!saved || string.IsNullOrEmpty(path)) return false;
            var ext = Path.GetExtension(path).ToLowerInvariant();
            if (!(host == "Excel" ? new[] { ".xlsx", ".xlsm", ".xltx", ".xltm" } : new[] { ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm" }).Contains(ext))
            { reason = "unsupported-package-format"; return false; }
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    return TryCapturePackage(stream, host, partIdentity, shapeName, widthPt, heightPt, maxBytes,
                        out png, out reason, out sourceWidth, out sourceHeight);
            }
            catch (ImageCaptureLimitException) { throw; }
            catch (Exception ex) { reason = "package-read-" + ex.GetType().Name; return false; }
        }

        public static bool TryCapturePackage(Stream stream, string host, string partIdentity, string shapeName,
            float widthPt, float heightPt, long maxBytes, out byte[] png, out string reason, out int sourceWidth, out int sourceHeight)
        {
            png = Array.Empty<byte>(); sourceWidth = sourceHeight = 0; reason = "package-location-not-proven";
            try
            {
                if (string.IsNullOrEmpty(shapeName) || string.IsNullOrEmpty(partIdentity)) return false;
                using (var zip = new ZipArchive(stream, ZipArchiveMode.Read, true))
                {
                    string root, part;
                    if (host == "Excel")
                    {
                        XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                        root = "xl/workbook.xml";
                        var sheet = ReadXml(zip, root).Descendants(s + "sheet").SingleOrDefault(e => (string?)e.Attribute("name") == partIdentity);
                        if (sheet == null) return false;
                        part = Resolve(zip, root, (string?)sheet.Attribute(R + "id"), "worksheet");
                        var drawings = ReadXml(zip, part).Descendants(s + "drawing").ToArray();
                        if (drawings.Length != 1) { reason = "drawing-not-unique"; return false; }
                        part = Resolve(zip, part, (string?)drawings[0].Attribute(R + "id"), "drawing");
                    }
                    else if (host == "PowerPoint")
                    {
                        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
                        root = "ppt/presentation.xml";
                        var slide = ReadXml(zip, root).Descendants(p + "sldId").SingleOrDefault(e => (string?)e.Attribute("id") == partIdentity);
                        if (slide == null) return false;
                        part = Resolve(zip, root, (string?)slide.Attribute(R + "id"), "slide");
                    }
                    else return false;
                    // Name scope is the proven drawing/slide, including other
                    // shapes. A duplicate anywhere in it is ambiguous.
                    var xml = ReadXml(zip, part);
                    var names = xml.Descendants().Where(e => e.Name.LocalName == "cNvPr" && (string?)e.Attribute("name") == shapeName).ToArray();
                    if (names.Length != 1) { reason = "picture-name-not-unique"; return false; }
                    var picture = names[0].Ancestors().FirstOrDefault(e => e.Name.LocalName == "pic");
                    if (picture == null || !ValidatePicture(picture, out var rid, out reason)) return false;
                    var media = Resolve(zip, part, rid, "image");
                    var mediaExt = Path.GetExtension(media).ToLowerInvariant();
                    if (!new[] { ".png", ".jpg", ".jpeg", ".bmp" }.Contains(mediaExt))
                    { reason = "unsupported-raster-format"; return false; }
                    var entry = UniqueEntry(zip, media);
                    if (entry.Length > ImageCaptureBudget.MaxSinglePngBytes) { reason = "source-bytes-limit"; return false; }
                    byte[] data;
                    using (var source = entry.Open()) using (var output = new MemoryStream())
                    {
                        var buffer = new byte[81920]; int count;
                        while ((count = source.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            if (output.Length + count > ImageCaptureBudget.MaxSinglePngBytes) { reason = "source-bytes-limit"; return false; }
                            output.Write(buffer, 0, count);
                        }
                        data = output.ToArray();
                    }
                    return TryEncode(data, widthPt, heightPt, maxBytes, out png, out reason, out sourceWidth, out sourceHeight);
                }
            }
            catch (ImageCaptureLimitException) { throw; }
            catch (Exception ex) { reason = "package-extraction-" + ex.GetType().Name; return false; }
        }

        internal static bool TryNormalizeExifOrientation(Image image, out string reason)
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

        internal static bool HasTransparency(Image image)
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

        private static ZipArchiveEntry UniqueEntry(ZipArchive zip, string name) =>
            zip.Entries.Single(e => e.FullName == name);

        private static XDocument ReadXml(ZipArchive zip, string name)
        {
            var entry = UniqueEntry(zip, name);
            if (entry.Length > MaxXmlBytes) throw new InvalidDataException("XML size limit");
            using (var stream = entry.Open())
            using (var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = MaxXmlBytes }))
                return XDocument.Load(reader);
        }

        private static string Resolve(ZipArchive zip, string part, string? rid, string kind)
        {
            if (string.IsNullOrEmpty(rid)) throw new InvalidDataException("Missing relationship");
            var slash = part.LastIndexOf('/');
            var relPath = part.Substring(0, slash + 1) + "_rels/" + part.Substring(slash + 1) + ".rels";
            var relationship = ReadXml(zip, relPath).Descendants(Rel + "Relationship").Single(e => (string?)e.Attribute("Id") == rid);
            if ((string?)relationship.Attribute("TargetMode") == "External" || !((string?)relationship.Attribute("Type") ?? "").EndsWith("/" + kind, StringComparison.Ordinal))
                throw new InvalidDataException("Unsupported relationship");
            var uri = new Uri(new Uri("http://package/" + part), (string?)relationship.Attribute("Target"));
            if (uri.Host != "package" || uri.Query.Length != 0 || uri.Fragment.Length != 0) throw new InvalidDataException("External target");
            return Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/'));
        }
    }
}
