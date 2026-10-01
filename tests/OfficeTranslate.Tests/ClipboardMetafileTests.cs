using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    // 4ca3b46 review P2: controllable clipboard abstraction + pure
    // conversion tests for the EMF fallback path. Covers: real EMF,
    // temporarily unavailable data, no data, unexpected exceptions,
    // cancellation propagation, the oversize guard, and multi-attempt
    // evidence chaining. Fully hermetic: the fake clipboard never touches
    // the real clipboard (clear/save/restore are part of the abstraction).
    [TestClass]
    public class ClipboardMetafileTests
    {
        private sealed class FakeClipboard : IClipboardReader
        {
            public Func<(Metafile? mf, string diag)>? EnhMetafile;
            public Exception? EnhMetafileError;
            public List<string>? NativeFormats = new List<string>();

            public void Clear() { }
            public IDataObject? GetDataObject() => null;
            public void SetDataObject(IDataObject data) { }
            public uint GetSequenceNumber() => 42;
            public bool ContainsImage() => false;
            public Image? GetImage() => null;
            public string[]? GetFormats() => Array.Empty<string>();
            public List<string>? GetNativeFormats() => NativeFormats;
            public Metafile? GetNativeEnhMetafile(out string diag)
            {
                if (EnhMetafileError != null) throw EnhMetafileError;
                if (EnhMetafile != null)
                {
                    var r = EnhMetafile();
                    diag = r.diag;
                    return r.mf;
                }
                diag = "absent";
                return null;
            }
        }

        // Builds a real in-memory EMF via GDI+. The frame defines the
        // metafile bounds WITHOUT allocating pixels, so the oversize case
        // tests the guard without a multi-hundred-MB allocation. The stream
        // must stay alive for the Metafile's lifetime; the caller disposes
        // both.
        private static (Metafile mf, MemoryStream stream) CreateTestEmf(int w, int h)
        {
            var stream = new MemoryStream();
            using (var bmp = new Bitmap(1, 1))
            using (var g = Graphics.FromImage(bmp))
            {
                IntPtr hdc = g.GetHdc();
                try
                {
                    using (var mf = new Metafile(stream, hdc, new RectangleF(0, 0, w, h), MetafileFrameUnit.Pixel))
                    using (var mg = Graphics.FromImage(mf))
                    {
                        mg.FillRectangle(Brushes.Red, 0, 0, w, h);
                    }
                }
                finally { g.ReleaseHdc(hdc); }
            }
            stream.Position = 0;
            return (new Metafile(stream), stream);
        }

        [TestMethod]
        public void Rasterize_RealEmf_ReturnsImageWithExpectedSize()
        {
            var (mf, stream) = CreateTestEmf(120, 60);
            try
            {
                var img = ClipboardImageCapture.RasterizeMetafile(mf, out var diag);
                Assert.AreEqual("ok", diag);
                Assert.IsNotNull(img);
                using (img)
                {
                    Assert.AreEqual(120, img.Width);
                    Assert.AreEqual(60, img.Height);
                }
            }
            finally { mf.Dispose(); stream.Dispose(); }
        }

        [TestMethod]
        public void Rasterize_OversizeMetafile_RejectedWithoutAllocation()
        {
            var (mf, stream) = CreateTestEmf(20000, 20000);
            try
            {
                var img = ClipboardImageCapture.RasterizeMetafile(mf, out var diag);
                Assert.IsNull(img);
                Assert.IsTrue(diag.StartsWith("over-limit", StringComparison.Ordinal),
                    "diag was: " + diag);
            }
            finally { mf.Dispose(); stream.Dispose(); }
        }

        [TestMethod]
        public void Capture_EmfFallback_SucceedsWhenBitmapAbsent()
        {
            var (mf, stream) = CreateTestEmf(120, 60);
            try
            {
                var fake = new FakeClipboard { EnhMetafile = () => (mf, "offered") };
                var png = ClipboardImageCapture.CapturePng(_ => { }, null, CancellationToken.None, fake);
                Assert.IsTrue(png.Length > 100, "png length " + png.Length);
                Assert.IsTrue(PngDimensions.TryRead(png, out var w, out var h));
                Assert.AreEqual(120, w);
                Assert.AreEqual(60, h);
            }
            finally { mf.Dispose(); stream.Dispose(); }
        }

        [TestMethod]
        public void Capture_EmfTemporarilyUnavailable_RetriesThenSucceeds()
        {
            var (mf, stream) = CreateTestEmf(80, 40);
            try
            {
                int calls = 0;
                var fake = new FakeClipboard
                {
                    EnhMetafile = () => (++calls < 3)
                        ? ((Metafile?)null, "unavailable")
                        : (mf, "offered")
                };
                var png = ClipboardImageCapture.CapturePng(_ => { }, null, CancellationToken.None, fake);
                Assert.IsTrue(png.Length > 100);
                Assert.IsTrue(calls >= 3, "calls " + calls);
            }
            finally { mf.Dispose(); stream.Dispose(); }
        }

        [TestMethod]
        public void Capture_EmfAbsent_ReportsNoImageWithEmfDiag()
        {
            var fake = new FakeClipboard { EnhMetafile = () => ((Metafile?)null, "absent") };
            try
            {
                ClipboardImageCapture.CapturePng(_ => { }, null, CancellationToken.None, fake);
                Assert.Fail("expected NoImage failure");
            }
            catch (InvalidOperationException ex)
            {
                StringAssert.Contains(ex.Message, "EMF：absent");
                StringAssert.Contains(ex.Message, "原生格式：无");
            }
        }

        [TestMethod]
        public void Capture_LongEvidence_IsTruncatedAfterBuild()
        {
            // 3b2a252 review: one very long round must not push the combined
            // evidence over the cap -- truncation is checked after building.
            var manyFormats = new List<string>();
            for (int i = 0; i < 64; i++)
                manyFormats.Add("49" + i.ToString("000") + "=SomeVeryLongCustomClipboardFormatName_" + i);
            var fake = new FakeClipboard
            {
                EnhMetafile = () => ((Metafile?)null, "absent"),
                NativeFormats = manyFormats
            };
            try
            {
                ClipboardImageCapture.CapturePng(_ => { }, null, CancellationToken.None, fake);
                Assert.Fail("expected NoImage failure");
            }
            catch (InvalidOperationException ex)
            {
                Assert.IsTrue(ex.Message.Length <= 2100, "evidence not capped: " + ex.Message.Length);
                StringAssert.Contains(ex.Message, "已截断");
            }
        }

        [TestMethod]
        public void Capture_NativeFormatReadFailure_ReportedDistinctFromEmpty()
        {
            var fake = new FakeClipboard
            {
                EnhMetafile = () => ((Metafile?)null, "absent"),
                NativeFormats = null // read failed: must not read as "none"
            };
            try
            {
                ClipboardImageCapture.CapturePng(_ => { }, null, CancellationToken.None, fake);
                Assert.Fail("expected NoImage failure");
            }
            catch (InvalidOperationException ex)
            {
                StringAssert.Contains(ex.Message, "原生格式：读取失败");
            }
        }

        [TestMethod]
        public void Capture_EmfUnexpectedException_RecordedNotDisguised()
        {
            var fake = new FakeClipboard { EnhMetafileError = new InvalidOperationException("boom") };
            try
            {
                ClipboardImageCapture.CapturePng(_ => { }, null, CancellationToken.None, fake);
                Assert.Fail("expected NoImage failure");
            }
            catch (InvalidOperationException ex)
            {
                // The unexpected error must be RECORDED, not disguised as a
                // plain "no EMF".
                StringAssert.Contains(ex.Message, "EMF：error:InvalidOperationException");
            }
        }

        [TestMethod]
        public void Capture_EmfPollCancellation_PropagatesNotSwallowed()
        {
            var cts = new CancellationTokenSource();
            int calls = 0;
            var fake = new FakeClipboard
            {
                EnhMetafile = () =>
                {
                    if (++calls == 1) cts.Cancel();
                    return ((Metafile?)null, "unavailable");
                }
            };
            // The bitmap poll sees no cancellation; the EMF poll must observe
            // the cancelled token on its next iteration and propagate
            // OperationCanceledException instead of turning it into NoImage.
            Assert.ThrowsException<OperationCanceledException>(() =>
                ClipboardImageCapture.CapturePng(_ => { }, null, cts.Token, fake));
        }

        [TestMethod]
        public void Capture_HardFailureAfterNoImage_ChainsAllAttempts()
        {
            var fake = new FakeClipboard { EnhMetafile = () => ((Metafile?)null, "absent") };
            int copies = 0;
            try
            {
                ClipboardImageCapture.CapturePng(
                    _ =>
                    {
                        if (++copies == 2)
                            throw new COMException("此命令无效。", unchecked((int)0x800A11FD));
                    },
                    null, CancellationToken.None, fake);
                Assert.Fail("expected hard failure");
            }
            catch (InvalidOperationException ex)
            {
                StringAssert.Contains(ex.Message, "0x800A11FD");
                // Attempt 1's NoImage evidence must survive the attempt-2
                // hard failure (4ca3b46 review P2: previously only the most
                // recent attempt was chained).
                StringAssert.Contains(ex.Message, "[第1轮 ");
                StringAssert.Contains(ex.Message, "EMF：absent");
            }
        }
    }
}
