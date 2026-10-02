using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ImageCaptureBudgetTests
    {
        [TestMethod]
        public void Real512MiBBoundary_AcceptsExactLimit_RejectsOneMoreByteWithoutAllocation()
        {
            var budget = new ImageCaptureBudget();
            budget.Add(ImageCaptureBudget.DefaultLimitBytes - 1);
            budget.Add(1);
            Assert.AreEqual(0L, budget.RemainingBytes);
            Assert.ThrowsException<ImageCaptureLimitException>(() => budget.Add(1));
            Assert.AreEqual(ImageCaptureBudget.DefaultLimitBytes, budget.UsedBytes);
        }
        [TestMethod]
        public void HugeReservation_CannotOverflow_OrModifyBudget()
        {
            var budget = new ImageCaptureBudget();
            budget.Add(10);
            Assert.ThrowsException<ImageCaptureLimitException>(() => budget.Add(long.MaxValue));
            Assert.AreEqual(10L, budget.UsedBytes);
        }
        [TestMethod]
        public void InvalidArguments_AreRejected()
        {
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new ImageCaptureBudget(0));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new ImageCaptureBudget().Add(-1));
        }
        [TestMethod]
        public void PngEncoder_UsesRealGdiAndReportsBudgetInsteadOfGenericError()
        {
            using (var image = new Bitmap(32, 32))
            {
                Assert.ThrowsException<ImageCaptureLimitException>(() => ImageCaptureBudget.EncodePng(image, 8));
                var png = ImageCaptureBudget.EncodePng(image, 10240);
                using (var stream = new MemoryStream(png))
                using (var roundTrip = Image.FromStream(stream)) Assert.AreEqual(32, roundTrip.Width);
            }
        }
        [TestMethod]
        public void Stream_WriteByteAndSetLengthCannotBypassCap()
        {
            using (var stream = new ImageCaptureBudget.LimitedPngStream(5))
            {
                stream.Write(new byte[5], 0, 5);
                Assert.IsTrue(stream.Capacity <= 5);
                Assert.ThrowsException<ImageCaptureLimitException>(() => stream.WriteByte(1));
                Assert.ThrowsException<ImageCaptureLimitException>(() => stream.SetLength(6));
                Assert.AreEqual(5L, stream.Length);
            }
        }
        [TestMethod]
        public async Task ExactBatchLimit_ReturnsAllImagesWithRemainingBudget()
        {
            var seen = new System.Collections.Generic.List<long>();
            var result = await ImagePreCapture.CaptureAllAsync(new[] { 1, 2 },
                (image, ordinal, remaining) => { seen.Add(remaining); return Task.FromResult(new byte[5]); },
                CancellationToken.None, limitBytes: 10);
            CollectionAssert.AreEqual(new long[] { 10, 5 }, seen);
            Assert.AreEqual(2, result.Count);
        }
        [TestMethod]
        public async Task BatchOverflow_DoesNotReturnPartialBatchOrInvokeDownstream()
        {
            var downstream = 0;
            var captured = 0;
            try
            {
                await ImagePreCapture.CaptureAllAsync(new[] { 1, 2, 3 },
                    (image, ordinal, remaining) => { captured++; return Task.FromResult(new byte[6]); },
                    CancellationToken.None, limitBytes: 10);
                downstream++;
                Assert.Fail("overflow must abort before the downstream stage");
            }
            catch (ImageCaptureLimitException) { }
            Assert.AreEqual(2, captured);
            Assert.AreEqual(0, downstream);
        }
        [TestMethod]
        public async Task ExhaustedBatch_DoesNotCaptureNextImage()
        {
            var calls = 0;
            await Assert.ThrowsExceptionAsync<ImageCaptureLimitException>(async () =>
                await ImagePreCapture.CaptureAllAsync(new[] { 1, 2 },
                    (image, ordinal, remaining) => { calls++; return Task.FromResult(new byte[10]); },
                    CancellationToken.None, limitBytes: 10));
            Assert.AreEqual(1, calls);
        }
        [TestMethod]
        public async Task CancellationAfterCapture_AbortsBeforeAcceptingBatch()
        {
            using (var cancel = new CancellationTokenSource())
            {
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(async () =>
                    await ImagePreCapture.CaptureAllAsync(new[] { 1 },
                        (image, ordinal, remaining) => { cancel.Cancel(); return Task.FromResult(new byte[1]); }, cancel.Token));
            }
        }
    }
}
