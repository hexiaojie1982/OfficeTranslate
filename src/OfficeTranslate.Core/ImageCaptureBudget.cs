using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeTranslate.Core
{
    // Limits retained PNG data, not the memory that Office/Clipboard.GetImage
    // itself allocates. Keep that distinction explicit in UI and diagnostics.
    public sealed class ImageCaptureBudget
    {
        public const long DefaultLimitBytes = 512L * 1024 * 1024;
        public const long MaxSinglePngBytes = 128L * 1024 * 1024;
        public long LimitBytes { get; }
        public long UsedBytes { get; private set; }
        public long RemainingBytes => LimitBytes - UsedBytes;

        public ImageCaptureBudget(long limitBytes = DefaultLimitBytes)
        {
            if (limitBytes <= 0) throw new ArgumentOutOfRangeException(nameof(limitBytes));
            LimitBytes = limitBytes;
        }

        // Subtraction avoids overflow; rejected reservations never alter state.
        public void Add(long bytes)
        {
            if (bytes < 0) throw new ArgumentOutOfRangeException(nameof(bytes));
            if (bytes > RemainingBytes) throw new ImageCaptureLimitException();
            UsedBytes += bytes;
        }

        public static byte[] EncodePng(Image image, long maxBytes)
        {
            if (image == null) throw new ArgumentNullException(nameof(image));
            if (maxBytes <= 0) throw new ImageCaptureLimitException();
            if (image.Width > 8000 || image.Height > 8000 || (long)image.Width * image.Height > 32_000_000)
                throw new ImageCaptureLimitException("单张图片像素过大，请降低分辨率后重试。");
            using (var stream = new LimitedPngStream(Math.Min(maxBytes, MaxSinglePngBytes)))
            {
                try { image.Save(stream, ImageFormat.Png); }
                catch (Exception) when (stream.LimitExceeded) { throw new ImageCaptureLimitException(); }
                // Some GDI+ encoders swallow IStream write failures and return
                // success with a truncated stream. Never accept that output.
                if (stream.LimitExceeded) throw new ImageCaptureLimitException();
                return stream.ToArray();
            }
        }

        // GDI+ sometimes wraps stream exceptions in ExternalException. The
        // sticky flag lets EncodePng report the actual cause, not a vague GDI error.
        internal sealed class LimitedPngStream : MemoryStream
        {
            private readonly long _limit;
            public bool LimitExceeded { get; private set; }
            public LimitedPngStream(long limit)
            {
                if (limit <= 0 || limit > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(limit));
                _limit = limit;
            }
            private void Check(long end)
            {
                if (end < 0 || end > _limit)
                {
                    LimitExceeded = true;
                    throw new ImageCaptureLimitException();
                }
                // Avoid MemoryStream's automatic capacity doubling above the cap.
                if (end > Capacity) Capacity = (int)Math.Min(_limit, Math.Max(end, Math.Max(256L, (long)Capacity * 2)));
            }
            public override void Write(byte[] buffer, int offset, int count)
            {
                if (buffer == null) throw new ArgumentNullException(nameof(buffer));
                if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException();
                Check(Position + count);
                base.Write(buffer, offset, count);
            }
            public override void WriteByte(byte value) { Check(Position + 1); base.WriteByte(value); }
            public override void SetLength(long value) { Check(value); base.SetLength(value); }
        }
    }

    public sealed class ImageCaptureLimitException : InvalidOperationException
    {
        public ImageCaptureLimitException(string message = "图片捕获数据达到容量上限，已停止当前捕获。请分批翻译或降低图片分辨率。") : base(message) { }
    }

    public static class ImagePreCapture
    {
        // Capture is injectable per CALL, not a static production hook. The
        // host cannot send requests or commit until this entire batch succeeds.
        public static async Task<List<Tuple<T, byte[]>>> CaptureAllAsync<T>(IReadOnlyList<T> images,
            Func<T, int, long, Task<byte[]>> capture, CancellationToken token,
            Action<int, bool>? progress = null, long limitBytes = ImageCaptureBudget.DefaultLimitBytes)
        {
            if (images == null) throw new ArgumentNullException(nameof(images));
            if (capture == null) throw new ArgumentNullException(nameof(capture));
            var budget = new ImageCaptureBudget(limitBytes);
            var result = new List<Tuple<T, byte[]>>();
            for (var i = 0; i < images.Count; i++)
            {
                token.ThrowIfCancellationRequested();
                if (budget.RemainingBytes == 0) throw new ImageCaptureLimitException();
                progress?.Invoke(i + 1, false);
                var bytes = await capture(images[i], i + 1, budget.RemainingBytes);
                token.ThrowIfCancellationRequested();
                if (bytes == null) throw new InvalidOperationException("图片捕获返回空数据。");
                budget.Add(bytes.LongLength);
                result.Add(Tuple.Create(images[i], bytes));
                progress?.Invoke(i + 1, true);
            }
            return result;
        }
    }
}
