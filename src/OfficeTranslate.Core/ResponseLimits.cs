using System;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeTranslate.Core
{
    internal static class ResponseLimits
    {
        internal const int Chat = 8 * 1024 * 1024;
        internal const int Image = 16 * 1024 * 1024;

        internal static async Task<Stream> OpenStreamAsync(HttpContent content, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            var pending = content.ReadAsStreamAsync();
            if (pending.IsCompleted) return await pending.ConfigureAwait(false);
            // HttpClient's headers-only completion can detach the handler's token.
            // Keep OUR deadline effective even while custom content is opening its stream.
            var cancelled = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            using (token.Register(() => cancelled.TrySetResult(true)))
            {
                if (await Task.WhenAny(pending, cancelled.Task).ConfigureAwait(false) != pending)
                {
                    _ = pending.ContinueWith(task => {
                        if (task.Status == TaskStatus.RanToCompletion) task.Result.Dispose();
                        else if (task.IsFaulted) _ = task.Exception;
                    }, CancellationToken.None, TaskContinuationOptions.ExecuteSynchronously, TaskScheduler.Default);
                    token.ThrowIfCancellationRequested();
                }
                return await pending.ConfigureAwait(false);
            }
        }

        // Count actual bytes, including chunked bodies and inaccurate/missing Content-Length.
        internal static async Task BufferImageAsync(HttpResponseMessage response, CancellationToken token)
        {
            var content = response.Content;
            if (content.Headers.ContentLength > Image) throw ImageTooLarge();
            using (var stream = await OpenStreamAsync(content, token).ConfigureAwait(false))
            using (token.Register(() => stream.Dispose()))
            using (var limited = new LimitedReadStream(stream, Image, ImageTooLarge))
            using (var buffer = new MemoryStream())
            {
                await limited.CopyToAsync(buffer, 8192, token).ConfigureAwait(false);
                token.ThrowIfCancellationRequested();
                var bounded = new ByteArrayContent(buffer.ToArray());
                foreach (var header in content.Headers) bounded.Headers.TryAddWithoutValidation(header.Key, header.Value);
                response.Content = bounded;
            }
            content.Dispose();
        }

        private static Exception ImageTooLarge() => new InvalidOperationException("图片响应超过 16 MiB 上限，未写回。请分批识别或更换图片模型。");
    }

    internal sealed class LimitedReadStream : Stream
    {
        private readonly Stream _inner;
        private readonly int _limit;
        private readonly Func<Exception> _tooLarge;
        private int _read;
        internal LimitedReadStream(Stream inner, int limit, Func<Exception> tooLarge)
        { _inner = inner; _limit = limit; _tooLarge = tooLarge; }
        private int Count(int requested) => Math.Min(requested, _limit - _read + 1);
        private int Record(int count) { _read += count; if (_read > _limit) throw _tooLarge(); return count; }
        public override int Read(byte[] buffer, int offset, int count) => Record(_inner.Read(buffer, offset, Count(count)));
        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token) =>
            Record(await _inner.ReadAsync(buffer, offset, Count(count), token).ConfigureAwait(false));
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }
}
