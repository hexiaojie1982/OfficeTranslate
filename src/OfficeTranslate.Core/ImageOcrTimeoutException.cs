using System;
using System.Globalization;
using System.Threading;

namespace OfficeTranslate.Core
{
    // Common to all hosts; no OCR text, image bytes, endpoint or credentials.
    public sealed class ImageOcrTimeoutException : TimeoutException
    {
        public TimeSpan RequestTimeout { get; }
        public ImageOcrTimeoutException(TimeSpan requestTimeout, Exception inner)
            : base("图片识别请求等待 " + requestTimeout.TotalSeconds.ToString("0.###", CultureInfo.InvariantCulture) +
                " 秒后超时，服务器/模型未返回结果，未写入当前图片的任何内容；此前已完成的内容保留。" +
                "可重试、减小图片或改用更快的图片模型。", inner)
        { RequestTimeout = requestTimeout; }
    }

    public static class TranslationFailurePolicy
    {
        // OCE alone is not proof that the user clicked Cancel. In particular,
        // the final text HTTP timeout must not silently disappear either.
        public static bool IsUserCancellation(Exception error, CancellationToken userToken) =>
            error is OperationCanceledException && userToken.IsCancellationRequested;
    }
}
