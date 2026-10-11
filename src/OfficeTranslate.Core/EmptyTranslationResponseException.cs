using System;

namespace OfficeTranslate.Core
{
    // A successful model response with no text, not a transport/auth/cancel
    // failure. Only the paragraph fallback boundary may recover from this.
    internal sealed class EmptyTranslationResponseException : InvalidOperationException
    {
        public EmptyTranslationResponseException() : base("翻译模型连续返回空内容，已停止写回。") { }
    }
}
