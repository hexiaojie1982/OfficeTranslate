using System;
using System.Runtime.InteropServices;

namespace OfficeTranslate.Core
{
    // Keep a failed COM overlay local to its region. Callers can deliver its
    // text through the normal side-note/recovery path without losing siblings.
    public static class ImageOverlayWriteback
    {
        public static bool TryCreate<T>(Func<T> create, Action<T> initialize, Action<T> remove,
            out T? shape, out string reason) where T : class
        {
            shape = null;
            reason = string.Empty;
            try
            {
                shape = create();
                initialize(shape);
                return true;
            }
            catch (Exception ex)
            {
                var cleanupFailed = false;
                if (shape != null)
                {
                    try { remove(shape); }
                    catch { cleanupFailed = true; }
                }
                shape = null;
                // Cancellation and non-COM programming/transport errors retain
                // their original meaning. No blanket "success" or retry.
                if (!(ex is COMException)) throw;
                reason = "创建译文覆盖框失败，已降级为旁注：" + ex.Message
                    + (cleanupFailed ? "（未能删除失败的框，请检查原图区域。）" : string.Empty);
                return false;
            }
        }
    }
}
