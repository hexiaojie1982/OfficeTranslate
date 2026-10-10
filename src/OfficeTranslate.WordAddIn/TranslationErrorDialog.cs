using System;
using System.Windows.Forms;
using OfficeTranslate.Core;

namespace OfficeTranslate.WordAddIn
{
    // Linked into Excel/PowerPoint like the existing progress/result forms.
    // Hosts resolve their own live owner and dispatch here on the UI thread.
    internal static class TranslationErrorDialog
    {
        public static void Show(Exception error, IWin32Window? owner)
        {
            var icon = error is ImageOcrTimeoutException ? MessageBoxIcon.Warning : MessageBoxIcon.Error;
            if (owner != null) MessageBox.Show(owner, error.Message, "OfficeTranslate", MessageBoxButtons.OK, icon);
            else MessageBox.Show(error.Message, "OfficeTranslate", MessageBoxButtons.OK, icon);
        }
    }
}
