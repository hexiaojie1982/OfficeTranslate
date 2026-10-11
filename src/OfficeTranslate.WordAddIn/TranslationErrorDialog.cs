using System;
using System.Threading;
using System.Windows.Forms;
using OfficeTranslate.Core;

namespace OfficeTranslate.WordAddIn
{
    // Linked into Excel/PowerPoint like the existing progress/result forms.
    // Hosts resolve their own live owner and dispatch here on the UI thread.
    internal static class TranslationErrorDialog
    {
        public static void ShowTaskFailure(Exception error, TranslationTaskSummary? summary,
            CancellationToken token, string language, IWin32Window? owner)
        {
            var failure = new TranslationTaskFailure(error, summary, token);
            if (failure.HasRecoverableNotes)
                TranslationResultForm.ShowResult(language, summary!, owner?.Handle ?? IntPtr.Zero, failure);
            else if (failure.ShowErrorDialog) Show(error, owner);
        }

        public static void Show(Exception error, IWin32Window? owner)
        {
            var icon = error is ImageOcrTimeoutException ? MessageBoxIcon.Warning : MessageBoxIcon.Error;
            if (owner != null) MessageBox.Show(owner, error.Message, "OfficeTranslate", MessageBoxButtons.OK, icon);
            else MessageBox.Show(error.Message, "OfficeTranslate", MessageBoxButtons.OK, icon);
        }
    }
}
