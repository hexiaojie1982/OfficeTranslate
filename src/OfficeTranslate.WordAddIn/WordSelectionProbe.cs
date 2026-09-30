using Microsoft.Office.Interop.Word;
using OfficeTranslate.Core;
using System;
using WordApplication = Microsoft.Office.Interop.Word.Application;

namespace OfficeTranslate.WordAddIn
{
    // S1: task-ending boundary probes. The service finally can log
    // selection_restore_diag matched=1 while the selection observed when
    // the error dialog appears is already back on the failed image, so
    // something in the AddIn teardown moves it after the restore. These
    // metadata-only snapshots bracket each teardown step (before/after
    // the progress close, before/after the error dialog) so the review
    // can pinpoint which step moves the selection. Never throws, never
    // touches or changes the selection.
    internal static class WordSelectionProbe
    {
        public static void Log(WordApplication? word, string stage, IntPtr taskWindow)
        {
            try
            {
                string state;
                try { state = Snapshot(word); }
                catch { state = "?"; }
                ImageOverlayDiagnostics.LogCaptureFailure("Word",
                    "selection_boundary_probe stage=" + stage
                    + " taskHwnd=" + taskWindow.ToInt64().ToString("X")
                    + " sel=[" + state + "]");
            }
            catch { }
        }

        private static string Snapshot(WordApplication? word)
        {
            // Same field order as WordTranslationService.SnapshotSelectionState
            // so the two logs compare directly.
            if (word == null) return "?";
            string doc = "?", story = "?", range = "?-?", type = "?";
            Selection? sel = SafeGet<Selection?>(() => word.Selection, null);
            if (sel != null)
            {
                doc = SafeGet(() => (sel.Range.Parent as Document)?.Name ?? "?", "?");
                story = SafeGet(() => ((int)sel.Range.StoryType).ToString(), "?");
                range = SafeGet(() => sel.Range.Start, -1) + "-" + SafeGet(() => sel.Range.End, -1);
                type = SafeGet(() => sel.Type.ToString(), "?");
            }
            string view = SafeGet(() => word.ActiveWindow.View.Type.ToString(), "?");
            string winCap = SafeGet(() => Flatten(word.ActiveWindow.Caption), "?");
            string wins = SafeGet(() => word.Windows.Count.ToString(), "?");
            string docs = SafeGet(() => word.Documents.Count.ToString(), "?");
            return "doc=" + doc + " story=" + story + " range=" + range + " type=" + type
                + " view=" + view + " winCap=" + winCap + " wins=" + wins + " docs=" + docs;
        }

        private static T SafeGet<T>(Func<T> read, T fallback)
        {
            try { return read(); }
            catch { return fallback; }
        }

        private static string Flatten(string? caption)
        {
            if (string.IsNullOrEmpty(caption)) return "?";
            string flat = caption.Replace("\r", " ").Replace("\n", " ");
            return flat.Length > 60 ? flat.Substring(0, 60) : flat;
        }
    }
}
