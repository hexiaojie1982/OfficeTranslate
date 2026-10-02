# Real Word queue acceptance helper

This project is NOT in the release solution or MSI. Never install/register this
assembly permanently: it deliberately reuses Word's existing COM GUID under a
temporary HKCU override. The runner refuses a running Word instance, exports any
existing HKCU key before changes and restores/verifies it on exit. It does not
write HKLM, change Office trust/policy settings, or use real customer documents.

Build this project against the same candidate Core as the production add-in.
In Windows PowerShell STA, run `RunOfficeQueueProbe.ps1 -EvidenceDirectory <dir>`.
When `WAITING_FOR_NORMAL_WORD_UI_STARTUP` appears, start Word normally from the
desktop within 60 seconds (not `New-Object Word.Application`: automation startup
may suppress COM add-ins). Keep the desktop unlocked. Use a fresh evidence folder.

The helper captures the dispatcher in Word's `OnConnection`, verifies the managed
callback is on Word's actual HWND native thread, and invokes the production
`UiRestoreGate`. It tests: 2.4s blocked queue then a newer real range; superseded
generation; and a normal one-shot restore. Metadata evidence includes the Core
MVID, native/UI thread IDs, callback count and final selection offsets.

This validates the shared gate on a real Office queue. It is not a substitute
for a complete translation/menu/OCR quality regression or a test of all Word
header/footer view transitions. Only synthetic temporary documents are created.
