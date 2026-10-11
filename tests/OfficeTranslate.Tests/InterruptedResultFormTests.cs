using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using OfficeTranslate.WordAddIn;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class InterruptedResultFormTests
    {
        [DataTestMethod]
        [DataRow("zh-CN", false)]
        [DataRow("zh-HK", true)]
        [DataRow("ko", false)]
        [DataRow("ja", true)]
        [DataRow("en", false)]
        public void RecoveryWindow_PreservesFullLongTextAndErrorWithNoCountdown(string language, bool cancel)
        {
            OnSta(() => {
                var text = "BEGIN\r\n" + new string('译', 40000) + "\rEND";
                var summary = new TranslationTaskSummary(2);
                summary.RecordImageNoteFailure("owner-1", text, "note placement failed");
                using (var token = new CancellationTokenSource())
                {
                    if (cancel) token.Cancel();
                    Exception error = cancel ? (Exception)new OperationCanceledException(token.Token) : new TimeoutException("second image failed");
                    TranslationErrorDialog.ShowTaskFailure(error, summary, token.Token, language, null);
                    var form = Application.OpenForms.Cast<Form>().OfType<TranslationResultForm>().Single();
                    try
                    {
                        var controls = Descendants(form).ToArray();
                        var details = controls.OfType<TextBox>().Single();
                        StringAssert.Contains(details.Text, text.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r\n"));
                        if (!cancel) StringAssert.Contains(details.Text, error.Message);
                        Assert.IsTrue(controls.OfType<Label>().Any(l => l.Text == UiText.Get(language, cancel ? "SummaryCancelled" : "SummaryFailed")));
                        Assert.IsFalse(controls.OfType<Label>().Any(l => l.Text == string.Format(UiText.Get(language, "SummaryCompleted"), 2)));
                        Assert.AreEqual(UiText.Get(language, "Close"), controls.OfType<Button>().Single().Text);
                        Assert.IsTrue(details.ReadOnly); Assert.AreEqual(0, details.MaxLength);
                        var timer = (System.Windows.Forms.Timer)typeof(TranslationResultForm).GetField("_timer", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(form)!;
                        Assert.IsFalse(timer.Enabled, "Recovery content must not disappear after five seconds.");
                    }
                    finally { form.Close(); form.Dispose(); }
                }
            });
        }

        [TestMethod]
        public void OrdinarySuccess_KeepsFiveSecondCountdown()
        {
            OnSta(() => {
                TranslationResultForm.ShowResult("en", new TranslationTaskSummary(1), IntPtr.Zero);
                var form = Application.OpenForms.Cast<Form>().OfType<TranslationResultForm>().Single();
                try
                {
                    var timer = (System.Windows.Forms.Timer)typeof(TranslationResultForm).GetField("_timer", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(form)!;
                    Assert.IsTrue(timer.Enabled);
                    Assert.AreEqual("OK (closes in 5s)", Descendants(form).OfType<Button>().Single().Text);
                }
                finally { form.Close(); form.Dispose(); }
            });
        }

        [TestMethod]
        public void CancelWithoutNotes_DoesNotOpenAnyDialog()
        {
            OnSta(() => {
                using (var token = new CancellationTokenSource())
                {
                    token.Cancel();
                    var before = Application.OpenForms.Count;
                    TranslationErrorDialog.ShowTaskFailure(new OperationCanceledException(token.Token), null, token.Token, "en", null);
                    Assert.AreEqual(before, Application.OpenForms.Count);
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word", false)]
        [DataRow("Word", true)]
        [DataRow("Excel", false)]
        [DataRow("Excel", true)]
        [DataRow("PowerPoint", false)]
        [DataRow("PowerPoint", true)]
        public void CompiledHostDll_FailurePresenterShowsRecoverableText(string host, bool cancel)
        {
            // Build the complete solution first. Exercise each actual host DLL,
            // not just the linked test copy; no COM or real documents involved.
            var configuration = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent!.Name;
            var path = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "../../../../../src/OfficeTranslate." + host + "AddIn/bin/" + configuration + "/net48/OfficeTranslate." + host + "AddIn.dll"));
            Assert.IsTrue(File.Exists(path), "Run build.ps1 -SkipInstaller first: " + path);
            var assembly = Assembly.LoadFrom(path);
            OnSta(() => {
                var summary = new TranslationTaskSummary(2);
                summary.RecordImageNoteFailure("image-1", "FULL RECOVERABLE TEXT\r\n第二行", "layout failure");
                using (var token = new CancellationTokenSource())
                {
                    if (cancel) token.Cancel();
                    Exception error = cancel ? (Exception)new OperationCanceledException(token.Token) : new TimeoutException("next image failed");
                    var presenter = assembly.GetType("OfficeTranslate.WordAddIn.TranslationErrorDialog", true)!;
                    presenter.GetMethod("ShowTaskFailure", BindingFlags.Public | BindingFlags.Static)!.Invoke(null,
                        new object?[] { error, summary, token.Token, "en", null });
                    var form = Application.OpenForms.Cast<Form>().Single(f => f.GetType().Assembly == assembly);
                    try
                    {
                        StringAssert.Contains(Descendants(form).OfType<TextBox>().Single().Text, "FULL RECOVERABLE TEXT\r\n第二行");
                        Assert.IsTrue(Descendants(form).OfType<Label>().Any(l => l.Text == UiText.Get("en", cancel ? "SummaryCancelled" : "SummaryFailed")));
                        var timer = (System.Windows.Forms.Timer)form.GetType().GetField("_timer", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(form)!;
                        Assert.IsFalse(timer.Enabled);
                    }
                    finally { form.Close(); form.Dispose(); }
                }
            });
        }

        private static IEnumerable<Control> Descendants(Control control)
        {
            foreach (Control child in control.Controls)
            {
                yield return child;
                foreach (var nested in Descendants(child)) yield return nested;
            }
        }

        private static void OnSta(Action action)
        {
            Exception? error = null;
            var thread = new Thread(() => { try { action(); } catch (Exception ex) { error = ex; } });
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA); thread.Start();
            Assert.IsTrue(thread.Join(15000), "STA form check must finish without modal blocking.");
            if (error != null) throw error;
        }
    }
}
