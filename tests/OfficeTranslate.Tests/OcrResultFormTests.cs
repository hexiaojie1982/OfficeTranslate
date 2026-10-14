using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using OfficeTranslate.WordAddIn;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class OcrResultFormTests
    {
        [DataTestMethod]
        [DataRow("Word", "zh-CN")]
        [DataRow("Excel", "en")]
        [DataRow("PowerPoint", "en")]
        public void SuccessWithoutNotes_StartsCountdownAndClosesAfterFiveSeconds(string host, string language)
        {
            OnSta(() => {
                var labels = new HashSet<string>();
                var watch = Stopwatch.StartNew();
                RunModal(host, language, new TranslationTaskSummary(1), null, form => {
                    Assert.IsTrue(Timer(form).Enabled);
                    Assert.AreEqual(1000, Timer(form).Interval);
                    labels.Add(Descendants(form).OfType<Button>().Single().Text);
                }, true);
                Assert.IsTrue(watch.Elapsed.TotalSeconds >= 4.5, "Must keep the result visible for five seconds.");
                Assert.IsTrue(watch.Elapsed.TotalSeconds < 10, "Must close automatically.");
                foreach (var seconds in Enumerable.Range(1, 5))
                    Assert.IsTrue(labels.Contains(string.Format(UiText.Get(language, "OcrResultCloseCountdown"), seconds)),
                        "Missing countdown label for " + seconds);
            });
        }

        [DataTestMethod]
        [DataRow("Word", "zh-CN", "notes")]
        [DataRow("Word", "en", "error")]
        [DataRow("Word", "en", "cancel")]
        [DataRow("Word", "en", "no-summary")]
        [DataRow("Excel", "en", "notes")]
        [DataRow("Excel", "zh-CN", "error")]
        [DataRow("Excel", "en", "cancel")]
        [DataRow("Excel", "en", "no-summary")]
        [DataRow("PowerPoint", "en", "notes")]
        [DataRow("PowerPoint", "zh-CN", "error")]
        [DataRow("PowerPoint", "en", "cancel")]
        [DataRow("PowerPoint", "en", "no-summary")]
        public void RecoveryResult_DoesNotStartCountdown(string host, string language, string state)
        {
            OnSta(() => {
                var summary = state == "no-summary" ? null : new TranslationTaskSummary(1);
                if (state == "notes") summary!.RecordImageNoteFailure("owner", "COPY ALL SOURCE TEXT", "layout failure");
                Exception? error = state == "notes" ? null : state == "cancel"
                    ? (Exception)new OperationCanceledException() : new InvalidOperationException("OCR failure");
                RunModal(host, language, summary, error, form => {
                    Assert.IsFalse(Timer(form).Enabled, "Errors and unplaced notes must remain copyable.");
                    Assert.AreEqual(UiText.Get(language, "Close"), Descendants(form).OfType<Button>().Single().Text);
                    if (state == "notes")
                        StringAssert.Contains(Descendants(form).OfType<TextBox>().Single().Text, "COPY ALL SOURCE TEXT");
                });
            });
        }

        [DataTestMethod]
        [DataRow("zh-CN")]
        [DataRow("zh-HK")]
        [DataRow("en")]
        [DataRow("ko")]
        [DataRow("ja")]
        public void AllHosts_LocalizeSuccessCancellationFailureAndStatistics(string language)
        {
            OnSta(() => {
                foreach (var host in new[] { "Word", "Excel", "PowerPoint" })
                foreach (var state in new[] { "success", "cancel", "error", "notes" })
                {
                    var summary = new TranslationTaskSummary(1);
                    summary.RecordImageRecognition(ImageRegionParser.Parse("{\"regions\":[{\"source\":\"SOURCE\",\"bbox\":[0,0,100,100]}]}"));
                    summary.RecordImageRecognition(ImageRegionParser.Parse("{\"regions\":[]}"));
                    summary.RecordImageNeedsReview();
                    if (state == "notes") summary.RecordImageNoteFailure("owner", "SOURCE\nSECOND LINE", "layout failure");
                    Exception? error = state == "cancel" ? (Exception)new OperationCanceledException()
                        : state == "error" ? new InvalidOperationException("OCR failure") : null;
                    var type = Host(host).GetType("OfficeTranslate.WordAddIn.OcrResultForm", true)!;
                    using (var form = (Form)Activator.CreateInstance(type, BindingFlags.Instance | BindingFlags.NonPublic,
                        null, new object?[] { language, summary, error }, null)!)
                    {
                        Assert.AreEqual(UiText.Get(language, "OcrResultTitle"), form.Text);
                        var headingKey = state == "cancel" ? "OcrResultCancelled" : state == "error"
                            ? "OcrResultFailed" : "OcrResultCompleted";
                        Assert.IsTrue(Descendants(form).OfType<Label>().Any(l => l.Text == UiText.Get(language, headingKey)));
                        var details = Descendants(form).OfType<TextBox>().Single();
                        StringAssert.Contains(details.Text, string.Format(UiText.Get(language, "OcrResultDetails"),
                            summary.OcrImages, summary.OcrKeptRegions, summary.OcrNoText, summary.ImageNeedsReview));
                        if (state == "cancel") StringAssert.Contains(details.Text, UiText.Get(language, "OcrResultCancelledDetails"));
                        if (state == "error") StringAssert.Contains(details.Text, string.Format(UiText.Get(language, "OcrResultFailedDetails"), error!.Message));
                        if (state == "notes")
                        {
                            StringAssert.Contains(details.Text, UiText.Get(language, "OcrResultUnplacedNotes"));
                            StringAssert.Contains(details.Text, "SOURCE\r\nSECOND LINE");
                        }
                        Assert.IsTrue(details.ReadOnly);
                        Assert.AreEqual(0, details.MaxLength);
                        if (language == "ko" || language == "ja")
                        {
                            Assert.AreEqual("OCR results", form.Text);
                            Assert.IsTrue(Descendants(form).OfType<Label>().Any(l => l.Text.StartsWith("OCR ")));
                        }
                    }
                }
            });
        }

        private static Assembly Host(string host)
        {
            var configuration = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent!.Name;
            return Assembly.LoadFrom(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "../../../../../src/OfficeTranslate." + host + "AddIn/bin/" + configuration + "/net48/OfficeTranslate." + host + "AddIn.dll")));
        }

        private static System.Windows.Forms.Timer Timer(Form form) =>
            (System.Windows.Forms.Timer)form.GetType().GetField("_timer", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(form)!;

        private static void RunModal(string host, string language, TranslationTaskSummary? summary, Exception? error,
            Action<Form> inspect, bool waitForAutomaticClose = false)
        {
            var type = Host(host).GetType("OfficeTranslate.WordAddIn.OcrResultForm", true)!;
            Exception? assertionError = null;
            Form? observed = null;
            var watch = Stopwatch.StartNew();
            using (var observer = new System.Windows.Forms.Timer { Interval = 50 })
            {
                observer.Tick += (s, e) => {
                    var form = Application.OpenForms.Cast<Form>().SingleOrDefault(f => f.GetType() == type);
                    if (form == null) return;
                    observed = form;
                    try
                    {
                        inspect(form);
                        if (watch.Elapsed.TotalSeconds >= 10) Assert.Fail("Result window did not close automatically.");
                    }
                    catch (Exception ex) { assertionError = ex; }
                    if (!waitForAutomaticClose || assertionError != null) form.Close();
                };
                observer.Start();
                type.GetMethod("ShowResult", BindingFlags.Public | BindingFlags.Static)!.Invoke(null,
                    new object?[] { language, summary, IntPtr.Zero, error });
            }
            Assert.IsNotNull(observed, "Must exercise the actual ShowResult entry point.");
            Assert.IsTrue(observed!.IsDisposed, "Modal result and timer must be disposed after closing.");
            Assert.IsFalse(Timer(observed).Enabled);
            if (assertionError != null) throw assertionError;
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
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            Assert.IsTrue(thread.Join(20000), "OCR form check must finish without modal blocking.");
            if (error != null) throw error;
        }
    }
}
