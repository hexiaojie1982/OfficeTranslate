using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class PureOcrHostTests
    {
        private static Assembly Host(string host)
        {
            var configuration = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent!.Name;
            return Assembly.LoadFrom(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "../../../../../src/OfficeTranslate." + host + "AddIn/bin/" + configuration + "/net48/OfficeTranslate." + host + "AddIn.dll")));
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void CompiledRibbon_ContainsOcrButtonAndCallableEntry(string host)
        {
            var assembly = Host(host);
            using (var stream = assembly.GetManifestResourceStream("OfficeTranslate." + host + "AddIn.Ribbon.xml"))
            {
                Assert.IsNotNull(stream);
                var xml = new XmlDocument(); xml.Load(stream!);
                var ns = new XmlNamespaceManager(xml.NameTable); ns.AddNamespace("r", xml.DocumentElement!.NamespaceURI);
                var button = (XmlElement)xml.SelectSingleNode("//r:button[@id='OfficeTranslate.RecognizeImage']", ns)!;
                Assert.IsNotNull(button); Assert.AreEqual("OCR", button.GetAttribute("label"));
                Assert.AreEqual("GetRibbonImage", button.GetAttribute("getImage"));
                Assert.IsFalse(button.HasAttribute("imageMso"));
                var addin = assembly.GetType("OfficeTranslate." + host + "AddIn.AddIn", true)!;
                Assert.IsNotNull(addin.GetMethod(button.GetAttribute("getImage"), BindingFlags.Public | BindingFlags.Instance));
                Assert.IsNotNull(addin.GetMethod(button.GetAttribute("onAction"), BindingFlags.Public | BindingFlags.Instance));
                Assert.IsNotNull(xml.SelectSingleNode("//r:toggleButton[@id='OfficeTranslate.ImageOcr']", ns));
            }
        }

        [DataTestMethod]
        [DataRow("Word", false)] [DataRow("Excel", false)] [DataRow("PowerPoint", false)]
        [DataRow("Word", true)] [DataRow("Excel", true)] [DataRow("PowerPoint", true)]
        public void CompiledOcrSummary_KeepsFullSourceAndDoesNotClaimTranslation(string host, bool interrupted)
        {
            Exception? failure = null;
            var thread = new Thread(() =>
            {
                try
                {
                    var source = "BEGIN\r\n" + new string('文', 40000) + "\r\nEND";
                    var summary = new TranslationTaskSummary(2);
                    summary.RecordImageNoteFailure("owner", source, "layout failure");
                    var type = Host(host).GetType("OfficeTranslate.WordAddIn.OcrResultForm", true)!;
                    using (var form = (Form)Activator.CreateInstance(type, BindingFlags.Instance | BindingFlags.NonPublic,
                        null, new object?[] { summary, interrupted ? new InvalidOperationException("next image failed") : null }, null)!)
                    {
                        Assert.AreEqual("OCR 识别结果", form.Text);
                        var details = Descendants(form).OfType<TextBox>().Single();
                        StringAssert.Contains(details.Text, source);
                        Assert.AreEqual(0, details.MaxLength); Assert.IsTrue(details.ReadOnly);
                        Assert.IsFalse(details.Text.Contains("已翻译"));
                        if (interrupted) StringAssert.Contains(details.Text, "next image failed");
                    }
                }
                catch (Exception ex) { failure = ex; }
            });
            thread.IsBackground = true; thread.SetApartmentState(ApartmentState.STA); thread.Start();
            Assert.IsTrue(thread.Join(15000));
            if (failure != null) throw failure;
        }

        [DataTestMethod]
        [DataRow("Word", false)] [DataRow("Excel", false)] [DataRow("PowerPoint", false)]
        [DataRow("Word", true)] [DataRow("Excel", true)] [DataRow("PowerPoint", true)]
        public void CompiledOcrSummary_MatchesTranslationResultStyle(string host, bool unplaced)
        {
            Exception? failure = null;
            var thread = new Thread(() =>
            {
                try
                {
                    var summary = new TranslationTaskSummary(2);
                    if (unplaced) summary.RecordImageNoteFailure("owner", "原文旁注", "layout failure");
                    var assembly = Host(host);
                    var ocrType = assembly.GetType("OfficeTranslate.WordAddIn.OcrResultForm", true)!;
                    var translationType = assembly.GetType("OfficeTranslate.WordAddIn.TranslationResultForm", true)!;
                    using (var ocr = (Form)Activator.CreateInstance(ocrType, BindingFlags.Instance | BindingFlags.NonPublic,
                        null, new object?[] { summary, null }, null)!)
                    using (var translation = (Form)Activator.CreateInstance(translationType, BindingFlags.Instance | BindingFlags.NonPublic,
                        null, new object?[] { "zh-CN", summary, null }, null)!)
                    {
                        Assert.AreEqual(translation.Size, ocr.Size);
                        Assert.AreEqual(translation.MinimumSize, ocr.MinimumSize);
                        Assert.AreEqual(translation.FormBorderStyle, ocr.FormBorderStyle);
                        Assert.AreEqual(translation.MaximizeBox, ocr.MaximizeBox);
                        Assert.AreEqual(translation.MinimizeBox, ocr.MinimizeBox);
                        Assert.AreEqual(translation.ShowInTaskbar, ocr.ShowInTaskbar);
                        AssertSameStyle(translation, ocr);

                        var ocrRoot = ocr.Controls.OfType<TableLayoutPanel>().Single();
                        var referenceRoot = translation.Controls.OfType<TableLayoutPanel>().Single();
                        AssertSameStyle(referenceRoot, ocrRoot);
                        Assert.AreEqual(referenceRoot.RowCount, ocrRoot.RowCount);
                        Assert.AreEqual(referenceRoot.ColumnCount, ocrRoot.ColumnCount);
                        for (var i = 0; i < referenceRoot.ColumnStyles.Count; i++)
                        {
                            Assert.AreEqual(referenceRoot.ColumnStyles[i].SizeType, ocrRoot.ColumnStyles[i].SizeType);
                            Assert.AreEqual(referenceRoot.ColumnStyles[i].Width, ocrRoot.ColumnStyles[i].Width);
                        }
                        for (var i = 0; i < referenceRoot.RowStyles.Count; i++)
                            Assert.AreEqual(referenceRoot.RowStyles[i].SizeType, ocrRoot.RowStyles[i].SizeType);

                        var ocrControls = Descendants(ocr).ToArray();
                        var referenceControls = Descendants(translation).ToArray();
                        var ocrLabels = ocrControls.OfType<Label>().ToArray();
                        var referenceLabels = referenceControls.OfType<Label>().ToArray();
                        Assert.AreEqual(referenceLabels.Length, ocrLabels.Length);
                        for (var i = 0; i < referenceLabels.Length; i++)
                            AssertSameStyle(referenceLabels[i], ocrLabels[i]);
                        Assert.AreEqual(referenceLabels[0].Text, ocrLabels[0].Text);

                        var details = ocrControls.OfType<TextBox>().Single();
                        var referenceDetails = referenceControls.OfType<TextBox>().Single();
                        AssertSameStyle(referenceDetails, details);
                        Assert.AreEqual(referenceDetails.BorderStyle, details.BorderStyle);
                        Assert.AreEqual(referenceDetails.ScrollBars, details.ScrollBars);
                        Assert.AreEqual(referenceDetails.Dock, details.Dock);
                        var close = ocrControls.OfType<Button>().Single();
                        var referenceClose = referenceControls.OfType<Button>().Single();
                        AssertSameStyle(referenceClose, close);
                        Assert.AreEqual(referenceClose.FlatStyle, close.FlatStyle);
                        Assert.AreEqual(referenceClose.FlatAppearance.BorderSize, close.FlatAppearance.BorderSize);
                        Assert.AreEqual(referenceClose.Anchor, close.Anchor);
                        Assert.AreEqual(referenceClose.AutoSize, close.AutoSize);
                        Assert.AreEqual("关闭", close.Text);
                        Assert.AreSame(close, ocr.AcceptButton);
                        Assert.AreSame(close, ocr.CancelButton);
                        Assert.AreEqual(DialogResult.OK, close.DialogResult);

                        ocr.PerformLayout();
                        ocrRoot.PerformLayout();
                        Assert.IsTrue(details.Height > 0 && details.Width > 0);
                        Assert.IsTrue(ocrLabels[1].Right <= ocrRoot.ClientSize.Width - ocrRoot.Padding.Right);
                        Assert.IsTrue(ocrLabels[0].Bottom <= details.Top);
                        Assert.IsTrue(details.Bottom <= close.Top);
                        Assert.IsTrue(close.Bottom <= ocrRoot.ClientSize.Height - ocrRoot.Padding.Bottom);
                    }
                }
                catch (Exception ex) { failure = ex; }
            });
            thread.IsBackground = true; thread.SetApartmentState(ApartmentState.STA); thread.Start();
            Assert.IsTrue(thread.Join(15000));
            if (failure != null) throw failure;
        }

        private static void AssertSameStyle(Control reference, Control actual)
        {
            Assert.AreEqual(reference.Font, actual.Font);
            Assert.AreEqual(reference.ForeColor, actual.ForeColor);
            Assert.AreEqual(reference.BackColor, actual.BackColor);
            Assert.AreEqual(reference.Margin, actual.Margin);
            Assert.AreEqual(reference.Padding, actual.Padding);
        }

        private static IEnumerable<Control> Descendants(Control control)
        {
            foreach (Control child in control.Controls)
            {
                yield return child;
                foreach (var descendant in Descendants(child)) yield return descendant;
            }
        }
    }
}
