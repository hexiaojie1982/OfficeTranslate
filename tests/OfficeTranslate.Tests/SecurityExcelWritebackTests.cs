using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class SecurityExcelWritebackTests
    {
        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        [TestCategory("OfficeIntegration")]
        public async Task RealTranslationTargets_WriteLiteralValuesInExcelWithoutFormulas(bool bilingual)
        {
            using (var pump = OfficeUiDispatcherTests.StartPumpingSta())
            {
                object? application = null, workbook = null;
                try
                {
                    await pump.Dispatcher.InvokeAsync(() => {
                        dynamic excel = application = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)!)!;
                        excel.Visible = false; excel.DisplayAlerts = false; excel.EnableEvents = false;
                        workbook = excel.Workbooks.Add();
                    });
                    foreach (var translation in new[] { "=2+2", "+2+2", "-2+2", "@SUM(A1:A2)", "正常译文", "00123", "1/2", "'leading quote" })
                    {
                        var targets = await pump.Dispatcher.InvokeAsync(() => {
                            dynamic sheet = ((dynamic)workbook!).Worksheets[1];
                            // Bilingual results beginning with = must also remain literal.
                            sheet.Range["A1"].Value2 = bilingual ? "'=original" : "original";
                            var method = DocumentChatHostTests.Host("Excel").GetType("OfficeTranslate.ExcelAddIn.ExcelTranslationService", true)!
                                .GetMethod("ReadRangeTargets", BindingFlags.NonPublic | BindingFlags.Static)!;
                            return (List<ExcelTextTarget>)method.Invoke(null, new object[] { sheet.Range["A1"] })!;
                        });
                        Assert.AreEqual(1, targets.Count);
                        await ExcelTextTranslation.RunAsync(targets, new TranslationSettings { BilingualMode = bilingual },
                            new TranslationTaskSummary(1), (_, token) => Task.FromResult(TranslationResult.Translated(translation)),
                            CancellationToken.None, _ => { }, pump.Dispatcher);
                        await pump.Dispatcher.InvokeAsync(() => {
                            dynamic cell = ((dynamic)workbook!).Worksheets[1].Range["A1"];
                            Assert.AreEqual(false, (bool)cell.HasFormula, translation);
                            Assert.AreEqual(bilingual ? "=original\n" + translation : translation, ((string)cell.Value2).Replace("\r\n", "\n"));
                            if (bilingual) Assert.AreEqual(true, (bool)cell.WrapText);
                        });
                    }
                }
                finally
                {
                    await pump.Dispatcher.InvokeAsync(() => {
                        if (workbook != null) { ((dynamic)workbook).Close(false); Marshal.FinalReleaseComObject(workbook); }
                        if (application != null) { ((dynamic)application).Quit(); Marshal.FinalReleaseComObject(application); }
                    });
                }
            }
        }

        [TestMethod]
        public void SharedLiteralHelper_PreservesLengthBoundaryAndRejectsOverflow()
        {
            Assert.AreEqual("'" + new string('文', 32767), ExcelCellText.Literal(new string('文', 32767)));
            Assert.ThrowsException<InvalidOperationException>(() => ExcelCellText.Literal(new string('文', 32768)));
        }
    }
}
