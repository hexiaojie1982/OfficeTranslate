using System;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class SecurityImageRegionTests
    {
        [DataTestMethod]
        [DataRow(0)] [DataRow(511)] [DataRow(512)] [DataRow(513)] [DataRow(1500)]
        public void Regions_AreCappedInOrder_AndExcessCountsSurviveOcrAdaptation(int count)
        {
            var content = "{\"regions\":[" + string.Join(",", Enumerable.Range(0, count).Select(i =>
                "{\"bbox\":[0,0,100,100],\"source\":\"source" + i + "\",\"translation\":\"译文\"}")) + "]}";
            var parsed = ImageRegionParser.Parse(content);
            foreach (var result in new[] { parsed, ImageRecognitionText.FromImageResponse(parsed) })
            {
                var retained = Math.Min(count, ImageRegionParser.MaximumRegions);
                var dropped = count - retained;
                Assert.AreEqual(retained, result.Regions.Count); Assert.AreEqual(count, result.Statistics.RawCount);
                Assert.AreEqual(retained, result.Statistics.RetainedCount);
                Assert.AreEqual(dropped, result.Statistics.RegionLimitDroppedCount);
                Assert.AreEqual(dropped, result.Statistics.DroppedCount); Assert.AreEqual(dropped > 0, result.RegionsTruncated);
                Assert.IsFalse(result.Truncated, "Region limiting must not overwrite model finish_reason semantics.");
                CollectionAssert.AreEqual(Enumerable.Range(0, retained).Select(i => "source" + i).ToArray(), result.Regions.Select(r => r.Source).ToArray());
                var summary = new TranslationTaskSummary(1); summary.RecordImageRecognition(result);
                Assert.AreEqual(dropped, summary.OcrDroppedRegions); Assert.AreEqual(dropped > 0 ? 1 : 0, summary.ImageNeedsReview);
                var log = ImageOverlayDiagnostics.FormatOcrSummary("Word", 10, 10, result.Statistics, "stop", false, "parsed", 1);
                StringAssert.Contains(log, "region_limit_dropped=" + dropped);
            }
        }

        [TestMethod]
        public void Regions_LimitCountsRawEntriesIncludingInvalidOnes()
        {
            var parsed = ImageRegionParser.Parse("{\"regions\":[" + string.Join(",", Enumerable.Repeat("null", 600)) + "]}");
            Assert.AreEqual(0, parsed.Regions.Count); Assert.AreEqual(512, parsed.Statistics.NonDictionaryCount);
            Assert.AreEqual(88, parsed.Statistics.RegionLimitDroppedCount); Assert.AreEqual(600, parsed.Statistics.DroppedCount);
            var summary = new TranslationTaskSummary(1); summary.RecordImageRecognition(parsed);
            Assert.AreEqual(1, summary.ImageNeedsReview);
        }

        [TestMethod]
        public void Regions_DirectParserRejectsOversizeBeforeDeserialization()
        {
            Assert.ThrowsException<InvalidOperationException>(() => ImageRegionParser.Parse(new string(' ', ResponseLimits.Image + 1)));
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void HostNoteWritebackCallbacks_StopOnCancellationBeforeTouchingCom(string host)
        {
            var service = DocumentChatHostTests.Host(host).GetType("OfficeTranslate." + host + "AddIn." + host + "TranslationService", true)!;
            Type Root(Type t) => t.DeclaringType == null ? t : Root(t.DeclaringType);
            var callbacks = service.Assembly.GetTypes().Where(t => Root(t) == service)
                .SelectMany(t => t.GetMethods(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.DeclaredOnly))
                .Where(m => m.ReturnType == typeof(float) && m.GetParameters().Select(p => p.ParameterType)
                    .SequenceEqual(new[] { typeof(ImageOverlayNoteChunk), typeof(float) })).ToArray();
            Assert.AreEqual(2, callbacks.Length, "Both OCR and image-translation note loops must be covered.");
            using (var cancel = new CancellationTokenSource())
            {
                cancel.Cancel();
                object Closure(Type type)
                {
                    var closure = FormatterServices.GetUninitializedObject(type);
                    foreach (var field in type.GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
                    {
                        if (field.FieldType == typeof(CancellationToken)) field.SetValue(closure, cancel.Token);
                        else if (field.FieldType.Name.Contains("DisplayClass")) field.SetValue(closure, Closure(field.FieldType));
                    }
                    return closure;
                }
                foreach (var callback in callbacks)
                {
                    var error = Assert.ThrowsException<TargetInvocationException>(() => callback.Invoke(Closure(callback.DeclaringType!), new object?[] { null, 0f }));
                    Assert.IsInstanceOfType<OperationCanceledException>(error.InnerException);
                }
            }
        }

        [DataTestMethod]
        [DataRow("Word", false)] [DataRow("Excel", false)] [DataRow("PowerPoint", false)]
        [DataRow("Word", true)] [DataRow("Excel", true)] [DataRow("PowerPoint", true)]
        public void TruncatedRegions_AreVisibleAndRequireAcknowledgement(string host, bool ocr)
        {
            DocumentChatHostTests.Sta(() => {
                var result = ImageRegionParser.Parse("{\"regions\":[" + string.Join(",", Enumerable.Repeat("null", 513)) + "]}");
                var summary = new TranslationTaskSummary(1); summary.RecordImageRecognition(result);
                var type = DocumentChatHostTests.Host(host).GetType("OfficeTranslate.WordAddIn." + (ocr ? "OcrResultForm" : "TranslationResultForm"), true)!;
                using (var form = (Form)Activator.CreateInstance(type, BindingFlags.NonPublic | BindingFlags.Instance, null,
                    new object?[] { "en", summary, null }, null)!)
                {
                    var root = form.Controls.OfType<TableLayoutPanel>().Single();
                    var details = root.Controls.OfType<TextBox>().Single().Text;
                    StringAssert.Contains(details, "first 512 regions"); StringAssert.Contains(details, "1 excess regions were not written");
                    var field = type.GetField(ocr ? "_keepOpen" : "_requiresAcknowledgement", BindingFlags.Instance | BindingFlags.NonPublic)!;
                    Assert.AreEqual(true, field.GetValue(form));
                }
            });
        }
    }
}
