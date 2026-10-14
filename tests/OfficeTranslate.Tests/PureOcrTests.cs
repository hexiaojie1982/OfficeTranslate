using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class PureOcrTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private sealed class Handler : HttpMessageHandler
        {
            public readonly List<string> Bodies = new List<string>();
            public readonly List<string> Paths = new List<string>();
            public Func<int, CancellationToken, Task<HttpResponseMessage>> Send = null!;
            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
            {
                Bodies.Add(await request.Content.ReadAsStringAsync());
                Paths.Add(request.RequestUri.AbsolutePath);
                return await Send(Bodies.Count, token);
            }
        }
        private static byte[] Png()
        { using (var b = new Bitmap(400, 200)) return ImageCaptureBudget.EncodePng(b, 102400); }
        private static TranslationSettings Settings(ProviderKind provider) => new TranslationSettings
        {
            Provider = provider, BaseUrl = "http://ocr-" + Guid.NewGuid().ToString("N") + ".invalid/v1", Model = "text-model", ImageModel = "vision-model",
            SourceLanguage = "SOURCE_SENTINEL", TargetLanguage = "TARGET_SENTINEL", BilingualMode = true,
            ImageOcrEnabled = false, Glossary = "GLOSSARY_SENTINEL", CustomInstructions = "CUSTOM_SENTINEL"
        };
        private static HttpResponseMessage Reply(ProviderKind provider, string content, string finish = "stop") =>
            new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(provider == ProviderKind.Ollama
                ? Json.Serialize(new { message = new { content }, done_reason = finish })
                : Json.Serialize(new { choices = new[] { new { message = new { content }, finish_reason = finish } } })) };
        private static string Regions(string source, string translation = "UNWANTED_TRANSLATION") =>
            Json.Serialize(new { regions = new[] { new { bbox = new[] { 100, 100, 800, 400 }, source, translation } } });
        private static ImageRecognitionResult Parse(string content) => ImageRecognitionText.FromImageResponse(ImageRegionParser.Parse(content));

        [DataTestMethod]
        [DataRow(ProviderKind.OpenAiCompatible)] [DataRow(ProviderKind.Ollama)]
        public async Task OneImageRequest_PreservesSourceAndIgnoresTranslationSettings(ProviderKind provider)
        {
            const string source = "한글 English 中文\nCase 123!";
            var handler = new Handler { Send = (n, t) => Task.FromResult(Reply(provider, Regions(source))) };
            using (var client = new TranslationClient(handler))
            {
                var result = await client.RecognizeImageAsync(Png(), Settings(provider), CancellationToken.None);
                Assert.AreEqual(source, result.Regions[0].Source);
                Assert.AreEqual(string.Empty, result.Regions[0].Translation);
                Assert.IsFalse(result.Regions[0].RequiresReview);
                Assert.AreEqual(0, result.Statistics.EmptyTranslationCount);
                var summary = new TranslationTaskSummary(1); summary.RecordImageRecognition(result);
                Assert.AreEqual(0, summary.Translated);
                Assert.AreEqual(1, summary.OcrKeptRegions);
            }
            Assert.AreEqual(1, handler.Bodies.Count);
            Assert.AreEqual(provider == ProviderKind.Ollama ? "/v1/api/chat" : "/v1/chat/completions", handler.Paths[0]);
            var body = Json.Deserialize<Dictionary<string, object>>(handler.Bodies[0]);
            Assert.AreEqual("vision-model", body["model"]);
            Assert.AreEqual(8192, provider == ProviderKind.Ollama ? ((Dictionary<string, object>)body["options"])["num_predict"] : body["max_tokens"]);
            StringAssert.Contains(handler.Bodies[0], "禁止翻译");
            foreach (var sentinel in new[] { "SOURCE_SENTINEL", "TARGET_SENTINEL", "GLOSSARY_SENTINEL", "CUSTOM_SENTINEL" })
                Assert.IsFalse(handler.Bodies[0].Contains(sentinel));
        }

        [TestMethod]
        public async Task ImageModelOnly_DoesNotRequireTextTranslationConfiguration()
        {
            var settings = Settings(ProviderKind.OpenAiCompatible);
            settings.Model = ""; settings.TargetLanguage = ""; settings.MaxParallelCells = 0; settings.MaxCharactersPerChunk = 0;
            var handler = new Handler { Send = (n, t) => Task.FromResult(Reply(settings.Provider, "{\"regions\":[]}")) };
            using (var client = new TranslationClient(handler))
                Assert.AreEqual(0, (await client.RecognizeImageAsync(Png(), settings, CancellationToken.None)).Regions.Count);
            Assert.AreEqual(1, handler.Bodies.Count);
        }

        [TestMethod]
        public void SourceOnlyRegion_PlansOverlayWithoutMissingTranslationReview()
        {
            var result = Parse("{\"regions\":[{\"bbox\":[100,100,800,400],\"source\":\"原文 ABC 123\"}]}");
            var region = result.Regions[0];
            Assert.IsFalse(region.RequiresReview);
            var plan = ImageOverlayPlanner.Plan(region.X1, region.Y1, region.X2, region.Y2, 400, 200, 400, 200, region.Source, 0);
            Assert.AreEqual(ImageOverlayVerdict.Place, plan.Verdict);
            Assert.IsTrue(plan.Left >= 0 && plan.Top >= 0 && plan.Left + plan.Width <= 400 && plan.Top + plan.Height <= 200);
            Assert.IsTrue(plan.FontSize >= ImageOverlayLayout.MinFontSize);
        }

        [TestMethod]
        public void InvalidBoxAndMissingSource_ArePreservedAsNoteOrDroppedCorrectly()
        {
            var result = Parse("{\"regions\":[{\"source\":\"完整原文\"},{\"translation\":\"discard\",\"bbox\":[0,0,500,500]},null,{\"source\":\"\"}]}");
            Assert.AreEqual(4, result.Statistics.RawCount);
            Assert.AreEqual(1, result.Regions.Count);
            Assert.AreEqual(3, result.Statistics.DroppedCount);
            Assert.AreEqual(1, result.Statistics.InvalidBboxDowngradedCount);
            Assert.AreEqual(0, result.Statistics.EmptyTranslationCount);
            Assert.AreEqual("完整原文", result.Regions[0].Source);
            Assert.IsTrue(result.Regions[0].RequiresReview);
        }

        [TestMethod]
        public void UnfittableSource_FallsBackToNotesAndRetainsFullRecoveryText()
        {
            var source = new string('字', 18000);
            var region = Parse(Regions(source)).Regions[0];
            var plan = ImageOverlayPlanner.Plan(region.X1, region.Y1, region.X2, region.Y2, 400, 200, 40, 20, region.Source, 0);
            Assert.AreEqual(ImageOverlayVerdict.SideNote, plan.Verdict);
            var summary = new TranslationTaskSummary(1);
            var recovered = new List<string>();
            ImageOverlayNotes.Place(new[] { region.Source }, ImageOverlayNotes.NoteWidth(40), 26,
                (chunk, top) => throw new InvalidOperationException("host layout failure"),
                (chunk, error) => {
                    foreach (var fragment in chunk.Fragments) recovered.Add(fragment.Text);
                    summary.RecordImageNoteFailure("wd:test", ImageRecognitionText.NoteText(chunk), error.Message);
                });
            Assert.AreEqual(source, string.Concat(recovered));
            Assert.IsTrue(summary.UnplacedImageNotes.Count > 0);
            StringAssert.Contains(summary.UnplacedImageNotes[0], "【图片原文待检查】");
        }

        [DataTestMethod]
        [DataRow("wd:Picture 1")] [DataRow("xl:Picture 1")] [DataRow("ppt:Picture 1")]
        public void Markers_CleanupIsIsolatedInBothDirections(string owner)
        {
            foreach (var marker in new[] { ImageRecognitionTags.OverlayFor(owner), ImageRecognitionTags.NoteFor(owner) })
            {
                Assert.IsTrue(ImageRecognitionTags.WordExcelMarkerBelongsTo(marker, owner));
                Assert.IsFalse(ImageRecognitionTags.WordExcelMarkerBelongsTo(marker, owner + "2"));
                Assert.IsFalse(ImageOverlayTags.WordExcelMarkerBelongsTo(marker, owner));
                Assert.IsTrue(ImageOverlayTags.IsOwnMarker(marker));
            }
            foreach (var marker in new[] { ImageOverlayTags.OverlayFor(owner), ImageOverlayTags.NoteFor(owner) })
                Assert.IsFalse(ImageRecognitionTags.WordExcelMarkerBelongsTo(marker, owner));
            Assert.IsFalse(ImageOverlayTags.PowerPointTagBelongsTo(ImageRecognitionTags.PowerPointOverlayFor(owner), owner));
            Assert.IsFalse(ImageOverlayTags.PowerPointTagBelongsTo(ImageRecognitionTags.PowerPointNoteFor(owner), owner));
        }

        [DataTestMethod]
        [DataRow(ProviderKind.OpenAiCompatible)] [DataRow(ProviderKind.Ollama)]
        public async Task StructuredOutputFallback_UsesSameOcrPromptAndBudget(ProviderKind provider)
        {
            var handler = new Handler { Send = (n, t) => Task.FromResult(n == 1
                ? new HttpResponseMessage((HttpStatusCode)422) { Content = new StringContent(provider == ProviderKind.Ollama ? "unsupported \"format\"" : "unsupported response_format") }
                : Reply(provider, Regions("original"))) };
            using (var client = new TranslationClient(handler)) await client.RecognizeImageAsync(Png(), Settings(provider), CancellationToken.None);
            Assert.AreEqual(2, handler.Bodies.Count);
            var first = Json.Deserialize<Dictionary<string, object>>(handler.Bodies[0]);
            var second = Json.Deserialize<Dictionary<string, object>>(handler.Bodies[1]);
            Assert.AreEqual(Json.Serialize(first["messages"]), Json.Serialize(second["messages"]));
            Assert.IsFalse(second.ContainsKey(provider == ProviderKind.Ollama ? "format" : "response_format"));
        }

        [DataTestMethod]
        [DataRow(429)] [DataRow(503)]
        public async Task TransientImageFailure_UsesExistingRetry(int status)
        {
            var handler = new Handler { Send = (n, t) => {
                var response = n == 1 ? new HttpResponseMessage((HttpStatusCode)status) { Content = new StringContent("temporary") }
                    : Reply(ProviderKind.OpenAiCompatible, Regions("retry original"));
                response.Headers.RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromMilliseconds(10));
                return Task.FromResult(response);
            } };
            using (var client = new TranslationClient(handler))
                Assert.AreEqual("retry original", (await client.RecognizeImageAsync(Png(), Settings(ProviderKind.OpenAiCompatible), CancellationToken.None)).Regions[0].Source);
            Assert.AreEqual(2, handler.Bodies.Count);
        }

        [DataTestMethod]
        [DataRow(ProviderKind.OpenAiCompatible)] [DataRow(ProviderKind.Ollama)]
        public async Task Deadline_UsesImageTimeoutWithoutRetry(ProviderKind provider)
        {
            var handler = new Handler { Send = async (n, token) => { await Task.Delay(Timeout.Infinite, token); return Reply(provider, "{}"); } };
            using (var client = new TranslationClient(handler, TimeSpan.FromMilliseconds(60)))
                await Assert.ThrowsExceptionAsync<ImageOcrTimeoutException>(() => client.RecognizeImageAsync(Png(), Settings(provider), CancellationToken.None));
            Assert.AreEqual(1, handler.Bodies.Count);
        }

        [TestMethod]
        public async Task UserCancellation_DoesNotRetryOrReturnAResult()
        {
            var started = new TaskCompletionSource<bool>();
            var handler = new Handler { Send = async (n, token) => { started.SetResult(true); await Task.Delay(Timeout.Infinite, token); return Reply(ProviderKind.OpenAiCompatible, "{}"); } };
            using (var cts = new CancellationTokenSource())
            using (var client = new TranslationClient(handler))
            {
                var pending = client.RecognizeImageAsync(Png(), Settings(ProviderKind.OpenAiCompatible), cts.Token);
                await started.Task; cts.Cancel();
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => pending);
            }
            Assert.AreEqual(1, handler.Bodies.Count);
        }

        [DataTestMethod]
        [DataRow(ProviderKind.OpenAiCompatible)] [DataRow(ProviderKind.Ollama)]
        public async Task TruncatedResponse_IsRejectedBeforeWriteback(ProviderKind provider)
        {
            var handler = new Handler { Send = (n, t) => Task.FromResult(Reply(provider, Regions("partial"), "length")) };
            using (var client = new TranslationClient(handler))
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.RecognizeImageAsync(Png(), Settings(provider), CancellationToken.None));
        }

        [TestMethod]
        public async Task InvalidJson_IsRejectedBeforeWriteback()
        {
            var handler = new Handler { Send = (n, t) => Task.FromResult(Reply(ProviderKind.OpenAiCompatible, "{\"regions\":[")) };
            using (var client = new TranslationClient(handler))
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.RecognizeImageAsync(Png(), Settings(ProviderKind.OpenAiCompatible), CancellationToken.None));
        }
    }
}
