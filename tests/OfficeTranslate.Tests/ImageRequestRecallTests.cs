using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ImageRequestRecallTests
    {
        private sealed class FakeService : HttpMessageHandler
        {
            public readonly List<Dictionary<string, object>> Requests = new List<Dictionary<string, object>>();
            public readonly Queue<HttpResponseMessage> Responses = new Queue<HttpResponseMessage>();
            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
            {
                token.ThrowIfCancellationRequested();
                Requests.Add(new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(await request.Content.ReadAsStringAsync()));
                return Responses.Dequeue();
            }
        }
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private static HttpResponseMessage Response(ProviderKind provider, string content = "{\"regions\":[]}", string finish = "stop") =>
            new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(provider == ProviderKind.Ollama
                ? Json.Serialize(new { message = new { content }, done_reason = finish })
                : Json.Serialize(new { choices = new[] { new { message = new { content }, finish_reason = finish } } })) };
        private static byte[] Png()
        { using (var b = new Bitmap(16, 12)) return ImageCaptureBudget.EncodePng(b, 102400); }
        private static TranslationSettings Settings(ProviderKind provider) => new TranslationSettings {
            Provider = provider, BaseUrl = "http://not-a-real-server.invalid/v1", Model = "text-model",
            ImageModel = "image-model", SourceLanguage = "韩语", TargetLanguage = "简体中文",
            CustomInstructions = "STYLE_TEST", Glossary = "GLOSSARY_TEST" };
        [TestMethod]
        public async Task OpenAiImageRequest_Has8192BudgetFullScanAndExistingLanguageRules()
        {
            var handler = new FakeService(); handler.Responses.Enqueue(Response(ProviderKind.OpenAiCompatible));
            using (var client = new TranslationClient(handler))
            {
                var result = await client.TranslateImageAsync(Png(), Settings(ProviderKind.OpenAiCompatible), CancellationToken.None, "Word");
                Assert.AreEqual(0, result.Statistics.RawCount); Assert.AreEqual("stop", result.FinishReason);
            }
            Assert.AreEqual(1, handler.Requests.Count);
            var request = handler.Requests[0]; Assert.AreEqual(8192, request["max_tokens"]);
            Assert.AreEqual(0.1, Convert.ToDouble(request["temperature"])); Assert.AreEqual("image-model", request["model"]);
            Assert.IsTrue(request.ContainsKey("response_format"));
            var body = Json.Serialize(request);
            foreach (var needle in new[] { "逐行逐块", "小字", "表格各单元格", "边缘文字", "只翻译其中属于韩语", "STYLE_TEST", "GLOSSARY_TEST", "bbox" })
                StringAssert.Contains(body, needle);
        }
        [TestMethod]
        public async Task OllamaImageRequest_HasMatchingOptionsAndSingleRequest()
        {
            var handler = new FakeService(); handler.Responses.Enqueue(Response(ProviderKind.Ollama));
            using (var client = new TranslationClient(handler)) await client.TranslateImageAsync(Png(), Settings(ProviderKind.Ollama), CancellationToken.None, "Excel");
            Assert.AreEqual(1, handler.Requests.Count);
            var options = (Dictionary<string, object>)handler.Requests[0]["options"];
            Assert.AreEqual(8192, options["num_predict"]); Assert.AreEqual(0.1, Convert.ToDouble(options["temperature"]));
            Assert.AreEqual("json", handler.Requests[0]["format"]);
        }
        [TestMethod]
        public async Task StructuredFallback_PreservesBudgetForBothProviders()
        {
            foreach (var provider in new[] { ProviderKind.OpenAiCompatible, ProviderKind.Ollama })
            {
                var handler = new FakeService(); handler.Responses.Enqueue(new HttpResponseMessage((HttpStatusCode)422) {
                    Content = new StringContent(provider == ProviderKind.Ollama ? "unsupported \"format\"" : "unsupported response_format") });
                handler.Responses.Enqueue(Response(provider));
                using (var client = new TranslationClient(handler)) await client.TranslateImageAsync(Png(), Settings(provider), CancellationToken.None);
                Assert.AreEqual(2, handler.Requests.Count);
                var fallback = handler.Requests[1];
                Assert.IsFalse(fallback.ContainsKey(provider == ProviderKind.Ollama ? "format" : "response_format"));
                Assert.AreEqual(8192, provider == ProviderKind.Ollama
                    ? ((Dictionary<string, object>)fallback["options"])["num_predict"] : fallback["max_tokens"]);
            }
        }
        [TestMethod]
        public async Task TruncatedValidJson_IsRejectedInsteadOfWritingPartialRegions()
        {
            foreach (var provider in new[] { ProviderKind.OpenAiCompatible, ProviderKind.Ollama })
            {
                var handler = new FakeService(); handler.Responses.Enqueue(Response(provider,
                    "{\"regions\":[{\"source\":\"x\",\"translation\":\"y\",\"bbox\":[0,0,100,100]}]}", "length"));
                using (var client = new TranslationClient(handler))
                    await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.TranslateImageAsync(Png(), Settings(provider), CancellationToken.None));
                Assert.AreEqual(1, handler.Requests.Count);
            }
        }
        [TestMethod]
        public async Task MalformedTruncatedJson_LogsFinishAndExplicitUnknownCounts()
        {
            var handler = new FakeService(); handler.Responses.Enqueue(Response(ProviderKind.OpenAiCompatible, "{\"regions\":[", "length"));
            using (var client = new TranslationClient(handler))
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.TranslateImageAsync(Png(), Settings(ProviderKind.OpenAiCompatible), CancellationToken.None));
            var lines = File.ReadAllLines(Path.Combine(Path.GetTempPath(), "OfficeTranslate", "image-overlay.log"));
            var last = lines[lines.Length - 1];
            StringAssert.Contains(last, "finish_reason=length truncated=True");
            StringAssert.Contains(last, "parsed=False raw=-1");
        }
        [TestMethod]
        public async Task EmptyImageResponse_LogsRealZeroCountsAndPixelSize()
        {
            var handler = new FakeService(); handler.Responses.Enqueue(Response(ProviderKind.OpenAiCompatible));
            using (var client = new TranslationClient(handler)) await client.TranslateImageAsync(Png(), Settings(ProviderKind.OpenAiCompatible), CancellationToken.None, "PowerPoint");
            var lines = File.ReadAllLines(Path.Combine(Path.GetTempPath(), "OfficeTranslate", "image-overlay.log"));
            StringAssert.Contains(lines[lines.Length - 1], "host=PowerPoint event=ocr_summary png=16x12 parsed=True raw=0 kept=0 dropped=0");
            Assert.IsFalse(lines[lines.Length - 1].Contains("STYLE_TEST"));
        }
        [TestMethod]
        public async Task CancelledImageRequest_DoesNotRetryOrReturnRegions()
        {
            var handler = new FakeService();
            using (var cts = new CancellationTokenSource())
            using (var client = new TranslationClient(handler))
            {
                cts.Cancel();
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => client.TranslateImageAsync(Png(), Settings(ProviderKind.OpenAiCompatible), cts.Token));
            }
            Assert.AreEqual(0, handler.Requests.Count);
        }
    }
}
