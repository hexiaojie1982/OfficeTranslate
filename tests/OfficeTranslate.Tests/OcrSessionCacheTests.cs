using System;
using System.Drawing;
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
    public class OcrSessionCacheTests
    {
        private const string Content = "{\"regions\":[{\"bbox\":[100,100,800,400],\"source\":\"原文 ABC\\n123\"}]}";
        private sealed class Handler : HttpMessageHandler
        {
            public int Calls;
            public Func<int, CancellationToken, Task<HttpResponseMessage>> Send = null!;
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token) =>
                Send(Interlocked.Increment(ref Calls), token);
        }

        // Each test gets its own endpoint so the real static session cache cannot
        // bypass another test's transport failure/cancellation scenario.
        private static TranslationSettings Settings(ProviderKind provider = ProviderKind.OpenAiCompatible) =>
            new TranslationSettings
            {
                Provider = provider, BaseUrl = "http://cache-" + Guid.NewGuid().ToString("N") + ".invalid/v1",
                Model = "text-model", ImageModel = "vision-model"
            };

        private static byte[] Png(Color color)
        {
            using (var bitmap = new Bitmap(20, 10))
            {
                bitmap.SetPixel(0, 0, color);
                return ImageCaptureBudget.EncodePng(bitmap, 102400);
            }
        }

        private static HttpResponseMessage Reply(ProviderKind provider, string content = Content, string finish = "stop") =>
            new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(provider == ProviderKind.Ollama
                    ? new JavaScriptSerializer().Serialize(new { message = new { content }, done_reason = finish })
                    : new JavaScriptSerializer().Serialize(new { choices = new[] { new { message = new { content }, finish_reason = finish } } }))
            };

        [DataTestMethod]
        [DataRow(ProviderKind.OpenAiCompatible)] [DataRow(ProviderKind.Ollama)]
        public async Task SameImageAcrossClients_CallsModelOnceAndReparsesRawContent(ProviderKind provider)
        {
            var settings = Settings(provider);
            var bytes = Png(Color.Red);
            var firstHandler = new Handler { Send = (n, t) => Task.FromResult(Reply(provider)) };
            ImageRecognitionResult first;
            using (var client = new TranslationClient(firstHandler))
                first = await client.RecognizeImageAsync(bytes, settings, CancellationToken.None, "Word");
            Assert.AreEqual(Content, GetCachedContent(settings, bytes));
            first.Regions[0].Source = "host mutation";
            first.Regions[0].X1 = 999;

            var secondHandler = new Handler { Send = (n, t) => throw new AssertFailedException("Cache hit called the model") };
            using (var client = new TranslationClient(secondHandler))
            {
                var second = await client.RecognizeImageAsync((byte[])bytes.Clone(), settings, CancellationToken.None, "Excel");
                var third = await client.RecognizeImageAsync(bytes, settings, CancellationToken.None, "PowerPoint");
                Assert.AreEqual("原文 ABC\n123", second.Regions[0].Source);
                Assert.AreEqual(100F, second.Regions[0].X1);
                Assert.AreEqual(string.Empty, second.Regions[0].Translation);
                Assert.AreNotSame(first, second);
                Assert.AreNotSame(second.Regions[0], third.Regions[0]);
                Assert.AreEqual(first.Statistics.RawCount, second.Statistics.RawCount);
            }
            Assert.AreEqual(1, firstHandler.Calls);
            Assert.AreEqual(0, secondHandler.Calls);
        }

        private static string GetCachedContent(TranslationSettings settings, byte[] bytes)
        {
            Assert.IsTrue(OcrSessionCache.TryGet(OcrSessionCache.CreateKey(settings, bytes), out var content));
            return content;
        }

        [TestMethod]
        public async Task DifferentImageBytes_CallModelAgain()
        {
            var settings = Settings();
            var first = Png(Color.Red);
            var second = Png(Color.Blue);
            Assert.AreNotEqual(OcrSessionCache.CreateKey(settings, first), OcrSessionCache.CreateKey(settings, second));
            var handler = new Handler { Send = (n, t) => Task.FromResult(Reply(settings.Provider)) };
            using (var client = new TranslationClient(handler))
            {
                await client.RecognizeImageAsync(first, settings, CancellationToken.None);
                await client.RecognizeImageAsync(second, settings, CancellationToken.None);
                await client.RecognizeImageAsync(first, settings, CancellationToken.None);
            }
            Assert.AreEqual(2, handler.Calls);
        }

        [DataTestMethod]
        [DataRow("ImageModel")] [DataRow("FallbackModel")] [DataRow("Provider")]
        [DataRow("BaseUrl")] [DataRow("UrlPathCase")]
        public async Task ChangedOcrSettings_CallModelAgain(string change)
        {
            var settings = Settings();
            if (change == "FallbackModel") settings.ImageModel = "";
            var bytes = Png(Color.Red);
            var handler = new Handler { Send = (n, t) => Task.FromResult(Reply(settings.Provider)) };
            using (var client = new TranslationClient(handler))
            {
                await client.RecognizeImageAsync(bytes, settings, CancellationToken.None);
                switch (change)
                {
                    case "ImageModel": settings.ImageModel = "other-vision"; break;
                    case "FallbackModel": settings.Model = "other-fallback"; break;
                    case "Provider": settings.Provider = ProviderKind.Ollama; break;
                    case "BaseUrl": settings.BaseUrl += "/other"; break;
                    case "UrlPathCase": settings.BaseUrl = settings.BaseUrl.Replace("/v1", "/V1"); break;
                }
                await client.RecognizeImageAsync(bytes, settings, CancellationToken.None);
                await client.RecognizeImageAsync(bytes, settings, CancellationToken.None);
            }
            Assert.AreEqual(2, handler.Calls, change);
        }

        [TestMethod]
        public async Task TranslationAndLayoutSettings_DoNotInvalidatePureOcr()
        {
            var settings = Settings();
            var bytes = Png(Color.Red);
            var handler = new Handler { Send = (n, t) => Task.FromResult(Reply(settings.Provider)) };
            using (var client = new TranslationClient(handler))
            {
                await client.RecognizeImageAsync(bytes, settings, CancellationToken.None);
                settings.Model = "other-text-model"; // ImageModel still selects the OCR model.
                settings.SourceLanguage = "韩语"; settings.TargetLanguage = "英语";
                settings.TranslationStyle = "Custom"; settings.CustomInstructions = "custom";
                settings.Glossary = "term"; settings.BilingualMode = true;
                settings.BaseUrl += "/";
                await client.RecognizeImageAsync(bytes, settings, CancellationToken.None);
            }
            Assert.AreEqual(1, handler.Calls);
        }

        [DataTestMethod]
        [DataRow("HttpFailure")] [DataRow("InvalidJson")] [DataRow("Truncated")]
        public async Task FailedResponse_IsNotCachedAndSuccessfulRetryIsCached(string failure)
        {
            var settings = Settings();
            var bytes = Png(Color.Red);
            var handler = new Handler { Send = (n, t) => Task.FromResult(n != 1 ? Reply(settings.Provider)
                : failure == "HttpFailure" ? new HttpResponseMessage(HttpStatusCode.BadRequest) { Content = new StringContent("failed") }
                : failure == "InvalidJson" ? Reply(settings.Provider, "{\"regions\":[")
                : Reply(settings.Provider, Content, "length")) };
            using (var client = new TranslationClient(handler))
            {
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.RecognizeImageAsync(bytes, settings, CancellationToken.None));
                Assert.IsFalse(OcrSessionCache.TryGet(OcrSessionCache.CreateKey(settings, bytes), out _));
                await client.RecognizeImageAsync(bytes, settings, CancellationToken.None);
                await client.RecognizeImageAsync(bytes, settings, CancellationToken.None);
            }
            Assert.AreEqual(2, handler.Calls);
        }

        [TestMethod]
        public async Task CancelledAfterModelResponse_IsNotCachedAndCachedHitHonorsCancellation()
        {
            var settings = Settings();
            var bytes = Png(Color.Red);
            using (var cts = new CancellationTokenSource())
            {
                var handler = new Handler { Send = (n, t) => {
                    if (n == 1) cts.Cancel();
                    return Task.FromResult(Reply(settings.Provider));
                } };
                using (var client = new TranslationClient(handler))
                {
                    await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => client.RecognizeImageAsync(bytes, settings, cts.Token));
                    Assert.IsFalse(OcrSessionCache.TryGet(OcrSessionCache.CreateKey(settings, bytes), out _));
                    await client.RecognizeImageAsync(bytes, settings, CancellationToken.None);
                    await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => client.RecognizeImageAsync(bytes, settings, cts.Token));
                    await client.RecognizeImageAsync(bytes, settings, CancellationToken.None);
                }
                Assert.AreEqual(2, handler.Calls);
            }
        }

        [TestMethod]
        public async Task Timeout_IsNotCached()
        {
            var settings = Settings();
            var bytes = Png(Color.Red);
            var handler = new Handler { Send = async (n, t) => {
                if (n == 1) await Task.Delay(Timeout.Infinite, t);
                return Reply(settings.Provider);
            } };
            using (var client = new TranslationClient(handler, TimeSpan.FromMilliseconds(60)))
            {
                await Assert.ThrowsExceptionAsync<ImageOcrTimeoutException>(() => client.RecognizeImageAsync(bytes, settings, CancellationToken.None));
                Assert.IsFalse(OcrSessionCache.TryGet(OcrSessionCache.CreateKey(settings, bytes), out _));
                await client.RecognizeImageAsync(bytes, settings, CancellationToken.None);
            }
            Assert.AreEqual(2, handler.Calls);
        }

        [TestMethod]
        public async Task SuccessfulEmptyRecognition_IsCached()
        {
            var settings = Settings();
            var bytes = Png(Color.Red);
            var handler = new Handler { Send = (n, t) => Task.FromResult(Reply(settings.Provider, "{\"regions\":[]}")) };
            using (var client = new TranslationClient(handler))
                for (var i = 0; i < 2; i++)
                    Assert.AreEqual(0, (await client.RecognizeImageAsync(bytes, settings, CancellationToken.None)).Regions.Count);
            Assert.AreEqual(1, handler.Calls);
        }

        [TestMethod]
        public async Task ImageTranslation_DoesNotReadOrPopulateOcrCache()
        {
            var settings = Settings();
            var bytes = Png(Color.Red);
            var handler = new Handler { Send = (n, t) => Task.FromResult(Reply(settings.Provider,
                n == 1 || n == 4 ? Content : Content.Replace("\"source\":", "\"translation\":\"译文\",\"source\":"))) };
            using (var client = new TranslationClient(handler))
            {
                await client.RecognizeImageAsync(bytes, settings, CancellationToken.None);
                await client.TranslateImageAsync(bytes, settings, CancellationToken.None);
                await client.TranslateImageAsync(bytes, settings, CancellationToken.None);
                await client.RecognizeImageAsync(bytes, settings, CancellationToken.None);
                Assert.AreEqual(3, handler.Calls);
                settings.ImageModel = "other-vision";
                await client.TranslateImageAsync(bytes, settings, CancellationToken.None);
                Assert.IsFalse(OcrSessionCache.TryGet(OcrSessionCache.CreateKey(settings, bytes), out _));
                await client.RecognizeImageAsync(bytes, settings, CancellationToken.None);
            }
            Assert.AreEqual(5, handler.Calls);
        }

        [TestMethod]
        public void Capacity_EvictsOldestInsertionAt2000Entries()
        {
            var prefix = Guid.NewGuid().ToString("N");
            for (var i = 0; i < 2000; i++) OcrSessionCache.Store(prefix + i, Content);
            Assert.IsTrue(OcrSessionCache.TryGet(prefix + 0, out _));
            OcrSessionCache.Store(prefix + 0, "updated");
            OcrSessionCache.Store(prefix + 2000, Content);
            Assert.IsFalse(OcrSessionCache.TryGet(prefix + 0, out _));
            Assert.IsTrue(OcrSessionCache.TryGet(prefix + 1, out _));
            Assert.IsTrue(OcrSessionCache.TryGet(prefix + 2000, out _));
        }
    }
}
