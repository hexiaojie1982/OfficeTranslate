using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ImageRequestTimeoutTests
    {
        private sealed class Handler : HttpMessageHandler
        {
            private readonly Func<int, CancellationToken, Task<HttpResponseMessage>> _send;
            public int Calls;
            public readonly TaskCompletionSource<bool> Started = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            public Handler(Func<int, CancellationToken, Task<HttpResponseMessage>> send) { _send = send; }
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
            {
                var attempt = Interlocked.Increment(ref Calls);
                Started.TrySetResult(true); return _send(attempt, token);
            }
        }
        private static TranslationSettings Settings(ProviderKind provider = ProviderKind.OpenAiCompatible) => new TranslationSettings
        {
            Provider = provider, BaseUrl = "http://not-a-real-server.invalid/v1", Model = "text-model", ImageModel = "image-model",
            SourceLanguage = "自动检测", TargetLanguage = "简体中文"
        };
        private static HttpResponseMessage Ok(ProviderKind provider = ProviderKind.OpenAiCompatible) =>
            new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(provider == ProviderKind.Ollama
                ? "{\"message\":{\"content\":\"{\\\"regions\\\":[]}\"},\"done_reason\":\"stop\"}"
                : "{\"choices\":[{\"message\":{\"content\":\"{\\\"regions\\\":[]}\"},\"finish_reason\":\"stop\"}]}") };
        private static HttpResponseMessage Transient(int status = 429)
        {
            var response = new HttpResponseMessage((HttpStatusCode)status) { Content = new StringContent("temporary") };
            response.Headers.RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromMilliseconds(10)); return response;
        }
        private static byte[] Png()
        { using (var image = new Bitmap(16, 12)) return ImageCaptureBudget.EncodePng(image, 102400); }
        private static async Task<HttpResponseMessage> WaitForCancellation(CancellationToken token)
        { await Task.Delay(Timeout.Infinite, token); throw new AssertFailedException("An infinite send completed"); }
        private static string LastSummary()
        {
            var lines = File.ReadAllLines(Path.Combine(Path.GetTempPath(), "OfficeTranslate", "image-overlay.log"));
            return lines[lines.Length - 1];
        }

        [DataTestMethod]
        [DataRow(ProviderKind.OpenAiCompatible)]
        [DataRow(ProviderKind.Ollama)]
        public async Task ImageDeadline_IsVisibleTimeoutNotCancellation_AndNeverRetries(ProviderKind provider)
        {
            var handler = new Handler((attempt, token) => WaitForCancellation(token));
            using (var client = new TranslationClient(handler, TimeSpan.FromMilliseconds(80)))
            {
                var ex = await Assert.ThrowsExceptionAsync<ImageOcrTimeoutException>(() => client.TranslateImageAsync(Png(), Settings(provider), CancellationToken.None));
                Assert.AreEqual(TimeSpan.FromMilliseconds(80), ex.RequestTimeout);
                StringAssert.Contains(ex.Message, "等待 0.08 秒后超时");
                StringAssert.Contains(ex.Message, "未写入当前图片的任何内容");
                StringAssert.Contains(ex.Message, "减小图片或改用更快");
                Assert.IsFalse(TranslationFailurePolicy.IsUserCancellation(ex, CancellationToken.None));
            }
            Assert.AreEqual(1, handler.Calls);
            StringAssert.Contains(LastSummary(), "status=timeout"); StringAssert.Contains(LastSummary(), "parsed=False raw=-1");
        }

        [DataTestMethod]
        [DataRow(ProviderKind.OpenAiCompatible)]
        [DataRow(ProviderKind.Ollama)]
        public async Task InFlightUserCancel_IsSilentCancellationAndNeverRetries(ProviderKind provider)
        {
            var handler = new Handler((attempt, token) => WaitForCancellation(token));
            using (var cts = new CancellationTokenSource())
            using (var client = new TranslationClient(handler, TimeSpan.FromSeconds(5)))
            {
                var pending = client.TranslateImageAsync(Png(), Settings(provider), cts.Token);
                await handler.Started.Task; cts.Cancel();
                var ex = await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => pending);
                Assert.IsTrue(TranslationFailurePolicy.IsUserCancellation(ex, cts.Token));
            }
            Assert.AreEqual(1, handler.Calls); StringAssert.Contains(LastSummary(), "status=cancelled");
        }

        [DataTestMethod]
        [DataRow(ProviderKind.OpenAiCompatible)]
        [DataRow(ProviderKind.Ollama)]
        public async Task SlowImageResponse_IsNotLimitedByConfiguredTextTimeout(ProviderKind provider)
        {
            var handler = new Handler(async (attempt, token) => { await Task.Delay(1250, token); return Ok(provider); });
            var settings = Settings(provider); settings.TimeoutSeconds = 1;
            using (var client = new TranslationClient(handler, TimeSpan.FromSeconds(5)))
            {
                var result = await client.TranslateImageAsync(Png(), settings, CancellationToken.None);
                Assert.AreEqual(0, result.Regions.Count);
            }
            Assert.AreEqual(1, handler.Calls); StringAssert.Contains(LastSummary(), "status=parsed");
        }

        [TestMethod]
        public async Task ImageClient_DoesNotExtendTextTimeoutOrChangeItsThreeAttemptRetry()
        {
            var handler = new Handler((attempt, token) => attempt == 1 ? Task.FromResult(Ok()) : WaitForCancellation(token));
            var settings = Settings(); settings.TimeoutSeconds = 1;
            using (var client = new TranslationClient(handler, TimeSpan.FromSeconds(5)))
            {
                await client.TranslateImageAsync(Png(), settings, CancellationToken.None);
                await Assert.ThrowsExceptionAsync<TaskCanceledException>(() => client.TranslateDetailedAsync(
                    "Please translate this unique sentence " + Guid.NewGuid().ToString("N"), settings, CancellationToken.None));
            }
            // One image success, then three text attempts at 1s, with the
            // unchanged 1s/3s backoff. No image deadline leaks into text.
            Assert.AreEqual(4, handler.Calls);
        }

        [DataTestMethod]
        [DataRow(429)] [DataRow(500)] [DataRow(502)] [DataRow(503)] [DataRow(504)]
        public async Task TransientImageHttpStatus_StillRetriesAndCanSucceed(int status)
        {
            var handler = new Handler((attempt, token) => Task.FromResult(attempt == 1 ? Transient(status) : Ok()));
            using (var client = new TranslationClient(handler, TimeSpan.FromSeconds(5)))
                await client.TranslateImageAsync(Png(), Settings(), CancellationToken.None);
            Assert.AreEqual(2, handler.Calls); StringAssert.Contains(LastSummary(), "status=parsed");
        }

        [TestMethod]
        public async Task TimeoutAfterTransientHttpResponse_StopsWithoutThirdAttempt()
        {
            var handler = new Handler((attempt, token) => attempt == 1 ? Task.FromResult(Transient()) : WaitForCancellation(token));
            using (var client = new TranslationClient(handler, TimeSpan.FromMilliseconds(80)))
                await Assert.ThrowsExceptionAsync<ImageOcrTimeoutException>(() => client.TranslateImageAsync(Png(), Settings(), CancellationToken.None));
            Assert.AreEqual(2, handler.Calls); StringAssert.Contains(LastSummary(), "status=timeout");
        }

        [TestMethod]
        public async Task StructuredFallbackTimeout_DoesNotRestartImageGeneration()
        {
            var handler = new Handler((attempt, token) => attempt == 1
                ? Task.FromResult(new HttpResponseMessage((HttpStatusCode)422) { Content = new StringContent("unsupported response_format") })
                : WaitForCancellation(token));
            using (var client = new TranslationClient(handler, TimeSpan.FromMilliseconds(80)))
                await Assert.ThrowsExceptionAsync<ImageOcrTimeoutException>(() => client.TranslateImageAsync(Png(), Settings(), CancellationToken.None));
            Assert.AreEqual(2, handler.Calls); StringAssert.Contains(LastSummary(), "status=timeout");
        }

        [TestMethod]
        public async Task NonUserTransportCancellation_IsVisibleFailure_NotSilentCancelOrRetry()
        {
            var handler = new Handler((attempt, token) => throw new TaskCanceledException("fake non-user transport interruption"));
            using (var client = new TranslationClient(handler, TimeSpan.FromSeconds(5)))
            {
                var ex = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => client.TranslateImageAsync(Png(), Settings(), CancellationToken.None));
                StringAssert.Contains(ex.Message, "非用户取消"); Assert.IsFalse(TranslationFailurePolicy.IsUserCancellation(ex, CancellationToken.None));
            }
            Assert.AreEqual(1, handler.Calls); StringAssert.Contains(LastSummary(), "status=failed");
        }

        [TestMethod]
        public async Task UserCancelDuringHttpBackoff_DoesNotStartAnotherAttempt()
        {
            var handler = new Handler((attempt, token) =>
            {
                var response = Transient(); response.Headers.RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromSeconds(1));
                return Task.FromResult(response);
            });
            using (var cts = new CancellationTokenSource())
            using (var client = new TranslationClient(handler, TimeSpan.FromSeconds(5)))
            {
                var pending = client.TranslateImageAsync(Png(), Settings(), cts.Token);
                await handler.Started.Task; cts.Cancel();
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => pending);
            }
            Assert.AreEqual(1, handler.Calls); StringAssert.Contains(LastSummary(), "status=cancelled");
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public async Task WrappedTransportCancellation_PreservesTimeoutOrUserCancelWithoutRetry(bool userCancel)
        {
            var handler = new Handler(async (attempt, token) =>
            {
                try { return await WaitForCancellation(token); }
                catch (OperationCanceledException ex) { throw new HttpRequestException("fake wrapped interruption", ex); }
            });
            using (var cts = new CancellationTokenSource())
            using (var client = new TranslationClient(handler, userCancel ? TimeSpan.FromSeconds(5) : TimeSpan.FromMilliseconds(80)))
            {
                var pending = client.TranslateImageAsync(Png(), Settings(), cts.Token);
                if (userCancel)
                {
                    await handler.Started.Task; cts.Cancel();
                    await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => pending);
                }
                else await Assert.ThrowsExceptionAsync<ImageOcrTimeoutException>(() => pending);
            }
            Assert.AreEqual(1, handler.Calls); StringAssert.Contains(LastSummary(), userCancel ? "status=cancelled" : "status=timeout");
        }

        private sealed class DelayedContent : HttpContent
        {
            private readonly CancellationToken _token;
            public DelayedContent(CancellationToken token) { _token = token; }
            protected override async Task SerializeToStreamAsync(Stream stream, TransportContext context) => await Task.Delay(Timeout.Infinite, _token);
            protected override bool TryComputeLength(out long length) { length = 0; return false; }
        }
        [TestMethod]
        public async Task DeadlineCoversResponseBody_NotOnlyHttpHeaders()
        {
            var handler = new Handler((attempt, token) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new DelayedContent(token) }));
            using (var client = new TranslationClient(handler, TimeSpan.FromMilliseconds(80)))
                await Assert.ThrowsExceptionAsync<ImageOcrTimeoutException>(() => client.TranslateImageAsync(Png(), Settings(), CancellationToken.None));
            Assert.AreEqual(1, handler.Calls); StringAssert.Contains(LastSummary(), "status=timeout");
        }

        [TestMethod]
        public void SharedFailurePolicy_RequiresUserTokenAndKeepsDefaultBudgets()
        {
            Assert.AreEqual(600, TranslationClient.ImageRequestTimeoutSeconds);
            Assert.AreEqual(120, new TranslationSettings().TimeoutSeconds);
            var timeout = new ImageOcrTimeoutException(TimeSpan.FromSeconds(600), new TaskCanceledException());
            StringAssert.Contains(timeout.Message, "等待 600 秒后超时");
            using (var cts = new CancellationTokenSource())
            {
                cts.Cancel();
                Assert.IsFalse(TranslationFailurePolicy.IsUserCancellation(timeout, cts.Token));
                Assert.IsFalse(TranslationFailurePolicy.IsUserCancellation(new TaskCanceledException(), CancellationToken.None));
                Assert.IsTrue(TranslationFailurePolicy.IsUserCancellation(new OperationCanceledException(), cts.Token));
            }
        }
    }
}
