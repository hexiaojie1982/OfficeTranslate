using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ChatCompatibilityTests
    {
        private const string JsonReply = "{\"choices\":[{\"message\":{\"content\":\"完整回答\"},\"finish_reason\":\"stop\"}]}";
        private const string SseReply = "data: {\"choices\":[{\"index\":0,\"delta\":{\"content\":\"完整回答\"},\"finish_reason\":null}]}\n\ndata: {\"choices\":[{\"delta\":{},\"finish_reason\":\"stop\"}]}\n\ndata: [DONE]";
        private static TranslationSettings Settings(string url = "https://api.deepseek.com/v1/") => new TranslationSettings {
            Provider = ProviderKind.OpenAiCompatible, BaseUrl = url, Model = "deepseek-chat", ApiKey = "fixture-private-key"
        };
        private static HttpResponseMessage Response(string body, string type = "application/json", int status = 200) =>
            new HttpResponseMessage((HttpStatusCode)status) { Content = new StringContent(body, Encoding.UTF8, type) };

        private sealed class Request
        {
            public string Url = "", Auth = "", Accept = "", ContentType = "";
            public Dictionary<string, object> Body = null!;
        }
        private sealed class Handler : HttpMessageHandler
        {
            public readonly List<Request> Requests = new List<Request>();
            public Func<Request, int, CancellationToken, Task<HttpResponseMessage>> Reply = null!;
            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
            {
                var captured = new Request { Url = request.RequestUri!.ToString(), Auth = request.Headers.Authorization?.ToString() ?? "",
                    Accept = request.Headers.Accept.ToString(), ContentType = request.Content!.Headers.ContentType!.MediaType,
                    Body = (Dictionary<string, object>)new JavaScriptSerializer().DeserializeObject(await request.Content.ReadAsStringAsync()) };
                Requests.Add(captured); return await Reply(captured, Requests.Count, token);
            }
        }
        private sealed class HangingStream : Stream
        {
            private readonly TaskCompletionSource<int> _read = new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);
            public override bool CanRead => true;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();
            public override long Position { get => 0; set => throw new NotSupportedException(); }
            public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token) => _read.Task;
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override void Flush() { }
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            protected override void Dispose(bool disposing) { _read.TrySetException(new ObjectDisposedException(nameof(HangingStream))); base.Dispose(disposing); }
        }
        private static Task<ChatReply> Send(TranslationClient client, TranslationSettings settings, Action<string>? output = null, CancellationToken token = default) =>
            client.ChatAsync(new[] { new ChatMessage("system", "private-system-prompt"), new ChatMessage("user", "private-document") }, settings, 1024, output ?? (_ => { }), token);

        [DataTestMethod]
        [DataRow("https://api.deepseek.com", "https://api.deepseek.com/chat/completions", true)]
        [DataRow("https://api.deepseek.com/v1/", "https://api.deepseek.com/v1/chat/completions", true)]
        [DataRow("https://API.DEEPSEEK.COM/beta/", "https://api.deepseek.com/beta/chat/completions", true)]
        [DataRow("http://localhost:8000/v1/", "http://localhost:8000/v1/chat/completions", false)]
        [DataRow("https://api.deepseek.com.example/v1", "https://api.deepseek.com.example/v1/chat/completions", false)]
        public async Task Request_MatchesWorkingTranslationContract_AndRestrictsDeepSeekExtension(string url, string expected, bool deepseek)
        {
            var handler = new Handler { Reply = (r,n,t) => Task.FromResult(Response(SseReply, "text/event-stream")) };
            using (var client = new TranslationClient(handler)) Assert.AreEqual("完整回答", (await Send(client, Settings(url))).Text);
            var r = handler.Requests.Single();
            Assert.AreEqual(expected, r.Url); Assert.AreEqual("Bearer fixture-private-key", r.Auth);
            Assert.AreEqual("application/json", r.ContentType); Assert.AreEqual("text/event-stream", r.Accept);
            Assert.AreEqual("deepseek-chat", r.Body["model"]); Assert.AreEqual(true, r.Body["stream"]); Assert.AreEqual(1024, r.Body["max_tokens"]);
            Assert.AreEqual(deepseek, r.Body.ContainsKey("thinking"));
            if (deepseek) Assert.AreEqual("disabled", ((Dictionary<string, object>)r.Body["thinking"])["type"]);
            var messages = (object[])r.Body["messages"]; Assert.AreEqual(2, messages.Length);
            Assert.AreEqual("system", ((Dictionary<string, object>)messages[0])["role"]);
            Assert.AreEqual("private-document", ((Dictionary<string, object>)messages[1])["content"]);
        }

        [TestMethod]
        public async Task DeepSeek_660CharacterDocument_DoesNotSpendAnswerBudgetOnDefaultThinking()
        {
            var handler = new Handler { Reply = (r,n,t) => Task.FromResult(r.Body.ContainsKey("thinking") ? Response(SseReply, "text/event-stream") :
                (bool)r.Body["stream"] ? Response("data: {\"choices\":[{\"delta\":{\"reasoning_content\":\"thinking consumed output budget\"},\"finish_reason\":\"length\"}]}\n\ndata: [DONE]\n\n", "text/event-stream") :
                Response("{\"choices\":[{\"message\":{\"reasoning_content\":\"thinking consumed output budget\",\"content\":\"\"},\"finish_reason\":\"length\"}]}")) };
            using (var client = new TranslationClient(handler))
            {
                var session = new DocumentChatSession();
                var result = await session.SendAsync(new ChatContext(new string('文', 660), false), "总结", ChatMode.Question,
                    Settings(), client.ChatAsync, _ => { }, (i,n) => { }, CancellationToken.None);
                Assert.AreEqual("完整回答", result.Text); Assert.AreEqual(2, session.History.Count); Assert.AreEqual(1, handler.Requests.Count);
            }
        }

        [TestMethod]
        public async Task DeepSeek_RecoveryKeepsProviderPolicyAndReplacesPartialText()
        {
            var handler = new Handler { Reply = (r,n,t) => Task.FromResult(n == 1 ? Response(SseReply.Replace("data: [DONE]", ""), "text/event-stream") : Response(JsonReply)) };
            var updates = new List<string>();
            using (var client = new TranslationClient(handler)) Assert.IsTrue((await Send(client, Settings(), updates.Add)).UsedFallback);
            Assert.AreEqual(2, handler.Requests.Count); Assert.IsTrue(updates.Contains("")); Assert.AreEqual("完整回答", updates.Last());
            foreach (var r in handler.Requests) Assert.AreEqual("disabled", ((Dictionary<string, object>)r.Body["thinking"])["type"]);
            Assert.AreEqual(false, handler.Requests[1].Body["stream"]); Assert.AreEqual("application/json", handler.Requests[1].Accept);
        }

        [DataTestMethod]
        [DataRow("\n", "data:", false)] [DataRow("\r\n", "data: ", true)]
        public async Task Sse_HandlesDeepSeekReasoningKeepAliveUsageBomAndUnterminatedDone(string newline, string prefix, bool bom)
        {
            var wire = (bom ? "\uFEFF" : "") + ": keep-alive\n\nevent: message\nid: 1\n" +
                "data: {\"choices\":[{\"delta\":{\"role\":\"assistant\",\"content\":null,\"reasoning_content\":\"private reasoning\"}}]}\n\n" +
                SseReply.Replace("data: [DONE]", "data: {\"choices\":[],\"usage\":{\"total_tokens\":33}}\n\ndata: [DONE]");
            wire = wire.Replace("data: ", prefix).Replace("\n", newline);
            var updates = new List<string>();
            Assert.AreEqual("完整回答", await ChatSseParser.ReadAsync(new StringReader(wire), updates.Add, CancellationToken.None));
            CollectionAssert.AreEqual(new[] { "完整回答" }, updates);
        }

        [TestMethod]
        public async Task Sse_MultipleDataLinesFormOneJsonEvent()
        {
            var wire = "data: {\"choices\":\n" + "data: [{\"delta\":{\"content\":\"你好😀\"}}]}\n\ndata:[DONE]";
            Assert.AreEqual("你好😀", await ChatSseParser.ReadAsync(new StringReader(wire), _ => { }, CancellationToken.None));
        }

        [DataTestMethod]
        [DataRow(401, "invalid_api_key", ChatFailure.Authentication, "ChatAuthError")]
        [DataRow(403, "permission_denied", ChatFailure.Authentication, "ChatAuthError")]
        [DataRow(402, "insufficient_balance", ChatFailure.Quota, "ChatQuotaError")]
        [DataRow(429, "rate_limit", ChatFailure.RateLimit, "ChatRateError")]
        [DataRow(400, "context_length_exceeded", ChatFailure.Budget, "ChatBudgetError")]
        [DataRow(422, "invalid max_tokens", ChatFailure.Budget, "ChatBudgetError")]
        [DataRow(404, "model_not_found", ChatFailure.Model, "ChatModelError")]
        [DataRow(503, "server_error", ChatFailure.Model, "ChatModelError")]
        public async Task HttpErrors_AreActionableAndDoNotRetryOrEchoPrivateBody(int status, string code, ChatFailure failure, string key)
        {
            var handler = new Handler { Reply = (r,n,t) => Task.FromResult(Response(new JavaScriptSerializer().Serialize(new {
                error = new { code, message = "private-document private-system-prompt fixture-private-key" } }), status: status)) };
            using (var client = new TranslationClient(handler))
            {
                var error = await Assert.ThrowsExceptionAsync<ChatException>(() => Send(client, Settings()));
                Assert.AreEqual(failure, error.Failure); Assert.AreEqual(key, ChatErrors.UiKey(error));
                var detail = ChatErrors.Describe(error); StringAssert.Contains(detail, "HTTP " + status);
                Assert.IsFalse(detail.Contains("private")); StringAssert.StartsWith(detail, "ChatException:");
            }
            Assert.AreEqual(1, handler.Requests.Count);
        }

        [DataTestMethod]
        [DataRow("length", ChatFailure.Budget)] [DataRow("content_filter", ChatFailure.Model)]
        [DataRow("insufficient_system_resource", ChatFailure.Model)]
        public async Task Sse_FinishErrorsDoNotBlindlyRepeatSameBudget(string reason, ChatFailure failure)
        {
            var handler = new Handler { Reply = (r,n,t) => Task.FromResult(Response("data: {\"choices\":[{\"delta\":{\"content\":\"partial\"},\"finish_reason\":\"" + reason + "\"}]}\n\ndata: [DONE]", "text/event-stream")) };
            using (var client = new TranslationClient(handler))
            {
                var error = await Assert.ThrowsExceptionAsync<ChatException>(() => Send(client, Settings())); Assert.AreEqual(failure, error.Failure);
                StringAssert.Contains(error.Message, "finish_reason=" + reason);
            }
            Assert.AreEqual(1, handler.Requests.Count);
        }

        [TestMethod]
        public async Task Sse_ErrorObjectPreservesCategoryWithoutLeakingBody()
        {
            var error = await Assert.ThrowsExceptionAsync<ChatException>(() => ChatSseParser.ReadAsync(new StringReader(
                "event: error\ndata: {\"error\":{\"code\":\"invalid_api_key\",\"message\":\"private-key private-document\"}}\n\n"), _ => { }, CancellationToken.None));
            Assert.AreEqual(ChatFailure.Authentication, error.Failure); Assert.IsFalse(ChatErrors.Describe(error).Contains("private"));
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public async Task JsonResponseToStreamingRequest_IsUsedWithoutDuplicateRequest(bool malformed)
        {
            var handler = new Handler { Reply = (r,n,t) => Task.FromResult(Response(malformed && n == 1 ? "{ private-document fixture-private-key" : JsonReply)) };
            using (var client = new TranslationClient(handler)) { var result = await Send(client, Settings()); Assert.AreEqual("完整回答", result.Text); Assert.IsTrue(result.UsedFallback); }
            Assert.AreEqual(malformed ? 2 : 1, handler.Requests.Count);
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public async Task ExplicitStreamingRejection_FallsBackForJsonAndPlainText(bool plain)
        {
            var handler = new Handler { Reply = (r,n,t) => Task.FromResult(n == 1 ? Response(plain ? "stream is not supported" : "{\"error\":{\"message\":\"stream is not supported\"}}", status: 400) : Response(JsonReply)) };
            using (var client = new TranslationClient(handler)) Assert.IsTrue((await Send(client, Settings("http://localhost:8000/v1"))).UsedFallback);
            Assert.AreEqual(2, handler.Requests.Count); Assert.IsFalse(handler.Requests[1].Body.ContainsKey("thinking"));
        }

        [TestMethod]
        public async Task NetworkDetails_RetainExceptionTypeAndMessage_RedactKeyUrlAndPrompt()
        {
            var handler = new Handler { Reply = (r,n,t) => throw new HttpRequestException("TLS handshake failed fixture-private-key private-document https://host/path?key=hidden", new WebException("Name resolution failed")) };
            using (var client = new TranslationClient(handler))
            {
                var error = await Assert.ThrowsExceptionAsync<ChatException>(() => Send(client, Settings()));
                Assert.AreEqual(ChatFailure.Network, error.Failure);
                StringAssert.Contains(error.Message, "HttpRequestException: TLS handshake failed"); StringAssert.Contains(error.Message, "WebException: Name resolution failed");
                StringAssert.Contains(error.Message, "Streaming attempt:"); Assert.IsFalse(error.Message.Contains("private")); Assert.IsFalse(error.Message.Contains("hidden"));
            }
        }

        [TestMethod]
        public void LocalDiagnostics_KeepRealMessagesAndTypes_RedactDocumentParagraphsPromptsAndKeys()
        {
            var settings = Settings(); settings.ChatQuestionSystemPrompt = "private prompt paragraph";
            var error = new InvalidOperationException("Local failure: private paragraph; private prompt paragraph; fixture-private-key", new IOException("Read failed"));
            var detail = ChatErrors.Describe(error, settings, "private paragraph\nsecond paragraph");
            StringAssert.Contains(detail, "InvalidOperationException: Local failure:"); StringAssert.Contains(detail, "IOException: Read failed");
            Assert.IsFalse(detail.Contains("private")); Assert.IsFalse(detail.Contains("fixture-private-key"));
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public async Task TimeoutAndUserCancellation_AreDifferentAndBounded(bool userCancelled)
        {
            var handler = new Handler { Reply = async (r,n,t) => { await Task.Delay(10000, t); return Response(JsonReply); } };
            var settings = Settings(); settings.TimeoutSeconds = 1;
            using (var cts = new CancellationTokenSource()) using (var client = new TranslationClient(handler))
            {
                if (userCancelled)
                {
                    cts.CancelAfter(100);
                    try { await Send(client, settings, token: cts.Token); Assert.Fail("Expected cancellation"); } catch (OperationCanceledException) { }
                    Assert.AreEqual(1, handler.Requests.Count);
                }
                else
                {
                    var error = await Assert.ThrowsExceptionAsync<ChatException>(() => Send(client, settings));
                    Assert.AreEqual(ChatFailure.Network, error.Failure); StringAssert.Contains(error.Message, "timed out after 1 seconds"); Assert.AreEqual(2, handler.Requests.Count);
                }
            }
        }

        [TestMethod]
        public async Task LocalBudgetError_IsDistinctWithoutCallingServiceOrCommittingHistory()
        {
            var settings = Settings(); settings.ChatOutputTokens = settings.ChatContextTokens;
            var session = new DocumentChatSession();
            var error = await Assert.ThrowsExceptionAsync<ChatException>(() => session.SendAsync(new ChatContext("doc", false), "Q", ChatMode.Question,
                settings, (m,s,n,o,t) => throw new AssertFailedException("No HTTP expected"), _ => { }, (i,n) => { }, CancellationToken.None));
            Assert.AreEqual("ChatBudgetError", ChatErrors.UiKey(error)); Assert.AreEqual(0, session.History.Count);
        }

        [TestMethod]
        public async Task Vllm_WithoutApiKey_UsesConfiguredPathAndNoAuthorization()
        {
            var settings = Settings("http://localhost:8000/v1/"); settings.ApiKey = ""; settings.Model = "local-qwen";
            var handler = new Handler { Reply = (r,n,t) => Task.FromResult(Response(SseReply, "text/event-stream")) };
            using (var client = new TranslationClient(handler)) Assert.IsFalse((await Send(client, settings)).UsedFallback);
            var r = handler.Requests.Single(); Assert.AreEqual("", r.Auth); Assert.AreEqual("local-qwen", r.Body["model"]);
            Assert.AreEqual("http://localhost:8000/v1/chat/completions", r.Url); Assert.IsFalse(r.Body.ContainsKey("thinking"));
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public async Task BlockedSseRead_IsDisposedOnDeadlineOrCancellation(bool userCancelled)
        {
            var handler = new Handler { Reply = (r,n,t) => {
                if (n > 1) return Task.FromResult(Response(JsonReply));
                var response = new HttpResponseMessage(HttpStatusCode.OK) { Content = new StreamContent(new HangingStream()) };
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("text/event-stream");
                return Task.FromResult(response);
            } };
            var settings = Settings(); settings.TimeoutSeconds = 1;
            using (var cts = new CancellationTokenSource()) using (var client = new TranslationClient(handler))
            {
                if (userCancelled)
                {
                    cts.CancelAfter(100);
                    try { await Send(client, settings, token: cts.Token); Assert.Fail("Expected cancellation"); } catch (OperationCanceledException) { }
                    Assert.AreEqual(1, handler.Requests.Count);
                }
                else
                {
                    var result = await Send(client, settings); Assert.IsTrue(result.UsedFallback); Assert.AreEqual("完整回答", result.Text);
                    Assert.AreEqual(2, handler.Requests.Count);
                }
            }
        }
    }
}
