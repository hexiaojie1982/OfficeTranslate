using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class SecurityResponseBudgetTests
    {
        private sealed class Handler : HttpMessageHandler
        {
            internal int Calls;
            internal Func<int, HttpResponseMessage> Reply = null!;
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token) =>
                Task.FromResult(Reply(++Calls));
        }

        // Unknown-length transport: produces adversarial input without preallocating the body.
        private sealed class RepeatingStream : Stream
        {
            private readonly byte[] _pattern;
            private readonly int _size;
            internal int ReadBytes;
            internal bool Disposed;
            internal RepeatingStream(string pattern, int size) { _pattern = Encoding.UTF8.GetBytes(pattern); _size = size; }
            public override int Read(byte[] buffer, int offset, int count)
            {
                count = Math.Min(count, _size - ReadBytes);
                for (var i = 0; i < count; i++) buffer[offset + i] = _pattern[(ReadBytes + i) % _pattern.Length];
                ReadBytes += count; return count;
            }
            public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
            { token.ThrowIfCancellationRequested(); return Task.FromResult(Read(buffer, offset, count)); }
            protected override void Dispose(bool disposing) { Disposed = true; base.Dispose(disposing); }
            public override bool CanRead => true;
            public override bool CanWrite => false;
            public override bool CanSeek => false;
            public override long Length => throw new NotSupportedException();
            public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
            public override void Flush() => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        }
        private static TranslationSettings Settings(ProviderKind provider) => new TranslationSettings {
            Provider = provider, BaseUrl = "https://fixture.invalid/v1", Model = "fixture", TimeoutSeconds = 30
        };
        private static HttpResponseMessage Response(string body, string type = "application/json", int status = 200) =>
            new HttpResponseMessage((HttpStatusCode)status) { Content = new StringContent(body, Encoding.UTF8, type) };

        [DataTestMethod]
        [DataRow("x")] // One unterminated line.
        [DataRow(": keepalive\n\n")] // Many short lines, no completion text.
        [DataRow("data: \n")] // Many data lines, no event boundary.
        public async Task ChatSse_GlobalCharacterBudgetCoversLineAndEvent(string pattern)
        {
            var wire = string.Concat(Enumerable.Repeat(pattern, ResponseLimits.Chat / pattern.Length + 1));
            var ex = await Assert.ThrowsExceptionAsync<ChatException>(() =>
                ChatSseParser.ReadAsync(new StringReader(wire), _ => { }, CancellationToken.None));
            Assert.AreEqual(ChatFailure.Protocol, ex.Failure); Assert.IsTrue(ex.ResponseLimitExceeded);
        }

        [TestMethod]
        public async Task ChatSse_ExactBudgetAndAllNewlineStylesRemainValid()
        {
            var answer = "data: {\"choices\":[{\"delta\":{\"content\":\"正常回答😀\"}}]}\r\n\rdata: [DONE]\n\n";
            var wire = ":" + new string(' ', ResponseLimits.Chat - answer.Length - 2) + "\n" + answer;
            Assert.AreEqual(ResponseLimits.Chat, wire.Length);
            Assert.AreEqual("正常回答😀", await ChatSseParser.ReadAsync(new StringReader(wire), _ => { }, CancellationToken.None));
        }

        [TestMethod]
        public async Task ChatSse_CumulativeEventsCannotResetResponseBudget()
        {
            var content = new string('a', 1024 * 1024);
            var part = "data: {\"choices\":[{\"delta\":{\"content\":\"" + content + "\"}}]}\n\n";
            var updates = 0;
            var ex = await Assert.ThrowsExceptionAsync<ChatException>(() => ChatSseParser.ReadAsync(
                new StringReader(string.Concat(Enumerable.Repeat(part, 9)) + "data: [DONE]\n\n"), _ => updates++, CancellationToken.None));
            Assert.IsTrue(ex.ResponseLimitExceeded); Assert.IsTrue(updates > 0 && updates < 8);
        }

        [DataTestMethod]
        [DataRow(ProviderKind.OpenAiCompatible, "text/event-stream", false, false)]
        [DataRow(ProviderKind.OpenAiCompatible, "application/json", false, false)]
        [DataRow(ProviderKind.Ollama, "application/json", false, false)]
        [DataRow(ProviderKind.OpenAiCompatible, "application/json", true, false)]
        [DataRow(ProviderKind.OpenAiCompatible, "application/json", false, true)]
        public async Task Chat_ByteBudgetDisposesTransport_AndNeverRetriesOversize(ProviderKind provider, string type, bool recovery, bool error)
        {
            // Non-ASCII exceeds the byte budget before the character budget. Includes lying Content-Length.
            var stream = new RepeatingStream("文", ResponseLimits.Chat * 2);
            var handler = new Handler { Reply = n => {
                if (recovery && n == 1) return Response("{\"error\":{\"message\":\"stream not supported\"}}", status: 400);
                var response = new HttpResponseMessage(error ? HttpStatusCode.BadRequest : HttpStatusCode.OK) { Content = new StreamContent(stream) };
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(type);
                response.Content.Headers.ContentLength = 1;
                return response;
            } };
            using (var client = new TranslationClient(handler))
            {
                var ex = await Assert.ThrowsExceptionAsync<ChatException>(() => client.ChatAsync(
                    new[] { new ChatMessage("user", "fixture") }, Settings(provider), 1024, _ => { }, CancellationToken.None));
                Assert.AreEqual(ChatFailure.Protocol, ex.Failure);
            }
            Assert.AreEqual(recovery ? 2 : 1, handler.Calls);
            Assert.AreEqual(ResponseLimits.Chat + 1, stream.ReadBytes); Assert.IsTrue(stream.Disposed);
        }

        [TestMethod]
        public void ChatJson_DirectParserHasSameFiniteBudget()
        {
            var ex = Assert.ThrowsException<ChatException>(() => ChatErrors.ParseJson(new string(' ', ResponseLimits.Chat + 1)));
            Assert.AreEqual(ChatFailure.Protocol, ex.Failure);
            Assert.IsNotNull(ChatErrors.ParseJson("{}" + new string(' ', ResponseLimits.Chat - 2)));
        }

        [DataTestMethod]
        [DataRow(false, false)] [DataRow(true, false)] [DataRow(false, true)] [DataRow(true, true)]
        public async Task Image_ActualByteBudgetBoundsRecognitionAndTranslation_IncludingErrors(bool recognition, bool error)
        {
            var stream = new RepeatingStream(" ", ResponseLimits.Image * 2);
            var handler = new Handler { Reply = n => new HttpResponseMessage(error ? HttpStatusCode.BadRequest : HttpStatusCode.OK) {
                Content = new StreamContent(stream)
            } };
            using (var client = new TranslationClient(handler))
            {
                var ex = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => recognition
                    ? client.RecognizeImageAsync(new byte[0], Settings(ProviderKind.OpenAiCompatible), CancellationToken.None)
                    : client.TranslateImageAsync(new byte[0], Settings(ProviderKind.OpenAiCompatible), CancellationToken.None));
                StringAssert.Contains(ex.Message, "16 MiB");
            }
            Assert.AreEqual(1, handler.Calls); Assert.AreEqual(ResponseLimits.Image + 1, stream.ReadBytes); Assert.IsTrue(stream.Disposed);
        }

        [TestMethod]
        public async Task Image_AdvertisedOversizeIsRejectedBeforeBodyRead()
        {
            var stream = new RepeatingStream("x", ResponseLimits.Image * 2);
            var handler = new Handler { Reply = n => {
                var response = new HttpResponseMessage(HttpStatusCode.OK) { Content = new StreamContent(stream) };
                response.Content.Headers.ContentLength = ResponseLimits.Image + 1; return response;
            } };
            using (var client = new TranslationClient(handler)) await Assert.ThrowsExceptionAsync<InvalidOperationException>(() =>
                client.TranslateImageAsync(new byte[0], Settings(ProviderKind.OpenAiCompatible), CancellationToken.None));
            Assert.AreEqual(0, stream.ReadBytes); Assert.IsTrue(stream.Disposed); Assert.AreEqual(1, handler.Calls);
        }

        [DataTestMethod]
        [DataRow(ProviderKind.OpenAiCompatible)] [DataRow(ProviderKind.Ollama)]
        public async Task Image_ExactByteBudgetStillParses(ProviderKind provider)
        {
            var raw = provider == ProviderKind.Ollama ? "{\"message\":{\"content\":\"{\\\"regions\\\":[]}\"}}" :
                "{\"choices\":[{\"message\":{\"content\":\"{\\\"regions\\\":[]}\"},\"finish_reason\":\"stop\"}]}";
            var handler = new Handler { Reply = n => Response(raw + new string(' ', ResponseLimits.Image - raw.Length)) };
            using (var client = new TranslationClient(handler)) Assert.AreEqual(0,
                (await client.TranslateImageAsync(new byte[0], Settings(provider), CancellationToken.None)).Regions.Count);
        }
    }
}
