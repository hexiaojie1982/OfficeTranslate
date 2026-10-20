using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeTranslate.Core
{
    public static class ChatSseParser
    {
        public static async Task<string> ReadAsync(TextReader reader, Action<string> output, CancellationToken token)
        {
            var text = new StringBuilder();
            var data = new StringBuilder();
            var done = false;
            var firstLine = true;
            var line = new StringBuilder();
            var buffer = new char[4096];
            var total = 0;
            var afterCr = false;
            void ConsumeLine()
            {
                var value = line.ToString(); line.Clear();
                if (firstLine) { value = value.TrimStart('\uFEFF'); firstLine = false; }
                if (value.Length == 0) done = ReadEvent(data, text, output);
                else if (value.StartsWith("data:", StringComparison.Ordinal))
                {
                    value = value.Substring(5);
                    if (value.StartsWith(" ", StringComparison.Ordinal)) value = value.Substring(1);
                    if (data.Length + value.Length + 1 > ResponseLimits.Chat) throw ChatErrors.ResponseTooLarge();
                    data.Append(value).Append('\n');
                }
                // Comments/keep-alives, event, id and retry fields do not carry completion text.
            }
            while (!done)
            {
                token.ThrowIfCancellationRequested();
                var count = await reader.ReadAsync(buffer, 0, Math.Min(buffer.Length, ResponseLimits.Chat - total + 1)).ConfigureAwait(false);
                if (count == 0) break;
                total += count;
                if (total > ResponseLimits.Chat) throw ChatErrors.ResponseTooLarge();
                for (var i = 0; i < count && !done; i++)
                {
                    var ch = buffer[i];
                    if (afterCr && ch == '\n') { afterCr = false; continue; }
                    afterCr = ch == '\r';
                    if (ch == '\r' || ch == '\n') ConsumeLine();
                    else line.Append(ch);
                }
            }
            token.ThrowIfCancellationRequested();
            if (!done && line.Length > 0) ConsumeLine();
            if (!done && data.Length > 0) done = ReadEvent(data, text, output);
            if (!done || string.IsNullOrWhiteSpace(text.ToString()))
                throw new ChatException(ChatFailure.Protocol, "Interrupted or empty SSE response.");
            return text.ToString();
        }

        private static bool ReadEvent(StringBuilder data, StringBuilder text, Action<string> output)
        {
            if (data.Length == 0) return false;
            var payload = data.ToString().TrimEnd('\n'); data.Clear();
            if (payload.Trim() == "[DONE]") return true;
            var root = ChatErrors.ParseJson(payload);
            ChatErrors.ThrowServerError(root);
            if (!root.TryGetValue("choices", out var c) || !(c is object[] choices))
                throw new ChatException(ChatFailure.Protocol, "SSE event has no choices array.");
            // Usage-only events have no choices; reasoning_content is deliberately not answer text.
            var choice = choices.OfType<Dictionary<string, object>>().FirstOrDefault(item =>
                !item.TryGetValue("index", out var index) || Convert.ToString(index) == "0");
            if (choice != null)
            {
                if (choice.TryGetValue("finish_reason", out var reason)) ChatErrors.CheckFinish(Convert.ToString(reason));
                if (choice.TryGetValue("delta", out var d) && d is Dictionary<string, object> delta &&
                    delta.TryGetValue("content", out var value) && value is string content && content.Length > 0)
                {
                    if (text.Length + content.Length > ResponseLimits.Chat) throw ChatErrors.ResponseTooLarge();
                    text.Append(content); output(text.ToString());
                }
            }
            return false;
        }
    }

    public sealed partial class TranslationClient
    {
        public async Task<ChatReply> ChatAsync(IReadOnlyList<ChatMessage> messages, TranslationSettings settings, int outputTokens,
            Action<string> replaceOutput, CancellationToken token)
        {
            if (!Uri.TryCreate(settings.BaseUrl, UriKind.Absolute, out var uri) ||
                (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps) || string.IsNullOrWhiteSpace(settings.Model))
                throw new ChatException(ChatFailure.Configuration, "Chat requires an HTTP(S) base URL and a model.");
            if (outputTokens < 1 || settings.TimeoutSeconds < 1)
                throw new ChatException(ChatFailure.Budget, "Chat output budget and timeout must be positive.");
            // Match translation's configured timeout rather than HttpClient's implicit 100 seconds.
            Configure(settings);
            var ollama = settings.Provider == ProviderKind.Ollama;
            var endpoint = settings.BaseUrl.TrimEnd('/') + (ollama ? "/api/chat" : "/chat/completions");
            ChatException? streamingFailure = null;
            if (!ollama)
            {
                try { return await SendChatAttemptAsync(endpoint, messages, settings, outputTokens, true, replaceOutput, token).ConfigureAwait(false); }
                catch (ChatException ex) when (!ex.ResponseLimitExceeded && (ex.Failure == ChatFailure.Protocol || ex.Failure == ChatFailure.Network || ex.Failure == ChatFailure.StreamUnsupported))
                {
                    token.ThrowIfCancellationRequested();
                    streamingFailure = ex;
                    replaceOutput("");
                }
            }
            try { return await SendChatAttemptAsync(endpoint, messages, settings, outputTokens, false, replaceOutput, token).ConfigureAwait(false); }
            catch (ChatException ex) when (streamingFailure != null)
            {
                // Preserve both attempts using only controlled diagnostics, never response bodies.
                throw new ChatException(ex.Failure, ex.Message + " Streaming attempt: " + ChatErrors.Describe(streamingFailure));
            }
        }

        private async Task<ChatReply> SendChatAttemptAsync(string endpoint, IReadOnlyList<ChatMessage> messages,
            TranslationSettings settings, int outputTokens, bool streaming, Action<string> output, CancellationToken token)
        {
            var ollama = settings.Provider == ProviderKind.Ollama;
            var body = new Dictionary<string, object> {
                ["model"] = settings.Model,
                ["messages"] = messages.Select(m => new { role = m.Role, content = m.Content }).ToArray(),
                ["stream"] = streaming
            };
            if (ollama) body["options"] = new { num_predict = outputTokens, num_ctx = settings.ChatContextTokens, temperature = 0.2 };
            else
            {
                body["temperature"] = 0.2; body["max_tokens"] = outputTokens;
                // Same provider policy as translation, for BOTH stream and recovery requests.
                // Do not send this extension to vLLM or other compatible servers.
                if (IsOfficialDeepSeekEndpoint(settings.BaseUrl)) body["thinking"] = new { type = "disabled" };
            }
            using (var deadline = CancellationTokenSource.CreateLinkedTokenSource(token))
            {
                deadline.CancelAfter(TimeSpan.FromSeconds(settings.TimeoutSeconds));
                try
                {
                    using (var request = CreateRequest(endpoint, body, settings.ApiKey))
                    {
                        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue(streaming ? "text/event-stream" : "application/json"));
                        using (var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, deadline.Token).ConfigureAwait(false))
                        {
                            if (response.Content.Headers.ContentLength > ResponseLimits.Chat) throw ChatErrors.ResponseTooLarge();
                            using (var stream = await ResponseLimits.OpenStreamAsync(response.Content, deadline.Token).ConfigureAwait(false))
                            using (deadline.Token.Register(() => stream.Dispose()))
                            using (var limited = new LimitedReadStream(stream, ResponseLimits.Chat, ChatErrors.ResponseTooLarge))
                            using (var reader = new StreamReader(limited, Encoding.UTF8))
                            {
                                var sse = string.Equals(response.Content.Headers.ContentType?.MediaType, "text/event-stream", StringComparison.OrdinalIgnoreCase);
                                if (response.IsSuccessStatusCode && streaming && sse)
                                    return new ChatReply(await ChatSseParser.ReadAsync(reader, output, deadline.Token).ConfigureAwait(false), false);
                                var raw = await ReadChatJsonAsync(reader, deadline.Token).ConfigureAwait(false);
                                deadline.Token.ThrowIfCancellationRequested();
                                if (!response.IsSuccessStatusCode) throw ChatErrors.HttpFailure((int)response.StatusCode, raw, streaming);
                                var root = ChatErrors.ParseJson(raw);
                                ChatErrors.ThrowServerError(root);
                                string? text = null;
                                if (ollama)
                                {
                                    if (root.TryGetValue("done_reason", out var reason)) ChatErrors.CheckFinish(Convert.ToString(reason));
                                    if (root.TryGetValue("message", out var m) && m is Dictionary<string, object> message && message.TryGetValue("content", out var content)) text = content as string;
                                    if (root.TryGetValue("done", out var done) && done is bool completed && !completed)
                                        throw new ChatException(ChatFailure.Protocol, "Incomplete Ollama response.");
                                }
                                else if (root.TryGetValue("choices", out var c) && c is object[] choices && choices.FirstOrDefault() is Dictionary<string, object> choice)
                                {
                                    if (choice.TryGetValue("finish_reason", out var reason)) ChatErrors.CheckFinish(Convert.ToString(reason));
                                    if (choice.TryGetValue("message", out var m) && m is Dictionary<string, object> message && message.TryGetValue("content", out var content)) text = content as string;
                                }
                                if (string.IsNullOrWhiteSpace(text)) throw new ChatException(ChatFailure.Protocol, "Empty or unrecognized chat response.");
                                output(text!);
                                // Some compatible servers ignore stream=true and return a complete JSON reply.
                                return new ChatReply(text!, true);
                            }
                        }
                    }
                }
                catch (ChatException) { throw; }
                catch (Exception ex) when (ex is HttpRequestException || ex is IOException || ex is OperationCanceledException || ex is ObjectDisposedException)
                {
                    token.ThrowIfCancellationRequested();
                    if (deadline.IsCancellationRequested || ex is OperationCanceledException)
                        throw new ChatException(ChatFailure.Network, ex.GetType().Name + ": Chat request timed out after " + settings.TimeoutSeconds + " seconds.");
                    throw new ChatException(ChatFailure.Network, ChatErrors.NetworkDetails(ex, settings, messages));
                }
            }
        }

        private static async Task<string> ReadChatJsonAsync(TextReader reader, CancellationToken token)
        {
            var text = new StringBuilder();
            var buffer = new char[4096];
            while (true)
            {
                token.ThrowIfCancellationRequested();
                var count = await reader.ReadAsync(buffer, 0, Math.Min(buffer.Length, ResponseLimits.Chat - text.Length + 1)).ConfigureAwait(false);
                token.ThrowIfCancellationRequested();
                if (count == 0) return text.ToString();
                if (text.Length + count > ResponseLimits.Chat) throw ChatErrors.ResponseTooLarge();
                text.Append(buffer, 0, count);
            }
        }
    }
}
