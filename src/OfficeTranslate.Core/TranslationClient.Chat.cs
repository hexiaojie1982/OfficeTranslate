using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace OfficeTranslate.Core
{
    public static class ChatSseParser
    {
        public static async Task<string> ReadAsync(TextReader reader, Action<string> output, CancellationToken token)
        {
            var text = new StringBuilder(); var data = new StringBuilder(); var done = false;
            string? line;
            while ((line = await reader.ReadLineAsync().ConfigureAwait(false)) != null)
            {
                token.ThrowIfCancellationRequested();
                if (line.Length == 0)
                {
                    if (data.Length == 0) continue;
                    var payload = data.ToString().TrimEnd('\n'); data.Clear();
                    if (payload.Trim() == "[DONE]") { done = true; break; }
                    var root = new JavaScriptSerializer().DeserializeObject(payload) as Dictionary<string, object>;
                    if (root == null || root.ContainsKey("error") || !root.TryGetValue("choices", out var c) || !(c is object[] choices))
                        throw new InvalidDataException("Invalid SSE event.");
                    if (choices.FirstOrDefault() is Dictionary<string, object> choice)
                    {
                        if (choice.TryGetValue("finish_reason", out var reason) && reason != null && Convert.ToString(reason) != "stop")
                            throw new InvalidDataException("Incomplete chat response.");
                        if (choice.TryGetValue("delta", out var d) && d is Dictionary<string, object> delta && delta.TryGetValue("content", out var value) && value != null)
                        { text.Append(Convert.ToString(value)); output(text.ToString()); }
                    }
                }
                else if (line.StartsWith("data:", StringComparison.Ordinal)) data.Append(line.Substring(5).TrimStart(' ')).Append('\n');
            }
            if (!done || string.IsNullOrWhiteSpace(text.ToString())) throw new InvalidDataException("Interrupted or empty SSE response.");
            return text.ToString();
        }
    }
    public sealed partial class TranslationClient
    {
        public async Task<ChatReply> ChatAsync(IReadOnlyList<ChatMessage> messages, TranslationSettings settings, int outputTokens,
            Action<string> replaceOutput, CancellationToken token)
        {
            settings.ValidateEndpoint();
            if (string.IsNullOrWhiteSpace(settings.Model)) throw new InvalidOperationException("Missing model.");
            var wire = messages.Select(m => new { role = m.Role, content = m.Content }).ToArray();
            var ollama = settings.Provider == ProviderKind.Ollama;
            var endpoint = settings.BaseUrl.TrimEnd('/') + (ollama ? "/api/chat" : "/chat/completions");
            using (var deadline = CancellationTokenSource.CreateLinkedTokenSource(token))
            {
                deadline.CancelAfter(TimeSpan.FromSeconds(Math.Max(1, settings.TimeoutSeconds)));
                // Linked deadlines bound each attempt without changing translation defaults.
                {
                    // Use the injectable transport in tests and the configured text transport in production.
                    var transport = _http;
                    var fallback = ollama;
                    if (!ollama)
                    {
                        try
                        {
                            var body = new { model = settings.Model, messages = wire, stream = true, temperature = 0.2, max_tokens = outputTokens };
                            using (var request = CreateRequest(endpoint, body, settings.ApiKey))
                            using (var response = await transport.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, deadline.Token).ConfigureAwait(false))
                            {
                                response.EnsureSuccessStatusCode();
                                if (response.Content.Headers.ContentType?.MediaType != "text/event-stream") throw new InvalidDataException("SSE unsupported.");
                                using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                                using (deadline.Token.Register(() => stream.Dispose()))
                                using (var reader = new StreamReader(stream, Encoding.UTF8))
                                {
                                    var text = await ChatSseParser.ReadAsync(reader, replaceOutput, deadline.Token).ConfigureAwait(false);
                                    return new ChatReply(text, false);
                                }
                            }
                        }
                        catch (Exception ex) when (!token.IsCancellationRequested && (ex is HttpRequestException || ex is InvalidDataException || ex is IOException || ex is InvalidOperationException || ex is ArgumentException || ex is OperationCanceledException || ex is ObjectDisposedException))
                        { fallback = true; replaceOutput(""); }
                    }
                    token.ThrowIfCancellationRequested();
                    // A fresh bounded deadline permits recovery even when the streaming deadline expired.
                    deadline.CancelAfter(Timeout.Infinite);
                    using (var recovery = CancellationTokenSource.CreateLinkedTokenSource(token))
                    {
                        recovery.CancelAfter(TimeSpan.FromSeconds(Math.Max(1, settings.TimeoutSeconds)));
                        object body = ollama ? (object)new { model = settings.Model, messages = wire, stream = false,
                            options = new { num_predict = outputTokens, num_ctx = settings.ChatContextTokens, temperature = 0.2 } } :
                            new { model = settings.Model, messages = wire, stream = false, temperature = 0.2, max_tokens = outputTokens };
                        using (var request = CreateRequest(endpoint, body, settings.ApiKey))
                        using (var response = await transport.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, recovery.Token).ConfigureAwait(false))
                        {
                            response.EnsureSuccessStatusCode();
                            using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                            using (recovery.Token.Register(() => stream.Dispose()))
                            using (var reader = new StreamReader(stream, Encoding.UTF8))
                            {
                                var raw = await reader.ReadToEndAsync().ConfigureAwait(false);
                                var root = DeserializeJson(raw) as Dictionary<string, object>;
                                var result = ollama ? ReadOllama(root) : ReadOpenAi(root);
                                var finish = ReadImageFinishReason(root, settings.Provider);
                                if (result.Truncated || (finish != "unknown" && finish != "stop" && finish != "") || string.IsNullOrWhiteSpace(result.Content))
                                    throw new InvalidDataException("Incomplete chat response.");
                                replaceOutput(result.Content); return new ChatReply(result.Content, fallback);
                            }
                        }
                    }
                }
            }
        }
    }
}
