using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace OfficeTranslate.Core
{
    public enum ChatFailure { Network, Authentication, Model, Budget, Configuration, Protocol, StreamUnsupported, RateLimit, Quota }

    // Messages on this type are controlled locally. Never attach raw provider JSON or parser exceptions.
    public sealed class ChatException : Exception
    {
        public ChatException(ChatFailure failure, string message) : base(message) { Failure = failure; }
        public ChatFailure Failure { get; }
    }

    public static class ChatErrors
    {
        public static string UiKey(Exception error)
        {
            if (error is ChatException chat)
            {
                switch (chat.Failure)
                {
                    case ChatFailure.Authentication: return "ChatAuthError";
                    case ChatFailure.Network: return "ChatNetworkError";
                    case ChatFailure.Budget: return "ChatBudgetError";
                    case ChatFailure.Configuration: return "ChatConfigError";
                    case ChatFailure.RateLimit: return "ChatRateError";
                    case ChatFailure.Quota: return "ChatQuotaError";
                    default: return "ChatModelError";
                }
            }
            return error is IOException || error is HttpRequestException || error is OperationCanceledException ? "ChatNetworkError" : "ChatError";
        }

        public static string Describe(Exception error, TranslationSettings? settings = null, params string[] privateInputs)
        {
            var details = "";
            for (Exception? current = error; current != null; current = current.InnerException)
                details += (details.Length == 0 ? "" : " -> ") + current.GetType().Name + ": " + Redact(current.Message, settings, privateInputs);
            return details;
        }

        private static string Redact(string message, TranslationSettings? settings, IEnumerable<string> privateInputs)
        {
            var secrets = privateInputs.ToList();
            if (settings != null) secrets.AddRange(new[] { settings.ApiKey, settings.OpenAiApiKey, settings.OllamaApiKey,
                settings.ChatQuestionSystemPrompt, settings.ChatSummarySystemPrompt, settings.ChatRewriteSystemPrompt,
                TranslationStyleCatalog.GetPrompt(settings.TranslationStyle), settings.CustomInstructions, settings.Glossary });
            foreach (var secret in secrets.Where(s => !string.IsNullOrEmpty(s)).OrderByDescending(s => s.Length))
            {
                message = message.Replace(secret, "[redacted]");
                // Also remove echoed individual paragraphs rather than requiring the whole document to match.
                foreach (var line in secret.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Where(s => s.Length >= 4))
                    message = message.Replace(line, "[redacted]");
            }
            message = Regex.Replace(message, @"https?://[^\s]+", "[endpoint]", RegexOptions.IgnoreCase);
            return message.Replace('\r', ' ').Replace('\n', ' ');
        }

        internal static Dictionary<string, object> ParseJson(string payload)
        {
            try
            {
                return new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.DeserializeObject(payload) as Dictionary<string, object>
                    ?? throw new ChatException(ChatFailure.Protocol, "Chat response is not a JSON object.");
            }
            catch (Exception ex) when (ex is ArgumentException || ex is InvalidOperationException)
            { throw new ChatException(ChatFailure.Protocol, ex.GetType().Name + ": Invalid chat JSON response."); }
        }

        internal static void CheckFinish(string? reason)
        {
            if (string.IsNullOrEmpty(reason) || reason == "stop") return;
            if (reason == "length") throw new ChatException(ChatFailure.Budget, "Model finish_reason=length: generation exhausted its token limit.");
            if (reason == "content_filter") throw new ChatException(ChatFailure.Model, "Model finish_reason=content_filter: response was filtered.");
            if (reason == "insufficient_system_resource") throw new ChatException(ChatFailure.Model, "Model finish_reason=insufficient_system_resource.");
            throw new ChatException(ChatFailure.Model, "Model did not complete a text answer (non-stop finish reason).");
        }

        internal static void ThrowServerError(Dictionary<string, object> root)
        {
            if (root.TryGetValue("error", out var error)) throw ProviderFailure(0, error, false);
            if (root.ContainsKey("message") && !root.ContainsKey("choices") && !root.ContainsKey("done"))
                throw ProviderFailure(0, root, false);
        }

        internal static ChatException HttpFailure(int status, string raw, bool streaming)
        {
            object? error = null;
            try { var root = ParseJson(raw); error = root.TryGetValue("error", out var e) ? e : root; }
            catch (ChatException) { error = raw; }
            return ProviderFailure(status, error, streaming);
        }

        private static ChatException ProviderFailure(int status, object? error, bool streaming)
        {
            var fields = error as Dictionary<string, object>;
            string Field(string name) => fields != null && fields.TryGetValue(name, out var value) ? Convert.ToString(value) ?? "" : "";
            // Examine messages ONLY to classify; never retain/echo them (they may contain prompts or keys).
            var hint = (Field("code") + " " + Field("type") + " " + Field("message") + " " + (error as string)).ToLowerInvariant();
            ChatFailure kind; string message;
            if (status == 401 || status == 403) { kind = ChatFailure.Authentication; message = "Authentication or permission denied."; }
            else if (status == 402 || hint.Contains("insufficient_quota") || hint.Contains("insufficient_balance")) { kind = ChatFailure.Quota; message = "Account balance or quota exhausted."; }
            else if (status == 429 || hint.Contains("rate_limit")) { kind = ChatFailure.RateLimit; message = "Service rate limit reached."; }
            else if (hint.Contains("context_length") || hint.Contains("context length") || hint.Contains("maximum context") ||
                hint.Contains("max_tokens") || hint.Contains("token limit")) { kind = ChatFailure.Budget; message = "Model rejected the context or output token budget."; }
            else if (hint.Contains("invalid_api_key") || hint.Contains("authentication_error")) { kind = ChatFailure.Authentication; message = "Invalid API key."; }
            else if (streaming && (status == 400 || status == 422 || status == 415 || status == 501) && hint.Contains("stream") &&
                (hint.Contains("support") || hint.Contains("unknown") || hint.Contains("unrecognized"))) { kind = ChatFailure.StreamUnsupported; message = "Service does not support streaming."; }
            else if (hint.Contains("model_not_found") || status == 404) { kind = ChatFailure.Model; message = "Model or API path not found."; }
            else if (hint.Contains("reasoning_content")) { kind = ChatFailure.Model; message = "Model requires reasoning history; check non-thinking mode support."; }
            else { kind = ChatFailure.Model; message = "Model/service rejected the chat request."; }
            return new ChatException(kind, (status > 0 ? "HTTP " + status + ": " : "SSE error: ") + message);
        }

        internal static string NetworkDetails(Exception error, TranslationSettings settings, IReadOnlyList<ChatMessage> messages)
        {
            return Describe(error, settings, messages.Select(m => m.Content).ToArray());
        }
    }
}
