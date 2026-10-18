using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeTranslate.Core
{
    public enum ChatMode { Question, Summarize, Rewrite, Translate }
    public enum ChatWriteAction { ReplaceSelection, InsertBelow, InsertAtStart }
    public sealed class ChatMessage
    {
        public ChatMessage(string role, string content) : this(role, content, DateTimeOffset.Now) { }
        public ChatMessage(string role, string content, DateTimeOffset timestamp) { Role = role; Content = content; Timestamp = timestamp; }
        public string Role { get; }
        public string Content { get; }
        public DateTimeOffset Timestamp { get; }
    }
    // The host captures the target before any await. Document context is never logged or persisted.
    public sealed class ChatContext
    {
        public ChatContext(string text, bool selected, Action<ChatWriteAction, string>? write = null, bool canReplace = false, string documentName = "", string host = "")
        { Text = text ?? ""; IsSelection = selected; Write = write; CanReplace = selected && canReplace; DocumentName = documentName; Host = host; }
        public string Text { get; }
        public string DocumentName { get; }
        public string Host { get; }
        public bool IsSelection { get; }
        public bool CanReplace { get; }
        public Action<ChatWriteAction, string>? Write { get; }
        public static ChatContext Capture(string? selection, Func<string> fullText, Action<ChatWriteAction, string>? write = null, bool canReplace = false, string documentName = "", string host = "") =>
            !string.IsNullOrWhiteSpace(selection) ? new ChatContext(selection!, true, write, canReplace, documentName, host) : new ChatContext(fullText(), false, write, documentName: documentName, host: host);
        public bool Allows(ChatMode mode, ChatWriteAction action, string result) => Write != null && !string.IsNullOrWhiteSpace(result) &&
            (mode == ChatMode.Summarize ? action == ChatWriteAction.InsertAtStart : (mode == ChatMode.Rewrite || mode == ChatMode.Translate) &&
             (action == ChatWriteAction.InsertBelow || action == ChatWriteAction.ReplaceSelection && CanReplace));
        public void Apply(ChatMode mode, ChatWriteAction action, string result)
        {
            if (!Allows(mode, action, result)) throw new InvalidOperationException("Write action unavailable.");
            Write!(action, result);
        }
    }
    public static class ChatBudget
    {
        // Conservative tokenizer-independent upper bound: UTF-8 bytes. Providers may use different tokenizers.
        public static int Estimate(string text) => Encoding.UTF8.GetByteCount(text ?? "");
        public static IReadOnlyList<string> Split(string text, int budget)
        {
            if (budget < 4) throw new ArgumentOutOfRangeException(nameof(budget));
            var chunks = new List<string>(); var start = 0; var size = 0;
            for (var i = 0; i < text.Length; i++)
            {
                var count = char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]) ? 2 : 1;
                var bytes = Encoding.UTF8.GetByteCount(text.Substring(i, count));
                if (size + bytes > budget) { chunks.Add(text.Substring(start, i - start)); start = i; size = 0; }
                size += bytes; i += count - 1;
            }
            if (start < text.Length) chunks.Add(text.Substring(start));
            return chunks;
        }
        public static int Cost(IEnumerable<ChatMessage> messages) => messages.Sum(m => Estimate(m.Content) + 16) + 32;
    }
    public sealed class ChatReply
    {
        public ChatReply(string text, bool fallback, bool summarized = false) { Text = text; UsedFallback = fallback; UsedSummaries = summarized; }
        public string Text { get; }
        public bool UsedFallback { get; }
        public bool UsedSummaries { get; }
    }
    public delegate Task<ChatReply> ChatTransport(IReadOnlyList<ChatMessage> messages, TranslationSettings settings, int outputTokens,
        Action<string> replaceOutput, CancellationToken token);
    public sealed class DocumentChatSession
    {
        private readonly List<ChatMessage> _history = new List<ChatMessage>();
        private readonly object _historySync = new object();
        private readonly ChatHistoryStore? _store;
        private readonly List<ChatHistoryTurn> _unsaved = new List<ChatHistoryTurn>();
        private string _sessionId = Guid.NewGuid().ToString("N"), _documentName = "", _host = "";
        private DateTimeOffset _createdAt;
        private int _generation;
        public DocumentChatSession(ChatHistoryStore? store = null) { _store = store; }
        public bool HistorySaveFailed { get { lock (_historySync) return _unsaved.Count > 0; } }
        public IReadOnlyList<ChatMessage> History { get { lock (_historySync) return _history.ToArray(); } }
        // A failed save must not silently discard completed turns on Clear.
        public bool Clear()
        {
            lock (_historySync)
            {
                if (!TrySaveHistory()) return false;
                _generation++; _history.Clear(); _sessionId = Guid.NewGuid().ToString("N"); _createdAt = default; _documentName = ""; _host = "";
                return true;
            }
        }
        public bool TrySaveHistory()
        {
            lock (_historySync)
            {
                if (_store == null || _unsaved.Count == 0) return true;
                try { _store.Append(_sessionId, _createdAt, _documentName, _unsaved, _host); _unsaved.Clear(); return true; }
                catch (Exception ex) when (ChatHistoryStore.IsStorageError(ex)) { return false; }
            }
        }
        public async Task<ChatReply> SendAsync(ChatContext context, string question, ChatMode mode, TranslationSettings settings,
            ChatTransport transport, Action<string> output, Action<int, int> progress, CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(context.Text) || string.IsNullOrWhiteSpace(question)) throw new InvalidOperationException("Empty chat input.");
            if (settings.ChatContextTokens < 2048 || settings.ChatContextTokens > 1048576 || settings.ChatOutputTokens < 256 ||
                settings.ChatOutputTokens > settings.ChatContextTokens / 2) throw new ChatException(ChatFailure.Budget, "Invalid chat context/output budget.");
            var system = mode == ChatMode.Translate ? settings.ChatTranslateSystemPrompt : mode == ChatMode.Summarize ? settings.ChatSummarySystemPrompt : mode == ChatMode.Rewrite ? settings.ChatRewriteSystemPrompt : settings.ChatQuestionSystemPrompt;
            system += "\n" + ChatPrompts.Safety(settings.UiLanguage, mode);
            system = ChatPrompts.AppendTranslationSettings(system, settings);
            var messages = new List<ChatMessage> { new ChatMessage("system", system) };
            var inputBudget = settings.ChatContextTokens - settings.ChatOutputTokens;
            var fixedCost = ChatBudget.Cost(messages) + ChatBudget.Estimate(question) + 128;
            if (fixedCost > inputBudget / 2) throw new ChatException(ChatFailure.Budget, "Chat question and system prompt exceed the input budget.");
            // Keep recent complete turns only; reserve at least half the remaining input for context.
            List<ChatMessage> history; int generation;
            lock (_historySync) { history = _history.ToList(); generation = _generation; }
            while (history.Count > 0 && ChatBudget.Cost(history) > (inputBudget - fixedCost) / 2) history.RemoveRange(0, 2);
            messages.AddRange(history);
            var contextBudget = inputBudget - ChatBudget.Cost(messages) - ChatBudget.Estimate(question) - 128;
            var text = context.Text; var usedFallback = false; var summarized = false;
            var mapSystem = ChatPrompts.AppendTranslationSettings(ChatPrompts.Map(settings.UiLanguage), settings);
            var mapBudget = settings.ChatContextTokens - 256 - ChatBudget.Estimate(mapSystem) - ChatBudget.Estimate(question) - 160;
            while (ChatBudget.Estimate(text) > contextBudget)
            {
                token.ThrowIfCancellationRequested(); summarized = true;
                var chunks = ChatBudget.Split(text, Math.Min(contextBudget, mapBudget));
                var summaries = new List<string>();
                for (var i = 0; i < chunks.Count; i++)
                {
                    progress(i + 1, chunks.Count);
                    var map = new[] { new ChatMessage("system", mapSystem), new ChatMessage("user", question + "\n<fragment>\n" + chunks[i] + "\n</fragment>") };
                    var part = await transport(map, settings, 256, _ => { }, token).ConfigureAwait(false);
                    usedFallback |= part.UsedFallback; summaries.Add(part.Text);
                }
                var reduced = string.Join("\n", summaries);
                if (ChatBudget.Estimate(reduced) >= ChatBudget.Estimate(text)) throw new ChatException(ChatFailure.Budget, "Summaries did not reduce context to fit the input budget.");
                text = reduced;
            }
            messages.Add(new ChatMessage("user", question + "\n<document>\n" + text + "\n</document>"));
            if (ChatBudget.Cost(messages) > inputBudget) throw new ChatException(ChatFailure.Budget, "Chat input exceeds the context budget.");
            progress(0, 0);
            var reply = await transport(messages, settings, settings.ChatOutputTokens, output, token).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            // Commit one canonical answer after streaming completes, before any
            // history write or final bubble/export/copy consumes the response.
            var answer = ChatCjkSpacing.Clean(reply.Text);
            lock (_historySync)
            {
                token.ThrowIfCancellationRequested();
                if (generation != _generation) throw new OperationCanceledException("Chat session cleared.");
                var timestamp = DateTimeOffset.Now;
                _history.Add(new ChatMessage("user", question, timestamp)); _history.Add(new ChatMessage("assistant", answer, timestamp));
                if (_history.Count > 20) _history.RemoveRange(0, _history.Count - 20);
                if (_store != null)
                {
                    if (_createdAt == default) { _createdAt = timestamp; _documentName = context.DocumentName; _host = context.Host; }
                    _unsaved.Add(new ChatHistoryTurn { Timestamp = timestamp.ToString("O", System.Globalization.CultureInfo.InvariantCulture), Question = question, Answer = answer });
                    TrySaveHistory();
                }
            }
            return new ChatReply(answer, usedFallback || reply.UsedFallback, summarized);
        }
    }
}
