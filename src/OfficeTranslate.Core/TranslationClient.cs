using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace OfficeTranslate.Core
{
    public sealed class TranslationClient : IDisposable
    {
        private readonly HttpClient _http = new HttpClient();
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer();
        private bool _configured;

        public async Task<string> TranslateAsync(string text, TranslationSettings settings, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;
            settings.Validate();
            Configure(settings);
            var chunks = Split(text, settings.MaxCharactersPerChunk);
            var output = new StringBuilder();
            foreach (var chunk in chunks)
            {
                cancellationToken.ThrowIfCancellationRequested();
                output.Append(await TranslateChunkAsync(chunk, settings, cancellationToken).ConfigureAwait(false));
            }
            return output.ToString();
        }

        public async Task<IReadOnlyList<string>> GetModelsAsync(TranslationSettings settings, CancellationToken cancellationToken)
        {
            settings.ValidateEndpoint();
            Configure(settings);
            var endpoint = settings.Provider == ProviderKind.Ollama
                ? settings.BaseUrl.TrimEnd('/') + "/api/tags"
                : settings.BaseUrl.TrimEnd('/') + "/models";
            using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint))
            {
                if (!string.IsNullOrWhiteSpace(settings.ApiKey)) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                using (var response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false))
                {
                    var raw = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode) throw new InvalidOperationException($"获取模型列表失败 {(int)response.StatusCode}: {raw}");
                    var root = _json.DeserializeObject(raw) as Dictionary<string, object>;
                    var arrayKey = settings.Provider == ProviderKind.Ollama ? "models" : "data";
                    var nameKey = settings.Provider == ProviderKind.Ollama ? "name" : "id";
                    if (root == null || !root.TryGetValue(arrayKey, out var value) || !(value is object[] items)) throw new InvalidOperationException("模型服务返回了无法识别的数据。");
                    return items.OfType<Dictionary<string, object>>()
                        .Select(item => item.TryGetValue(nameKey, out var name) ? Convert.ToString(name) : null)
                        .Where(name => !string.IsNullOrWhiteSpace(name)).Cast<string>()
                        .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToArray();
                }
            }
        }

        private void Configure(TranslationSettings settings)
        {
            if (_configured) return;
            _http.Timeout = TimeSpan.FromSeconds(settings.TimeoutSeconds);
            _configured = true;
        }

        private async Task<string> TranslateChunkAsync(string text, TranslationSettings settings, CancellationToken token)
        {
            var endpoint = settings.Provider == ProviderKind.Ollama
                ? settings.BaseUrl.TrimEnd('/') + "/api/chat"
                : settings.BaseUrl.TrimEnd('/') + "/chat/completions";
            var prompt = BuildPrompt(settings);
            object body = settings.Provider == ProviderKind.Ollama
                ? new { model = settings.Model, stream = false, messages = Messages(prompt, text) }
                : new { model = settings.Model, temperature = 0.2, messages = Messages(prompt, text) };
            using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
                request.Content = new StringContent(_json.Serialize(body), Encoding.UTF8, "application/json");
                if (!string.IsNullOrWhiteSpace(settings.ApiKey)) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                using (var response = await _http.SendAsync(request, token).ConfigureAwait(false))
                {
                    var raw = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode) throw new InvalidOperationException($"翻译服务返回 {(int)response.StatusCode}: {raw}");
                    var root = _json.DeserializeObject(raw) as Dictionary<string, object>;
                    return settings.Provider == ProviderKind.Ollama ? ReadOllama(root) : ReadOpenAi(root);
                }
            }
        }

        private static object[] Messages(string prompt, string text) => new object[] {
            new { role = "system", content = prompt }, new { role = "user", content = text }
        };

        private static string BuildPrompt(TranslationSettings s)
        {
            var glossary = string.IsNullOrWhiteSpace(s.Glossary) ? "" : "\n必须遵循以下术语表（每行 source=target）：\n" + s.Glossary;
            var source = string.IsNullOrWhiteSpace(s.SourceLanguage) || s.SourceLanguage == "自动检测"
                ? "自动识别输入的源语言"
                : $"输入语言是{s.SourceLanguage}";
            return $"你是专业翻译。{source}，将输入完整翻译成{s.TargetLanguage}。只输出译文，不解释，不添加标题；保留换行、编号和占位符。{s.CustomInstructions}{glossary}";
        }

        private static string ReadOllama(Dictionary<string, object>? root)
        {
            if (root != null && root.TryGetValue("message", out var m) && m is Dictionary<string, object> msg && msg.TryGetValue("content", out var c)) return Convert.ToString(c) ?? "";
            throw new InvalidOperationException("无法解析 Ollama 响应。");
        }

        private static string ReadOpenAi(Dictionary<string, object>? root)
        {
            if (root != null && root.TryGetValue("choices", out var value) && value is object[] choices && choices.FirstOrDefault() is Dictionary<string, object> choice && choice["message"] is Dictionary<string, object> msg) return Convert.ToString(msg["content"]) ?? "";
            throw new InvalidOperationException("无法解析 OpenAI-compatible 响应。");
        }

        private static IEnumerable<string> Split(string text, int max)
        {
            var position = 0;
            while (position < text.Length)
            {
                var length = Math.Min(max, text.Length - position);
                if (position + length < text.Length)
                {
                    var breakAt = text.LastIndexOfAny(new[] { '\n', '。', '.', '！', '？' }, position + length - 1, length);
                    if (breakAt >= position + max / 2) length = breakAt - position + 1;
                }
                yield return text.Substring(position, length);
                position += length;
            }
        }

        public void Dispose() => _http.Dispose();
    }
}
