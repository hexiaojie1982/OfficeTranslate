using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Net;
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
        private static readonly ConcurrentDictionary<string, PromptDescriptor> PromptDescriptors =
            new ConcurrentDictionary<string, PromptDescriptor>(StringComparer.Ordinal);
        private static readonly ConcurrentDictionary<string, bool> PromptCacheKeySupport =
            new ConcurrentDictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        private bool _configured;

        public async Task<string> TranslateAsync(string text, TranslationSettings settings, CancellationToken cancellationToken)
        {
            return (await TranslateDetailedAsync(text, settings, cancellationToken).ConfigureAwait(false)).Text;
        }

        public async Task<TranslationResult> TranslateDetailedAsync(string text, TranslationSettings settings, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(text)) return TranslationResult.Skipped(text);
            settings.Validate();
            Configure(settings);
            var cacheKey = TranslationSessionCache.CreateKey(settings, text);
            if (TranslationSessionCache.TryGet(cacheKey, out var cached))
                return TranslationResult.Cached(cached);

            var chunks = Split(text, settings.MaxCharactersPerChunk);
            var output = new StringBuilder();
            var anyChanged = false;
            var anyTranslated = false;
            var anySkipped = false;
            var needsReview = false;
            foreach (var chunk in chunks)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var result = await TranslateChunkAsync(chunk, settings, cancellationToken).ConfigureAwait(false);
                output.Append(result.Text);
                anyChanged |= result.Changed;
                anyTranslated |= !result.SkippedNoSource;
                anySkipped |= result.SkippedNoSource;
                needsReview |= result.NeedsReview;
            }
            var completed = output.ToString();
            if (anyChanged && !needsReview)
            {
                TranslationSessionCache.Store(cacheKey, completed);
                return TranslationResult.Translated(completed);
            }
            if (anyChanged) return TranslationResult.Translated(completed, true);
            if (!anyTranslated && anySkipped) return TranslationResult.Skipped(text);
            return TranslationResult.Unchanged(text);
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

        public async Task<IReadOnlyList<ImageTranslationRegion>> TranslateImageAsync(byte[] imageBytes, TranslationSettings settings, CancellationToken token)
        {
            settings.Validate(); Configure(settings);
            var base64 = Convert.ToBase64String(imageBytes);
            var model = string.IsNullOrWhiteSpace(settings.ImageModel) ? settings.Model : settings.ImageModel;
            var automaticImageSource = SourceLanguageProtector.IsAutomatic(settings.SourceLanguage);
            var imageSource = automaticImageSource
                ? "自动识别文字语言，并将所有识别出的语言翻译成目标语言"
                : "图片可能包含多种语言。只翻译其中属于" + settings.SourceLanguage + "的文字；其他语言必须在 translation 字段中逐字原样保留，不得翻译、改写、调整大小写或删除";
            var imageGlossary = string.IsNullOrWhiteSpace(settings.Glossary) ? "" : " 必须遵循术语表：" + settings.Glossary;
            var prompt = "图片本身就是本次需要处理的内容，请立即同时完成 OCR 和翻译，不要要求用户再发送文字。" +
                imageSource + "，目标语言是" + settings.TargetLanguage + "。对于混合语言区域，只替换属于指定源语言的部分。" +
                "每个文字区域必须同时填写 source 原文和 translation 译文，translation 不得留空。" +
                "只返回严格 JSON，不要 Markdown：{\"regions\":[{\"bbox\":[x1,y1,x2,y2],\"source\":\"原文\",\"translation\":\"译文\"}]}。" +
                "坐标必须是相对于图片宽高的 0 到 1000 整数；按阅读顺序返回；没有文字时返回 {\"regions\":[]}。" +
                settings.CustomInstructions + imageGlossary;
            var endpoint = settings.Provider == ProviderKind.Ollama
                ? settings.BaseUrl.TrimEnd('/') + "/api/chat"
                : settings.BaseUrl.TrimEnd('/') + "/chat/completions";
            object body;
            if (settings.Provider == ProviderKind.Ollama)
                body = new { model, stream = false, format = "json", messages = new[] { new { role = "user", content = prompt, images = new[] { base64 } } } };
            else
                body = new { model, temperature = 0.1, messages = new object[] { new { role = "user", content = new object[] { new { type = "text", text = prompt }, new { type = "image_url", image_url = new { url = "data:image/png;base64," + base64 } } } } } };
            using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
                request.Content = new StringContent(_json.Serialize(body), Encoding.UTF8, "application/json");
                if (!string.IsNullOrWhiteSpace(settings.ApiKey)) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                using (var response = await _http.SendAsync(request, token).ConfigureAwait(false))
                {
                    var raw = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode) throw new InvalidOperationException($"图片识别服务返回 {(int)response.StatusCode}: {raw}");
                    var root = _json.DeserializeObject(raw) as Dictionary<string, object>;
                    var content = settings.Provider == ProviderKind.Ollama ? ReadOllama(root) : ReadOpenAi(root);
                    return ParseImageRegions(content);
                }
            }
        }

        private IReadOnlyList<ImageTranslationRegion> ParseImageRegions(string content)
        {
            var start = content.IndexOf('{'); var end = content.LastIndexOf('}');
            if (start < 0 || end <= start) throw new InvalidOperationException("图片模型没有返回有效的 JSON 坐标。");
            var root = _json.DeserializeObject(content.Substring(start, end - start + 1)) as Dictionary<string, object>;
            if (root == null || !root.TryGetValue("regions", out var value) || !(value is object[] items)) throw new InvalidOperationException("图片模型返回的 regions 格式无效。");
            var result = new List<ImageTranslationRegion>();
            foreach (var item in items.OfType<Dictionary<string, object>>())
            {
                if (!item.TryGetValue("bbox", out var boxValue) || !(boxValue is object[] box) || box.Length != 4) continue;
                var region = new ImageTranslationRegion {
                    X1 = Clamp(Convert.ToSingle(box[0])), Y1 = Clamp(Convert.ToSingle(box[1])),
                    X2 = Clamp(Convert.ToSingle(box[2])), Y2 = Clamp(Convert.ToSingle(box[3])),
                    Source = item.TryGetValue("source", out var source) ? Convert.ToString(source) ?? "" : "",
                    Translation = item.TryGetValue("translation", out var translation) ? Convert.ToString(translation) ?? "" : ""
                };
                if (region.X2 > region.X1 && region.Y2 > region.Y1 && !string.IsNullOrWhiteSpace(region.Translation)) result.Add(region);
            }
            return result;
        }

        private static float Clamp(float value)
        {
            if (float.IsNaN(value) || float.IsInfinity(value)) return 0;
            return Math.Max(0, Math.Min(1000, value));
        }

        private void Configure(TranslationSettings settings)
        {
            if (_configured) return;
            _http.Timeout = TimeSpan.FromSeconds(settings.TimeoutSeconds);
            _configured = true;
        }

        private async Task<TranslationResult> TranslateChunkAsync(string text, TranslationSettings settings, CancellationToken token)
        {
            var protectedText = SourceLanguageProtector.Protect(text, settings.SourceLanguage);
            if (protectedText.IsFullyProtected) return TranslationResult.Skipped(text);

            var prompt = GetPromptDescriptor(settings, false, false);
            string translated;
            try
            {
                translated = await SendTranslationRequestAsync(protectedText.Text, prompt, settings, token).ConfigureAwait(false);
                if (string.IsNullOrWhiteSpace(translated))
                    return protectedText.HasProtectedSegments
                        ? await TranslateProtectedPartsAsync(protectedText, settings, token).ConfigureAwait(false)
                        : await RetryUnchangedChunkAsync(text, protectedText, settings, token).ConfigureAwait(false);
                translated = protectedText.Restore(translated);
            }
            catch (ProtectedContentException)
            {
                return await TranslateProtectedPartsAsync(protectedText, settings, token).ConfigureAwait(false);
            }

            if (!Equivalent(text, translated)) return TranslationResult.Translated(translated);
            if (protectedText.HasProtectedSegments)
                return await TranslateProtectedPartsAsync(protectedText, settings, token).ConfigureAwait(false);
            return await RetryUnchangedChunkAsync(text, protectedText, settings, token).ConfigureAwait(false);
        }

        private async Task<TranslationResult> RetryUnchangedChunkAsync(string original, ProtectedTranslationText protectedText, TranslationSettings settings, CancellationToken token)
        {
            var strictPrompt = GetPromptDescriptor(settings, false, true);
            var retried = await SendTranslationRequestAsync(protectedText.Text, strictPrompt, settings, token).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(retried))
                throw new InvalidOperationException("翻译模型连续返回空内容，已停止写回。");
            try { retried = protectedText.Restore(retried); }
            catch (ProtectedContentException)
            {
                if (protectedText.HasProtectedSegments)
                    return await TranslateProtectedPartsAsync(protectedText, settings, token).ConfigureAwait(false);
                throw;
            }
            return Equivalent(original, retried) ? TranslationResult.Unchanged(original) : TranslationResult.Translated(retried);
        }

        private async Task<TranslationResult> TranslateProtectedPartsAsync(ProtectedTranslationText protectedText, TranslationSettings settings, CancellationToken token)
        {
            var output = new StringBuilder();
            var original = new StringBuilder();
            var prompt = GetPromptDescriptor(settings, true, false);
            var strictPrompt = GetPromptDescriptor(settings, true, true);
            var needsReview = false;
            foreach (var part in protectedText.GetParts())
            {
                token.ThrowIfCancellationRequested();
                original.Append(part.Text);
                if (!part.ShouldTranslate || !part.Text.Any(char.IsLetter))
                {
                    output.Append(part.Text);
                    continue;
                }

                var start = 0;
                while (start < part.Text.Length && char.IsWhiteSpace(part.Text[start])) start++;
                var end = part.Text.Length;
                while (end > start && char.IsWhiteSpace(part.Text[end - 1])) end--;
                if (start == end)
                {
                    output.Append(part.Text);
                    continue;
                }

                output.Append(part.Text.Substring(0, start));
                var source = part.Text.Substring(start, end - start);
                var translated = await SendTranslationRequestAsync(source, prompt, settings, token).ConfigureAwait(false);
                if (string.IsNullOrWhiteSpace(translated) || Equivalent(source, translated))
                    translated = await SendTranslationRequestAsync(source, strictPrompt, settings, token).ConfigureAwait(false);
                if (string.IsNullOrWhiteSpace(translated))
                    throw new InvalidOperationException("翻译模型连续返回空内容，已停止写回。");
                if (Equivalent(source, translated))
                {
                    needsReview = true;
                    output.Append(source);
                }
                else output.Append(translated.Trim());
                output.Append(part.Text.Substring(end));
            }

            var completed = output.ToString();
            if (Equivalent(original.ToString(), completed)) return TranslationResult.Unchanged(original.ToString());
            return TranslationResult.Translated(completed, needsReview);
        }

        private async Task<string> SendTranslationRequestAsync(string text, PromptDescriptor prompt, TranslationSettings settings, CancellationToken token)
        {
            var endpoint = settings.Provider == ProviderKind.Ollama
                ? settings.BaseUrl.TrimEnd('/') + "/api/chat"
                : settings.BaseUrl.TrimEnd('/') + "/chat/completions";
            if (settings.Provider == ProviderKind.Ollama)
            {
                var body = new { model = settings.Model, stream = false, messages = Messages(prompt.Text, text) };
                using (var request = CreateRequest(endpoint, body, settings.ApiKey))
                using (var response = await _http.SendAsync(request, token).ConfigureAwait(false))
                {
                    var raw = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode) throw new InvalidOperationException($"翻译服务返回 {(int)response.StatusCode}: {raw}");
                    return ReadOllama(_json.DeserializeObject(raw) as Dictionary<string, object>);
                }
            }

            var capabilityKey = settings.BaseUrl.TrimEnd('/');
            var useCacheKey = !PromptCacheKeySupport.TryGetValue(capabilityKey, out var supported) || supported;
            for (var attempt = 0; attempt < 2; attempt++)
            {
                var body = new Dictionary<string, object>
                {
                    ["model"] = settings.Model,
                    ["temperature"] = 0.2,
                    ["messages"] = Messages(prompt.Text, text)
                };
                if (useCacheKey) body["prompt_cache_key"] = prompt.CacheKey;
                using (var request = CreateRequest(endpoint, body, settings.ApiKey))
                using (var response = await _http.SendAsync(request, token).ConfigureAwait(false))
                {
                    var raw = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (response.IsSuccessStatusCode)
                    {
                        if (useCacheKey) PromptCacheKeySupport[capabilityKey] = true;
                        return ReadOpenAi(_json.DeserializeObject(raw) as Dictionary<string, object>);
                    }
                    if (useCacheKey && IsUnsupportedCacheKey(response.StatusCode, raw))
                    {
                        PromptCacheKeySupport[capabilityKey] = false;
                        useCacheKey = false;
                        continue;
                    }
                    throw new InvalidOperationException($"翻译服务返回 {(int)response.StatusCode}: {raw}");
                }
            }
            throw new InvalidOperationException("翻译服务不支持提示词缓存参数，兼容回退失败。");
        }

        private HttpRequestMessage CreateRequest(string endpoint, object body, string apiKey)
        {
            var request = new HttpRequestMessage(HttpMethod.Post, endpoint) {
                Content = new StringContent(_json.Serialize(body), Encoding.UTF8, "application/json")
            };
            if (!string.IsNullOrWhiteSpace(apiKey)) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
            return request;
        }

        private static bool IsUnsupportedCacheKey(HttpStatusCode statusCode, string response)
        {
            if ((int)statusCode != 400 && (int)statusCode != 422) return false;
            var raw = response ?? string.Empty;
            return raw.IndexOf("prompt_cache_key", StringComparison.OrdinalIgnoreCase) >= 0 ||
                raw.IndexOf("unknown field", StringComparison.OrdinalIgnoreCase) >= 0 ||
                raw.IndexOf("unrecognized", StringComparison.OrdinalIgnoreCase) >= 0 ||
                raw.IndexOf("extra input", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool Equivalent(string source, string translated) =>
            string.Equals(NormalizeForComparison(source), NormalizeForComparison(translated), StringComparison.Ordinal);

        private static string NormalizeForComparison(string text) =>
            (text ?? string.Empty).Trim().Replace("\r\n", "\n").Replace('\r', '\n');

        private static object[] Messages(string prompt, string text) => new object[] {
            new { role = "system", content = prompt }, new { role = "user", content = text }
        };

        private static PromptDescriptor GetPromptDescriptor(TranslationSettings settings, bool sourcePart, bool strict)
        {
            var descriptorKey = string.Join("\u001F", new[]
            {
                sourcePart ? "part" : "full",
                strict ? "strict" : "normal",
                settings.SourceLanguage ?? string.Empty,
                settings.TargetLanguage ?? string.Empty,
                settings.CustomInstructions ?? string.Empty,
                settings.Glossary ?? string.Empty
            });
            return PromptDescriptors.GetOrAdd(descriptorKey, _ =>
            {
                var prompt = sourcePart ? BuildSourcePartPrompt(settings) : BuildPrompt(settings);
                if (strict)
                    prompt += "\n重要：上一次结果没有完成翻译。本次必须将输入中的源语言内容实际转换为目标语言；不得原样返回输入，不得返回空内容。";
                return new PromptDescriptor(prompt, "ot-" + TranslationSessionCache.Hash(prompt));
            });
        }

        private static string BuildPrompt(TranslationSettings s)
        {
            var glossary = string.IsNullOrWhiteSpace(s.Glossary) ? "" : "\n必须遵循以下术语表（每行 source=target）：\n" + s.Glossary;
            if (SourceLanguageProtector.IsAutomatic(s.SourceLanguage))
                return $"你是专业翻译。自动识别输入的源语言，将输入完整翻译成{s.TargetLanguage}。只输出译文，不解释，不添加标题；保留换行、编号和占位符。{s.CustomInstructions}{glossary}";

            return $"你是专业翻译。输入可能包含多种语言，只翻译其中属于{s.SourceLanguage}的文本片段，将其翻译成{s.TargetLanguage}。" +
                "所有非源语言内容必须逐字原样保留，包括其他语言的单词和句子、产品名称、型号、缩写、网址、邮箱、代码及大小写；不得翻译、改写、解释、移动或删除。" +
                "形如 ⟦OT_KEEP_0001⟧ 的保护占位符必须完整、原样、按原位置输出，绝对不能修改。" +
                $"只输出处理后的完整文本，不解释，不添加标题；保留换行和编号。{s.CustomInstructions}{glossary}";
        }

        private static string BuildSourcePartPrompt(TranslationSettings s)
        {
            var glossary = string.IsNullOrWhiteSpace(s.Glossary) ? "" : "\n必须遵循以下术语表（每行 source=target）：\n" + s.Glossary;
            return $"你是专业翻译。当前输入只包含需要翻译的{s.SourceLanguage}文本片段，请将其翻译成{s.TargetLanguage}。" +
                $"只输出译文，不解释，不添加标题或原文，不得输出 OT_KEEP 等内部标记。{s.CustomInstructions}{glossary}";
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

        private sealed class PromptDescriptor
        {
            public PromptDescriptor(string text, string cacheKey) { Text = text; CacheKey = cacheKey; }
            public string Text { get; }
            public string CacheKey { get; }
        }

        public void Dispose() => _http.Dispose();
    }

    public sealed class ImageTranslationRegion
    {
        public float X1 { get; set; } public float Y1 { get; set; } public float X2 { get; set; } public float Y2 { get; set; }
        public string Source { get; set; } = string.Empty; public string Translation { get; set; } = string.Empty;
    }
}
