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
        // Image OCR embeds a whole PNG as base64 in the request body, which can be
        // megabytes long. Raising MaxJsonLength removes the serializer as a
        // possible failure point for large payloads. (The exact default
        // threshold was not measured in the target .NET Framework 4.8 runtime;
        // verify there before quoting a number.)
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
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
            try
            {
                return await RequestImageRegionsAsync(endpoint, model, prompt, base64, settings, token, true).ConfigureAwait(false);
            }
            catch (ImageFormatNotSupportedException)
            {
                // Older OpenAI-compatible servers may reject the structured-output
                // parameter; fall back to asking for JSON through the prompt only.
                return await RequestImageRegionsAsync(endpoint, model, prompt, base64, settings, token, false).ConfigureAwait(false);
            }
        }

        private sealed class ImageFormatNotSupportedException : InvalidOperationException
        {
            public ImageFormatNotSupportedException(string message) : base(message) { }
        }

        private async Task<IReadOnlyList<ImageTranslationRegion>> RequestImageRegionsAsync(string endpoint, string model, string prompt, string base64, TranslationSettings settings, CancellationToken token, bool structuredOutput)
        {
            object body;
            if (settings.Provider == ProviderKind.Ollama)
            {
                var messages = new[] { new { role = "user", content = prompt, images = new[] { base64 } } };
                body = structuredOutput
                    ? (object)new { model, stream = false, format = "json", messages }
                    : new { model, stream = false, messages };
            }
            else
            {
                var messages = new object[] { new { role = "user", content = new object[] { new { type = "text", text = prompt }, new { type = "image_url", image_url = new { url = "data:image/png;base64," + base64 } } } } };
                body = structuredOutput
                    ? (object)new { model, temperature = 0.1, response_format = new { type = "json_object" }, messages }
                    : new { model, temperature = 0.1, messages };
            }
            using (var response = await SendWithRetryAsync(() => CreateRequest(endpoint, body, settings.ApiKey), token).ConfigureAwait(false))
            {
                var raw = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                {
                    if (ShouldFallbackFromStructuredOutput((int)response.StatusCode, raw))
                        throw new ImageFormatNotSupportedException($"图片识别服务不支持结构化输出参数，已回退为纯提示词模式: {raw}");
                    throw new InvalidOperationException($"图片识别服务返回 {(int)response.StatusCode}: {raw}");
                }
                var root = _json.DeserializeObject(raw) as Dictionary<string, object>;
                var imageResponse = settings.Provider == ProviderKind.Ollama ? ReadOllama(root) : ReadOpenAi(root);
                if (imageResponse.Truncated)
                    throw new InvalidOperationException("图片模型的输出因长度限制被截断，未生成译文覆盖框。请更换视觉模型或缩小图片后重试。");
                return ParseImageRegions(imageResponse.Content);
            }
        }

        // Some OpenAI-compatible servers (e.g. vLLM front-ends) answer 422 instead
        // of 400 when they do not understand response_format.
        internal static bool ShouldFallbackFromStructuredOutput(int statusCode, string raw) =>
            (statusCode == 400 || statusCode == 422) && IsFormatRejection(raw);

        private static bool IsFormatRejection(string raw) =>
            raw != null && (raw.IndexOf("response_format", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            raw.IndexOf("\"format\"", StringComparison.OrdinalIgnoreCase) >= 0);

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
                // Keep the model's raw coordinates (no 0-1000 clamping): range and
                // sanity validation is the overlay planner's job, so abnormal
                // values stay visible in diagnostics instead of being masked.
                var region = new ImageTranslationRegion {
                    X1 = ToCoordinate(box[0]), Y1 = ToCoordinate(box[1]),
                    X2 = ToCoordinate(box[2]), Y2 = ToCoordinate(box[3]),
                    Source = item.TryGetValue("source", out var source) ? Convert.ToString(source) ?? "" : "",
                    Translation = item.TryGetValue("translation", out var translation) ? Convert.ToString(translation) ?? "" : ""
                };
                // A region without translated text has nothing to overlay; a
                // geometrically invalid bbox is passed through so the planner
                // can downgrade it to an explicit side-note instead of being
                // silently dropped here.
                if (!string.IsNullOrWhiteSpace(region.Translation)) result.Add(region);
            }
            return result;
        }

        private static float ToCoordinate(object value)
        {
            // Non-finite or unparsable coordinates become NaN, which the
            // planner's range check rejects with a side-note. One corrupt
            // region no longer aborts the whole image.
            try
            {
                var parsed = Convert.ToSingle(value);
                return float.IsNaN(parsed) || float.IsInfinity(parsed) ? float.NaN : parsed;
            }
            catch { return float.NaN; }
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

            if (protectedText.HasProtectedSegments)
                return await TranslateMixedLanguageChunkAsync(text, protectedText, settings, token).ConfigureAwait(false);

            var prompt = GetPromptDescriptor(settings, PromptMode.Direct);
            var response = await SendTranslationRequestAsync(text, prompt, settings, token).ConfigureAwait(false);
            var translated = response.Content;
            if (string.IsNullOrWhiteSpace(translated))
                return await RetryUnchangedChunkAsync(text, protectedText, settings, token).ConfigureAwait(false);

            // A response cut off by the output token limit is never treated as a
            // complete translation; the strict retry below uses an expanded budget.
            if (!response.Truncated && IsCompleteTranslation(text, translated, settings)) return TranslationResult.Translated(translated);
            return await RetryUnchangedChunkAsync(text, protectedText, settings, token).ConfigureAwait(false);
        }

        private async Task<TranslationResult> TranslateMixedLanguageChunkAsync(string original, ProtectedTranslationText protectedText, TranslationSettings settings, CancellationToken token)
        {
            // Translating the intact paragraph gives the model enough context and is much faster than
            // translating dozens of tiny pieces. Protected foreign terms are then verified in order.
            var directPrompt = GetPromptDescriptor(settings, PromptMode.Direct);
            var directResponse = await SendTranslationRequestAsync(original, directPrompt, settings, token).ConfigureAwait(false);
            var direct = directResponse.Content;
            if (!directResponse.Truncated && IsCompleteTranslation(original, direct, settings) && protectedText.PreservesProtectedSegmentsInOrder(direct))
                return TranslationResult.Translated(direct);

            // Retry at most once. The fallback protects foreign terms with placeholders while still
            // translating the whole paragraph, so it keeps context without multiplying API calls.
            var protectedPrompt = GetPromptDescriptor(settings, PromptMode.Protected);
            var protectedResponse = await SendTranslationRequestAsync(protectedText.Text, protectedPrompt, settings, token, true).ConfigureAwait(false);
            var protectedResult = protectedResponse.Content;
            if (!protectedResponse.Truncated && !string.IsNullOrWhiteSpace(protectedResult))
            {
                try
                {
                    var restored = protectedText.Restore(protectedResult, NeedsCjkLoanwordSpacing(settings));
                    if (IsCompleteTranslation(original, restored, settings))
                        return TranslationResult.Translated(restored);
                }
                catch (ProtectedContentException) { }
            }

            // A single stubborn paragraph must not abort the whole task. Keep the
            // original text and let the task summary flag it for manual review.
            return TranslationResult.Unchanged(original);
        }

        private async Task<TranslationResult> RetryUnchangedChunkAsync(string original, ProtectedTranslationText protectedText, TranslationSettings settings, CancellationToken token)
        {
            var strictPrompt = GetPromptDescriptor(settings, PromptMode.Strict);
            var response = await SendTranslationRequestAsync(protectedText.Text, strictPrompt, settings, token, true).ConfigureAwait(false);
            var retried = response.Content;
            if (string.IsNullOrWhiteSpace(retried))
                throw new InvalidOperationException("翻译模型连续返回空内容，已停止写回。");
            string restored;
            try
            {
                restored = protectedText.Restore(retried, NeedsCjkLoanwordSpacing(settings));
            }
            catch (ProtectedContentException)
            {
                return TranslationResult.Unchanged(original);
            }
            // Even the expanded output budget can be exhausted by a very long chunk.
            // Never write back a truncated translation as if it were complete.
            return !response.Truncated && IsCompleteTranslation(original, restored, settings)
                ? TranslationResult.Translated(restored)
                : TranslationResult.Unchanged(original);
        }

        private async Task<ModelResponse> SendTranslationRequestAsync(string text, PromptDescriptor prompt, TranslationSettings settings, CancellationToken token, bool expandedOutputBudget = false)
        {
            var endpoint = settings.Provider == ProviderKind.Ollama
                ? settings.BaseUrl.TrimEnd('/') + "/api/chat"
                : settings.BaseUrl.TrimEnd('/') + "/chat/completions";
            if (settings.Provider == ProviderKind.Ollama)
            {
                var body = new
                {
                    model = settings.Model,
                    stream = false,
                    messages = Messages(prompt.Text, text),
                    options = new { temperature = 0.1, num_predict = EstimateMaxOutputTokens(text, settings.TargetLanguage, expandedOutputBudget) }
                };
                using (var response = await SendWithRetryAsync(() => CreateRequest(endpoint, body, settings.ApiKey), token).ConfigureAwait(false))
                {
                    var raw = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode) throw new InvalidOperationException($"翻译服务返回 {(int)response.StatusCode}: {raw}");
                    return ReadOllama(_json.DeserializeObject(raw) as Dictionary<string, object>);
                }
            }

            var capabilityKey = settings.BaseUrl.TrimEnd('/');
            var useCacheKey = IsOfficialOpenAiEndpoint(settings.BaseUrl) &&
                (!PromptCacheKeySupport.TryGetValue(capabilityKey, out var supported) || supported);
            for (var attempt = 0; attempt < 2; attempt++)
            {
                var body = new Dictionary<string, object>
                {
                    ["model"] = settings.Model,
                    ["temperature"] = 0.1,
                    ["max_tokens"] = EstimateMaxOutputTokens(text, settings.TargetLanguage, expandedOutputBudget),
                    ["messages"] = Messages(prompt.Text, text)
                };
                if (IsOfficialDeepSeekEndpoint(settings.BaseUrl))
                    body["thinking"] = new Dictionary<string, object> { ["type"] = "disabled" };
                if (useCacheKey) body["prompt_cache_key"] = prompt.CacheKey;
                using (var response = await SendWithRetryAsync(() => CreateRequest(endpoint, body, settings.ApiKey), token).ConfigureAwait(false))
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

        // Retries transient failures (network errors, timeouts, 429 and 5xx) with
        // exponential backoff. User cancellation is never retried. A fresh
        // HttpRequestMessage is created per attempt because a sent request
        // cannot be re-sent.
        private async Task<HttpResponseMessage> SendWithRetryAsync(Func<HttpRequestMessage> requestFactory, CancellationToken token)
        {
            var backoff = new[] { TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(3) };
            for (var attempt = 0; ; attempt++)
            {
                token.ThrowIfCancellationRequested();
                var lastAttempt = attempt >= backoff.Length;
                HttpResponseMessage? response = null;
                try
                {
                    using (var request = requestFactory())
                        response = await _http.SendAsync(request, token).ConfigureAwait(false);
                }
                catch (HttpRequestException) when (!lastAttempt)
                {
                    await Task.Delay(backoff[attempt], token).ConfigureAwait(false);
                    continue;
                }
                catch (TaskCanceledException) when (!token.IsCancellationRequested && !lastAttempt)
                {
                    // HttpClient timeout, not user cancellation.
                    await Task.Delay(backoff[attempt], token).ConfigureAwait(false);
                    continue;
                }

                if (response != null && IsTransientStatusCode(response.StatusCode) && !lastAttempt)
                {
                    var wait = GetRetryDelay(response, backoff[attempt]);
                    response.Dispose();
                    await Task.Delay(wait, token).ConfigureAwait(false);
                    continue;
                }
                if (response == null)
                    throw new InvalidOperationException("翻译服务无响应。");
                return response;
            }
        }

        // Note: HttpStatusCode has no TooManyRequests member on .NET Framework 4.8
        // (CS0117); compare the numeric code instead.
        internal static bool IsTransientStatusCode(HttpStatusCode status)
        {
            var code = (int)status;
            return code == 429 ||
                status == HttpStatusCode.InternalServerError ||
                status == HttpStatusCode.BadGateway ||
                status == HttpStatusCode.ServiceUnavailable ||
                status == HttpStatusCode.GatewayTimeout;
        }

        private static TimeSpan GetRetryDelay(HttpResponseMessage response, TimeSpan fallback)
        {
            // Honor Retry-After when the server tells us how long to wait.
            if (response.Headers.RetryAfter != null)
            {
                if (response.Headers.RetryAfter.Delta is TimeSpan delta && delta > TimeSpan.Zero)
                    return delta > TimeSpan.FromSeconds(30) ? TimeSpan.FromSeconds(30) : delta;
                if (response.Headers.RetryAfter.Date is DateTimeOffset date)
                {
                    var wait = date - DateTimeOffset.UtcNow;
                    if (wait > TimeSpan.Zero)
                        return wait > TimeSpan.FromSeconds(30) ? TimeSpan.FromSeconds(30) : wait;
                }
            }
            return fallback;
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

        private static bool IsOfficialOpenAiEndpoint(string baseUrl)
        {
            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var uri)) return false;
            return uri.Host.Equals("api.openai.com", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsOfficialDeepSeekEndpoint(string baseUrl)
        {
            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var uri)) return false;
            return uri.Host.Equals("api.deepseek.com", StringComparison.OrdinalIgnoreCase);
        }

        private static int EstimateMaxOutputTokens(string text, string targetLanguage, bool expanded)
        {
            var length = string.IsNullOrEmpty(text) ? 0 : text.Length;
            if (IsCjkTargetLanguage(targetLanguage))
                return expanded
                    ? Math.Max(512, Math.Min(8192, (int)Math.Ceiling(length * 2.25) + 256))
                    : Math.Max(256, Math.Min(4096, (int)Math.Ceiling(length * 1.25) + 128));
            return expanded
                ? Math.Max(512, Math.Min(12288, length * 3 + 256))
                : Math.Max(256, Math.Min(8192, length * 2 + 128));
        }

        private static bool IsCjkTargetLanguage(string language) =>
            language == "简体中文" || language == "繁體中文" || language == "日语" || language == "韩语";

        private static bool Equivalent(string source, string translated) =>
            string.Equals(NormalizeForComparison(source), NormalizeForComparison(translated), StringComparison.Ordinal);

        private static bool IsCompleteTranslation(string source, string translated, TranslationSettings settings)
        {
            if (string.IsNullOrWhiteSpace(translated) || Equivalent(source, translated)) return false;
            var maximumLength = IsCjkTargetLanguage(settings.TargetLanguage)
                ? Math.Max(800, (int)Math.Ceiling(source.Length * 1.8))
                : Math.Max(1200, source.Length * 3);
            if (translated.Length > maximumLength || HasSuspiciousRepetition(translated)) return false;
            if (!HasRequiredTargetLanguageOutput(source, translated, settings)) return false;
            if (!ShouldCheckSourceResidual(settings.SourceLanguage, settings.TargetLanguage)) return true;
            var sourceLetters = SourceLanguageProtector.CountDistinctiveSourceLetters(source, settings.SourceLanguage);
            if (sourceLetters < 10) return true;
            var remainingLetters = SourceLanguageProtector.CountDistinctiveSourceLetters(translated, settings.SourceLanguage);
            var maximumRemaining = Math.Max(3, (int)Math.Ceiling(sourceLetters * 0.02));
            return remainingLetters <= maximumRemaining;
        }

        private static bool ShouldCheckSourceResidual(string sourceLanguage, string targetLanguage)
        {
            if (SourceLanguageProtector.IsAutomatic(sourceLanguage)) return false;
            switch (sourceLanguage)
            {
                case "英语":
                case "法语":
                case "德语":
                case "西班牙语":
                case "葡萄牙语":
                case "意大利语":
                    // These languages share the Latin writing system, and product names or
                    // abbreviations can legitimately remain in a correct translation.
                    return false;
                case "简体中文":
                case "繁體中文":
                    // Japanese and both Chinese variants legitimately retain Han characters.
                    return targetLanguage != "简体中文" && targetLanguage != "繁體中文" && targetLanguage != "日语";
                default:
                    return true;
            }
        }

        private static bool HasRequiredTargetLanguageOutput(string source, string translated, TranslationSettings settings)
        {
            var sourceLetterCount = source.Count(char.IsLetter);
            if (sourceLetterCount == 0) return true;

            var sourceTargetLetters = SourceLanguageProtector.CountDistinctiveSourceLetters(source, settings.TargetLanguage);
            var lettersNeedingConversion = Math.Max(0, sourceLetterCount - sourceTargetLetters);
            if (lettersNeedingConversion == 0) return true;

            // A result that contains no meaningful amount of the requested writing system is not a
            // translation. This catches, for example, a Chinese paragraph returned almost unchanged
            // when the target is Korean, even if punctuation or whitespace changed.
            var translatedTargetLetters = SourceLanguageProtector.CountDistinctiveSourceLetters(translated, settings.TargetLanguage);
            var requiredTargetLetters = lettersNeedingConversion < 8
                ? 1
                : Math.Max(2, Math.Min(12, (int)Math.Ceiling(lettersNeedingConversion * 0.05)));
            if (translatedTargetLetters < requiredTargetLetters) return false;

            if (!SourceLanguageProtector.IsAutomatic(settings.SourceLanguage)) return true;

            // Automatic detection has no explicit source script for the regular residual check.
            // Reject a result that still contains most of a clearly detectable non-target script.
            var detectableScripts = new[] { "简体中文", "日语", "韩语", "俄语", "阿拉伯语" };
            foreach (var script in detectableScripts)
            {
                if (script == settings.TargetLanguage ||
                    (settings.TargetLanguage == "繁體中文" && script == "简体中文") ||
                    (settings.TargetLanguage == "日语" && script == "简体中文")) continue;
                var originalCount = SourceLanguageProtector.CountDistinctiveSourceLetters(source, script);
                if (originalCount < 10) continue;
                var remainingCount = SourceLanguageProtector.CountDistinctiveSourceLetters(translated, script);
                var maximumRemaining = Math.Max(3, (int)Math.Ceiling(originalCount * 0.05));
                if (remainingCount > maximumRemaining) return false;
            }
            return true;
        }

        private static bool HasSuspiciousRepetition(string text)
        {
            const int sampleLength = 100;
            if (string.IsNullOrEmpty(text) || text.Length < sampleLength * 3) return false;
            for (var start = 0; start + sampleLength <= text.Length; start += sampleLength)
            {
                var sample = text.Substring(start, sampleLength);
                var second = text.IndexOf(sample, start + sampleLength, StringComparison.Ordinal);
                if (second < 0) continue;
                if (text.IndexOf(sample, second + sampleLength, StringComparison.Ordinal) >= 0) return true;
            }
            return false;
        }

        private static string NormalizeForComparison(string text) =>
            (text ?? string.Empty).Trim().Replace("\r\n", "\n").Replace('\r', '\n');

        private static object[] Messages(string prompt, string text) => new object[] {
            new { role = "system", content = prompt }, new { role = "user", content = text }
        };

        private static PromptDescriptor GetPromptDescriptor(TranslationSettings settings, PromptMode mode)
        {
            var descriptorKey = string.Join("\u001F", new[]
            {
                mode.ToString(),
                settings.SourceLanguage ?? string.Empty,
                settings.TargetLanguage ?? string.Empty,
                settings.CustomInstructions ?? string.Empty,
                settings.Glossary ?? string.Empty
            });
            return PromptDescriptors.GetOrAdd(descriptorKey, _ =>
            {
                var prompt = BuildPrompt(settings, mode == PromptMode.Protected);
                if (mode == PromptMode.Strict)
                    prompt += "\n重要：上一次结果没有完成翻译。本次必须将输入中的源语言内容实际转换为目标语言；不得原样返回输入，不得返回空内容。";
                return new PromptDescriptor(prompt, "ot-" + TranslationSessionCache.Hash(prompt));
            });
        }

        private static string BuildPrompt(TranslationSettings s, bool protectedPlaceholders)
        {
            var glossary = string.IsNullOrWhiteSpace(s.Glossary) ? "" : "\n必须遵循以下术语表（每行 source=target）：\n" + s.Glossary;
            var spacing = CjkLoanwordSpacingInstruction(s);
            var targetOutput = TargetLanguageOutputInstruction(s.TargetLanguage);
            if (SourceLanguageProtector.IsAutomatic(s.SourceLanguage))
            {
                var automaticProtection = protectedPlaceholders
                    ? "形如 ⟦OT_KEEP_0001⟧ 的内容代表受保护的换行或版式标记，必须原样、按原顺序和原位置保留。"
                    : string.Empty;
                return $"你是专业翻译。自动识别输入的源语言，将输入完整翻译成{s.TargetLanguage}。{automaticProtection}{targetOutput}只输出译文，不解释，不添加标题；保留换行、编号和占位符。{spacing}{s.CustomInstructions}{glossary}";
            }

            var protection = protectedPlaceholders
                ? "形如 ⟦OT_KEEP_0001⟧ 的内容是已保护原文，必须原样、按原顺序和原位置保留。"
                : "已经是其他语言的单词、缩写、型号、数字、单位、网址和代码保持原样。";
            return $"你是专业翻译。将输入中的{s.SourceLanguage}内容翻译成{s.TargetLanguage}；{protection}" +
                $"{targetOutput}只输出完整译文，不解释，不附带原文或标题；保留换行和编号。{spacing}{s.CustomInstructions}{glossary}";
        }

        private static string TargetLanguageOutputInstruction(string targetLanguage)
        {
            switch (targetLanguage)
            {
                case "韩语":
                    return "译文正文必须使用韩文（Hangul/한글）书写，不能继续使用中文充当韩语译文。";
                case "日语":
                    return "译文正文必须使用自然的日语书写，并包含必要的平假名或片假名，不能继续使用中文充当日语译文。";
                case "简体中文":
                    return "译文正文必须使用简体中文书写。";
                case "繁體中文":
                    return "译文正文必须使用繁體中文书写。";
                default:
                    return string.Empty;
            }
        }

        private static string CjkLoanwordSpacingInstruction(TranslationSettings settings)
        {
            return NeedsCjkLoanwordSpacing(settings)
                ? "翻译韩语或日语时，如果术语表或语义要求输出英文词，英文词与相邻的拉丁字母、英文缩写或数字之间必须保留一个空格，不能直接粘连。"
                : string.Empty;
        }

        private static bool NeedsCjkLoanwordSpacing(TranslationSettings settings)
        {
            return settings.SourceLanguage == "韩语" || settings.SourceLanguage == "日语";
        }

        private static void AppendWithBoundary(StringBuilder output, string value, bool enabled)
        {
            if (string.IsNullOrEmpty(value)) return;
            if (enabled && output.Length > 0 && ProtectedTranslationText.NeedsBoundarySpace(output[output.Length - 1], value[0]))
                output.Append(' ');
            output.Append(value);
        }

        private static ModelResponse ReadOllama(Dictionary<string, object>? root)
        {
            if (root != null && root.TryGetValue("message", out var m) && m is Dictionary<string, object> msg && msg.TryGetValue("content", out var c))
            {
                var doneReason = root.TryGetValue("done_reason", out var reason) ? Convert.ToString(reason) : string.Empty;
                return new ModelResponse(Convert.ToString(c) ?? "", IsLengthTruncation(doneReason));
            }
            throw new InvalidOperationException("无法解析 Ollama 响应。");
        }

        private static ModelResponse ReadOpenAi(Dictionary<string, object>? root)
        {
            if (root != null && root.TryGetValue("choices", out var value) && value is object[] choices &&
                choices.FirstOrDefault() is Dictionary<string, object> choice &&
                choice.TryGetValue("message", out var messageValue) && messageValue is Dictionary<string, object> msg)
            {
                var finishReason = choice.TryGetValue("finish_reason", out var reason) ? Convert.ToString(reason) : string.Empty;
                var content = msg.TryGetValue("content", out var c) ? Convert.ToString(c) ?? string.Empty : string.Empty;
                return new ModelResponse(content, IsLengthTruncation(finishReason));
            }
            throw new InvalidOperationException("无法解析 OpenAI-compatible 响应。");
        }

        private static bool IsLengthTruncation(string? reason) =>
            string.Equals(reason, "length", StringComparison.OrdinalIgnoreCase);

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

        private sealed class ModelResponse
        {
            public ModelResponse(string content, bool truncated) { Content = content; Truncated = truncated; }
            public string Content { get; }
            public bool Truncated { get; }
        }

        private enum PromptMode { Direct, Strict, Protected }

        public void Dispose() => _http.Dispose();
    }

    public sealed class ImageTranslationRegion
    {
        public float X1 { get; set; } public float Y1 { get; set; } public float X2 { get; set; } public float Y2 { get; set; }
        public string Source { get; set; } = string.Empty; public string Translation { get; set; } = string.Empty;
    }
}
