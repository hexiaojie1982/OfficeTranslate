using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
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

        private async Task<string> TranslateChunkAsync(string text, TranslationSettings settings, CancellationToken token)
        {
            var endpoint = settings.Provider == ProviderKind.Ollama
                ? settings.BaseUrl.TrimEnd('/') + "/api/chat"
                : settings.BaseUrl.TrimEnd('/') + "/chat/completions";
            var prompt = BuildPrompt(settings);
            var protectedText = SourceLanguageProtector.Protect(text, settings.SourceLanguage);
            object body = settings.Provider == ProviderKind.Ollama
                ? new { model = settings.Model, stream = false, messages = Messages(prompt, protectedText.Text) }
                : new { model = settings.Model, temperature = 0.2, messages = Messages(prompt, protectedText.Text) };
            using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
                request.Content = new StringContent(_json.Serialize(body), Encoding.UTF8, "application/json");
                if (!string.IsNullOrWhiteSpace(settings.ApiKey)) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                using (var response = await _http.SendAsync(request, token).ConfigureAwait(false))
                {
                    var raw = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode) throw new InvalidOperationException($"翻译服务返回 {(int)response.StatusCode}: {raw}");
                    var root = _json.DeserializeObject(raw) as Dictionary<string, object>;
                    var translated = settings.Provider == ProviderKind.Ollama ? ReadOllama(root) : ReadOpenAi(root);
                    return protectedText.Restore(translated);
                }
            }
        }

        private static object[] Messages(string prompt, string text) => new object[] {
            new { role = "system", content = prompt }, new { role = "user", content = text }
        };

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

    public sealed class ImageTranslationRegion
    {
        public float X1 { get; set; } public float Y1 { get; set; } public float X2 { get; set; } public float Y2 { get; set; }
        public string Source { get; set; } = string.Empty; public string Translation { get; set; } = string.Empty;
    }
}
