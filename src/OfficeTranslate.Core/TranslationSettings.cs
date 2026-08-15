using System;

namespace OfficeTranslate.Core
{
    public enum ProviderKind { OpenAiCompatible, Ollama }

    public sealed class TranslationSettings
    {
        public ProviderKind Provider { get; set; } = ProviderKind.Ollama;
        public string BaseUrl { get; set; } = "http://localhost:11434";
        public string Model { get; set; } = "qwen2.5:7b";
        public string ApiKey { get; set; } = string.Empty;
        public string SourceLanguage { get; set; } = "自动检测";
        public string TargetLanguage { get; set; } = "简体中文";
        public string UiLanguage { get; set; } = "zh-CN";
        public bool BilingualMode { get; set; }
        public string CustomInstructions { get; set; } = string.Empty;
        public string Glossary { get; set; } = string.Empty;
        public int MaxCharactersPerChunk { get; set; } = 5000;
        public int TimeoutSeconds { get; set; } = 120;

        public void Validate()
        {
            ValidateEndpoint();
            if (string.IsNullOrWhiteSpace(Model)) throw new InvalidOperationException("模型名称不能为空。");
            if (string.IsNullOrWhiteSpace(TargetLanguage)) throw new InvalidOperationException("目标语言不能为空。");
            if (MaxCharactersPerChunk < 500 || MaxCharactersPerChunk > 30000) throw new InvalidOperationException("分块字符数必须在 500–30000 之间。");
        }

        public void ValidateEndpoint()
        {
            if (!Uri.TryCreate(BaseUrl, UriKind.Absolute, out _)) throw new InvalidOperationException("Base URL 无效。");
        }
    }
}
