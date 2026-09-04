using System;
using System.Web.Script.Serialization;

namespace OfficeTranslate.Core
{
    public enum ProviderKind { OpenAiCompatible, Ollama }

    public sealed class TranslationSettings
    {
        public ProviderKind Provider { get; set; } = ProviderKind.Ollama;
        public string BaseUrl { get; set; } = "http://localhost:11434";
        public string Model { get; set; } = "qwen2.5:7b";
        public string ApiKey { get; set; } = string.Empty;
        public string OllamaBaseUrl { get; set; } = "http://localhost:11434";
        public string OllamaApiKey { get; set; } = string.Empty;
        public string OllamaModel { get; set; } = "qwen2.5:7b";
        public string OpenAiBaseUrl { get; set; } = "https://api.openai.com/v1";
        public string OpenAiApiKey { get; set; } = string.Empty;
        public string OpenAiModel { get; set; } = string.Empty;
        public string ImageModel { get; set; } = string.Empty;
        [ScriptIgnore]
        public string SourceLanguage { get; set; } = "自动检测";
        [ScriptIgnore]
        public string TargetLanguage { get; set; } = "简体中文";
        public string UiLanguage { get; set; } = "zh-CN";
        [ScriptIgnore]
        public bool BilingualMode { get; set; }
        [ScriptIgnore]
        public bool ImageOcrEnabled { get; set; }
        public string TranslationStyle { get; set; } = "ProfessionalReport";
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

    public static class TranslationStyleCatalog
    {
        public static readonly string[] Ids =
        {
            "ProfessionalReport", "AcademicPaper", "Technology", "News", "FreeTranslation", "Custom"
        };

        public static string GetPrompt(string id)
        {
            switch (id)
            {
                case "AcademicPaper":
                    return "采用严谨、客观的学术论文风格翻译。准确保留专业概念、论证关系、引文、公式、变量、单位和章节结构；术语前后一致，避免口语化、夸张表达和无依据的增译。";
                case "Technology":
                    return "采用准确、清晰的科技类文本风格翻译。正确处理技术术语、产品名称、代码、命令、参数、缩写、数字和单位；表达简洁明确，保留原文的逻辑层级与操作步骤。";
                case "News":
                    return "采用客观、简洁、自然的新闻报道风格翻译。准确呈现人物、机构、地点、时间、数字和事件关系；使用符合目标语言新闻习惯的表达，不渲染、不评论、不改变事实立场。";
                case "FreeTranslation":
                    return "以传达原文含义和语气为优先进行意译。允许调整句式、语序和惯用表达，使译文自然流畅、符合目标语言习惯；不得遗漏关键信息、改变事实或擅自扩写。";
                case "Custom":
                    return string.Empty;
                default:
                    return "采用正式、专业的报告风格翻译。准确传达事实、结论、数据和逻辑关系；语言清晰凝练、结构严谨、术语一致，避免口语化、歧义、夸张和无依据的增译。";
            }
        }
    }
}
