using System.Collections.Generic;

namespace OfficeTranslate.WordAddIn
{
    internal static class UiText
    {
        public static readonly string[] LanguageCodes = { "zh-CN", "zh-HK", "ko", "ja", "en" };
        public static readonly string[] LanguageNames = { "简体中文", "繁體中文", "한국어", "日本語", "English" };

        private static readonly Dictionary<string, string[]> Texts = new Dictionary<string, string[]>
        {
            ["LanguageGroup"] = new[] { "语言", "語言", "언어", "言語", "Language" },
            ["TranslateGroup"] = new[] { "翻译", "翻譯", "번역", "翻訳", "Translate" },
            ["ToolsGroup"] = new[] { "工具", "工具", "도구", "ツール", "Tools" },
            ["Source"] = new[] { "源语言", "來源語言", "원문 언어", "原文言語", "Source" },
            ["Target"] = new[] { "目标语言", "目標語言", "대상 언어", "翻訳言語", "Target" },
            ["Swap"] = new[] { "交换", "交換", "전환", "入替", "Swap" },
            ["Selection"] = new[] { "选区", "選取", "선택", "選択", "Selection" },
            ["Document"] = new[] { "全文", "全文", "전체", "全文", "Document" },
            ["Bilingual"] = new[] { "双语", "雙語", "이중", "対訳", "Bilingual" },
            ["Cancel"] = new[] { "取消", "取消", "취소", "取消", "Cancel" },
            ["TranslationProgress"] = new[] { "翻译进度", "翻譯進度", "번역 진행률", "翻訳の進行状況", "Translation progress" },
            ["PreparingTranslation"] = new[] { "正在准备翻译…", "正在準備翻譯…", "번역을 준비하는 중…", "翻訳を準備しています…", "Preparing translation…" },
            ["CancelTranslation"] = new[] { "取消翻译", "取消翻譯", "번역 취소", "翻訳をキャンセル", "Cancel translation" },
            ["CancellingTranslation"] = new[] { "正在取消，已完成的内容将保留…", "正在取消，已完成的內容將保留…", "취소하는 중입니다. 완료된 내용은 유지됩니다…", "キャンセルしています。完了した内容は保持されます…", "Cancelling… Completed content will be kept." },
            ["Settings"] = new[] { "设置", "設定", "설정", "設定", "Settings" },
            ["SettingsTitle"] = new[] { "OfficeTranslate 设置", "OfficeTranslate 設定", "OfficeTranslate 설정", "OfficeTranslate 設定", "OfficeTranslate Settings" },
            ["Heading"] = new[] { "翻译服务设置", "翻譯服務設定", "번역 서비스 설정", "翻訳サービス設定", "Translation service" },
            ["UiLanguage"] = new[] { "界面语言", "介面語言", "UI 언어", "表示言語", "UI language" },
            ["Provider"] = new[] { "服务类型", "服務類型", "서비스 유형", "サービス種類", "Provider" },
            ["Model"] = new[] { "模型", "模型", "모델", "モデル", "Model" },
            ["Chunk"] = new[] { "分块字符数", "分段字元數", "청크 문자 수", "分割文字数", "Chunk size" },
            ["TranslationStyle"] = new[] { "翻译风格", "翻譯風格", "번역 스타일", "翻訳スタイル", "Translation style" },
            ["Instructions"] = new[] { "额外要求", "額外要求", "추가 지침", "追加指示", "Instructions" },
            ["Glossary"] = new[] { "术语表", "術語表", "용어집", "用語集", "Glossary" },
            ["Refresh"] = new[] { "↻  更新模型", "↻  更新模型", "↻  모델 새로 고침", "↻  モデル更新", "↻  Refresh models" },
            ["Save"] = new[] { "✓  保存设置", "✓  儲存設定", "✓  설정 저장", "✓  設定を保存", "✓  Save" },
            ["CancelButton"] = new[] { "×  取消", "×  取消", "×  취소", "×  キャンセル", "×  Cancel" },
            ["InitialStatus"] = new[] { "先填写 Base URL 和 API Key，再点击“更新模型”。", "請先填寫 Base URL 和 API Key，再按「更新模型」。", "Base URL과 API Key를 입력한 후 모델을 새로 고치세요.", "Base URL と API Key を入力してモデルを更新してください。", "Enter Base URL and API Key, then refresh models." },
            ["ProviderChanged"] = new[] { "服务类型已更改，请更新模型列表。", "服務類型已變更，請更新模型清單。", "서비스 유형이 변경되었습니다. 모델 목록을 새로 고치세요.", "サービス種類が変更されました。モデル一覧を更新してください。", "Provider changed. Refresh the model list." },
            ["LoadingModels"] = new[] { "正在连接服务并获取模型…", "正在連線服務並取得模型…", "서비스에 연결하여 모델을 가져오는 중…", "サービスに接続してモデルを取得中…", "Connecting and fetching models…" },
            ["ModelsLoaded"] = new[] { "✓ 已获取 {0} 个模型", "✓ 已取得 {0} 個模型", "✓ 모델 {0}개를 가져왔습니다", "✓ {0} 個のモデルを取得しました", "✓ Loaded {0} models" },
            ["NoModels"] = new[] { "服务没有返回任何可用模型。", "服務未傳回任何可用模型。", "사용 가능한 모델이 없습니다.", "利用可能なモデルがありません。", "The service returned no available models." },
            ["SaveError"] = new[] { "无法保存设置", "無法儲存設定", "설정을 저장할 수 없음", "設定を保存できません", "Unable to save settings" },
            ["About"] = new[] { "ⓘ  关于", "ⓘ  關於", "ⓘ  정보", "ⓘ  情報", "ⓘ  About" },
            ["AboutTitle"] = new[] { "关于 OfficeTranslate", "關於 OfficeTranslate", "OfficeTranslate 정보", "OfficeTranslate について", "About OfficeTranslate" },
            ["Changelog"] = new[] { "更新日志", "更新記錄", "변경 기록", "更新履歴", "Changelog" },
            ["Guide"] = new[] { "使用说明", "使用說明", "사용 방법", "使用方法", "User guide" },
            ["Close"] = new[] { "关闭", "關閉", "닫기", "閉じる", "Close" },
            ["SourceTip"] = new[] { "当前源语言：{0}", "目前來源語言：{0}", "현재 원문 언어: {0}", "現在の原文言語：{0}", "Current source: {0}" },
            ["TargetTip"] = new[] { "当前目标语言：{0}", "目前目標語言：{0}", "현재 대상 언어: {0}", "現在の翻訳言語：{0}", "Current target: {0}" }
        };

        public static string Get(string language, string key)
        {
            var index = System.Array.IndexOf(LanguageCodes, language);
            if (index < 0) index = 0;
            return Texts.TryGetValue(key, out var values) ? values[index] : key;
        }

        public static string Language(string uiLanguage, string internalName)
        {
            var names = new Dictionary<string, string[]>
            {
                ["自动检测"] = new[] { "自动检测", "自動偵測", "자동 감지", "自動検出", "Auto detect" },
                ["简体中文"] = new[] { "简体中文", "簡體中文", "중국어 간체", "簡体字中国語", "Chinese (Simplified)" },
                ["繁體中文"] = new[] { "繁體中文", "繁體中文", "중국어 번체", "繁体字中国語", "Chinese (Traditional)" },
                ["英语"] = new[] { "英语", "英語", "영어", "英語", "English" }, ["日语"] = new[] { "日语", "日語", "일본어", "日本語", "Japanese" },
                ["韩语"] = new[] { "韩语", "韓語", "한국어", "韓国語", "Korean" }, ["法语"] = new[] { "法语", "法語", "프랑스어", "フランス語", "French" },
                ["德语"] = new[] { "德语", "德語", "독일어", "ドイツ語", "German" }, ["西班牙语"] = new[] { "西班牙语", "西班牙語", "스페인어", "スペイン語", "Spanish" },
                ["俄语"] = new[] { "俄语", "俄語", "러시아어", "ロシア語", "Russian" }, ["葡萄牙语"] = new[] { "葡萄牙语", "葡萄牙語", "포르투갈어", "ポルトガル語", "Portuguese" },
                ["意大利语"] = new[] { "意大利语", "義大利語", "이탈리아어", "イタリア語", "Italian" }, ["阿拉伯语"] = new[] { "阿拉伯语", "阿拉伯語", "아랍어", "アラビア語", "Arabic" }
            };
            var index = System.Array.IndexOf(LanguageCodes, uiLanguage); if (index < 0) index = 0;
            return names.TryGetValue(internalName, out var values) ? values[index] : internalName;
        }

        public static string TranslationStyle(string uiLanguage, string id)
        {
            var names = new Dictionary<string, string[]>
            {
                ["ProfessionalReport"] = new[] { "专业报告翻译", "專業報告翻譯", "전문 보고서 번역", "専門レポート翻訳", "Professional report" },
                ["AcademicPaper"] = new[] { "学术论文翻译", "學術論文翻譯", "학술 논문 번역", "学術論文翻訳", "Academic paper" },
                ["Technology"] = new[] { "科技类翻译", "科技類翻譯", "기술 번역", "科学技術翻訳", "Technology" },
                ["News"] = new[] { "新闻翻译", "新聞翻譯", "뉴스 번역", "ニュース翻訳", "News" },
                ["FreeTranslation"] = new[] { "意译", "意譯", "의역", "意訳", "Free translation" },
                ["Custom"] = new[] { "自定义风格", "自訂風格", "사용자 지정", "カスタム", "Custom" }
            };
            var index = System.Array.IndexOf(LanguageCodes, uiLanguage); if (index < 0) index = 0;
            return names.TryGetValue(id, out var values) ? values[index] : id;
        }
    }
}
