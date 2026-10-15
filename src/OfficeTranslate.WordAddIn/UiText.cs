using System.Collections.Generic;

namespace OfficeTranslate.WordAddIn
{
    internal static class UiText
    {
        public static readonly string[] LanguageCodes = { "zh-CN", "zh-HK", "ko", "ja", "en" };
        public static readonly string[] LanguageNames = { "简体中文", "繁體中文", "한국어", "日本語", "English" };

        private static readonly Dictionary<string, string[]> Texts = new Dictionary<string, string[]>
        {
            ["Chat"] = new[] { "文档内 Chat", "文档内 Chat", "Document Chat", "Document Chat", "Document Chat" },
            ["ChatSummary"] = new[] { "总结", "总结", "Summarize", "Summarize", "Summarize" },
            ["ChatRewrite"] = new[] { "改写", "改写", "Rewrite", "Rewrite", "Rewrite" },
            ["ChatClear"] = new[] { "清空", "清空", "Clear", "Clear", "Clear" },
            ["ChatRefresh"] = new[] { "刷新上下文", "刷新上下文", "Refresh context", "Refresh context", "Refresh context" },
            ["ChatSettings"] = new[] { "Chat 设置", "Chat 设置", "Chat settings", "Chat settings", "Chat settings" },
            ["ChatSend"] = new[] { "发送", "发送", "Send", "Send", "Send" },
            ["ChatEmpty"] = new[] { "选中内容或打开文档，提问、总结或改写。Ctrl+Enter 发送。会话只在内存保留最近10轮。", "选中内容或打开文档，提问、总结或改写。Ctrl+Enter 发送。会话只在内存保留最近10轮。", "Select content or open a document to ask, summarize or rewrite. Ctrl+Enter to send. Only the last 10 turns are kept in memory.", "Select content or open a document to ask, summarize or rewrite. Ctrl+Enter to send. Only the last 10 turns are kept in memory.", "Select content or open a document to ask, summarize or rewrite. Ctrl+Enter to send. Only the last 10 turns are kept in memory." },
            ["ChatContext"] = new[] { "上下文：{0} · 约{1:N0}字", "上下文：{0} · 约{1:N0}字", "Context: {0} · approx. {1:N0} characters", "Context: {0} · approx. {1:N0} characters", "Context: {0} · approx. {1:N0} characters" },
            ["ChatSelected"] = new[] { "选中内容", "选中内容", "Selection", "Selection", "Selection" },
            ["ChatFull"] = new[] { "全文", "全文", "Entire document", "Entire document", "Entire document" },
            ["ChatNoDocument"] = new[] { "上下文不可用，请打开文档并刷新。", "上下文不可用，请打开文档并刷新。", "Context unavailable. Open a document and refresh.", "Context unavailable. Open a document and refresh.", "Context unavailable. Open a document and refresh." },
            ["ChatCleared"] = new[] { "会话已清空。", "会话已清空。", "Conversation cleared.", "Conversation cleared.", "Conversation cleared." },
            ["ChatSummaryRequest"] = new[] { "请总结当前文档上下文。", "请总结当前文档上下文。", "Please summarize the current document context.", "Please summarize the current document context.", "Please summarize the current document context." },
            ["ChatRewriteRequest"] = new[] { "请改写当前文档上下文，保留原意和事实，只返回改写后的正文。", "请改写当前文档上下文，保留原意和事实，只返回改写后的正文。", "Please rewrite the current document context, preserving meaning and facts. Return only the rewritten text.", "Please rewrite the current document context, preserving meaning and facts. Return only the rewritten text.", "Please rewrite the current document context, preserving meaning and facts. Return only the rewritten text." },
            ["ChatNeedInput"] = new[] { "请输入问题。", "请输入问题。", "Enter a question.", "Enter a question.", "Enter a question." },
            ["ChatNoText"] = new[] { "当前上下文没有可读取的文字。", "当前上下文没有可读取的文字。", "The current context contains no readable text.", "The current context contains no readable text.", "The current context contains no readable text." },
            ["ChatYou"] = new[] { "你", "你", "You", "You", "You" },
            ["ChatAssistant"] = new[] { "助手", "助手", "Assistant", "Assistant", "Assistant" },
            ["ChatLoading"] = new[] { "正在生成回答…可点击取消。", "正在生成回答…可点击取消。", "Generating a response… You can cancel.", "Generating a response… You can cancel.", "Generating a response… You can cancel." },
            ["ChatChunkProgress"] = new[] { "长文档摘要：正在处理第{0}/{1}块…", "长文档摘要：正在处理第{0}/{1}块…", "Long document: summarizing chunk {0}/{1}…", "Long document: summarizing chunk {0}/{1}…", "Long document: summarizing chunk {0}/{1}…" },
            ["ChatRecovering"] = new[] { "流式输出中断，正在尝试普通请求…", "流式輸出中斷，正在嘗試普通請求…", "Stream interrupted. Retrying without streaming…", "Stream interrupted. Retrying without streaming…", "Stream interrupted. Retrying without streaming…" },
            ["ChatFallback"] = new[] { "流式输出不可用或已中断，已降级为普通请求并完成回答。", "流式输出不可用或已中断，已降级为普通请求并完成回答。", "Streaming was unavailable or interrupted. Completed using a non-streaming request.", "Streaming was unavailable or interrupted. Completed using a non-streaming request.", "Streaming was unavailable or interrupted. Completed using a non-streaming request." },
            ["ChatReduced"] = new[] { "长文档已分块摘要；本次回答或改写基于摘要。", "長文件已分塊摘要；本次回答或改寫基於摘要。", "Long document summarized in chunks; this answer or rewrite is based on summaries.", "Long document summarized in chunks; this answer or rewrite is based on summaries.", "Long document summarized in chunks; this answer or rewrite is based on summaries." },
            ["ChatDone"] = new[] { "已完成。写回需点击下方按钮并确认。", "已完成。写回需点击下方按钮并确认。", "Completed. To write back, choose an action and confirm.", "Completed. To write back, choose an action and confirm.", "Completed. To write back, choose an action and confirm." },
            ["ChatCancelled"] = new[] { "已取消；未完成回答不会加入会话或写回。", "已取消；未完成回答不会加入会话或写回。", "Cancelled. Incomplete responses are not added to history or written back.", "Cancelled. Incomplete responses are not added to history or written back.", "Cancelled. Incomplete responses are not added to history or written back." },
            ["ChatError"] = new[] { "请求失败。请检查模型、服务、网络和 Chat token 预算后重试；未写回文档。", "请求失败。请检查模型、服务、网络和 Chat token 预算后重试；未写回文档。", "Request failed. Check the model, service, network and Chat token budgets, then retry. The document was not changed.", "Request failed. Check the model, service, network and Chat token budgets, then retry. The document was not changed.", "Request failed. Check the model, service, network and Chat token budgets, then retry. The document was not changed." },
            ["ChatReplace"] = new[] { "替换选中", "替换选中", "Replace selection", "Replace selection", "Replace selection" },
            ["ChatInsertBelow"] = new[] { "插入到下方", "插入到下方", "Insert below", "Insert below", "Insert below" },
            ["ChatInsertStart"] = new[] { "插入到文档开头", "插入到文档开头", "Insert at start", "Insert at start", "Insert at start" },
            ["ChatWriteConfirm"] = new[] { "将此结果写入发送时的原文档？替换操作会修改原文，请确认。", "将此结果写入发送时的原文档？替换操作会修改原文，请确认。", "Write this result to the original document? Replacing changes the original text. Please confirm.", "Write this result to the original document? Replacing changes the original text. Please confirm.", "Write this result to the original document? Replacing changes the original text. Please confirm." },
            ["ChatWritten"] = new[] { "已写入原文档。", "已写入原文档。", "Written to the original document.", "Written to the original document.", "Written to the original document." },
            ["ChatWriteError"] = new[] { "无法写回：原文档可能已关闭、目标已更改、受保护或超出宿主限制。请复制结果后手动处理。", "无法写回：原文档可能已关闭、目标已更改、受保护或超出宿主限制。请复制结果后手动处理。", "Cannot write back: the original document may be closed, changed, protected, or beyond host limits. Copy the result for manual use.", "Cannot write back: the original document may be closed, changed, protected, or beyond host limits. Copy the result for manual use.", "Cannot write back: the original document may be closed, changed, protected, or beyond host limits. Copy the result for manual use." },
            ["ChatReplaceUnavailable"] = new[] { "当前上下文无法整体替换；可插入结果。Excel 替换仅支持单个未合并单元格或单个文本框。", "当前上下文无法整体替换；可插入结果。Excel 替换仅支持单个未合并单元格或单个文本框。", "This context cannot be replaced as a whole; insert the result instead. Excel replacement supports a single unmerged cell or text box.", "This context cannot be replaced as a whole; insert the result instead. Excel replacement supports a single unmerged cell or text box.", "This context cannot be replaced as a whole; insert the result instead. Excel replacement supports a single unmerged cell or text box." },
            ["ChatQuestionPrompt"] = new[] { "问答系统提示词", "问答系统提示词", "Question system prompt", "Question system prompt", "Question system prompt" },
            ["ChatSummaryPrompt"] = new[] { "总结系统提示词", "总结系统提示词", "Summary system prompt", "Summary system prompt", "Summary system prompt" },
            ["ChatRewritePrompt"] = new[] { "改写系统提示词", "改写系统提示词", "Rewrite system prompt", "Rewrite system prompt", "Rewrite system prompt" },
            ["ChatBudget"] = new[] { "模型上下文 token 预算（按模型实际上限设置；保守按 UTF-8 字节估算）", "模型上下文 token 预算（按模型实际上限设置；保守按 UTF-8 字节估算）", "Model context token budget (use the model limit; estimated conservatively by UTF-8 bytes)", "Model context token budget (use the model limit; estimated conservatively by UTF-8 bytes)", "Model context token budget (use the model limit; estimated conservatively by UTF-8 bytes)" },
            ["ChatOutputBudget"] = new[] { "回答最大 token 数（不超过上下文预算一半）", "回答最大 token 数（不超过上下文预算一半）", "Maximum output tokens (at most half the context budget)", "Maximum output tokens (at most half the context budget)", "Maximum output tokens (at most half the context budget)" },
            ["ChatInvalidSettings"] = new[] { "请填写三个提示词；回答 token 数不得超过上下文预算的一半。", "请填写三个提示词；回答 token 数不得超过上下文预算的一半。", "Fill in all three prompts. Output tokens must not exceed half the context budget.", "Fill in all three prompts. Output tokens must not exceed half the context budget.", "Fill in all three prompts. Output tokens must not exceed half the context budget." },
            ["ChatSettingsError"] = new[] { "无法保存 Chat 设置，请检查配置文件权限。", "无法保存 Chat 设置，请检查配置文件权限。", "Cannot save Chat settings. Check configuration file permissions.", "Cannot save Chat settings. Check configuration file permissions.", "Cannot save Chat settings. Check configuration file permissions." },
            ["ChatPaneError"] = new[] { "无法打开 Chat 侧边栏。请打开文档，并确认已安装含 Chat 控件注册的 2.3.0 加载项。", "无法打开 Chat 侧边栏。请打开文档，并确认已安装含 Chat 控件注册的 2.3.0 加载项。", "Cannot open the Chat pane. Open a document and ensure the 2.3.0 add-in and Chat control registration are installed.", "Cannot open the Chat pane. Open a document and ensure the 2.3.0 add-in and Chat control registration are installed.", "Cannot open the Chat pane. Open a document and ensure the 2.3.0 add-in and Chat control registration are installed." },
            ["ChatChangelog"] = new[] { "OfficeTranslate 2.3.0.0\r\n• 三宿主新增右侧文档内 Chat：流式问答、总结、改写、长文档分块摘要和最近10轮内存会话。\r\n• 确认后写回原目标；Excel 总结插入当前表顶部一行，PowerPoint 总结新建为第一页。\r\n• 独立 Chat 提示词与 token 预算；沿用现有模型配置。\r\n\r\n", "OfficeTranslate 2.3.0.0\r\n• 三宿主新增右侧文档内 Chat：流式问答、总结、改写、长文档分块摘要和最近10轮内存会话。\r\n• 确认后写回原目标；Excel 总结插入当前表顶部一行，PowerPoint 总结新建为第一页。\r\n• 独立 Chat 提示词与 token 预算；沿用现有模型配置。\r\n\r\n", "OfficeTranslate 2.3.0.0\r\n• Right-docked document Chat in all three hosts: streaming Q&A, summaries, rewrites, chunked long documents and 10 in-memory turns.\r\n• Confirmed writeback to the original target; Excel summaries insert a top row, PowerPoint summaries create a new first slide.\r\n• Separate Chat prompts and token budgets; shared model configuration.\r\n\r\n", "OfficeTranslate 2.3.0.0\r\n• Right-docked document Chat in all three hosts: streaming Q&A, summaries, rewrites, chunked long documents and 10 in-memory turns.\r\n• Confirmed writeback to the original target; Excel summaries insert a top row, PowerPoint summaries create a new first slide.\r\n• Separate Chat prompts and token budgets; shared model configuration.\r\n\r\n", "OfficeTranslate 2.3.0.0\r\n• Right-docked document Chat in all three hosts: streaming Q&A, summaries, rewrites, chunked long documents and 10 in-memory turns.\r\n• Confirmed writeback to the original target; Excel summaries insert a top row, PowerPoint summaries create a new first slide.\r\n• Separate Chat prompts and token budgets; shared model configuration.\r\n\r\n" },
            ["LanguageGroup"] = new[] { "语言", "語言", "언어", "言語", "Language" },
            ["TranslateGroup"] = new[] { "翻译", "翻譯", "번역", "翻訳", "Translate" },
            ["ToolsGroup"] = new[] { "工具", "工具", "도구", "ツール", "Tools" },
            ["Source"] = new[] { "源语言", "來源語言", "원문 언어", "原文言語", "Source" },
            ["Target"] = new[] { "目标语言", "目標語言", "대상 언어", "翻訳言語", "Target" },
            ["Swap"] = new[] { "交换", "交換", "전환", "入替", "Swap" },
            ["Selection"] = new[] { "选区", "選取", "선택", "選択", "Selection" },
            ["Document"] = new[] { "全文", "全文", "전체", "全文", "Document" },
            ["Bilingual"] = new[] { "双语", "雙語", "이중", "対訳", "Bilingual" },
            ["ImageOcr"] = new[] { "图片", "圖片", "이미지", "画像", "Image" },
            ["Cancel"] = new[] { "取消", "取消", "취소", "取消", "Cancel" },
            ["TranslationProgress"] = new[] { "翻译进度", "翻譯進度", "번역 진행률", "翻訳の進行状況", "Translation progress" },
            ["PreparingTranslation"] = new[] { "正在准备翻译…", "正在準備翻譯…", "번역을 준비하는 중…", "翻訳を準備しています…", "Preparing translation…" },
            ["CancelTranslation"] = new[] { "取消翻译", "取消翻譯", "번역 취소", "翻訳をキャンセル", "Cancel translation" },
            ["CancellingTranslation"] = new[] { "正在取消，已完成的内容将保留…", "正在取消，已完成的內容將保留…", "취소하는 중입니다. 완료된 내용은 유지됩니다…", "キャンセルしています。完了した内容は保持されます…", "Cancelling… Completed content will be kept." },
            ["SummaryTitle"] = new[] { "翻译摘要", "翻譯摘要", "번역 요약", "翻訳概要", "Translation summary" },
            ["OcrResultTitle"] = new[] { "OCR 识别结果", "OCR 識別結果", "OCR results", "OCR results", "OCR results" },
            ["OcrResultCompleted"] = new[] { "OCR 识别完成", "OCR 識別完成", "OCR complete", "OCR complete", "OCR complete" },
            ["OcrResultCancelled"] = new[] { "OCR 识别已取消", "OCR 識別已取消", "OCR cancelled", "OCR cancelled", "OCR cancelled" },
            ["OcrResultFailed"] = new[] { "OCR 识别未全部完成", "OCR 識別未全部完成", "OCR did not finish", "OCR did not finish", "OCR did not finish" },
            ["OcrResultCancelledDetails"] = new[] { "OCR 识别已取消；已写回的内容保留。", "OCR 識別已取消；已寫回的內容保留。", "OCR cancelled; written content remains.", "OCR cancelled; written content remains.", "OCR cancelled; written content remains." },
            ["OcrResultFailedDetails"] = new[] { "OCR 识别未全部完成：{0}", "OCR 識別未全部完成：{0}", "OCR did not finish: {0}", "OCR did not finish: {0}", "OCR did not finish: {0}" },
            ["OcrResultDetails"] = new[] {
                "已处理图片 {0} 张，保留原文区域 {1} 个。\r\n未识别到文字 {2} 张；需要检查排版/旁注 {3} 项。",
                "已處理圖片 {0} 張，保留原文區域 {1} 個。\r\n未識別到文字 {2} 張；需要檢查排版/旁註 {3} 項。",
                "Images processed: {0}; source regions kept: {1}.\r\nImages with no text: {2}; layout/notes to review: {3}.",
                "Images processed: {0}; source regions kept: {1}.\r\nImages with no text: {2}; layout/notes to review: {3}.",
                "Images processed: {0}; source regions kept: {1}.\r\nImages with no text: {2}; layout/notes to review: {3}." },
            ["OcrResultUnplacedNotes"] = new[] { "以下旁注未放置，请复制保存；本窗口不会自动关闭。", "以下旁註未放置，請複製儲存；本視窗不會自動關閉。", "Copy and save the unplaced notes below; this window will not close automatically.", "Copy and save the unplaced notes below; this window will not close automatically.", "Copy and save the unplaced notes below; this window will not close automatically." },
            ["OcrResultCloseCountdown"] = new[] { "关闭（{0}s）", "關閉（{0}s）", "Close ({0}s)", "Close ({0}s)", "Close ({0}s)" },
            ["SummaryCancelled"] = new[] { "翻译已取消：有旁注待保存", "翻譯已取消：有旁註待儲存", "번역 취소: 저장할 주석이 있습니다", "翻訳をキャンセル：注記の保存が必要です", "Translation cancelled: notes to save" },
            ["SummaryFailed"] = new[] { "翻译已中止：有旁注待保存", "翻譯已中止：有旁註待儲存", "번역 중단: 저장할 주석이 있습니다", "翻訳を中断：注記の保存が必要です", "Translation stopped: notes to save" },
            ["SummaryInterruptedHint"] = new[] { "任务未全部完成。已写回的内容保留；未放置旁注全文如下，可复制保存，窗口不会自动关闭。", "工作未全部完成。已寫回內容保留；未放置旁註全文如下，可複製儲存，視窗不會自動關閉。", "작업이 완료되지 않았습니다. 기록된 내용은 유지됩니다. 아래 미배치 주석 전문을 복사해 저장하세요. 창은 자동으로 닫히지 않습니다.", "処理は未完了です。書き込み済みの内容は保持されます。以下の未配置注記をコピーして保存してください。自動では閉じません。", "The task did not finish. Written content remains. Copy and save the full unplaced notes below; this window will not close automatically." },
            ["OcrRegionDetails"] = new[] {
                "图片 OCR：{0} 张；模型条目 {1}，保留 {2}，无可用内容丢弃 {3}\r\n坏坐标转旁注 {4}，空译文 {5}；图片待检查 {6} 项（含排版降级，不代表文本严格重试失败）。",
                "圖片 OCR：{0} 張；模型條目 {1}，保留 {2}，無可用內容捨棄 {3}\r\n壞座標轉旁註 {4}，空譯文 {5}；圖片待檢查 {6} 項（含排版降級，並非文字重試失敗）。",
                "이미지 OCR {0}개: 모델 항목 {1}, 보존 {2}, 사용 불가 {3}\r\n좌표 오류 주석 {4}, 번역 누락 {5}, 이미지 확인 {6}개 (레이아웃 대체 포함; 텍스트 재시도 실패와 별개).",
                "画像 OCR {0} 枚：モデル項目 {1}、保持 {2}、使用不可 {3}\r\n座標不良の注記 {4}、訳文なし {5}、画像要確認 {6} 件（配置の代替を含む。テキスト再試行失敗とは別）。",
                "Image OCR: {0} images; model entries {1}, kept {2}, unusable {3}\r\nBad-box notes {4}, missing translations {5}; image review {6} (includes layout fallback, not text retry failures)." },
            ["SummaryUnplaced"] = new[] { "需处理：{0} 条图片旁注未放置", "需處理：{0} 條圖片旁註未放置", "확인 필요: 이미지 주석 {0}개 미배치", "要確認：画像注記 {0} 件未配置", "Action needed: {0} image notes not placed" },
            ["UnplacedNoteDetails"] = new[] { "以下旁注未放置；已成功的译文保留。完整内容可选中复制，本窗口不会自动关闭。", "以下旁註未放置；已成功的譯文保留。完整內容可選取複製，本視窗不會自動關閉。", "아래 주석을 배치하지 못했습니다. 완료된 번역은 유지됩니다. 전체 내용을 복사할 수 있으며 창은 자동으로 닫히지 않습니다.", "以下の注記は配置できませんでした。完了分は保持されます。全文をコピーできます。このウィンドウは自動で閉じません。", "The notes below were not placed. Successful translations remain. Select and copy the full text; this window will not close automatically." },
            ["SummaryCompleted"] = new[] { "翻译完成：共 {0} 项", "翻譯完成：共 {0} 項", "번역 완료: 총 {0}개", "翻訳完了：全 {0} 件", "Translation complete: {0} items" },
            ["SummaryDetails"] = new[] { "已翻译 {0}　缓存命中 {1}\r\n无源语言跳过 {2}　需要检查 {3}　OCR 未识别 {4}", "已翻譯 {0}　快取命中 {1}\r\n無來源語言而略過 {2}　需要檢查 {3}　OCR 未識別 {4}", "번역 {0}　캐시 적중 {1}\r\n원문 언어 없음 {2}　확인 필요 {3}　OCR 미인식 {4}", "翻訳済み {0}　キャッシュ命中 {1}\r\n原文言語なし {2}　要確認 {3}　OCR 未認識 {4}", "Translated {0}   Cache hits {1}\r\nNo source language {2}   Review {3}   OCR empty {4}" },
            ["SummaryReviewHint"] = new[] { "其中 {0} 项严格重试后仍未变化，已保留原文。", "其中 {0} 項嚴格重試後仍未變更，已保留原文。", "{0}개 항목은 재시도 후에도 변경되지 않아 원문을 유지했습니다.", "{0} 件は再試行後も変化しなかったため、原文を保持しました。", "{0} items were still unchanged after retry; the originals were kept." },
            ["SummaryPartialHint"] = new[] { "其中 {0} 项已部分翻译；未通过校验的部分保留原文，需要检查。", "其中 {0} 項已部分翻譯；未通過校驗的部分保留原文，需要檢查。", "{0}개 항목은 부분 번역되었습니다. 검증에 실패한 부분은 원문을 유지했으며 확인이 필요합니다.", "{0} 件は部分的に翻訳済みです。検証できない部分は原文を保持し、確認が必要です。", "{0} items were partially translated; unvalidated parts retain their originals and need review." },
            ["SummaryParagraphFallback"] = new[] { "逐段降级：已翻译 {0} 段，保留原文待检查 {1} 段。", "逐段降級：已翻譯 {0} 段，保留原文待檢查 {1} 段。", "문단별 대체: 번역 {0}개, 원문 유지 및 확인 {1}개.", "段落別の代替処理：翻訳済み {0} 段落、原文保持・要確認 {1} 段落。", "Paragraph fallback: {0} translated, {1} kept as original for review." },
            ["ResultCloseCountdown"] = new[] { "确定（{0} 秒后自动关闭）", "確定（{0} 秒後自動關閉）", "확인 ({0}초 후 자동 닫기)", "確認（{0} 秒後に自動で閉じる）", "OK (closes in {0}s)" },
            ["Settings"] = new[] { "设置", "設定", "설정", "設定", "Settings" },
            ["SettingsTitle"] = new[] { "OfficeTranslate 设置", "OfficeTranslate 設定", "OfficeTranslate 설정", "OfficeTranslate 設定", "OfficeTranslate Settings" },
            ["Heading"] = new[] { "翻译服务设置", "翻譯服務設定", "번역 서비스 설정", "翻訳サービス設定", "Translation service" },
            ["UiLanguage"] = new[] { "界面语言", "介面語言", "UI 언어", "表示言語", "UI language" },
            ["Provider"] = new[] { "服务类型", "服務類型", "서비스 유형", "サービス種類", "Provider" },
            ["Model"] = new[] { "模型", "模型", "모델", "モデル", "Model" },
            ["ImageModel"] = new[] { "图片模型", "圖片模型", "이미지 모델", "画像モデル", "Image model" },
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
