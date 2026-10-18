using System.Collections.Generic;

namespace OfficeTranslate.WordAddIn
{
    internal static class UiText
    {
        public static readonly string[] LanguageCodes = { "zh-CN", "zh-HK", "ko", "ja", "en" };
        public static readonly string[] LanguageNames = { "简体中文", "繁體中文", "한국어", "日本語", "English" };

        private static readonly Dictionary<string, string[]> Texts = new Dictionary<string, string[]>
        {
            ["WordTextConcurrency"] = new[] { "Word 文本并发数", "Word 文字並行數", "Word 텍스트 동시 요청", "Word テキスト同時実行数", "Word text concurrency" },
            ["WordConcurrencyInvalid"] = new[] { "Word 文本并发数必须在 1–32 之间。", "Word 文字並行數必須介於 1–32。", "Word 텍스트 동시 요청 수는 1–32여야 합니다.", "Word テキスト同時実行数は1–32で指定してください。", "Word text concurrency must be between 1 and 32." },
            ["TextTranslating"] = new[] { "OfficeTranslate：正在翻译 {0}/{1}", "OfficeTranslate：正在翻譯 {0}/{1}", "OfficeTranslate: 번역 중 {0}/{1}", "OfficeTranslate：翻訳中 {0}/{1}", "OfficeTranslate: translating {0}/{1}" },
            ["TextWriting"] = new[] { "OfficeTranslate：正在写回 {0}/{1}", "OfficeTranslate：正在寫回 {0}/{1}", "OfficeTranslate: 쓰는 중 {0}/{1}", "OfficeTranslate：書き込み中 {0}/{1}", "OfficeTranslate: writing {0}/{1}" },
            ["TextCompleted"] = new[] { "OfficeTranslate：已完成 {0}/{1}", "OfficeTranslate：已完成 {0}/{1}", "OfficeTranslate: 완료 {0}/{1}", "OfficeTranslate：完了 {0}/{1}", "OfficeTranslate: completed {0}/{1}" },
            ["SummaryTextFailures"] = new[] { "处理结束，{0} 个文本块失败", "處理結束，{0} 個文字區塊失敗", "처리 완료, 텍스트 블록 {0}개 실패", "処理終了、{0}個のテキストブロックが失敗", "Processing finished; {0} text blocks failed" },
            ["TextBlockFailure"] = new[] { "第 {0} 块：{1}失败（{2}）。请重试此块。", "第 {0} 區塊：{1}失敗（{2}）。請重試此區塊。", "블록 {0}: {1} 실패 ({2}). 이 블록을 다시 시도하세요.", "ブロック {0}：{1}失敗（{2}）。再試行してください。", "Block {0}: {1} failed ({2}). Retry this block." },
            ["TextWriteStage"] = new[] { "写回", "寫回", "쓰기", "書き込み", "write" },
            ["TextTranslateStage"] = new[] { "翻译", "翻譯", "번역", "翻訳", "translation" },
            ["ParallelChangelog"] = new[] {
                "OfficeTranslate 2.3.4.0\r\n• Word 文本默认 4 块并行翻译，可在翻译设置中调整为 1–32；按原文顺序逐块写回。\r\n• 取消保留已写回内容；单块失败继续处理并汇总。图片 OCR 翻译仍串行。\r\n\r\n",
                "OfficeTranslate 2.3.4.0\r\n• Word 文字預設 4 區塊並行翻譯，可在翻譯設定調整為 1–32；依原文順序逐塊寫回。\r\n• 取消保留已寫回內容；單塊失敗繼續處理並彙總。圖片 OCR 翻譯仍循序執行。\r\n\r\n",
                "OfficeTranslate 2.3.4.0\r\n• Word 텍스트를 기본 4개 블록씩 병렬 번역합니다. 설정에서 1–32로 조절하며 원문 순서대로 기록합니다.\r\n• 취소 시 완료된 기록을 보존하며 블록 오류는 계속 처리한 뒤 보고합니다. 이미지 OCR 번역은 순차 실행합니다.\r\n\r\n",
                "OfficeTranslate 2.3.4.0\r\n• Word テキストを既定で4ブロック並列翻訳し、原文順に書き込みます。設定で1–32に変更できます。\r\n• キャンセル時は書き込み済みの内容を保持し、ブロックの失敗は処理を続けて集計します。画像OCR翻訳は順次実行します。\r\n\r\n",
                "OfficeTranslate 2.3.4.0\r\n• Word text translates in parallel (default 4, configurable 1–32) and commits each block in source order.\r\n• Cancellation keeps committed blocks; individual failures are reported after processing the remaining blocks. Image OCR translation remains serial.\r\n\r\n" },
            ["Chat"] = new[] { "文档内 Chat","文件內 Chat","문서 Chat","文書 Chat","Document Chat" },
            ["ChatCopy"] = new[] { "复制", "複製", "복사", "コピー", "Copy" },
            ["ChatCopied"] = new[] { "已复制", "已複製", "복사됨", "コピーしました", "Copied" },
            ["ChatCopyFailed"] = new[] { "复制失败", "複製失敗", "복사 실패", "コピーに失敗しました", "Copy failed" },
            ["ChatWritebackChangelog"] = new[] {
                "OfficeTranslate 2.3.8.0\r\n• Chat 所有模式发送开始即清空输入框，回答完成时保留下一条草稿。\r\n• 手动提问完成后可确认替换选中或插入下方，不提供插入到开头。\r\n• 包含 2.3.7 翻译按钮和 2.3.6 CJK 空格清理。\r\n\r\n",
                "OfficeTranslate 2.3.8.0\r\n• Chat 所有模式開始傳送即清空輸入框，回答完成時保留下一則草稿。\r\n• 手動提問完成後可確認取代選取或插入下方，不提供插入到開頭。\r\n• 包含 2.3.7 翻譯按鈕和 2.3.6 CJK 空格清理。\r\n\r\n",
                "OfficeTranslate 2.3.8.0\r\n• 모든 Chat 모드에서 전송 즉시 입력란을 비우고 답변 완료 시 다음 초안을 유지합니다.\r\n• 질문 답변을 확인 후 선택 영역에 바꾸거나 아래에 삽입할 수 있습니다. 문서 시작에는 삽입하지 않습니다.\r\n• 2.3.7 번역 버튼과 2.3.6 CJK 공백 정리를 포함합니다.\r\n\r\n",
                "OfficeTranslate 2.3.8.0\r\n• すべてのChatモードで送信開始時に入力欄をクリアし、回答完了時に次の下書きを保持します。\r\n• 質問の回答は確認後に選択範囲を置換するか下に挿入できます。文書先頭への挿入は行いません。\r\n• 2.3.7の翻訳ボタンと2.3.6のCJK空白除去を含みます。\r\n\r\n",
                "OfficeTranslate 2.3.8.0\r\n• All Chat modes clear input when sending starts and preserve the next draft on completion.\r\n• Question answers can replace the selection or insert below after confirmation, without inserting at the document start.\r\n• Includes the 2.3.7 translation button and 2.3.6 CJK spacing cleanup.\r\n\r\n" },
            ["TranslateChangelog"] = new[] {
                "OfficeTranslate 2.3.7.0\r\n• Chat 新增翻译按钮，直接使用翻译服务的目标语言，流式显示译文。\r\n• 可自定义翻译系统提示词；译文支持替换选中或插入下方。\r\n\r\n",
                "OfficeTranslate 2.3.7.0\r\n• Chat 新增翻譯按鈕，直接使用翻譯服務的目標語言，串流顯示譯文。\r\n• 可自訂翻譯系統提示詞；譯文支援取代選取或插入下方。\r\n\r\n",
                "OfficeTranslate 2.3.7.0\r\n• Chat 번역 버튼이 번역 서비스의 대상 언어로 번역하고 스트리밍으로 표시합니다.\r\n• 번역 시스템 프롬프트를 편집하고 선택 영역을 바꾸거나 아래에 삽입할 수 있습니다.\r\n\r\n",
                "OfficeTranslate 2.3.7.0\r\n• Chat の翻訳ボタンで翻訳サービスの翻訳先言語に翻訳し、ストリーミング表示します。\r\n• 翻訳システムプロンプトの編集、選択範囲の置換、下への挿入に対応します。\r\n\r\n",
                "OfficeTranslate 2.3.7.0\r\n• Chat translates into the translation service's target language with streamed output.\r\n• Customize the translation system prompt; replace the selection or insert below.\r\n\r\n" },
            ["CjkSpacingChangelog"] = new[] {
                "OfficeTranslate 2.3.6.0\r\n• Chat 回复完成后自动删除中日文字符之间的多余空格，显示、历史、导出与复制保持一致。\r\n• 保留中英、中数、韩文间距，以及代码块和行内代码。\r\n\r\n",
                "OfficeTranslate 2.3.6.0\r\n• Chat 回覆完成後自動刪除中日文字元之間的多餘空格，顯示、歷史、匯出與複製保持一致。\r\n• 保留中英、中數、韓文間距，以及程式碼區塊和行內程式碼。\r\n\r\n",
                "OfficeTranslate 2.3.6.0\r\n• Chat 답변 완료 후 중국어·일본어 문자 사이의 불필요한 공백을 제거하고 화면, 기록, 내보내기, 복사에 동일하게 적용합니다.\r\n• 한국어, 영문·숫자 경계, 코드 블록과 인라인 코드의 공백을 보존합니다.\r\n\r\n",
                "OfficeTranslate 2.3.6.0\r\n• Chat回答の完了後、中国語・日本語の文字間の余分な空白を除去し、表示・履歴・エクスポート・コピーに反映します。\r\n• 韓国語、英数字との間隔、コードブロックとインラインコードを保持します。\r\n\r\n",
                "OfficeTranslate 2.3.6.0\r\n• Completed Chat answers remove extra spaces between Chinese/Japanese characters consistently across display, history, export and copy.\r\n• Preserve Korean spacing, English/number boundaries, fenced and inline code.\r\n\r\n" },
            ["CopyChangelog"] = new[] {
                "OfficeTranslate 2.3.5.0\r\n• 聊天气泡和历史消息支持一键复制原文，AI 回复保留 Markdown。\r\n• 流式输出完成后显示复制按钮，提供复制成功或失败提示。\r\n\r\n",
                "OfficeTranslate 2.3.5.0\r\n• 聊天氣泡與歷史訊息支援一鍵複製原文，AI 回覆保留 Markdown。\r\n• 串流輸出結束後顯示複製按鈕，提供複製成功或失敗提示。\r\n\r\n",
                "OfficeTranslate 2.3.5.0\r\n• 대화 말풍선과 기록에서 원문을 복사합니다. AI 답변의 Markdown을 보존합니다.\r\n• 스트리밍 종료 후 복사 버튼과 성공 또는 실패 안내를 표시합니다.\r\n\r\n",
                "OfficeTranslate 2.3.5.0\r\n• 会話の吹き出しと履歴から原文をコピーできます。AI回答のMarkdownを保持します。\r\n• ストリーミング終了後にコピーボタンを表示し、成功または失敗を通知します。\r\n\r\n",
                "OfficeTranslate 2.3.5.0\r\n• Copy original messages from chat bubbles and history, preserving AI Markdown.\r\n• Copy buttons appear after streaming ends, with success or failure feedback.\r\n\r\n" },
            ["ChatUserAvatar"] = new[] { "你", "你", "나", "私", "You" },
            ["ChatAiAvatar"] = new[] { "AI", "AI", "AI", "AI", "AI" },
            ["ChatTranslate"] = new[] { "翻译","翻譯","번역","翻訳","Translate" },
            ["ChatSummary"] = new[] { "总结","總結","요약","要約","Summarize" },
            ["ChatRewrite"] = new[] { "改写","改寫","다시 작성","書き直し","Rewrite" },
            ["ChatClear"] = new[] { "清空","清空","지우기","クリア","Clear" },
            ["ChatExport"] = new[] { "导出", "匯出", "내보내기", "エクスポート", "Export" },
            ["ChatExportTitle"] = new[] { "导出聊天内容", "匯出聊天內容", "대화 내보내기", "会話をエクスポート", "Export conversation" },
            ["ChatExportFilter"] = new[] { "Markdown 文件 (*.md)|*.md|纯文本文件 (*.txt)|*.txt", "Markdown 檔案 (*.md)|*.md|純文字檔案 (*.txt)|*.txt", "Markdown 파일 (*.md)|*.md|텍스트 파일 (*.txt)|*.txt", "Markdown ファイル (*.md)|*.md|テキストファイル (*.txt)|*.txt", "Markdown files (*.md)|*.md|Plain text files (*.txt)|*.txt" },
            ["ChatExportFileName"] = new[] { "文档内Chat_{0:yyyyMMdd_HHmmss}.md", "文件內Chat_{0:yyyyMMdd_HHmmss}.md", "문서Chat_{0:yyyyMMdd_HHmmss}.md", "文書Chat_{0:yyyyMMdd_HHmmss}.md", "DocumentChat_{0:yyyyMMdd_HHmmss}.md" },
            ["ChatExportHeader"] = new[] { "文档内 Chat 聊天记录", "文件內 Chat 聊天記錄", "문서 Chat 대화 기록", "文書 Chat 会話履歴", "Document Chat transcript" },
            ["ChatExportTimestamp"] = new[] { "时间", "時間", "시간", "時刻", "Time" },
            ["ChatExportQuestion"] = new[] { "用户提问", "使用者提問", "사용자 질문", "ユーザーの質問", "User question" },
            ["ChatExportAnswer"] = new[] { "AI 回答", "AI 回答", "AI 답변", "AI の回答", "AI answer" },
            ["ChatExportEmpty"] = new[] { "暂无可导出的完整问答。", "暫無可匯出的完整問答。", "내보낼 완료된 대화가 없습니다.", "エクスポートできる完了済みの会話がありません。", "No completed turns to export." },
            ["ChatExportSuccess"] = new[] { "聊天内容已导出。", "聊天內容已匯出。", "대화를 내보냈습니다.", "会話をエクスポートしました。", "Conversation exported." },
            ["ChatExportFailed"] = new[] { "导出失败。", "匯出失敗。", "내보내기에 실패했습니다.", "エクスポートに失敗しました。", "Export failed." },
            ["ChatExportAccessError"] = new[] { "没有文件写入权限，请选择可写位置。", "沒有檔案寫入權限，請選擇可寫入的位置。", "파일 쓰기 권한이 없습니다. 쓰기 가능한 위치를 선택하세요.", "ファイルの書き込み権限がありません。書き込み可能な場所を選択してください。", "File access denied. Choose a writable location." },
            ["ChatExportIoError"] = new[] { "文件无法写入，请检查磁盘空间、文件占用和目标位置。", "無法寫入檔案，請檢查磁碟空間、檔案是否被佔用及目標位置。", "파일을 쓸 수 없습니다. 디스크 공간, 파일 사용 여부 및 저장 위치를 확인하세요.", "ファイルを書き込めません。空き容量、ファイルの使用状況と保存先を確認してください。", "Cannot write the file. Check disk space, file locks and the destination." },
            ["ChatExportPathError"] = new[] { "文件名或路径无效，请重新选择。", "檔案名稱或路徑無效，請重新選擇。", "파일 이름 또는 경로가 잘못되었습니다. 다시 선택하세요.", "ファイル名またはパスが無効です。選び直してください。", "Invalid file name or path. Choose another destination." },
            ["ChatExportUnknownError"] = new[] { "无法保存聊天内容，请重试或选择其他位置。", "無法儲存聊天內容，請重試或選擇其他位置。", "대화를 저장할 수 없습니다. 다시 시도하거나 다른 위치를 선택하세요.", "会話を保存できません。再試行するか、別の場所を選択してください。", "Cannot save the conversation. Retry or choose another location." },
            ["ChatRefresh"] = new[] { "更新上下文","更新內容","문맥 업데이트","文脈を更新","Update context" },
            ["ChatPaneTitle"] = new[] { "OfficeTranslate", "OfficeTranslate", "OfficeTranslate", "OfficeTranslate", "OfficeTranslate" },
            ["ChatRefreshTip"] = new[] { "重新读取当前选中内容或全文，更新上下文预览。每次发送前也会自动重新读取。", "重新讀取目前選取內容或全文，更新內容預覽。每次傳送前也會自動重新讀取。", "현재 선택 영역 또는 전체 문서를 다시 읽어 문맥 미리보기를 업데이트합니다. 보내기 전에도 자동으로 다시 읽습니다.", "現在の選択範囲または文書全体を再読込して文脈のプレビューを更新します。送信前にも自動で再読込します。", "Reread the current selection or entire document to update the context preview. Context is also reread automatically before each send." },
            ["ChatContextTip"] = new[] { "显示最近读取的上下文。发送前自动重新读取当前选中内容；未选中时读取全文。", "顯示最近讀取的內容。傳送前自動重新讀取目前選取內容；未選取時讀取全文。", "최근 읽은 문맥입니다. 보내기 전 현재 선택 영역을 자동으로 다시 읽으며 선택이 없으면 전체 문서를 읽습니다.", "最後に読み取った文脈です。送信前に現在の選択範囲を再読込し、選択がなければ文書全体を読み取ります。", "Shows the last context preview. Before sending, the current selection is reread automatically, or the entire document if nothing is selected." },
            ["ChatInputHint"] = new[] { "输入问题（Enter 发送，Shift+Enter 换行）", "輸入問題（Enter 傳送，Shift+Enter 換行）", "질문 입력 (Enter로 보내기, Shift+Enter로 줄 바꿈)", "質問を入力（Enterで送信、Shift+Enterで改行）", "Enter a question (Enter to send, Shift+Enter for a new line)" },
            ["ChatTranscript"] = new[] { "会话内容", "對話內容", "대화 내용", "会話内容", "Conversation" },
            ["ChatStage"] = new[] { "阶段", "階段", "단계", "段階", "Stage" },
            ["ChatRibbonTip"] = new[] { "根据选中内容或全文提问、总结或改写。", "根據選取內容或全文提問、總結或改寫。", "선택 영역 또는 전체 문서에 대해 질문하거나 요약 또는 다시 작성합니다.", "選択範囲や文書全体について質問、要約、書き直しを行います。", "Ask about, summarize or rewrite the selection or entire document." },
            ["ChatSettings"] = new[] { "Chat 设置","Chat 設定","Chat 설정","Chat 設定","Chat settings" },
            ["ChatUseTranslationSettings"] = new[] { "使用翻译设置（风格/额外要求/术语表）", "使用翻譯設定（風格/額外要求/術語表）", "번역 설정 사용 (스타일/추가 요구 사항/용어집)", "翻訳設定を使用（スタイル/追加要件/用語集）", "Use translation settings (style/requirements/glossary)" },
            ["ChatHistory"] = new[] { "历史会话", "歷史會話", "대화 기록", "会話履歴", "Chat history" },
            ["ChatHistoryTitle"] = new[] { "历史会话（只读）", "歷史會話（唯讀）", "대화 기록 (읽기 전용)", "会話履歴（読み取り専用）", "Chat history (read only)" },
            ["ChatHistoryTime"] = new[] { "创建时间", "建立時間", "생성 시간", "作成日時", "Created" },
            ["ChatHistoryHost"] = new[] { "来源应用", "來源應用", "원본 앱", "作成元アプリ", "Source app" },
            ["ChatHistoryDocument"] = new[] { "文档名", "文件名稱", "문서 이름", "文書名", "Document name" },
            ["ChatHistoryQuestion"] = new[] { "首个问题预览", "首個問題預覽", "첫 질문 미리 보기", "最初の質問", "First question preview" },
            ["ChatHistoryDetails"] = new[] { "会话详情（原文，只读）", "會話詳情（原文，唯讀）", "대화 내용 (원문, 읽기 전용)", "会話の詳細（原文、読み取り専用）", "Conversation details (original text, read only)" },
            ["ChatHistoryDelete"] = new[] { "删除会话", "刪除會話", "대화 삭제", "会話を削除", "Delete conversation" },
            ["ChatHistoryDeleteAll"] = new[] { "清空全部", "清空全部", "모두 삭제", "すべて削除", "Clear all" },
            ["ChatHistoryDeleteConfirm"] = new[] { "确定永久删除选中的历史会话？", "確定永久刪除選取的歷史會話？", "선택한 대화 기록을 영구 삭제할까요?", "選択した会話履歴を完全に削除しますか？", "Permanently delete the selected conversation?" },
            ["ChatHistoryDeleteAllConfirm"] = new[] { "确定永久删除全部历史会话？此操作无法撤销。", "確定永久刪除全部歷史會話？此操作無法復原。", "모든 대화 기록을 영구 삭제할까요? 취소할 수 없습니다.", "すべての会話履歴を完全に削除しますか？元に戻せません。", "Permanently delete all chat history? This cannot be undone." },
            ["ChatHistoryEmpty"] = new[] { "暂无历史会话。", "暫無歷史會話。", "대화 기록이 없습니다.", "会話履歴はありません。", "No chat history yet." },
            ["ChatHistoryHint"] = new[] { "选择会话查看完整问答原文。历史文件仅保存在本地。", "選取會話以查看完整問答原文。歷史檔案僅儲存在本機。", "대화를 선택하여 전체 원문을 확인하세요. 기록은 로컬에만 저장됩니다.", "会話を選択すると全文を表示します。履歴はローカルにのみ保存されます。", "Select a conversation to view all original questions and answers. History is stored locally." },
            ["ChatHistoryUnreadable"] = new[] { "部分历史文件无法读取，已跳过；可用“清空全部”删除。", "部分歷史檔案無法讀取，已略過；可用「清空全部」刪除。", "일부 기록을 읽을 수 없어 건너뛰었습니다. 모두 삭제로 제거할 수 있습니다.", "読み取れない履歴をスキップしました。「すべて削除」で削除できます。", "Some history files could not be read and were skipped. Clear all can remove them." },
            ["ChatHistoryLoadFailed"] = new[] { "无法读取历史会话，请检查文件权限后重新打开。", "無法讀取歷史會話，請檢查檔案權限後重新開啟。", "기록을 읽을 수 없습니다. 파일 권한을 확인하고 다시 여세요.", "履歴を読み取れません。ファイルの権限を確認して開き直してください。", "Cannot read chat history. Check file permissions and reopen." },
            ["ChatHistoryDeleteFailed"] = new[] { "删除未全部完成，请检查文件权限后重试。", "刪除未全部完成，請檢查檔案權限後重試。", "삭제를 완료하지 못했습니다. 파일 권한을 확인하고 다시 시도하세요.", "削除が完了しませんでした。ファイルの権限を確認して再試行してください。", "Deletion did not fully complete. Check file permissions and retry." },
            ["ChatHistorySaveFailed"] = new[] { "历史保存失败，请检查磁盘空间和权限。问答仍在内存，下次问答或清空时重试；关闭前请重试或导出。", "歷史儲存失敗，請檢查磁碟空間及權限。問答仍在記憶體，下次問答或清空時重試；關閉前請重試或匯出。", "기록을 저장하지 못했습니다. 디스크와 권한을 확인하세요. 다음 대화 또는 지우기에서 재시도합니다. 닫기 전에 재시도하거나 내보내세요.", "履歴を保存できません。容量と権限を確認してください。次の会話またはクリア時に再試行します。閉じる前に再試行またはエクスポートしてください。", "History save failed. Check disk space and permissions. Turns remain in memory; the next turn or Clear retries. Retry or export before closing." },
            ["ChatSend"] = new[] { "发送","傳送","보내기","送信","Send" },
            ["ChatEmpty"] = new[] { "选中内容或打开文档，提问、总结或改写。Enter 发送，Shift+Enter 换行。活跃会话保留最近10轮，完整问答自动保存到本地历史。","選取內容或開啟文件，提問、總結或改寫。Enter 傳送，Shift+Enter 換行。活躍會話保留最近10輪，完整問答自動儲存至本機歷史。","내용을 선택하거나 문서를 열어 질문, 요약 또는 다시 작성을 시작하세요. Enter로 보내고 Shift+Enter로 줄을 바꿉니다. 최근 10회 대화는 메모리에, 전체 대화는 로컬 기록에 자동 저장됩니다.","内容を選択するか文書を開き、質問、要約、書き直しを行います。Enterで送信、Shift+Enterで改行します。メモリは直近10往復を保持し、完了した会話全体はローカル履歴に自動保存します。","Select content or open a document to ask, summarize or rewrite. Enter to send, Shift+Enter for a new line. Active memory keeps 10 turns; all completed turns are saved to local history." },
            ["ChatContext"] = new[] { "上下文：{0} · 约{1:N0}字","內容：{0} · 約{1:N0}字","문맥: {0} · 약 {1:N0}자","文脈：{0} · 約{1:N0}文字","Context: {0} · approx. {1:N0} characters" },
            ["ChatSelected"] = new[] { "选中内容","選取內容","선택 영역","選択範囲","Selection" },
            ["ChatFull"] = new[] { "全文","全文","전체 문서","文書全体","Entire document" },
            ["ChatNoDocument"] = new[] { "上下文不可用，请打开文档并更新上下文。","無法讀取內容，請開啟文件並更新內容。","문맥을 사용할 수 없습니다. 문서를 열고 문맥을 업데이트하세요.","文脈を読み取れません。文書を開いて文脈を更新してください。","Context unavailable. Open a document and update the context." },
            ["ChatCleared"] = new[] { "当前会话已清空，已保存的历史仍保留。","目前會話已清空，已儲存的歷史仍保留。","현재 대화를 지웠습니다. 저장된 기록은 유지됩니다.","現在の会話をクリアしました。保存済みの履歴は保持されます。","Current conversation cleared. Saved history is retained." },
            ["ChatTranslateRequest"] = new[] { "请将以下内容翻译成{0}","請將以下內容翻譯成{0}","다음 내용을 {0}(으)로 번역하세요.","以下の内容を{0}に翻訳してください。","Please translate the following content into {0}." },
            ["ChatSummaryRequest"] = new[] { "请总结当前文档上下文。","請總結目前文件內容。","현재 문서 문맥을 요약하세요.","現在の文書の文脈を要約してください。","Please summarize the current document context." },
            ["ChatRewriteRequest"] = new[] { "请改写当前文档上下文，保留原意和事实，只返回改写后的正文。","請改寫目前文件內容，保留原意和事實，只傳回改寫後的正文。","의미와 사실을 보존하여 현재 문서 문맥을 다시 작성하세요. 다시 작성한 본문만 반환하세요.","意味と事実を保ち、現在の文書の文脈を書き直してください。書き直した本文のみ返してください。","Please rewrite the current document context, preserving meaning and facts. Return only the rewritten text." },
            ["ChatNeedInput"] = new[] { "请输入问题。","請輸入問題。","질문을 입력하세요.","質問を入力してください。","Enter a question." },
            ["ChatNoText"] = new[] { "当前上下文没有可读取的文字。","目前內容中沒有可讀取的文字。","현재 문맥에 읽을 수 있는 텍스트가 없습니다.","現在の文脈には読み取れる文字がありません。","The current context contains no readable text." },
            ["ChatYou"] = new[] { "你","你","사용자","あなた","You" },
            ["ChatErrorDetails"] = new[] { "错误详情（可复制）","錯誤詳情（可複製）","오류 세부 정보 (복사 가능)","エラーの詳細（コピー可能）","Error details (copyable)" },
            ["ChatAssistant"] = new[] { "助手","助手","도우미","アシスタント","Assistant" },
            ["ChatLoading"] = new[] { "正在生成回答…可点击取消。","正在產生回答…可按取消。","답변 생성 중… 취소할 수 있습니다.","回答を生成中…キャンセルできます。","Generating a response… You can cancel." },
            ["ChatChunkProgress"] = new[] { "长文档摘要：正在处理第{0}/{1}块…","長文件摘要：正在處理第{0}/{1}塊…","긴 문서: {0}/{1} 조각 요약 중…","長い文書：{0}/{1}個目の断片を要約中…","Long document: summarizing chunk {0}/{1}…" },
            ["ChatRecovering"] = new[] { "流式输出中断，正在尝试普通请求…","串流輸出中斷，正在嘗試一般請求…","스트리밍이 중단되었습니다. 일반 요청으로 다시 시도 중…","ストリーミングが中断しました。通常のリクエストで再試行中…","Stream interrupted. Retrying without streaming…" },
            ["ChatFallback"] = new[] { "流式输出不可用或已中断，已降级为普通请求并完成回答。","串流輸出無法使用或已中斷，已改用一般請求並完成回答。","스트리밍을 사용할 수 없거나 중단되어 일반 요청으로 답변을 완료했습니다.","ストリーミングが利用できないか中断したため、通常のリクエストで回答を完了しました。","Streaming was unavailable or interrupted. Completed using a non-streaming request." },
            ["ChatReduced"] = new[] { "长文档已分块摘要；本次回答或改写基于摘要。","長文件已分塊摘要；本次回答或改寫根據摘要產生。","긴 문서를 조각별로 요약했습니다. 이번 답변 또는 다시 작성은 요약을 기반으로 합니다.","長い文書を断片ごとに要約しました。今回の回答または書き直しは要約に基づいています。","Long document summarized in chunks; this answer or rewrite is based on summaries." },
            ["ChatDone"] = new[] { "已完成。写回需点击下方按钮并确认。","已完成。寫回時需按下方按鈕並確認。","완료했습니다. 문서에 쓰려면 아래 작업을 선택하고 확인하세요.","完了しました。書き込むには下の操作を選び、確認してください。","Completed. To write back, choose an action and confirm." },
            ["ChatCancelled"] = new[] { "已取消；未完成回答不会加入会话或写回。","已取消；未完成的回答不會加入對話或寫回文件。","취소했습니다. 미완성 답변은 대화에 추가하거나 문서에 쓰지 않습니다.","キャンセルしました。未完成の回答は会話に追加せず、文書にも書き込みません。","Cancelled. Incomplete responses are not added to history or written back." },
            ["ChatNetworkError"] = new[] { "网络连接失败或请求超时；未写回文档。请检查网络、代理及服务可达性，必要时增加请求超时后重试。","網絡連線失敗或請求逾時；未寫回文件。請檢查網絡、代理及服務連線，必要時增加請求逾時後重試。","연결에 실패했거나 시간이 초과되었습니다. 문서는 변경되지 않았습니다. 네트워크, 프록시 및 서비스를 확인하거나 요청 제한 시간을 늘리세요.","接続に失敗したかタイムアウトしました。文書は変更されていません。ネットワーク、プロキシ、サービスを確認するか、待機時間を延ばしてください。","Connection failed or timed out; the document was not changed. Check the network, proxy and service, or increase the request timeout." },
            ["ChatAuthError"] = new[] { "API 认证或权限失败；未写回文档。请检查 API Key、账号权限及所选服务。","API 認證或權限失敗；未寫回文件。請檢查 API Key、帳戶權限及所選服務。","API 인증 또는 권한 오류입니다. 문서는 변경되지 않았습니다. API 키, 계정 권한 및 서비스를 확인하세요.","API認証または権限のエラーです。文書は変更されていません。APIキー、アカウント権限、サービスを確認してください。","API authentication or permission failed; the document was not changed. Check the API key, account permissions and service." },
            ["ChatModelError"] = new[] { "模型服务拒绝请求或返回不完整数据；未写回文档。请检查模型能力和服务状态，参考下方详情后重试。","模型服務拒絕請求或傳回不完整資料；未寫回文件。請檢查模型能力和服務狀態，參考下方詳情後重試。","모델이 요청을 거부했거나 불완전한 데이터를 반환했습니다. 문서는 변경되지 않았습니다. 모델 기능과 서비스 상태를 확인하고 아래 세부 정보를 참고하여 다시 시도하세요.","モデルがリクエストを拒否したか、不完全なデータを返しました。文書は変更されていません。モデルの機能とサービスの状態を確認し、下の詳細を参考に再試行してください。","The model rejected the request or returned incomplete data; the document was not changed. Check model capabilities and service status, then retry using the details below." },
            ["ChatBudgetError"] = new[] { "Chat 输入或回答达到 token 限制；未写回文档。输入超限时缩短问题/选区，回答截断时增加回答预算，并按模型实际上限设置上下文预算。","Chat 輸入或回答達到 token 限制；未寫回文件。輸入過長時縮短問題或選取內容，回答截斷時增加回答預算，並按模型實際上限設定內容預算。","Chat 입력 또는 출력이 토큰 제한에 도달했습니다. 문서는 변경되지 않았습니다. 입력을 줄이거나 출력 토큰을 늘리고 모델의 실제 문맥 한도를 설정하세요.","Chatの入力または出力がトークン制限に達しました。文書は変更されていません。入力を短くするか出力トークン数を増やし、モデルの実際の文脈上限を設定してください。","Chat input or output reached a token limit; the document was not changed. Shorten input if too large, increase output tokens if truncated, and set the context budget to the actual model limit." },
            ["ChatConfigError"] = new[] { "Chat 配置无效；未写回文档。请设置有效的 HTTP(S) Base URL 和模型名。","Chat 設定無效；未寫回文件。請設定有效的 HTTP(S) Base URL 和模型名稱。","Chat 설정이 잘못되었습니다. 문서는 변경되지 않았습니다. 유효한 HTTP(S) 기본 URL과 모델을 설정하세요.","Chatの設定が無効です。文書は変更されていません。有効なHTTP(S)のベースURLとモデルを設定してください。","Invalid Chat configuration; the document was not changed. Set a valid HTTP(S) base URL and model." },
            ["ChatRateError"] = new[] { "模型服务限流；未写回文档。请稍后重试或检查 API 速率限制。","模型服務限流；未寫回文件。請稍後重試或檢查 API 速率限制。","서비스 요청 제한에 도달했습니다. 문서는 변경되지 않았습니다. 잠시 후 다시 시도하거나 API 속도 제한을 확인하세요.","サービスのレート制限に達しました。文書は変更されていません。しばらくして再試行するか、APIのレート制限を確認してください。","The service rate limit was reached; the document was not changed. Wait and retry or check API rate limits." },
            ["ChatQuotaError"] = new[] { "API 账号余额或额度不足；未写回文档。请检查服务商账号余额和可用额度。","API 帳戶餘額或配額不足；未寫回文件。請檢查服務商帳戶餘額和可用配額。","API 잔액 또는 할당량이 부족합니다. 문서는 변경되지 않았습니다. 제공업체 계정의 잔액과 할당량을 확인하세요.","APIの残高または割り当てが不足しています。文書は変更されていません。プロバイダーのアカウント残高と割り当てを確認してください。","API balance or quota exhausted; the document was not changed. Check your provider account balance and quota." },
            ["ChatError"] = new[] { "Chat 请求失败；未写回文档。请复制下方错误详情用于排查。","Chat 請求失敗；未寫回文件。請複製下方錯誤詳情以便排查。","Chat 요청에 실패했습니다. 문서는 변경되지 않았습니다. 아래 오류 세부 정보를 복사하여 진단하세요.","Chatのリクエストに失敗しました。文書は変更されていません。下のエラー詳細をコピーして調査してください。","Chat failed; the document was not changed. Copy the error details below for diagnosis." },
            ["ChatReplace"] = new[] { "替换选中","取代選取內容","선택 영역 바꾸기","選択範囲を置換","Replace selection" },
            ["ChatInsertBelow"] = new[] { "插入到下方","插入到下方","아래에 삽입","下に挿入","Insert below" },
            ["ChatInsertStart"] = new[] { "插入到文档开头","插入到文件開頭","문서 시작에 삽입","文書の先頭に挿入","Insert at start" },
            ["ChatWriteConfirm"] = new[] { "将此结果写入发送时的原文档？替换操作会修改原文，请确认。","將此結果寫入傳送時的原文件？取代操作會修改原文，請確認。","결과를 원래 문서에 쓸까요? 바꾸기 작업은 원문을 변경합니다. 확인하세요.","この結果を送信時の元の文書に書き込みますか？置換すると元の文章が変更されます。確認してください。","Write this result to the original document? Replacing changes the original text. Please confirm." },
            ["ChatWritten"] = new[] { "已写入原文档。","已寫入原文件。","원래 문서에 썼습니다.","元の文書に書き込みました。","Written to the original document." },
            ["ChatWriteError"] = new[] { "无法写回：原文档可能已关闭、目标已更改、受保护或超出宿主限制。请复制结果后手动处理。","無法寫回：原文件可能已關閉、目標已變更、受保護或超出宿主限制。請複製結果後手動處理。","원래 문서가 닫혔거나 변경 또는 보호되었거나 호스트 제한을 초과하여 쓸 수 없습니다. 결과를 복사하여 수동으로 처리하세요.","書き込めません。元の文書が閉じられた、対象が変更された、保護されている、またはホストの制限を超えています。結果をコピーして手動で処理してください。","Cannot write back: the original document may be closed, changed, protected, or beyond host limits. Copy the result for manual use." },
            ["ChatReplaceUnavailable"] = new[] { "当前上下文无法整体替换；可插入结果。Excel 替换仅支持单个未合并单元格或单个文本框。","目前內容無法整體取代；可插入結果。Excel 取代僅支援單一未合併儲存格或單一文字方塊。","이 문맥 전체를 바꿀 수 없습니다. 결과를 삽입하세요. Excel에서는 병합되지 않은 단일 셀 또는 단일 텍스트 상자만 바꿀 수 있습니다.","この文脈全体は置換できません。結果を挿入してください。Excelでは結合されていない単一セルまたは単一テキストボックスのみ置換できます。","This context cannot be replaced as a whole; insert the result instead. Excel replacement supports a single unmerged cell or text box." },
            ["ChatQuestionPrompt"] = new[] { "问答系统提示词","問答系統提示詞","질문 시스템 프롬프트","質問のシステムプロンプト","Question system prompt" },
            ["ChatTranslatePrompt"] = new[] { "翻译系统提示词","翻譯系統提示詞","번역 시스템 프롬프트","翻訳のシステムプロンプト","Translation system prompt" },
            ["ChatSummaryPrompt"] = new[] { "总结系统提示词","總結系統提示詞","요약 시스템 프롬프트","要約のシステムプロンプト","Summary system prompt" },
            ["ChatRewritePrompt"] = new[] { "改写系统提示词","改寫系統提示詞","다시 작성 시스템 프롬프트","書き直しのシステムプロンプト","Rewrite system prompt" },
            ["ChatBudget"] = new[] { "模型上下文 token 预算（按模型实际上限设置；保守按 UTF-8 字节估算）","模型內容 token 預算（按模型實際上限設定；保守按 UTF-8 位元組估算）","모델 문맥 토큰 예산 (모델의 실제 한도 사용, UTF-8 바이트로 보수적 추정)","モデルの文脈トークン予算（モデルの実際の上限を使用、UTF-8バイト数で保守的に推定）","Model context token budget (use the model limit; estimated conservatively by UTF-8 bytes)" },
            ["ChatOutputBudget"] = new[] { "回答最大 token 数（不超过上下文预算一半）","回答最大 token 數（不超過內容預算一半）","최대 출력 토큰 (문맥 예산의 절반 이하)","最大出力トークン数（文脈予算の半分以下）","Maximum output tokens (at most half the context budget)" },
            ["ChatInvalidSettings"] = new[] { "请填写三个提示词；回答 token 数不得超过上下文预算的一半。","請填寫三個提示詞；回答 token 數不得超過內容預算的一半。","세 가지 프롬프트를 모두 입력하세요. 출력 토큰은 문맥 예산의 절반 이하여야 합니다.","3つのプロンプトを入力してください。出力トークン数は文脈予算の半分以下にしてください。","Fill in all three prompts. Output tokens must not exceed half the context budget." },
            ["ChatSettingsError"] = new[] { "无法保存 Chat 设置，请检查配置文件权限。","無法儲存 Chat 設定，請檢查設定檔權限。","Chat 설정을 저장할 수 없습니다. 설정 파일 권한을 확인하세요.","Chatの設定を保存できません。設定ファイルの権限を確認してください。","Cannot save Chat settings. Check configuration file permissions." },
            ["ChatPaneError"] = new[] { "无法打开 Chat 侧边栏。请打开文档后重试，并复制下方错误详情用于排查。","無法開啟 Chat 側邊欄。請開啟文件後重試，並複製下方錯誤詳情以便排查。","Chat 창을 열 수 없습니다. 문서를 열고 다시 시도한 다음 아래 오류 세부 정보를 복사하여 진단하세요.","Chatペインを開けません。文書を開いて再試行し、下のエラー詳細をコピーして調査してください。","Cannot open the Chat pane. Open a document and retry, then copy the error details below for diagnosis." },
            ["ChatChangelog"] = new[] { "OfficeTranslate 2.3.3.0\r\n• 每轮完整问答保存本地 JSON 历史；支持只读查看、删除单条及清空全部，活跃会话仍保留最近10轮。\r\n\r\nOfficeTranslate 2.3.2.0\r\n• 三宿主 Chat 改为左右聊天气泡，带头像与时间戳；自适应宽度、长文换行，保留 Markdown、50ms 流式合批和原文导出。\r\n\r\nOfficeTranslate 2.3.1.0\r\n• AI 回答支持离线 Markdown 排版，保留流式合批和导出原文。\r\n• 新增聊天记录导出：Markdown/纯文本、轮次时间戳和多语言提示。\r\n• 修复 Chat 流式闪烁，统一按钮分组和上下文说明，补齐界面及默认提示词多语言，回答语言跟随提示词。\r\n• 修复 DeepSeek Chat 非思考模式请求、SSE 兼容和错误分类，显示安全诊断详情。\r\n• 修复 PowerPoint Chat 侧边栏打开失败，提供侧栏初始化安全错误详情。\r\n\r\nOfficeTranslate 2.3.0.0\r\n• 三宿主新增右侧文档内 Chat：流式问答、总结、改写、长文档分块摘要和最近10轮内存会话。\r\n• 确认后写回原目标；Excel 总结插入当前表顶部一行，PowerPoint 总结新建为第一页。\r\n• 独立 Chat 提示词与 token 预算；沿用现有模型配置。\r\n\r\n","OfficeTranslate 2.3.3.0\r\n• 每輪完整問答儲存本機 JSON 歷史；支援唯讀檢視、刪除單筆及清空全部，活躍會話仍保留最近10輪。\r\n\r\nOfficeTranslate 2.3.2.0\r\n• 三宿主 Chat 改為左右聊天氣泡，附頭像與時間戳記；自動調整寬度及長文換行，保留 Markdown、50ms 串流批次更新與原文匯出。\r\n\r\nOfficeTranslate 2.3.1.0\r\n• AI 回答支援離線 Markdown 排版，保留串流批次更新與匯出原文。\r\n• 新增聊天記錄匯出：Markdown/純文字、對話時間及多語言提示。\r\n• 修復 Chat 串流閃爍，統一按鈕分組及內容說明，補齊介面及預設提示詞的多語言支援，回答語言跟隨提示詞。\r\n• 修复 DeepSeek Chat 非思考模式请求、SSE 兼容和错误分类，显示安全诊断详情。\r\n• 修复 PowerPoint Chat 侧边栏打开失败，提供侧栏初始化安全错误详情。\r\n\r\nOfficeTranslate 2.3.0.0\r\n• 三宿主新增右侧文档内 Chat：流式问答、总结、改写、长文档分块摘要和最近10轮内存会话。\r\n• 确认后写回原目标；Excel 总结插入当前表顶部一行，PowerPoint 总结新建为第一页。\r\n• 独立 Chat 提示词与 token 预算；沿用现有模型配置。\r\n\r\n","OfficeTranslate 2.3.3.0\r\n• 완료된 대화를 로컬 JSON 기록으로 저장합니다. 읽기 전용 보기, 개별 및 전체 삭제를 지원하며 메모리는 최근 10회로 유지합니다.\r\n\r\nOfficeTranslate 2.3.2.0\r\n• 세 호스트의 Chat에 아바타와 시간이 있는 좌우 말풍선을 적용했습니다. 자동 너비와 줄 바꿈, Markdown, 50ms 스트리밍 일괄 갱신 및 원문 내보내기를 지원합니다.\r\n\r\nOfficeTranslate 2.3.1.0\r\n• AI 답변에 오프라인 Markdown 서식을 적용하며 스트리밍 일괄 갱신과 원문 내보내기를 유지합니다.\r\n• 대화 기록을 시간 및 다국어 안내와 함께 Markdown/텍스트로 내보냅니다.\r\n• Chat 스트리밍 깜박임, 버튼 배치 및 문맥 설명을 개선하고 UI와 기본 프롬프트를 현지화했습니다. 답변은 프롬프트 언어를 사용합니다.\r\n• Fix DeepSeek Chat request policy, SSE compatibility and error classification; show safe diagnostic details.\r\n• Fix PowerPoint Chat pane activation and show safe initialization diagnostics.\r\n\r\nOfficeTranslate 2.3.0.0\r\n• Right-docked document Chat in all three hosts: streaming Q&A, summaries, rewrites, chunked long documents and 10 in-memory turns.\r\n• Confirmed writeback to the original target; Excel summaries insert a top row, PowerPoint summaries create a new first slide.\r\n• Separate Chat prompts and token budgets; shared model configuration.\r\n\r\n","OfficeTranslate 2.3.3.0\r\n• 完了した会話をローカルのJSON履歴に保存します。読み取り専用の表示、個別・一括削除に対応し、メモリは直近10往復を維持します。\r\n\r\nOfficeTranslate 2.3.2.0\r\n• 3つのホストのChatを、アバターと時刻付きの左右の吹き出しに変更しました。幅の自動調整と折り返し、Markdown、50msの一括更新、原文のエクスポートを維持します。\r\n\r\nOfficeTranslate 2.3.1.0\r\n• AIの回答をオフラインでMarkdown表示し、ストリーミングの一括更新と原文のエクスポートを維持します。\r\n• 会話履歴を時刻と多言語の案内付きでMarkdown/テキストにエクスポートできます。\r\n• Chatのストリーミングのちらつき、ボタン配置、文脈の説明を改善し、UIと既定プロンプトを各言語に対応しました。回答はプロンプトの言語に従います。\r\n• Fix DeepSeek Chat request policy, SSE compatibility and error classification; show safe diagnostic details.\r\n• Fix PowerPoint Chat pane activation and show safe initialization diagnostics.\r\n\r\nOfficeTranslate 2.3.0.0\r\n• Right-docked document Chat in all three hosts: streaming Q&A, summaries, rewrites, chunked long documents and 10 in-memory turns.\r\n• Confirmed writeback to the original target; Excel summaries insert a top row, PowerPoint summaries create a new first slide.\r\n• Separate Chat prompts and token budgets; shared model configuration.\r\n\r\n","OfficeTranslate 2.3.3.0\r\n• Save completed turns to local JSON history with read-only viewing and individual/all deletion; active memory still keeps 10 turns.\r\n\r\nOfficeTranslate 2.3.2.0\r\n• Chat now uses left/right bubbles with avatars and timestamps in all three hosts. Adaptive widths and wrapping preserve Markdown, 50ms batched streaming and original-source exports.\r\n\r\nOfficeTranslate 2.3.1.0\r\n• Render AI answers as offline Markdown, preserving batched streaming and original-source exports.\r\n• Export conversation history to Markdown/plain text with turn timestamps and localized prompts.\r\n• Smooth Chat streaming, grouped buttons and clearer context guidance; localized UI and default prompts, with answers in the prompt language.\r\n• Fix DeepSeek Chat request policy, SSE compatibility and error classification; show safe diagnostic details.\r\n• Fix PowerPoint Chat pane activation and show safe initialization diagnostics.\r\n\r\nOfficeTranslate 2.3.0.0\r\n• Right-docked document Chat in all three hosts: streaming Q&A, summaries, rewrites, chunked long documents and 10 in-memory turns.\r\n• Confirmed writeback to the original target; Excel summaries insert a top row, PowerPoint summaries create a new first slide.\r\n• Separate Chat prompts and token budgets; shared model configuration.\r\n\r\n" },
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
            if (key == "ChatQuestionDefaultPrompt") return OfficeTranslate.Core.ChatPrompts.Default(language, OfficeTranslate.Core.ChatMode.Question);
            if (key == "ChatTranslateDefaultPrompt") return OfficeTranslate.Core.ChatPrompts.Default(language, OfficeTranslate.Core.ChatMode.Translate);
            if (key == "ChatSummaryDefaultPrompt") return OfficeTranslate.Core.ChatPrompts.Default(language, OfficeTranslate.Core.ChatMode.Summarize);
            if (key == "ChatRewriteDefaultPrompt") return OfficeTranslate.Core.ChatPrompts.Default(language, OfficeTranslate.Core.ChatMode.Rewrite);
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
