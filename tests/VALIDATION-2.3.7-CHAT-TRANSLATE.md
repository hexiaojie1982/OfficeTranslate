# 2.3.7 Chat 翻译按钮验证

日期：2026-10-18。产品版本：2.3.7.0。

## 分支隔离

- 工作树：`D:\pythonProject\office-translate\chat-translate-button`。
- 分支：`feat/chat-translate-button`，从 main `a6bbdfa`（2.3.5）创建。
- 原工作树 `native-word` 仍在 `feat/cjk-space-cleanup` 的 `34023c9`，未修改其文件或分支引用；未引入 2.3.6 的清理实现或测试。
- 未合并、未打 tag、未 push、未发布或安装。

## 实现与三宿主一致性

- `DocumentChat.cs`：新增 Translate 枚举（追加以保留原枚举值）、专用 system prompt、与改写一致的写回权限。
- `ChatPrompts.cs`：五语言默认翻译提示词；翻译的安全约束仍将文档视为数据，回答语言遵从请求目标语言；复用 `AppendTranslationSettings`。
- `TranslationSettings.cs`：新增 `ChatTranslateSystemPrompt`，由现有 SettingsStore JSON 自动保存/读取，旧文件缺省为本地化默认值。
- `ChatErrors.cs`：翻译提示词及其各行参与现有错误脱敏。
- `DocumentChatPane.cs`：蓝色功能组顺序为翻译、总结、改写、Chat 设置；显式分支构造请求 `string.Format(T("ChatTranslateRequest"), settings.TargetLanguage)`，不额外映射目标语言名；发送时重新捕获上下文。
- `AddIn.Chat.cs`：真实宿主初始化改为使用 `LoadForTranslation`，合并最新持久化设置与当前功能区的会话目标语言。直接调用 SettingsStore.Load 会重置目标语言，已用三宿主真实 AddIn/COM 接口代理回归覆盖。
- `DocumentChatPane.cs` 中 ChatSettingsForm 增加独立提示词输入框与保存/空值校验，扩大默认高度并允许滚动；`SettingsForm.cs` 保留新提示词，避免保存翻译服务设置时丢失。
- `UiText.cs`：简中、繁中、英语、韩语、日语的按钮、请求模板、设置标签、默认提示词及 2.3.7 更新日志；`AboutForm.cs` 显示新更新记录。
- Excel/PowerPoint 工程通过 Compile Link 共享 Word 的面板、设置、UiText，以及 AddInShared 接线代码；没有三份独立按钮实现。三宿主程序集均更新为 2.3.7.0，Core 版本随 VERSION.txt。
- `README.md`、`APPLY.md`、`CHANGELOG.md`、`VERSION.txt` 与三处 AssemblyInfo 同步；历史更新记录中的 2.3.5 保留原版本号。

## 回归覆盖

新增 `ChatTranslateTests.cs` 共 44 个数据用例：

| 范围 | 用例数 | 验证内容 |
| --- | ---: | --- |
| 三宿主 × 五语言面板 | 15 | 按钮顺序、重新本地化后仍为主按钮色、320/440/600 宽度不溢出、点击即发送、最新原始目标语言名、选中/全文、流式中禁写、完成后写回按钮、只能写入一次、本地历史 |
| 写回策略 | 4 | 与改写一致、不可插入开头、空结果/无写入委托禁写、全文不替换 |
| 三宿主 × 五语言设置窗 | 15 | 旧 JSON 缺省、提示词输入标签、6000 字多行保存/读取、重开表单、保存翻译服务设置仍保留原值及其他设置 |
| 默认提示词与本地化 | 5 | 三个新增 UiText 键、占位符、本地化默认值随语言切换、自定义值保存后保留 |
| 错误脱敏 | 1 | 翻译自定义提示词按行脱敏 |
| 实际 Chat SSE 传输 | 2 | SSE 成功/中断回退、相同翻译请求、完整历史、最终原文 |
| 长文与设置开关 | 2 | 沿用现有摘要缩减、请求预算、开关控制风格/额外要求/术语表 |

另外新增 11 个现有测试的数据用例：翻译设置开关 2 项、三宿主选中/全文真实适配器 6 项、真实 AddIn 接线取得会话目标语言及最新设置 3 项。合计比 main 的 803 项增加 55 项。

原有全量回归覆盖气泡、逐条复制、Markdown、SSE/回退/取消、历史、导出、Enter/Shift+Enter、多语言、OCR、并行翻译及原有写回流程。翻译专项断言 `中 文 **AI**` 的历史和写回原文不被改动，确保未引入另一个分支的 CJK 空格清理。

## 执行结果

Release 构建：`./build.ps1 -Configuration Release` 成功，**0 错误、6 条既有 nullable 警告**（警告所在文件均未修改）。构建与 WiX 打包日志：`TestResults/release-build-final.log`。

最终全量回归：**858/858 通过，0 失败、0 跳过**，比 main 增加 **55 项数据用例**。执行于已登录桌面 Session 1，耗时约 2 分 29 秒，退出码 0；未删减原有测试或更改其焦点/模态断言。

命令：`dotnet test tests/OfficeTranslate.Tests/OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger "trx;LogFileName=full-regression-interactive.trx" --results-directory TestResults`。

结果证据：`TestResults/full-regression-interactive.trx`、`TestResults/full-regression-interactive.log`、`TestResults/interactive-run-state.json`。首次从 Session 0 运行出现 48 个焦点/模态环境失败（807 通过），已保留 `full-regression.trx`；改在交互会话后全部通过，最终结论依据上述 858 项完整执行结果。临时计划任务已在完成后删除。

安装包：`D:\pythonProject\office-translate\chat-translate-button\artifacts\OfficeTranslate.Office.x64.msi`。

- 字节数：**958464**。
- SHA-256：`3907F1338283685D002189BEB6B651888FF88202E3D4DE0A928A91FE8EE34E94`。
- 与 2.3.5 正式版 SHA-256 `C43658B34CB02948C39DE1A1E92A7905EB67E84519F1ED6CF77459A4EE09CAF0` **不同**。
- Windows Installer 数据库实际 ProductVersion：`2.3.7.0`；6 条加载项/Chat 控件 COM Assembly 注册均为 2.3.7.0。
- 用 WiX dark 解包后，Word、Excel、PowerPoint、Core 四个 DLL 版本均为 2.3.7.0，SHA-256 逐一与本次 Release 输出一致。元数据：`TestResults/msi-final-metadata.json`；解包：`TestResults/msi-final/`。dark 的 DARK1059 为 WiX 标准对话框反编译关系警告，不影响上述实际数据库/载荷核验。

完整改动文件（23 个）：

| 目录 | 文件 |
| --- | --- |
| 根目录 | `VERSION.txt`、`README.md`、`APPLY.md`、`CHANGELOG.md` |
| `src/OfficeTranslate.AddInShared/` | `AddIn.Chat.cs` |
| `src/OfficeTranslate.Core/` | `DocumentChat.cs`、`TranslationSettings.cs`、`ChatPrompts.cs`、`ChatErrors.cs` |
| `src/OfficeTranslate.WordAddIn/` | `DocumentChatPane.cs`、`SettingsForm.cs`、`UiText.cs`、`AboutForm.cs`、`Properties/AssemblyInfo.cs` |
| `src/OfficeTranslate.ExcelAddIn/` | `Properties/AssemblyInfo.cs`（功能由共享源文件同步编译） |
| `src/OfficeTranslate.PowerPointAddIn/` | `Properties/AssemblyInfo.cs`（功能由共享源文件同步编译） |
| `tests/OfficeTranslate.Tests/` | 新增 `ChatTranslateTests.cs`；更新 `ChatPaneActivationTests.cs`、`ChatTranslationSettingsTests.cs`、`ChatUiRegressionTests.cs`、`DocumentChatHostTests.cs`、`DocumentChatTests.cs` |
| `tests/` | 本验证报告 `VALIDATION-2.3.7-CHAT-TRANSLATE.md` |

## 验证边界

- 自动化测试加载三宿主真实 DLL，使用 COM 接口代理、动态宿主模型及本地 HTTP/SSE fixture；未连接真实远程模型，不证明模型译文质量。
- 面板/设置窗已作位图检查，布局与按钮顺序正常；截图位于工作树 `TestResults/chat-translate-2.3.7/`。
- 超长上下文按既有 Chat 行为先分段摘要再请求最终结果，界面提示基于摘要；不保证无损逐段翻译。
- 未执行真实 Office 安装、COM 注册后的真机验收或邮件发送。生产路径未增加正文日志，测试只使用合成 fixture 内容。
