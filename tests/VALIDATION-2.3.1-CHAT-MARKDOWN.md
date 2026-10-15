# 2.3.1 Chat 离线 Markdown 渲染

日期：2026-10-15。直接修改本地 `feat/doc-chat-2.3.0`；按用户要求保持 **2.3.1.0**。保留工作区原有的 DeepSeek、PowerPoint、流式 UI、多语言和聊天导出修复。本轮未提交、push、合并、发布、安装或发送文件；未重新生成 MSI。

## 技术路线与行为

- 内置轻量 Markdown 解析器生成安全 RTF，用原生 `RichTextBox` 显示。无新增 NuGet 依赖，无浏览器、HTML 执行、CDN、JS 或联网渲染组件。Markdown 中的 HTML、图片及链接只显示文字，不访问外部资源；反斜杠、花括号、Unicode 统一转义，输入不能注入 RTF 控制字。
- 支持 `**加粗**`、`*斜体*`、嵌套强调、行内反引号、反引号/波浪线围栏代码块、1–6 级 ATX 标题、有序/无序列表、引用、换行和空段落。代码块中的 Markdown 保持字面文字；围栏未闭合时也渐进显示代码。未闭合的行内强调暂按字面显示，闭合后即时排版。这是覆盖所需语法的轻量子集，不宣称实现完整 CommonMark。
- 正文、提问和代码沿用插件的 Microsoft YaHei UI、9pt；标题加粗放大，代码使用浅灰背景，列表悬挂缩进，引用灰色缩进。用户提问、角色标签和错误诊断始终纯文本。
- 网络回调仍只写入最新快照；原有 50ms UI Timer、取消/完成/迟到回调保护保持不变。渲染比较已显示段落的 RTF，只替换第一个变化段落起的后缀；已完成回答在流式更新期间不重排。仅冻结当前文本控件的 `WM_SETREDRAW`，在 finally 恢复；不使按钮、上下文或根布局失效。
- `RichTextBox` 自动显示滚动条会改变客户区宽度，触发父布局。改为 `ForcedVertical` 提前预留宽度，保留现有“长回答不触发布局”的回归要求。
- 会话历史、`_result` 和 `ChatTranscriptExport` 保持原始 Markdown；渲染仅产生显示副本。导出 Markdown / 纯文本两种格式均按原有格式保存源内容，写回同样保持原文。
- 无新增操作文案；更新说明通过现有 `UiText.ChatChangelog` 补齐五种语言。切换 UI 语言时重建角色标签，同时保留 Markdown 样式和用户提问原文。

## 本轮修改文件

| 文件 | 修改 |
| --- | --- |
| `src/OfficeTranslate.WordAddIn/ChatMarkdown.cs`（新增） | 纯内存解析、Markdown 段落模型和安全 RTF 输出 |
| `src/OfficeTranslate.WordAddIn/ChatTranscriptBox.cs`（新增） | 原生富文本显示、变化段落后缀替换、固定滚动条宽度与局部重绘 |
| `src/OfficeTranslate.WordAddIn/DocumentChatPane.cs` | 接入渲染器；保持合批、最终回复、取消、回退、语言刷新和错误处理 |
| `src/OfficeTranslate.ExcelAddIn/OfficeTranslate.ExcelAddIn.csproj` | 链接共用解析器和富文本控件 |
| `src/OfficeTranslate.PowerPointAddIn/OfficeTranslate.PowerPointAddIn.csproj` | 链接共用解析器和富文本控件 |
| `src/OfficeTranslate.WordAddIn/UiText.cs` | 五语言 2.3.1 Markdown 更新说明 |
| `tests/OfficeTranslate.Tests/ChatMarkdownTests.cs`（新增） | 解析器及原生 RTF 样式、Unicode、注入、增量替换回归 |
| `tests/OfficeTranslate.Tests/ChatUiRegressionTests.cs` | 适配 RichTextBox 的 LF 文本；新增三宿主 Markdown 渐进渲染与原文导出回归 |
| `tests/OfficeTranslate.Tests/DocumentChatHostTests.cs` | 原有宿主回归的控件类型更新为 RichTextBox |
| `tests/OfficeTranslate.Tests/OfficeTranslate.Tests.csproj` | 链接纯解析器和实际渲染控件以检查原生 RTF 行为 |
| `README.md`、本文 | 功能与验证说明 |

本轮未改动 Core 的 ChatTranscriptExport、会话存储、传输、提示词或版本文件；这些文件的既有未提交变更属于前序工作。

## 验证范围

- 新增 28 项回归：22 个数据化语法/边界用例、3 个原生富文本用例、3 个宿主集成用例。检查实际 RichEdit 的粗体/斜体、字体/字号、灰底、缩进和显示文本，避免只验证生成字符串。
- 三宿主后台 600 次快照合批为 1 次 UI 更新，期间父布局与相邻控件失效均为 0；原有回退、最终回复覆盖、实际 Timer、取消和迟到回调测试继续通过。新增测试在请求尚未完成时确认已经显示加粗效果，并显式检查 Timer 为 50ms。
- 三宿主用户 Markdown 标记保留为纯文本；五种语言切换后回答样式保留。Markdown / 纯文本文件实际导出均与 `ChatTranscriptExport.Format` 完全一致，含 `**`、反引号、围栏等原始字符，未输出 RTF/HTML。
- 已目视检查 `TestResults/chat-ui-2.3.1/screenshots/Word-markdown.png`：用户提问保持字面文本，回答标题、强调、列表、引用、灰底代码均已排版。
- 使用真实编译的三宿主 WinForms 侧栏和 STA 消息循环，模型传输及文档上下文为隔离夹具。未安装到真实 Office 做人眼动态验收；静态截图及布局/失效计数不代表所有 Office / DPI 环境的无闪烁人工验收。

## 最终结果与复现命令

- Release：`build.ps1 -Configuration Release -SkipInstaller`，**0 错误，6 条既有 nullable 警告**；新增代码无编译警告。日志：`TestResults/chat-markdown-2.3.1/build-release.log`。未重打包或安装 MSI。
- 专项：**86/86 通过**，包括既有三宿主 Chat UI / 宿主回归与新增 Markdown 回归；日志及 TRX：同目录 `markdown-desktop.log`、`markdown-desktop.trx`。
- 全量：**660/660 通过，0 失败、0 跳过**，约 1 分 27 秒，退出码 0。执行于已登录桌面 Session 1、`UserInteractive=True`：`dotnet test tests/OfficeTranslate.Tests/OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger "trx;LogFileName=all-desktop.trx" --results-directory TestResults/chat-markdown-2.3.1`。
- 全量证据：同目录 `all-desktop.trx`、`all-desktop.log`、`all-desktop.exit`、`desktop-context.txt`；测试均使用项目本地 `DOTNET_CLI_HOME=dotnet-home` 和 `NUGET_PACKAGES=packages`。
- Core、Word、Excel、PowerPoint 最终 DLL 的 AssemblyVersion / FileVersion 均为 **2.3.1.0**，见同目录 `versions.json`；`VERSION.txt` 与三宿主版本属性未因本轮修改而变化。
- `git diff --check` 通过。临时交互桌面测试任务 `Codex-OfficeTranslate-Markdown-Regression` 在完成后删除；测试脚本及日志保留在已忽略的 `TestResults` 目录。
