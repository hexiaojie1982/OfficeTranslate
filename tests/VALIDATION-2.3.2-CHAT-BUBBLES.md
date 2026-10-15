# 2.3.2 文档内 Chat 聊天气泡

日期：2026-10-16。本地分支 `feat/chat-bubbles`，基于干净的 `main`（`b6de418`，与 `v2.3.1` 的文件树相同）。本地验收版本按全局规则递增为 **2.3.2.0**；未 push、合并、创建 tag、发布、安装或发送文件。等待用户在真实 Office 中验收。

## 实现方案

- 三宿主共用原生 WinForms `ChatConversationView`：可滚动的会话视口，每条消息为独立的 `ChatBubbleRow` 和只读 `ChatTranscriptBox`。没有新增 NuGet 包、浏览器或图片资源。
- 用户在右侧，`#2D6CDF` 蓝底白字，问题始终按纯文本处理；AI 在左侧，白底深蓝字，复用现有离线 Markdown 解析器。圆角背景和圆形文字头像由 GDI+ 绘制，新增头像文案统一经 `UiText` 提供五种语言。微软雅黑、面板背景 `#F0F4FA`。
- 气泡先按不超过会话宽度 80% 的可用宽度排版，再通过 RichEdit 的实际行位置与 `ContentsResized` 测量收缩宽度和自然高度；长词、代码、中文、英文自动换行。留出头像、内边距及滚动条空间，每条消息下方显示较小字号的本地 `HH:mm` 时间戳；辅助名称包含完整本地日期/时间。历史采用已有会话时间戳；进行中采用本轮开始时间，成功后同步会话原有的完成时间，不改变导出的时间语义。
- 保留后台快照、50ms UI Timer、最终回复覆盖、回退、取消和迟到回调保护。每批只替换活动 AI 气泡的 RTF，并局部冻结/恢复其 HWND 重绘。为避免 RichEdit 在选择区替换时继承错误的末段格式，活动消息按完整 RTF 更新；已完成消息的控件、内容和尺寸缓存保持不动。气泡尺寸改变只在会话内部布局，不改变按钮或上下文区域。重复快照不重绘。
- 位于底部时跟随新增内容，翻看历史后保持阅读位置；消息正文处的滚轮转给外层会话。语言刷新保留时间戳和 Markdown 样式，清空时释放消息控件及字体。
- `DocumentChatSession`、`ChatTranscriptExport`、传输及写回不变。界面只渲染显示副本，Markdown 和纯文本导出继续使用原始 Markdown。
- 原生自然高度测量采用 Microsoft 文档中的 [EM_REQUESTRESIZE / EN_REQUESTRESIZE](https://learn.microsoft.com/en-us/windows/win32/controls/automatically-resize-rich-edit-controls) 机制。

## 修改文件

| 文件 | 用途 |
| --- | --- |
| `src/OfficeTranslate.WordAddIn/ChatConversationView.cs`（新增） | 气泡、文字头像、时间戳、尺寸测量、内部布局和滚动 |
| `src/OfficeTranslate.WordAddIn/ChatTranscriptBox.cs` | 每条消息的原生富文本渲染和局部重绘；移除旧的整段会话拼接 |
| `src/OfficeTranslate.WordAddIn/ChatMarkdown.cs` | RTF 正文颜色可指定，保留用户原文与代码中的 RTF 字面文本 |
| `src/OfficeTranslate.WordAddIn/DocumentChatPane.cs` | 接入气泡会话、主题色和本轮时间；沿用 50ms 合批及导出 |
| `src/OfficeTranslate.WordAddIn/UiText.cs` | 五语言头像文案及 2.3.2 更新记录 |
| `src/OfficeTranslate.ExcelAddIn/OfficeTranslate.ExcelAddIn.csproj` | 链接共享气泡控件 |
| `src/OfficeTranslate.PowerPointAddIn/OfficeTranslate.PowerPointAddIn.csproj` | 链接共享气泡控件 |
| `VERSION.txt`、三个宿主的 `Properties/AssemblyInfo.cs` | 同步 2.3.2.0；Core/package/MSI 继续从统一版本文件派生 |
| `tests/OfficeTranslate.Tests/ChatUiRegressionTests.cs` | 新增三宿主气泡、尺寸/换行、时间/语言、历史控件/滚动和释放回归；原有样式检查改为定位消息正文 |
| `tests/OfficeTranslate.Tests/ChatMarkdownTests.cs` | 单条消息渲染、末段引用样式、RTF 字面文本回归 |
| `tests/OfficeTranslate.Tests/DocumentChatHostTests.cs` | 适配会话容器、主题色和版本断言 |
| `README.md`、本文 | 功能说明与验证记录 |

## 验证

- Release（含 WiX MSI）：**0 错误，6 条既有 nullable 警告**，退出码 0。
- 专项：**96/96 通过，0 失败、0 跳过**，14 秒，退出码 0。新增 9 项气泡数据化回归和 1 项 RTF 字面文本回归，并扩展已有末段样式和正文背景色断言。
- 最终全量：**670/670 通过，0 失败、0 跳过**，1 分 30 秒，退出码 0。
- Core、Word、Excel、PowerPoint 的 AssemblyVersion / FileVersion、MSI ProductVersion 和 MSI 内四个 DLL 的版本均为 **2.3.2.0**。
- MSI SHA-256：`E7993F92829374D6A06864A26B663DA4DB2371A32B6141C966D228992EB270BB`。
- `git diff --check` 通过；临时交互桌面测试任务已在完成后删除。

使用真实编译的 Word / Excel / PowerPoint 侧栏控件，在已登录的 Windows 桌面 Session 1（`UserInteractive=True`）和 STA 消息循环中执行；文档上下文与模型传输使用隔离夹具，不读取用户凭据，不发送真实模型请求。

覆盖范围：

- 六组宿主/语言组合 × 260/320/440/600px 面板宽度：左右锚点、80% 上限、短消息收缩、长段落/代码/CJK/emoji 换行及末行无裁切，字号、正文颜色及时间位置。
- 三宿主的已完成气泡身份/文本/布局保持，进行中时间稳定，跟随底部与翻看历史的滚动位置，回退重置与清空释放。
- 600 次后台快照合并为 1 次会话更新，UI Timer 确认 50ms；渐进 Markdown 在完成前显示，最终结果、取消及迟到回调、五语言切换后的样式和正文保留。
- 原有 Markdown 语法、HTML/RTF 字面文本、错误诊断纯文本和实际文件导出；导出与 `ChatTranscriptExport.Format` 逐字一致，保留 `**`、反引号和代码围栏，不输出 RTF/HTML。
- `TestResults/chat-bubbles-2.3.2/screenshots/` 共 45 张控件截图，其中 24 张为新增气泡尺寸组合；已目视检查中文 320/440px、英文 600px 及完整 Markdown 对话预览。最终修正 RTF 颜色表：索引 0 保留自动背景，正文使用索引 3，避免同色底块遮字；不替换转义后的字面字符串。

构建：`build.ps1 -Configuration Release`（含 WiX MSI）。测试：`dotnet test tests/OfficeTranslate.Tests/OfficeTranslate.Tests.csproj -c Release --no-build --no-restore`；专项过滤 `ChatMarkdownTests|ChatUiRegressionTests|DocumentChatHostTests`。使用项目本地 `DOTNET_CLI_HOME=dotnet-home` 和 `NUGET_PACKAGES=packages`。

证据目录：`TestResults/chat-bubbles-2.3.2/`。最终以 `build-release.log`、`bubbles-desktop.trx` / `.log` / `.exit`、`all-desktop.trx` / `.log` / `.exit`、`desktop-context.txt`、`versions.json`、`msi-version.txt`、`msi-sha256.txt` 为准；带序号的 TRX 是开发迭代记录。

本地候选 MSI：`artifacts/OfficeTranslate.Office.x64.msi`。未安装到 Office；控件回归及静态截图不等于真实 Word / Excel / PowerPoint 中的动态视觉验收，也不代表已覆盖全部 Office 版本和 DPI。真机应检查多轮短/长消息、流式输出、面板拖动/滚动、语言切换、导出及总结/改写写回入口。
