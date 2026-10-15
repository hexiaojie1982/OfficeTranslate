# 文档内 Chat 导出验证（2.3.1.0）

本次仅在本地 `feat/doc-chat-2.3.0` 分支增量修改；保留原有未提交的 Chat、DeepSeek、侧栏及多语言修复。按本次明确要求保持 `2.3.1.0`，不 push、不合并、不发布、不安装。

## 行为

- Word、Excel、PowerPoint 共用 Chat 侧栏：灰色操作组增加「导出」，与清空、更新上下文、设置沿用相同按钮样式；窄幅自动换行。
- 使用原生 `SaveFileDialog`，提供 Markdown / 纯文本过滤器，默认 Markdown、`.md` 扩展名及 `文档内Chat_yyyyMMdd_HHmmss.md`；文件名模板、标题、过滤器随 UI 语言变化，保留覆盖确认。
- 只取 `DocumentChatSession.History` 的不可变消息快照，导出当前最多 10 轮完整问答；模型请求预算裁剪历史不影响内存导出。时间记录在完整回答成功加入会话时，同轮问答使用相同时间，含时区。
- Markdown 使用文件标题、时间标题、问答子标题和逐行引用；纯文本使用分隔线与问答标签。实际内容保留 Unicode、空行及多行结构；UTF-8（含 BOM）保存。
- 空会话和清空后仅显示提示，不打开保存框；取消不写文件。失败提示只包含本地化错误类别（权限、磁盘/占用/位置、无效路径或未知原因），不显示原始异常、路径、密钥或聊天内容。
- 不读取/写回 Office 文档，不调用模型，不记录聊天正文日志；按钮沿用操作组在生成期间禁用的行为，未完成/取消/失败的回答不属于可导出的完整轮次。
- 新增 15 个 UiText 键，简中、繁中、韩语、日语和英语均补齐，包括文件名、文件头、时间/问答标签、按钮和所有状态提示；更新记录的导出说明也覆盖五种语言。

## 本次文件

- `src/OfficeTranslate.Core/DocumentChat.cs`：内存消息时间戳，完整问答提交时记录同一时间。
- `src/OfficeTranslate.Core/ChatTranscriptExport.cs`（新增）：纯格式化，无文档、文件或网络访问。
- `src/OfficeTranslate.WordAddIn/DocumentChatPane.cs`：操作按钮、保存对话框、只读快照导出和安全状态提示；通过既有链接编译覆盖三宿主。
- `src/OfficeTranslate.WordAddIn/UiText.cs`：五语言导出文案及更新记录。
- `tests/OfficeTranslate.Tests/ChatTranscriptExportTests.cs`（新增）：10 个数据化回归。
- `tests/OfficeTranslate.Tests/ChatUiRegressionTests.cs`：增加 15 个三宿主导出回归，复用真实控件和 STA 消息循环。
- `README.md`、本文：功能与验证说明。

## 验证

- Release：`build.ps1 -Configuration Release -SkipInstaller`，**0 错误，6 条现有 nullable 警告**；新增代码无编译警告。最终日志：`TestResults/chat-export-2.3.1/build-final.log`。本次不重新生成 MSI。
- 全量：已登录桌面 Session 1、`UserInteractive=True` 下执行 `dotnet test tests/OfficeTranslate.Tests/OfficeTranslate.Tests.csproj -c Release --no-build --no-restore`；**632/632 通过，0 失败、0 跳过**，包含本次新增 25 项。最终证据：`TestResults/chat-export-2.3.1/all-desktop.trx`、`desktop-test.log`、`desktop-test.exit`。
- 回归覆盖：中英两种格式的精确输出（标题、引用、分隔线、时间、时区、多行、空行、Unicode）；全部五语言的导出键、过滤器和默认文件名；最近 10 轮、请求裁剪不丢内存记录、清空和快照稳定性；三宿主两种实际文件导出、保存框配置和切换语言；空会话、清空后提示、取消、四类安全错误、实际文件系统写入失败；导出不增加文档读取、不调用写回、不更改会话。
- 既有按钮布局回归覆盖三宿主中英文 320/440/600px，全部通过。目视检查更新后的 Word 中英文 320px 控件截图：导出按钮未裁切，设置按钮自然换行；截图在 `TestResults/chat-ui-2.3.1/screenshots/`。
- Core、Word、Excel、PowerPoint 最终 DLL 的 AssemblyVersion/FileVersion 均为 `2.3.1.0`，见 `TestResults/chat-export-2.3.1/versions.json`；`git diff --check` 通过。

测试使用实际三宿主编译的 WinForms 控件，模拟模型传输和文档上下文，不读取用户凭据或发送真实模型请求。保存对话框配置通过注入的选择回调验证，文件输出使用真实临时文件；未在真实 Office 内人工操作系统保存对话框，未安装候选 DLL/MSI。
