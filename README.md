# OfficeTranslate for Microsoft Office（原生 COM 加载项）

当前本地验收版本：**2.3.5.0**（`feat/copy-button`，尚未合并或发布）。Word / Excel / PowerPoint 的「OCR」识别按钮检测图片文字区域并转录原文，不执行翻译。PowerPoint 支持选中图片或幻灯片（单页、多页），逐页识别其中的图片；Word / Excel 仍仅处理选中图片。

2.3.5：用户和 AI 聊天气泡右下角常显一键复制按钮，AI 复制 Markdown 原文，用户复制原文；流式输出中和空消息隐藏。复制成功显示对勾及“已复制”，失败显示“复制失败”，约 1.5 秒恢复。提示覆盖简中、繁中、英语、韩语和日语。历史会话同样以只读气泡显示并可逐条复制；剪贴板仅在 UI STA 线程操作，不记录正文。验证见 [复制按钮验证记录](tests/VALIDATION-2.3.5-CHAT-COPY.md)。

Word 一键翻译使用有界并行文本请求，翻译设置中的「Word 文本并发数」默认为 4，可选 1–32，独立于 Excel 并发数。每个现有段落/选区目标为一个提交单位；超长目标继续由原客户端按分块字符数拆分请求，全部子块成功后才提交该目标。结果按原文顺序逐块写回，后面的块先返回时等待前面的块；不增加预览确认。取消保留已提交内容，并等待在途请求退出。单块失败继续处理其他块，在最终结果窗列出失败块序号、阶段和异常类型，保留至手动关闭。图片 OCR 翻译保持原有串行流程。验证见 [并行翻译验证记录](tests/VALIDATION-2.3.4-PARALLEL-TRANSLATE.md)。

## 文档内 Chat（2.3.3.0）

「Chat 设置」下方新增「使用翻译设置（风格/额外要求/术语表）」勾选框，默认关闭，切换后随 Chat 配置保存。开启后每次发送直接读取翻译服务设置的风格、额外要求、术语表，追加到问答、总结、改写及长文分段摘要的 system prompt，不另存一份内容；修改翻译服务设置后下次发送生效。关闭后保持原有 Chat prompt。保存翻译服务或 Chat 设置均保留开关，保存失败恢复原状态并显示安全提示；新增 prompt 内容参与预算计算和错误脱敏，不记录正文日志。验证见 [翻译设置开关验证记录](tests/VALIDATION-2.3.3-CHAT-TRANSLATION-SETTINGS.md)。

2.3.3 正式版：功能组按「总结 / 改写 / Chat 设置」排列，操作组增加「历史会话」。完整问答自动保存至 `%APPDATA%\OfficeTranslate\ChatHistory\<会话ID>.json`，历史对话框按创建时间倒序显示来源应用、文档名与首问预览；来源应用用蓝色 Word、绿色 Excel、橙色 PowerPoint 文字标识，无外部图片。三宿主显式保存 Host，JSON 格式为 SchemaVersion 2；旧记录缺少 Host 时按 `.doc/.docx`、`.xls/.xlsx`、`.ppt/.pptx` 推断，无法推断时留空。点击可只读查看问答原文，支持确认后删除单条或清空全部。活跃会话仅向模型提供最近 10 轮，历史不会自动恢复到模型上下文。「清空」结束当前会话并开始新的会话 ID，保留已存历史。历史正文不写日志、不自动上传；不存文档上下文或模型凭据。保存失败保留未保存轮次，下一轮或「清空」重试，并显示不含正文的错误提示。格式与验证见 [历史会话验证记录](tests/VALIDATION-2.3.3-CHAT-HISTORY.md)。

2.3.2.0 正式版：Word / Excel / PowerPoint 共用左右气泡聊天界面，用户提问为右侧蓝底白字纯文本，AI 回答为左侧白底深蓝 Markdown。每条消息带圆形文字头像与小字时间戳；气泡按内容收缩，最大为会话宽度的 80%，长文本自动换行。沿用 50ms 流式合批，完成的气泡不随 token 重排；翻看历史时不强制跳回底部。界面采用微软雅黑和 `#F0F4FA` 背景，头像文案支持五种 UI 语言。导出继续读取内存中的原始 Markdown。验证见 [聊天气泡验证记录](tests/VALIDATION-2.3.2-CHAT-BUBBLES.md)，发布变更见 [CHANGELOG](CHANGELOG.md)，应用步骤见 [APPLY](APPLY.md)。

2.3.1 修复：Chat 的 DeepSeek 官方请求与翻译保持一致，显式关闭思考模式；补齐 SSE 边界解析、配置超时与分类错误提示。聊天框显示可复制的安全诊断详情。DeepSeek 专用参数不发送给 vLLM；翻译/OCR 链路未修改。验证见 [2.3.1 修复记录](tests/VALIDATION-2.3.1-DEEPSEEK-CHAT.md)。

2.3.1 同时修复 PowerPoint 文档窗口句柄的动态 COM 调用失败；侧栏打开失败时可按 Ctrl+C 复制含阶段、异常类型、消息和关键堆栈方法的脱敏详情。验证见 [PowerPoint 侧栏修复记录](tests/VALIDATION-2.3.1-PPT-CHAT-PANE.md)。

2.3.1 真机反馈修复：流式输出每 50ms 合批，仅更新会话文本尾部；按钮分为蓝色功能组和灰色操作组。Chat 界面与默认系统提示词跟随插件 UI 语言，默认回答使用提示词语言，不跟随文档语言；已有自定义提示词保留。验证见 [Chat 界面与语言修复记录](tests/VALIDATION-2.3.1-CHAT-UI.md)。

三宿主功能区「工具 → 文档内 Chat」打开右侧原生 CustomTaskPane。支持多轮问答、总结、改写、Enter 发送、Shift+Enter 换行、取消和清空；活跃会话在内存中保留最近 10 轮；每次完整问答结束后将全部轮次增量保存至本地历史。退出宿主清理活跃内存，历史文件保留。顶部显示上下文来源和约字符数；进入侧边栏、点击「更新上下文」或发送时更新，不持续轮询大文档。「更新上下文」仅重新读取预览，ToolTip 说明每次发送前也会自动读取最新选中内容或全文。

2.3.1 新增操作组「导出」：通过保存对话框选择 Markdown（默认）或纯文本，导出内存中最近 10 轮完整问答，包含每轮完成时间（含时区）、实际提问和回答。默认文件名 `文档内Chat_yyyyMMdd_HHmmss.md` 随 UI 语言本地化；空会话提示后不打开保存框。导出不重新读取或写回文档，不记录聊天正文日志；取消或保存失败保留会话，错误提示仅显示本地化的安全原因。验证见 [聊天导出记录](tests/VALIDATION-2.3.1-CHAT-EXPORT.md)。

2.3.1 Chat 回答支持离线 Markdown 排版：加粗、斜体、行内代码、浅灰底代码块、标题、有序/无序列表、引用、换行和段落。使用内置轻量解析器和原生 RichTextBox/RTF，沿用微软雅黑与正文字号；每 50ms 渐进刷新，仅替换改变的段落及后续内容，保留局部重绘控制。用户提问按纯文本显示；导出和写回继续使用原始 Markdown，不导出 RTF/HTML。无需 CDN、浏览器或额外运行库，HTML/图片/链接仅作文本。验证见 [Markdown 渲染记录](tests/VALIDATION-2.3.1-CHAT-MARKDOWN.md)。

- 有可读取的选中文字/单元格/形状/幻灯片时使用选中内容，否则使用全文。Word 全文为正文及正文浮动文字图形；Excel 全文为整个工作簿各工作表的已用单元格及文字形状；PowerPoint 全文为全部幻灯片的文字形状/表格。图片像素、批注、页眉页脚、演讲者备注等不作为 Chat 文字上下文。
- 复用原有 TranslationSettings 的 provider、baseUrl、model 和密钥。OpenAI 兼容（包括 vLLM/Qwen）走 SSE；中断/不支持时清除半截输出并重试普通请求，界面提示降级。Ollama 原生接口使用普通请求并提示降级。用户主动取消不重试。
- 「Chat 设置」提供问答、总结、改写三个独立系统提示词，以及模型上下文/回答 token 预算。默认 8192 / 1024，需按实际服务上限设置。采用保守的 UTF-8 字节估算，扣除提示词、问题、最近完整历史轮次和输出预算后动态分块；先逐块摘要，必要时多级汇总，再回答。过程显示块进度，结果注明基于摘要；超长全文改写因此是基于摘要的改写，不能视为逐段无损改写。
- 结果仅显示于聊天窗。改写提供「替换选中」「插入到下方」；全文/多目标等无法整体替换时仅支持插入。Excel 替换只支持单个未合并单元格或单个文字形状；多单元格不把整段改写铺回网格。每次写回均需再次确认，绑定发送时的原窗口/文档/选区，检查选区文字是否被修改；单次结果只允许写入一次。
- 总结的「插入到文档开头」：Word 在正文开头插入；Excel 在发送时当前工作表顶部插入一行并写入 A1；PowerPoint 新建空白页作为第一页。改写「插入到下方」：Word 在选区后/正文末尾插入，Excel 在选区/已用区域下方新增一行写入 A 列，PowerPoint 在原文字选区后追加段落，形状/幻灯片多目标则在原页/最后选中页后新增一页。
- Excel 单元格结果上限 32,767 字符，按纯文本写入，避免将模型输出解释为公式。写回失败可从结果框复制。关闭宿主或卸载加载项会取消请求并清理活跃会话，已保存历史不受影响；新控件的 COM 注册包含在 MSI 中，仅复制 DLL 不足以打开侧边栏。

行为对齐清单、测试证据和待真机验收项目见 [2.3.0 Chat 验证记录](tests/VALIDATION-2.3.0-DOC-CHAT.md)。


Excel / PowerPoint 显式选中 OCR 生成的文本框（包括组合内文本框）后，可使用「选区」翻译。整表、选中幻灯片和全文扫描仍跳过 OCR 覆盖框，避免重复翻译。验证见 [2.2.3 Excel / PowerPoint 验证记录](tests/VALIDATION-2.2.3-EXCEL-PPT-OCR-SELECTION.md)。

OCR 结果窗支持中英文（繁体中文另有译文，韩语/日语暂用英文），成功且无未放置旁注时显示 5 秒关闭倒计时；发生错误、取消或有未放置旁注时保持打开，供用户复制保存。验证见 [2.2.2 验证记录](tests/VALIDATION-2.2.2-OCR-DIALOG-PARITY.md)。
识别结果复用现有不透明覆盖框、固定几何和旁注降级；与图片翻译分别标记和清理。无需打开「图片」翻译开关，双语、源/目标语言、术语表和翻译风格不影响纯 OCR。
选区内没有图片时保持提示「请先选中一张图片」。验证见 [2.2.1 验证记录](tests/VALIDATION-2.2.1-PPT-OCR-SLIDES.md)。

Excel 文本并发度可在 `%LOCALAPPDATA%\OfficeTranslate\settings.json` 设置 `"MaxParallelCells": 4`（整数 1–32，旧配置缺省为 4；1 为串行）。任务开始前修改，下一次加载配置生效；设置窗口保存会保留该值。此配置也适用于 Excel 图形文本，Word、PowerPoint 和图片 OCR 不使用此并发度。进度表示已处理的文本目标数，取消/失败时等待在途任务收尾后报告已写回结果。

面向完全禁止 Office Web Add-in 的 Windows 企业环境。该版本不使用 `manifest.xml`，通过 MSI 注册原生 Word COM 加载项。

Word「选区」翻译支持选中的浮动文本框及带文字的图形，包括 OCR 生成的原文覆盖框；正文与文本框混选时分别写回，文本框保留基础字体、字号和段落格式。验证见 [2.2.3 Word 验证记录](tests/VALIDATION-2.2.3-WORD-SHAPE-TEXT-SELECTION.md)。

## 当前功能

- 翻译并直接替换当前选区。
- 可切换双语模式：启用后，翻译选区或全文时保留原文，并在每段原文下方插入译文。
- 全文翻译按段实时写回：每完成一段立即更新文档，无需等待全部段落翻译完成。
- Excel：翻译选中单元格或当前工作表，跳过公式、数字与空白，逐单元格实时写回；双语模式在同一单元格换行追加译文。
- PowerPoint：翻译选中文字、形状、表格、组合形状、选中幻灯片或整份演示文稿，逐文本框实时写回。
- Word、Excel 和 PowerPoint 支持图片 OCR 翻译，使用配置的视觉模型识别文字，并将可编辑译文文本框覆盖到原文字区域。
- 图片翻译支持双语模式，在覆盖框中同时显示原文和译文；“图片”开关为会话设置，默认关闭。
- 逐段翻译正文并直接替换；按文档顺序写回，写回时通过 Word Range 实时定位，不存在位置漂移。
- Word 功能区快速选择源语言和目标语言，源语言支持自动检测，并可一键反转。
- 紧凑型语言工具栏和统一的自绘线性图标，不依赖不同 Office 版本的内置图标映射。
- Word 每段翻译使用独立的撤销记录，可按段撤销以恢复翻译前内容。
- Ollama `/api/chat`。
- OpenAI-compatible `/chat/completions`，支持 OpenAI、DeepSeek、LM Studio、vLLM、LocalAI 等。
- 根据 Base URL 和 API Key 从 `/api/tags` 或 `/models` 更新模型下拉列表。
- 目标语言、提示词、术语表、分块大小配置。
- API Key 使用当前 Windows 用户的 DPAPI 加密。
- 三个 Office 应用统一使用独立进度窗口；取消当前任务时保留已完成的翻译，不影响其他文档的翻译任务。
- 翻译结果会进行非空与原样返回校验；必要时自动严格重试，任务结束后显示已翻译、缓存命中、跳过、需要检查和 OCR 未识别摘要。
- 翻译结果会校验目标语言的书写系统；中英混合段落翻译为韩语、日语时，如果仍大量保留中文，不会再被误判为成功。
- Word、Excel 和 PowerPoint 文本中的换行符会被原位保护；即使模型忽略换行要求，也会通过一次安全重试恢复原有多行结构。
- 当前 Office 进程内缓存最多 2000 条有效译文；缓存键包含服务、模型、语言、Prompt、术语表和原文，关闭 Office 后自动清空。
- OpenAI-compatible 服务自动尝试稳定的 `prompt_cache_key`，不支持时自动回退；所有文本请求保持稳定 System Prompt 前缀，便于服务端复用 KV Cache。

## 构建环境

构建机需要：

1. Visual Studio 2022 Build Tools，安装“.NET 桌面生成工具”。
2. .NET Framework 4.8 Developer Pack。
3. .NET SDK 8.x。
4. WiX Toolset 3.14；脚本会自动检查 PATH、`WIX` 环境变量和默认安装目录。
5. NuGet 访问权限（首次恢复 `Microsoft.Office.Interop.Word`）。

注意：安装 `.NET Runtime` 不够，`dotnet --list-sdks` 必须至少显示一个 8.x SDK。安装 WiX 后应能在新 PowerShell 窗口中运行 `candle.exe -?` 和 `light.exe -?`。

在 x64 PowerShell 中执行：

```powershell
.\build.ps1 -Configuration Release
```

输出：`artifacts\OfficeTranslate.Office.x64.msi`，统一安装 Word、Excel 和 PowerPoint 加载项。

## 安装与卸载

当前 MSI 面向 **64 位 Office**：

```powershell
msiexec /i OfficeTranslate.Office.x64.msi /qn /l*v install.log
msiexec /x OfficeTranslate.Office.x64.msi /qn /l*v uninstall.log
```

安装后重新启动 Word、Excel 或 PowerPoint，功能区会出现“OfficeTranslate”。首次使用先打开“翻译设置”。

> Windows 位数不等于 Office 位数。若组织仍使用 32 位 Office，需要另行构建 x86 MSI，不能把 x64 MSI 部署到 32 位 Office。

## 企业部署注意事项

- MSI 应使用组织信任的代码签名证书签名，再通过 Intune、Configuration Manager 或组策略分发。
- 可通过防火墙仅允许 `WINWORD.EXE` 访问批准的模型地址。
- 使用 Ollama 时，模型服务应由 IT 统一安装、配置和更新。
- 使用 vLLM 时，应在服务端启用 Automatic Prefix Caching，例如启动参数加入 `--enable-prefix-caching`；客户端无法替代服务端开启该功能。
- 若 Word 将加载项放入“禁用项目”，请先检查事件日志和安装日志，不建议通过策略强制忽略崩溃。

## 当前边界

- 仅处理主文档正文；页眉页脚、脚注、尾注、文本框和批注暂不翻译。
- 全文按 Word 段落翻译。段落内部复杂混合格式可能被统一为该段落首个字符的格式；段落样式和表格结构通常保留。
- 密码保护、只读或受信息权限管理保护的文档无法写回。
- 当前安装定义仅覆盖 64 位 Office，尚未加入签名和自动更新。
