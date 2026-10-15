# OfficeTranslate 2.3.0：文档内 Chat

日期：2026-10-15。基线 `main / f759ec6 / v2.2.3`。本地分支 `feat/doc-chat-2.3.0`。项目四段版本 `2.3.0.0`。未 push、合并、发布或安装候选加载项。

## 实现与新增类型

- `src/OfficeTranslate.Core/DocumentChat.cs`：`ChatMessage`、`ChatContext`、`ChatBudget`、`ChatReply`、`DocumentChatSession`，以及 `ChatMode` / `ChatWriteAction` / `ChatTransport`。统一上下文选择、写回动作许可、token 预算、Unicode 分块、多级 map-reduce、最近 10 个完整对话轮次。历史不含文档上下文消息，不写入文件；清空/取消不会复活在途会话。
- `src/OfficeTranslate.Core/TranslationClient.Chat.cs`：`ChatSseParser` 和 `TranslationClient.ChatAsync`。按 OpenAI SSE 事件增量显示；检查 DONE、空响应和结束原因；流失败/不支持时清掉半截输出，普通请求恢复。主动取消不重试；恢复结果仍截断或被过滤则失败，不提供写回。每次尝试都有超时和取消绑定，不显示或记录服务原始错误体。
- `src/OfficeTranslate.Core/TranslationSettings.cs`：独立的问答/总结/改写系统提示词，`ChatContextTokens=8192`、`ChatOutputTokens=1024`。复用现有 provider、baseUrl、model、API key。
- `src/OfficeTranslate.WordAddIn/DocumentChatPane.cs`：COM 可见的 `DocumentChatPane` 和 `ChatSettingsForm`。作为共享源码编入三个宿主，分别使用独立 CLSID / ProgID。显式使用现有 `OfficeUiDispatcher` 把所有异步界面更新送回 STA，不信任 Office 环境的 SynchronizationContext。
- `src/OfficeTranslate.WordAddIn/DocumentChatHost.cs`：`DocumentChatHost`，按宿主编译上下文提取和写回。上下文/目标绑定发送时的原窗口及文档；替换前检查原文字、Excel 公式和值。Excel 输出以纯文本存储；拒绝超过单元格 32,767 字符的结果。
- `src/OfficeTranslate.AddInShared/AddIn.Chat.cs`：三个 AddIn 的共享 partial，实现 `ICustomTaskPaneConsumer`、工厂回调、按窗口创建/复用右停靠原生 CustomTaskPane，以及宿主退出/窗口关闭后的清理。使用 HWND 复用检查，避免误用关闭窗口的旧面板。
- 入口与配置：两个 `AddIn.cs`、共享 `Ribbon.xml`、`UiText.cs`、`SettingsForm.cs`、三个宿主 csproj。Chat 文案统一通过 UiText；中英文齐全，繁中暂沿用简中文案，韩文/日文沿用英文。原设置窗保存时保留独立 Chat 配置。
- 版本/注册：README、VERSION、build.ps1、三个 AssemblyInfo、About 更新日志、`installer/Product.wxs`。MSI 包含三个 Chat 控件的独立 COM 注册；仅替换 DLL 不足以激活新控件。
- 新增测试：`tests/OfficeTranslate.Tests/DocumentChatTests.cs`、`DocumentChatHostTests.cs`。本验证记录及 README 使用说明。

## 老功能可观测行为对齐清单

| 观察项 | TranslationResultForm / OcrResultForm | Chat 的对应行为与验证 |
| --- | --- | --- |
| 字体与卡片 | Microsoft YaHei UI，约 9pt，浅灰蓝底/白色正文 | 同字体、9pt、#F7F9FC、#203758 标题；三宿主控件断言和中英文 PNG 人工检查 |
| 主按钮 | Flat 蓝底白字 | #2D6CDF、无边框、白字；断言样式，窄侧栏自动换行 |
| 文案语言 | UiText 根据任务语言取值 | 所有新增可见文案通过 UiText，中文/英文状态及按钮覆盖 |
| 正常完成 | 明确完成状态，保留可读结果 | 最近 10 轮结果可复制；总结/改写展示相应写回操作，尚未生成完整结果时不展示写回按钮 |
| 空状态 | 无文字/无可用目标有说明 | 初始使用提示、空问题提示、无文档/无文字提示；不发送空请求 |
| 加载/取消 | 进度提示与取消 | 生成状态、分块 i/n 进度、取消按钮；同一面板禁止并行重复请求 |
| 错误/中断 | 保留可恢复内容供查看 | 原有会话保留，未完成回答不入历史、不允许写回；原始服务错误内容不显示 |
| 降级 | OCR 等有明确降级提示 | SSE 中断时清掉半截，提示正在恢复；恢复成功说明普通请求降级，map 阶段发生的降级也汇总提示 |
| 写回确认 | 结果展示不等于自动接受 | 按钮点击后再次确认，拒绝确认不写；确认一次后禁用本次写回，原目标变动则报错 |
| 无法替换 | 无目标不能写入 | 全文/多目标改写保留“替换选中”按钮但禁用，并说明可插入；Excel 多格/合并格不可整体替换 |
| 关闭生命周期 | 结果窗完成态有倒计时；错误态保留 | Chat 是持续会话侧栏，不使用结果窗的自动关闭倒计时；清空按钮、关闭文档/退出宿主清理内存 |

现有 `TranslationResultForm.cs`、`OcrResultForm.cs`、三宿主原翻译/OCR 服务、原 TranslationClient.cs / Recognition.cs 与 main 的比较无差异。全量旧测试继续执行，不以 Chat 的 happy path 替代旧行为验证。

## 写回含义与边界

- Word：选区后插入，未选文字时正文末尾插入；总结在正文开头插入。正文浮动文字形状也参与上下文。单个可写文字选区支持替换，多图形/组不做整组替换。
- Excel：全文读取整个工作簿；总结在发送时当前工作表顶部新增一行，写入 A1。改写插在原选区/已用区域下方新增一行的 A 列。替换仅限单个未合并单元格或单个文字形状，避免把一段改写误铺回多个单元格。
- PowerPoint：总结新建空白页作为第一页。文字选区的改写在原 TextRange 后追加段落；幻灯片/多形状上下文在原页或最后选中页后新建一页。新页结果按文本框尺寸自动适配文字。
- 长文档：采用保守的 UTF-8 字节上界估算而非某个模型的真实 tokenizer；扣除提示词、问题、可容纳的最近完整历史轮次及输出预算后动态分块，摘要未收敛则明确失败。最终状态说明基于摘要，全文改写不是逐段无损改写。
- “全文”指宿主可读取的主体文字：Word 正文及正文浮动文字形状；Excel 已用单元格/文字形状；PowerPoint 幻灯片文字形状及表格。不包含图片像素、批注、Word 页眉页脚、演讲者备注、嵌入对象内部文字等。

## 验证结果

1. `build.ps1 -Configuration Release` 成功：**0 错误**；7 条警告均位于原有代码/测试，新增 Chat 文件无编译警告。WiX 3.14 编译/链接成功，生成本地 MSI。
2. 交互桌面全量测试 **530/530 通过，0 失败，0 跳过**；其中新增 Chat 回归 **44/44** 通过。测试运行环境 Session 1 / UserInteractive=True。覆盖上下文选择和实际宿主适配方法、Unicode 分块/递归摘要、预算与提示词、10 轮历史/清空竞态、SSE/普通降级/取消/截断/过滤、写回动作许可、确认拒绝/确认一次、空状态/加载/错误状态、真实编译控件样式，以及不可信同步上下文下的跨线程回调。
3. Core + 三宿主 DLL 的 AssemblyVersion / FileVersion、三个 About.CurrentVersion、MSI ProductVersion、MSI 中四个产品 DLL 的版本，以及六项 COM Assembly 注册值，均为 **2.3.0.0**。
4. 版本全扫及 `git diff --check` 通过；旧版引用仅保留在历史说明/链接与历史更新日志。旧翻译/OCR 业务链路与 main 比较无差异。三张中英文侧栏渲染图已经查看，未见按钮或主体文字裁切。
5. 本地 MSI：`artifacts/OfficeTranslate.Office.x64.msi`；SHA-256：`48B119140F504D60BA598A1ACE21A4B56B88570AA7ED0A3A36939D98BEC2FC28`。

复现命令：

```powershell
.\build.ps1 -Configuration Release
# 已登录的交互桌面（当前工具 Session 0 不作为剪贴板/窗体验收环境）：
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=all-desktop.trx' --results-directory TestResults\doc-chat-2.3.0
```

证据位于 git 忽略目录 `TestResults/doc-chat-2.3.0/`：`build.log`、`chat.trx`、`all-desktop.trx`、`desktop-test.log`、`desktop-test.exit`、`desktop-context.txt`、`version-verification.log`、`version-scan.log`、`frozen-paths.log`、三张 `chat-*.png`。测试使用临时文档替身、模型 HTTP 处理器和隔离的 Chat 设置，不读取/发送真实文档或连接用户真实模型。

## 尚未覆盖的真实场景

- 未安装候选 MSI，尚未在真实 Word / Excel / PowerPoint 中人工点击 Ribbon、激活新控件、拖动/关闭/重新打开侧栏，或验证多窗口/不同 DPI/不同 Office 版本下的 COM 激活与生命周期。测试断言了编译后的真实宿主 DLL 中的入口、接口、ProgID，以及使用实际适配方法的 Office 对象替身。
- 未连接真实 vLLM/Qwen/Ollama 服务，尚未测真实 tokenizer、网络代理/断线/慢流、模型思考输出和回答质量。SSE/降级/取消/截断由可控 HTTP 响应验证；Ollama 原生接口使用普通返回并明确提示降级。
- 受保护文档/工作表、复杂混合格式、合并表格、巨量或被过度格式化的 UsedRange、大型组合/SmartArt/图表、撤销体验及新建 PowerPoint 页的实际排版需真机验证；不宣称已测吞吐量或所有 Office 对象的文字召回率。

原生 COM 工厂接入参考：[Microsoft ICustomTaskPaneConsumer](https://learn.microsoft.com/en-us/office/vba/api/office.icustomtaskpaneconsumer)。


测试临时计划任务已删除；最终测试退出码 0。
