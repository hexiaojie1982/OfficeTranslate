# 2.3.1 Chat 界面与语言修复

日期：2026-10-15。直接修改本地 `feat/doc-chat-2.3.0`，按本次要求保持 **2.3.1.0**。保留工作区已有的 DeepSeek 和 PowerPoint 修复。本轮未提交、push、合并、发布、发送或安装候选包。

## 五项问题的根因与修改

1. **流式闪烁**：原实现每次收到累计输出都同步调回 Office STA，重新赋值 `_transcript.Text = _prefix + text` 并滚动，会不断重置整个会话文本。改为线程安全的最新输出快照；50ms 的 WinForms Timer 在 UI 线程批量处理，正常流只追加新增部分，回退请求只替换当前回答尾部。文本框明确关闭 AutoSize，更新时只对该文本框使用 `WM_SETREDRAW` 并在 finally 恢复，仅使文本框失效；不对根布局、按钮、上下文做每 token 更新。完成、取消、异常、Stop/Dispose 均停止 Timer；不接收已结束请求的迟到回调。完成以最终回复为准，不依赖最后一次 token 回调。
2. **按钮布局**：原来所有按钮在同一个 FlowLayoutPanel 中自然挤排，使用 AutoSize 和相同蓝色，尺寸随文字变化且无功能区分。改为功能组（总结、改写）和操作组（清空、更新上下文、Chat 设置）两行；发送/取消在输入区下方。统一微软雅黑、32px 高、8px 间距、Flat 无边框，主按钮 `#2D6CDF`，次按钮灰色。按文字测量宽度，窄面板在组内换行。
3. **上下文按钮含义**：`RefreshContext` 只重新抓取当前选区/全文，更新来源和字符数预览；请求从不使用这个缓存，每次发送都会重新 Capture。原来在控件进入和显示时刷新，并未订阅持续的选区变化。保留手动预览操作，改名「更新上下文」，ToolTip 明确说明重新读取当前选中内容或全文，以及发送前仍会自动读取；上下文状态本身也有说明。不增加轮询大文档。
4. **界面多语言**：Ribbon 标签原本已有 GetRibbonLabel，但侧栏语言只在 Initialize 时保存一次；已打开的控件不会随插件语言更新，部分语言资源为简中或英文兜底。补齐简中、繁中、韩、日、英的 Chat 界面资源，标题、按钮、状态、输入提示、ToolTip、确认和错误提示统一经 `UiText.Get`。设置保存/重开侧栏后刷新已存在控件，保留会话和写回绑定；成功、取消、失败时的角色标签及错误详情标题也可重新本地化。Ribbon 新增本地化 ScreenTip。Office 的 CustomTaskPane.Title 只读，外框标题通过 UiText 使用产品名 OfficeTranslate，避免语言切换时销毁会话；内部 Chat 标题随语言更新。异常类型、HRESULT、HTTP 状态、阶段标识及脱敏诊断原文属于技术详情，保留可诊断信息。
5. **系统提示词语言**：三个旧默认值固定为英文，改写提示还要求保留原文语言；请求又附加了「按用户语言回答」，长文档摘要要求使用原文语言。这几条相互冲突。新增 Core `ChatPrompts` 作为 UI 与请求共享的默认提示词来源；默认值随 UiLanguage 取，中文明确「请用中文回答」，英文明确「Please answer in English」，其他 UI 语言同样补齐。删除跟随用户/文档语言的旧规则，安全后缀明确使用系统提示词语言；长文档中间摘要也使用 UI 语言。识别旧配置中序列化的三个内置英文值以及当前内置默认值，切换语言时自动解析为对应版本；自定义文本保留。已有 `ChatQuestionPrompt/ChatSummaryPrompt/ChatRewritePrompt` 是设置框标签，继续通过 UiText 获取；另提供三个 `Chat*DefaultPrompt` 键，委托同一 Core 默认提示词资源，避免 UI 与实际请求产生两套文案。

## 本轮修改文件

- `src/OfficeTranslate.WordAddIn/DocumentChatPane.cs`：三宿主共享的流式刷新、布局、ToolTip、输入提示、语言刷新和会话显示。
- `src/OfficeTranslate.WordAddIn/UiText.cs`：Chat 本地化资源、默认提示词资源入口和 2.3.1 更新记录。
- `src/OfficeTranslate.WordAddIn/Ribbon.xml`：三宿主共用的 Chat ScreenTip 回调。
- `src/OfficeTranslate.WordAddIn/AddIn.cs`、`src/OfficeTranslate.AddInShared/AddIn.cs`：重新加载持久化设置后刷新 Chat 语言。
- `src/OfficeTranslate.AddInShared/AddIn.Chat.cs`：Chat 入口同步设置、语言刷新、外框标题与诊断阶段本地化、Ribbon ToolTip。
- `src/OfficeTranslate.Core/ChatPrompts.cs`（新增）：多语言默认提示词、旧默认值兼容及安全/中间摘要提示。
- `src/OfficeTranslate.Core/TranslationSettings.cs`：语言感知的默认值，保留自定义提示词。
- `src/OfficeTranslate.Core/DocumentChat.cs`：移除跟随原文/用户语言的冲突规则。
- `tests/OfficeTranslate.Tests/ChatUiRegressionTests.cs`（新增）：24 项数据化回归，使用真实编译的三宿主控件，隔离设置文件与模拟传输。
- `tests/OfficeTranslate.Tests/ChatPaneActivationTests.cs`：原有技术诊断断言显式使用英文 UI，避免依赖默认语言。
- `README.md`、本文：行为说明及验证记录。

工作区中原有的版本、DeepSeek、PowerPoint 激活修复及其测试不是本轮新增改动；本轮未提高版本号。

## 验证与证据

- 最终结果：**607/607 通过，0 失败、0 跳过**，包含本轮新增 24 项回归；全量用时约 1 分 25 秒，退出码 0。Release/WiX 完成，**0 错误，6 条原有 nullable 警告**，新增代码无编译警告。
- 三宿主流式回归均记录：600 次后台快照合并成 1 次文本更新；控件及根布局事件为 0，根布局/按钮区/上下文区失效事件为 0（正常流阶段）。
- Release：`build.ps1 -Configuration Release`，含 WiX 本地 MSI 打包；构建日志位于 `TestResults/chat-ui-2.3.1/build-release.log`。
- 全量：在已登录桌面 Session 1（UserInteractive=True）运行 `dotnet test tests/OfficeTranslate.Tests/OfficeTranslate.Tests.csproj -c Release --no-build --no-restore`。以同目录最终 `all-desktop.trx`、`desktop-test.log`、`desktop-test.exit` 为准。
- 回归：三宿主的 600 次后台输出合批、增量追加、Unicode/换行、回退尾部替换、最终回复覆盖、Timer 实际刷新、取消/迟到回调；按钮分组、32px 高度、样式和 320/440/600px 宽度边界；手动预览与发送前重新读取；五种 UI 语言的现有侧栏、Ribbon、成功/取消/失败状态；三模式默认值、旧配置恢复、自定义保留、长文档中间摘要与设置对话框。
- 18 张中英文控件截图：`TestResults/chat-ui-2.3.1/screenshots/`；已目视检查中文 320px、英文 320/440px，窄幅可换行，无按钮裁切。
- Core 与三个宿主的 AssemblyVersion/FileVersion，以及 MSI ProductVersion/包内四个 DLL 的版本均核对为 **2.3.1.0**。
- `git diff --check`；临时交互桌面测试任务结束后删除。

本地候选：`artifacts/OfficeTranslate.Office.x64.msi`；SHA-256：`5030AAB29AF20BE73C1CBE3AFEE31DC20162555D3C9823B4FFF131BB10ADA783`。未安装、发送或发布。

测试使用实际 WinForms 侧栏、真实 UI 线程与消息循环，传输和文档上下文使用隔离夹具，不读取用户凭据、不发送真实模型请求。本轮未把候选 MSI 安装到 Office，也未执行真实服务速度下的 Office 人眼动态验收；因此不把合批/布局计数或静态截图表述为所有 Office/DPI 环境均已通过无闪烁人工验收。
