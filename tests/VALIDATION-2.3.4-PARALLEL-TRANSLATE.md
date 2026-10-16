# Word 并行翻译与按序增量写回验证

日期：2026-10-16。基线：`main` / `7f21359`，已发布版本 2.3.3。
本地分支：`feat/parallel-translate`；本地构建版本：2.3.4.0。
以下记录对应合并前的功能验证；验证当时未 push、合并、创建 tag、发布、安装或生成新 MSI。正式版沿用本次功能与测试代码，版本为 2.3.4.0（`v2.3.4`）；合并后的全量测试与完整 MSI 构建结果另见发布交付报告。

## 实现与兼容边界

- 使用现有 Word 段落/选区/文字图形目标及其源文本快照，一个目标为一个提交单位。新增 Core 调度器 `WordTextTranslation`，不访问 COM。
- `MaxParallelTextBlocks` 默认 4，范围 1–32，保存于共享翻译设置；旧 JSON 缺少该字段时默认为 4。独立于 Excel 的 `MaxParallelCells`。共用的翻译设置窗以「Word 文本并发数」显示该选项。
- 有界滚动窗口最多保留指定数量的任务，限制在途请求及待写回结果数量；HTTP 文本传输的每服务器连接上限同步设置，避免 .NET Framework 默认连接限制把并行请求串行化。仍使用原客户端的缓存、源语言保护、完整性校验、重试、超长分块及段落回退。
- 单一消费者按目标原顺序等待结果，逐块通过 `OfficeUiDispatcher` 写入原目标。后块先完成时等待前块；无需全篇完成或预览确认。COM、UndoRecord、summary 和进度回调均在同一 Office STA 队列执行。
- `Changed=false` 不写回。保留不可翻译文本跳过、缓存命中、待检查及段落回退等原有统计；每个已处理目标恰好一次开始和一次完成回调，Changed=true 时增加一次写回回调。图片数量继续作为正文进度的偏移量。
- 取消时停止新增任务，通过关联 token 取消在途请求，并等待所有已启动任务退出；UI 队列提交前再次检查取消。已开始的同步块提交完成并记账，已写回内容保留。超长目标全部子块成功后才提交，子块失败不会提交该目标的半截译文。
- 单块翻译或写回失败后继续其他块，记录失败块序号、阶段、异常类型并增加待检查统计。最终结果窗使用警示状态，列出失败块并保持打开供复制；中途取消且已有错误时也保留错误汇总。错误信息只保存异常类型，不保留可能含正文或密钥的服务端回显，不新增正文日志。
- 图片预捕获和图片 OCR 翻译循环保持原样，先串行处理图片，再启动正文并行流程。Excel 和 PowerPoint 翻译调度不变；共享设置/结果窗兼容原有行为。

## 自动化验证

Windows 交互桌面 Session 1 执行测试，`UserInteractive=True`。最初在非交互会话执行原始基线时，43 项窗口/键盘测试受桌面限制失败；切换交互会话后，原始基线全部 **740/740** 通过，无需修改或跳过这些测试。

最终运行：**768/768 通过，失败 0，跳过 0**，耗时约 1 分 43 秒。测试前后核对全部源码和测试代码 SHA-256，无变化。原有测试同步宿主程序集版本断言到 2.3.4.0，并将并行 HTTP 传输测试扩展为 Excel / Word 两个参数化用例；原有行为断言保留。

新增 `WordParallelTranslationTests.cs`：**27 个可执行用例**（包含参数化用例）；既有 `ParallelTranslationTransportTests.cs` 额外增加 Word 工厂用例 1 项，合计比基线增加 **28 项**。

| 覆盖 | 结果 |
| --- | --- |
| 并发数 1、4、32；故意反向返回；有界窗口；按序且逐块可见 | 通过 |
| 并发数 1 的请求/写回/完成事件顺序与原串行流程一致 | 通过 |
| 第一块提交后取消，保留已写回部分，取消在途任务，停止新增任务 | 通过 |
| 写回排队期间取消，忽略 token 的请求也等待收尾，禁止迟到写入 | 通过 |
| 提交内部收到取消，当前同步提交完成并统计，再停止后续写回 | 通过 |
| 开始进度回调中取消，不再发送请求；进度回调异常后等待任务收尾，保留原始错误 | 通过 |
| 翻译异常、写回异常、非用户取消的超时；其他块继续；错误序号/阶段/数量正确 | 通过 |
| Changed、缓存、跳过、部分译文、段落回退；进度次数、内容、图片偏移正确 | 通过 |
| 空目标、无可翻译文字、预先取消 | 通过 |
| 超长单块经原客户端拆成 500/500/500/100 字符；全部成功一次提交；第二子块失败不写入 | 通过 |
| 默认值、旧配置兼容、JSON 往返、非法并发数拒绝 | 通过 |
| 三宿主共享设置窗读取、编辑、持久化；Excel 并发数保持独立 | 通过 |
| 五种语言错误汇总可见、可复制、无自动关闭；取消后仍可查看；无服务端正文回显 | 通过 |
| Excel / Word 工厂分别通过本机 HTTP 服务验证 4 个真实连接并发、严格重试、共享缓存 | 通过 |

命令：

```powershell
$env:DOTNET_CLI_HOME = Join-Path $PWD 'dotnet-home'
$env:NUGET_PACKAGES = Join-Path $PWD 'packages'
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-restore --logger 'trx;LogFileName=parallel-verified-all.trx' --results-directory TestResults\parallel-verified-all
.\build.ps1 -Configuration Release -SkipInstaller
```

Release 构建 **0 错误**。本次最终增量构建为 0 警告；前面的重新编译仍有仓库原有的 7 条可空性警告，新增代码无编译警告。
从磁盘读取 Core、Word、Excel、PowerPoint 四个最终 DLL 的 AssemblyName，版本均为 **2.3.4.0**。MSI 版本参数继续由 `VERSION.txt` 派生；本任务未生成或安装 MSI。

本机证据：

- 基线 TRX：`TestResults/parallel-baseline-interactive/baseline.trx`
- 最终全量 TRX：`TestResults/parallel-verified-all/parallel-verified-all.trx`
- Release 构建日志：`TestResults/parallel-release-build.log`
- 复核全量 TRX：`TestResults/parallel-completion/all.trx`，**768/768**，退出码 `all.exit`；交互会话记录 `session.txt`。
- Release 全部重编：`TestResults/parallel-completion/release-build.log`（0 错误、7 条既有可空性警告）；版本校验及增量构建 `version-build.log`（0 错误、0 警告）。
- 设置窗截图：`TestResults/parallel-settings.png`；已检查默认值 4、新行及原有设置控件无重叠或裁切。

## 真实 Word Range 补查

通过新建的独立 Word COM 实例和临时文档，调用本次编译 DLL 的真实 `ReadDocumentParagraphs` / `WriteTarget`。覆盖 **正文和表格 × 替换和双语插入，共 4 个场景、12 次提交全部通过**。每场景有 3 个目标，第一段替换为明显更长的文本，每次提交后核对全部目标；后续目标未错位，表格单元格边界保持正确。探针仅记录版本、MVID、计数和结果，不输出正文。

测试实例与临时文档已关闭，不保存、不修改用户文档。四场景证据：`TestResults/parallel-completion/range-result.json`；可重跑脚本已纳入仓库：`tests/RunWordParallelRangeProbe.ps1`，在交互桌面的 Windows PowerShell 中以 `-STA` 执行。最初探针显式写入结尾回车，使 Word 额外保留一个空段，导致期望长度不符；修正测试夹具为由 Word 自动提供终段标记后，四场景均通过，无需修改生产写回逻辑。

该补查验证真实 Range 随前方文本长度变化的定位行为；乱序网络响应、取消和失败隔离由上述真实 STA 队列上的自动化测试验证。未连接外部模型做一键翻译端到端验收，也未测量实际 API 加速倍数。

## 修改文件清单

| 文件 | 用途 |
| --- | --- |
| `src/OfficeTranslate.Core/WordTextTranslation.cs`（新增） | 有界并行、顺序提交、取消收尾、单块错误隔离 |
| `src/OfficeTranslate.Core/TranslationClient.cs` | 共用有连接上限的并行文本客户端工厂，保留 Excel 入口 |
| `src/OfficeTranslate.Core/TranslationSettings.cs` | 新设置项、默认值和范围校验 |
| `src/OfficeTranslate.Core/TranslationResults.cs` | 安全的失败块元数据和汇总统计 |
| `src/OfficeTranslate.Core/TranslationTaskFailure.cs` | 中断后仍可呈现已有文本错误汇总 |
| `src/OfficeTranslate.Core/OfficeTranslate.Core.csproj` | Word 内部访问权限；Core 版本读取 VERSION.txt |
| `src/OfficeTranslate.WordAddIn/WordTranslationService.cs` | 原正文循环接入调度器，保留原目标、写回和 UndoRecord |
| `src/OfficeTranslate.WordAddIn/SettingsForm.cs` | 并发数控件、加载和保存 |
| `src/OfficeTranslate.WordAddIn/TranslationResultForm.cs` | 失败块警示与详情，手动关闭 |
| `src/OfficeTranslate.WordAddIn/UiText.cs` | 五语言设置、进度、错误和更新说明 |
| `src/OfficeTranslate.WordAddIn/AboutForm.cs` | 显示本次更新说明 |
| `src/OfficeTranslate.WordAddIn/Properties/AssemblyInfo.cs` | 2.3.4.0 |
| `src/OfficeTranslate.ExcelAddIn/Properties/AssemblyInfo.cs` | 2.3.4.0 |
| `src/OfficeTranslate.PowerPointAddIn/Properties/AssemblyInfo.cs` | 2.3.4.0 |
| `VERSION.txt` | 统一版本源 2.3.4.0 |
| `tests/OfficeTranslate.Tests/WordParallelTranslationTests.cs`（新增） | 新增 27 个回归用例 |
| `tests/OfficeTranslate.Tests/ParallelTranslationTransportTests.cs` | 保留 Excel 用例，额外验证 Word 并行 HTTP 工厂 |
| `tests/OfficeTranslate.Tests/DocumentChatHostTests.cs` | 同步三宿主版本断言 |
| `tests/RunWordParallelRangeProbe.ps1`（新增） | 真实 Word 正文/表格替换与双语定位补查 |
| `tests/VALIDATION-2.3.4-PARALLEL-TRANSLATE.md`（新增） | 本验证记录 |
| `CHANGELOG.md` | 2.3.4 正式版变更 |
| `README.md` | 使用方式、提交单位、顺序等待语义及版本状态 |
| `APPLY.md` | 本地版本及构建/验证范围 |
