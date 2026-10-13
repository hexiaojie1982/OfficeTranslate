# OfficeTranslate 2.1.38.0：Excel 文本有界并发

日期：2026-10-13。仅本地分支 `fix/excel-parallel-cells-2.1.38`。
基线：`b94d95f963b351029c2f0182369d3263111458fc`，与本次开始时的
`origin/main`（2.1.37 发布合并）一致；本地 `main` 引用落后且未修改。
未 push、合并、发布或安装。

## 实现与配置

- `TranslationSettings.MaxParallelCells` 默认 4，范围 1–32。配置文件
  `%LOCALAPPDATA%\OfficeTranslate\settings.json` 可添加 `"MaxParallelCells": 4`；
  老配置缺省为 4，1 恢复串行。设置窗口保存时保留该值。
- Excel 单元格和图形文本共享 `SemaphoreSlim` 有界调度。内存只保留在途 Task，
  不为整张工作表一次性创建所有任务。每个任务持有自己的目标对象，不捕获可变循环索引。
- 读取目标仍在原 UI 路径；网络并发，写回均通过 `OfficeUiDispatcher.InvokeAsync`。
  结果可乱序完成，双语拼接、尾部换行清理、单元格 WrapText 行为不变。
- 完成数用 `Interlocked.Increment`，递增与报告在同一 UI 回调中，避免进度倒退。
  未变化、跳过的目标也算已处理；计入 summary 的成功发生在写回之后。
- `TranslationTaskSummary.Record` 加锁，保护全部文本计数的复合更新。
  图片仍在所有文本任务完成后串行处理。
- 调度器每次调用使用独立 TaskCompletionSource；固定 WinForms 上下文负责将并发
  Post 串行送到捕获的 STA。异步续体不在 UI 回调内内联运行。
- 共享客户端首次 Timeout 配置和 JSON 序列化/反序列化加锁；没有跨网络 await 持锁。
  Excel 文本 HttpClientHandler 的 MaxConnectionsPerServer 对齐配置，避免连接数限制
  使应用层并发无法到达服务端。图片 HttpClient 及 OCR 超时/重试路径不变。
- 缓存和 PromptDescriptors 保持原并发集合；strict 重试、翻译完整性校验不变。
- 取消或首个失败停止新请求，取消在途请求，在排队写回内再次检查 token；所有在途任务
  收尾后才返回/抛出并释放客户端和信号量。保留首个失败，不让其触发的其他取消覆盖原因。

## 验证

- Release 解决方案构建及 WiX MSI 构建成功，0 错误。完整构建保留 7 个已有 nullable 警告。
- 新增并发专项：16/16 通过，覆盖 1/4 并发上限、逆序完成、单语/双语、UI 线程归属、
  单调进度、预取消、在途取消、排队写回取消、模型失败、写回失败、跳过/未变化和准确统计。
- 100 个线程池调用者并发向真实 WinForms STA 消息泵投递，返回值和线程归属全部正确。
- 50,000 条并发混合结果记录验证 summary 的全部文本分类及段落计数。
- 本机真实 HTTP 测试在收到 4 个请求后才释放响应，证明 4 个请求能同时到达服务端；
  同一客户端每格先返回原文再 strict 重试，共 8 次请求，之后重复翻译命中缓存。
  不连接外部模型，不读取密钥或用户文档。
- 非交互会话首轮：353 通过、22 失败；22 项均因 UserInteractive=false 无法显示模态
  窗口。保留原始 TRX，不更改断言或跳过用例；改在已登录的桌面会话重跑全量测试。
- 桌面全量结果：**375/375 通过，0 失败、0 跳过**，59 秒；
  Session=1、UserInteractive=True，CLI 退出码 0。临时交互测试计划任务运行后删除。
- 未实测真实 Excel + vLLM Qwen3 27B 的端到端吞吐，不将本机并发证据解释为模型加速倍数。

## 版本与文件

版本统一为 `2.1.38.0`：README、VERSION.txt、build.ps1 示例、三个宿主 AssemblyInfo、
About 当前版本日志。Core 的版本由 `src/Directory.Build.props` 从 VERSION.txt 派生。
`installer/Product.wxs` 已使用 `$(var.ProductVersion)`，无需改为硬编码；实际查询 MSI
ProductVersion、四个产品 DLL 的 File 表版本、三条 COM Assembly 注册值均为 2.1.38.0。
全仓版本扫描保留 2.1.37 历史日志、验证记录和旧探针的原始版本。

修改/新增文件：

- `src/OfficeTranslate.Core/ExcelTextTranslation.cs`（新增）、`TranslationSettings.cs`、
  `TranslationResults.cs`、`TranslationClient.cs`、`OfficeUiDispatcher.cs`。
- `src/OfficeTranslate.ExcelAddIn/ExcelTranslationService.cs`。
- `src/OfficeTranslate.WordAddIn/SettingsForm.cs`、`AboutForm.cs`。
- Word、Excel、PowerPoint 各自的 `Properties/AssemblyInfo.cs`。
- `tests/OfficeTranslate.Tests/ExcelParallelCellsTests.cs`、`ParallelTranslationTransportTests.cs`
  （新增）、`OfficeUiDispatcherTests.cs`（复用 STA 测试消息泵）。
- `README.md`、`VERSION.txt`、`build.ps1`、本验证记录。

## 复现与本机证据

```powershell
.\build.ps1 -Configuration Release
# 全量测试需要交互桌面；非交互后台会话不能执行原有模态窗口用例。
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=all-desktop.trx' --results-directory TestResults\excel-parallel-2.1.38
```

`TestResults/excel-parallel-2.1.38/`（忽略目录）保存 build.log、parallel.trx、
all-final.trx（非交互首轮）、all-desktop.trx（桌面全量）、desktop-context.txt、
desktop-test.log、desktop-test.exit 和 version-scan.log。
本地 MSI：`artifacts/OfficeTranslate.Office.x64.msi`，未安装或分发。
SHA-256：`7D68E48000B22756F77895F672FB93941F7E4829DABD2760954E7E8AEC2000C7`。
