# 2.3.9 安全审计修复验证

日期：2026-10-21。开发版本：**2.3.9.0**。工作目录为 `D:\pythonProject\office-translate\native-word`；开发分支 `fix/security-audit-findings`，基于 main `3fba5c2`（v2.3.8）。

## 修复对应

| 发现 | 实现 | 回归证据 |
| --- | --- | --- |
| B1：Excel 翻译公式注入 | Core `ExcelCellText.Literal` 抽取原 Chat 的撇号前缀和 32767 字符限制，翻译目标闭包和 Chat 共用。替换及双语模式均在最后一次 Value2 写入前中和 | 独立 Excel COM 实例、临时工作簿，经真实 `ReadRangeTargets` 闭包和 `ExcelTextTranslation.RunAsync` 写入。检查 `=2+2`、`+2+2`、`-2+2`、`@SUM(A1:A2)`、普通中文、数字/日期外观和前导撇号；`HasFormula=false`、Value2 保留准确文本、双语换行和 WrapText 正确 |
| A1：Chat 响应无界 | 传输 `ResponseHeadersRead` 加实际字节计数；8 MiB 字节及累计字符上限，含当前行、累计 data、生成文本、JSON 回退及错误响应。`MaxJsonLength` 同为 8 MiB。超限为 `ChatFailure.Protocol`，释放流并停止协议回退 | 超长无换行行、keepalive、无事件边界 data、累计多个合法事件、精确字符边界、混合换行、非 ASCII 字节越界、伪造 Content-Length、SSE/JSON/Ollama/错误/回退分支、直接 JSON 长度检查 |
| A2：公网 HTTP API Key 无警告 | Core `EndpointSecurity` 统一分类；三宿主链接同一 SettingsForm / ChatSettingsForm 及警告 UI。公网 HTTP 且非空 Key 时弹 Yes/No 警告，默认 No，明确说明 Key 将明文发送 | IPv4 loopback/RFC1918/链路本地及边界、IPv6 loopback/ULA/链路本地/mapped IPv4、HTTPS/无 Key、本地域名/混合解析/解析失败；三宿主两种设置窗实际保存入口，拒绝不改文件、确认后保存 |
| B2：图片响应和 regions 无界、写回不查取消 | 图片 HTTP 在默认全量缓冲之前改为有界读取，16 MiB 字节；外层响应和 regions JSON 解析均有限。仅处理前 512 个原始条目，超限计入 `RegionLimitDroppedCount` / DroppedCount，经纯 OCR 转换仍保留；结果窗显示数量并要求手动关闭。三宿主 OCR 和图片翻译的区域规划、覆盖框写回及旁注回调均查取消 | 分块/无长度响应、错误响应、Content-Length 提前拒绝、16 MiB 精确边界、0/511/512/513/1500 条区域、无效条目、统计与诊断、三宿主结果窗提示及取消后的真实已编译旁注回调 |

## 最小改动与现状差异

- `ImageRecognitionResult.Truncated` 已用于表示模型输出长度截断，并由调用方拒绝整张图片写回。保留此语义；新增独立 `RegionsTruncated`，表示只写前 512 条并将遗漏数量提示为待检查。
- 第 513 条及后续区域不再逐条建对象或 COM 形状；外层 JSON 仍会在 16 MiB 有界输入内反序列化。统计按原始前 512 条处理，不能用大量无效条目绕过条数预算。
- 图片默认传输原来是 `ResponseContentRead`，只替换末端 `ReadAsStringAsync` 无法限制网络缓冲。因此将预算前移至 `SendImageAttemptAsync`，保持原有每次图片请求完整读取 deadline 和失败分类；`ReadAsStringAsync` 只解码已经受限的缓冲内容。
- 图片读取增加异步边界后，既有 HTTP backoff 取消用例揭示了 `TaskCanceledException` 与既有精确类型断言的竞态。图片 API 在用户令牌取消时统一重新抛出该令牌的 `OperationCanceledException`，保留静默用户取消与超时错误的区别。
- 改为 headers-only 传输后，HttpClient 可能结束它传给 handler 的内部令牌生命周期；因此响应流获取阶段也由调用方自己的令牌/截止时间保护。既有延迟响应体测试仍使用原断言。包含图片超时、Chat 兼容及响应预算的传输专项 **72/72 通过**，证据 `TestResults/security-transport.trx`。
- 取消检查位于每次区域/旁注操作前；保留已完成写回，不宣称能中断正在执行的单次同步 COM 调用，也未增加回滚。
- 域名通过本机 DNS 分类，全部地址属于本地范围才豁免。不能解析或混有公网地址时仍警告；未改变 HTTP 请求头或后台请求策略。

## 新增用例

| 测试类 | 新增数据用例数 |
| --- | ---: |
| SecurityExcelWritebackTests | 3 |
| SecurityResponseBudgetTests | 18 |
| SecurityEndpointTests | 36 |
| SecurityImageRegionTests | 16 |
| 合计 | **73** |

专项测试：**73/73 通过，0 失败、0 跳过**，运行于已登录桌面 Session 1；证据 `TestResults/security-focused-interactive.trx` / `.log`，退出码文件为 0。真实 Excel COM 用例包含在全量测试中，无过滤跳过。

## 基线与最终全量

基线是未经本次修改构建的 `3fba5c2`（2.3.8），共 **929** 项。首次在服务 Session 0 启动时有 48 项真实窗口/键盘/剪贴板失败；未削弱原有断言，改为交互式 Session 1 运行。

Session 1 连续两轮基线均为 **923 通过、6 失败、0 跳过 / 929**。失败是三宿主各两项：

- `Copy_RealClipboardLockShowsFailure_ThenRecoversAfterRelease`：原生 `OpenClipboard` 无法取得锁。
- `Copy_RealUnicodeClipboardRoundTripsLongRawMarkdown`：剪贴板写入反馈为 Copy failed。

独立原生探针确认 Session 1、WinSta0、Default 输入桌面、UserInteractive=true；`OpenClipboard` 返回 false，Windows 错误码 **5（Access denied）**。调整到普通/提升运行权限以及重新打开窗口站句柄后仍复现。没有关闭用户程序、修改剪贴板权限/系统策略，或修改/跳过这 6 个测试。

最终全量结果：**996 通过、6 失败、0 跳过 / 1002（929 + 73）**，执行 1002 项，耗时约 2 分 50 秒，退出码 1。新增 73 项均通过，失败名单与两轮基线逐项相同。图片超时、响应体 deadline 和 HTTP backoff 用户取消用例已通过，无新增失败。**尚未满足“全量测试必须通过”**；PR 保持草稿，需在剪贴板访问恢复后重新取得完整通过结果。

证据：`TestResults/security-after-final.trx` / `.log` / `.exit`；基线位于 `tests/OfficeTranslate.Tests/TestResults/security-before-interactive.trx` 和 `security-before-interactive2.trx`。逐项比对结果在 `TestResults/security-validation-state.json`，`same_failed_as_baseline=true`。没有跳过或放宽失败断言。

执行命令（通过临时 InteractiveToken 计划任务在已登录桌面启动，PowerShell 使用隐藏窗口）：

```powershell
./build.ps1 -Configuration Release -SkipInstaller
dotnet test OfficeTranslate.Word.sln -c Release --no-build --no-restore --logger 'trx;LogFileName=security-after-final.trx' --results-directory TestResults --blame-hang-timeout 60s --blame-hang-dump-type none
```

## 版本与交付边界

- VERSION.txt、Word / Excel / PowerPoint AssemblyVersion / AssemblyFileVersion、Core 包/程序集版本均为 **2.3.9.0**；四个 Release DLL 的 AssemblyVersion / FileVersion 实际读取一致。About 新增五语言 2.3.9 条目，保留历史记录。
- Release 构建 **0 错误、7 条既有 nullable 警告**；`build.ps1` 版本一致性检查通过。WiX 的产品与 COM 注册版本继续从 VERSION.txt 注入。
- 本次没有构建 MSI；不将旧 artifacts 中的包作为 2.3.9 交付物。未安装、创建 tag、合并 PR、发布 Release、发邮件或发送附件。
- 其他 worktree 的位置和 HEAD 保持不变：`chat-translate-button` 为 main `3fba5c2`，`review-image-overlay` 为 `fix/validation-boundaries` `663a081`。main 重复检出仅在本目录使用 `--ignore-other-worktrees`，pull 已确认无更新。
