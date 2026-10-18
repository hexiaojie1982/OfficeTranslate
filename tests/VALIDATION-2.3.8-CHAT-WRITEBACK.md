# 2.3.8 Chat 输入与 Question 写回验证

日期：2026-10-18。产品、程序集、安装包及 COM 注册版本：**2.3.8.0**。

## 分支与合并

- 工作树：`D:\pythonProject\office-translate\chat-translate-button`。
- 分支：`feat/chat-writeback-fix`，从 `feat/chat-translate-button` 的 `7293983`（2.3.7）创建。
- 合并提交 `22f09cf` 的两个父提交为 `7293983` 和 `34023c9`：已合入 `feat/cjk-space-cleanup`（2.3.6）。保留翻译枚举、请求和配置，以及完成回答后、历史保存前的 CJK 清理。
- 版本、更新记录、About 和版本断言的冲突按功能叠加解决，About 同时保留 2.3.7 与 2.3.6 历史条目，并增加五语言 2.3.8 条目。
- 原 `native-word` 工作树仍在 `feat/cjk-space-cleanup` 的 `34023c9e5303f5948d9c5c8596e4998ac1ac8a24`，tracked 工作区干净；原翻译按钮分支仍为 `7293983`。
- 仅本地提交，未合并到 main、未打 tag、未 push，未安装或发送邮件。

## 修复

- `DocumentChatPane.Start` 读取并校验请求后，立即在 UI 线程清空输入框，四种模式一致；移除回答完成回调中的清空，避免删掉正在编辑的下一条草稿。空问题及 pending 请求仍由原入口检查拦截。
- Question 的写回按钮与 Rewrite / Translate 一致：替换选中与插入下方，不显示插入开头。Summarize 仍仅插入开头。
- 同步扩展 `ChatContext.Allows` 的 Question 写回权限，确保按钮点击后确实可执行；保留 `CanReplace`、非空结果及 writer 委托检查。全文及不可替换选区的替换按钮仍禁用，插入下方按 writer 可用性启用。
- 保留用户确认、拒绝确认不写入、成功后只写一次的行为。生产改动没有新增正文、提示词、文档内容或剪贴板日志。
- 2.3.7 翻译按钮、可编辑翻译系统提示词及设置开关完整保留；2.3.6 CJK 清理同时适用于完成的 Translate 回答，历史、气泡、导出、复制及写回使用同一结果。翻译专项测试由旧的保留 CJK 空格断言调整为合并后的清理行为，Markdown 原文仍保留。

## 测试

本次新增 **31 项数据用例**，另合入 2.3.6 的 **40 项**；相对 2.3.7 的 858 项共增加 **71 项**。

| 新增范围 | 数量 | 验证 |
| --- | ---: | --- |
| Question：三宿主 × 五语言 | 15 | 实际点击发送，transport 进入前已清空、流式未完成时仍清空、完成保留下一条草稿；完成显示两种写回按钮和对应提示；逐项确认并实际调用替换/插入委托，拒绝确认不写入，成功后禁用；CJK 清理后写回、历史一致 |
| 四种模式 × 三宿主 | 12 | 发送开始清空；可替换选区、不可替换选区、全文和无 writer 四种上下文；总结仅插入开头，其余两按钮；结果非空、writer 与 CanReplace 防护 |
| Core 写回策略 | 3 | 增加 Question 可替换选区及 Translate 选区/全文；原 Question 全文用例更新为允许插入下方 |
| Translate CJK 清理与历史重试 | 1 | Translate 完成清理、历史保存失败重试及后续请求历史一致 |

先新增 27 项面板测试再修改生产代码：**27/27 复现失败**，错误均为发送后输入框仍保留指令。证据：`TestResults/writeback-red.trx`、`writeback-red.log`。

修复后专项回归：**119/119 通过，0 失败、0 跳过**，覆盖新面板用例、Core 写回策略、完整 CJK 清理和翻译按钮测试。证据：`TestResults/writeback-focused.trx`、`writeback-focused.log`。

最终全量回归：**929/929 通过，0 失败、0 跳过**；Session 1，测试耗时约 **2 分 46 秒**，退出码 **0**。TRX 的 total / executed / passed 均为 929，failed / notExecuted 均为 0。相对 main 的 803 项共增加 126 项（2.3.7 新增 55 + 2.3.6 新增 40 + 本次新增 31）。运行结束后已删除临时交互计划任务。

执行命令：

```powershell
dotnet test tests/OfficeTranslate.Tests/OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=writeback-full.trx' --results-directory TestResults
```

全量覆盖原有翻译按钮、CJK 空格清理、气泡、复制按钮、Markdown、流式/SSE/回退/取消、历史、Markdown/文本导出、Enter/Shift+Enter、五语言、翻译设置开关，以及原有翻译/OCR/并行处理测试。为保留原焦点/模态断言，在已登录桌面 Session 1 运行，没有删减原有用例。证据：`TestResults/writeback-full.trx`、`writeback-full.log`、`writeback-full-state.json`。

Question 完成后的三宿主五语言截图：`TestResults/chat-writeback-2.3.8/`；翻译按钮回归截图：`TestResults/chat-translate-2.3.8/`。已检查 Word 简中、PowerPoint 韩文截图：两种写回按钮、完成提示和下一条草稿均正常显示。

## Release 与 MSI

执行 `dotnet clean OfficeTranslate.Word.sln -c Release` 后运行 `./build.ps1 -Configuration Release`：构建及 WiX 打包成功，**0 错误、7 条既有 nullable 警告**。警告所在源码/测试文件本次均未修改。证据：`TestResults/writeback-clean.log`、`writeback-release-build.log`。

- MSI 路径：`D:\pythonProject\office-translate\chat-translate-button\artifacts\OfficeTranslate.Office.x64.msi`。
- 字节数：**962560**。
- SHA-256：`9FDB30DFD5A0F0A23017EA16E1906C1D973C49A231EFA89FD64096961E21DEAC`。
- 与 2.3.7 的 `3907F1338283685D002189BEB6B651888FF88202E3D4DE0A928A91FE8EE34E94` **不同**。
- 实际读取 Windows Installer 数据库：ProductVersion 为 **2.3.8.0**，6 条加载项/Chat 控件 COM Assembly 注册均为 **2.3.8.0**。
- WiX dark 解包后，Core / Word / Excel / PowerPoint 四个 DLL 的 AssemblyVersion 与 FileVersion 均为 **2.3.8.0**，SHA-256 与对应 Release 输出逐一一致。
- 核验脚本及证据：`TestResults/inspect-writeback238-msi.ps1`、`writeback-msi-metadata.json`、`writeback-msi/`、`writeback-msi-dark.log`。dark 的 DARK1059 为 WiX 标准对话框反编译关系警告，不影响实际 MSI 数据库和载荷核验。

## 改动文件

相对 2.3.7 提交 `7293983`，共 20 个文件（含合并带入的清理实现和历史验证报告）：

| 目录 | 文件 |
| --- | --- |
| 根目录 | `VERSION.txt`、`README.md`、`APPLY.md`、`CHANGELOG.md` |
| `src/OfficeTranslate.Core/` | `ChatCjkSpacing.cs`、`DocumentChat.cs` |
| `src/OfficeTranslate.WordAddIn/` | `DocumentChatPane.cs`、`UiText.cs`、`AboutForm.cs`、`Properties/AssemblyInfo.cs` |
| `src/OfficeTranslate.ExcelAddIn/` | `Properties/AssemblyInfo.cs` |
| `src/OfficeTranslate.PowerPointAddIn/` | `Properties/AssemblyInfo.cs` |
| `tests/OfficeTranslate.Tests/` | `ChatWritebackFixTests.cs`、`ChatCjkSpacingTests.cs`、`ChatTranslateTests.cs`、`ChatUiRegressionTests.cs`、`DocumentChatHostTests.cs`、`DocumentChatTests.cs` |
| `tests/` | `VALIDATION-2.3.6-CHAT-CJK-SPACING.md`、本报告 |

版本扫描保留旧 Changelog、UiText 历史条目和历史验证报告的原版本号；当前版本源、程序集声明/断言、新 UI 更新记录和 MSI 统一为 2.3.8.0。Core/package 版本由 `src/Directory.Build.props` 读取 VERSION.txt 派生。

## 验证边界

自动化加载三宿主真实 DLL，使用真实 WinForms 控件、隔离历史目录、宿主接口代理及本地 HTTP/SSE fixture；不会连接远程模型，不代表模型翻译质量或安装注册后的 Office 真机验收。所有测试正文均为合成 fixture。未修改另一工作树的代码、构建输出或安装状态。
