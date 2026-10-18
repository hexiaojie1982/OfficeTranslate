# 2.3.6 Chat CJK 词间空格清理验证

日期：2026-10-18。基于干净的 `main`（`a6bbdfa`，merge #28）创建 `feat/cjk-space-cleanup`。产品、Core、Word / Excel / PowerPoint 程序集、MSI 及 COM 注册版本均为 **2.3.6.0**。未合并、未创建 tag、未 push，未执行安装。

## 实现与旧逻辑检查

先搜索全仓的 CJK、空格、Unicode 范围、Normalize/Remove spacing 相关内容，并检查普通翻译的 `TranslationClient.cs` 与 `SourceLanguageProtector.cs`。已有的是 `CjkLoanwordSpacingInstruction`、`NeedsBoundarySpace`、`Restore(..., separateLatinWordBoundaries)`：用于保护混合语言内容并给拉丁词边界补空格；**没有删除 CJK 字符间空白的旧 helper**，因此没有复用它们，也没有修改普通文字翻译路径。

新增仅供 Chat 使用的 `ChatCjkSpacing.Clean`，在 `DocumentChatSession.SendAsync` 最终 transport 返回后、提交历史前调用一次。清理后的同一份 answer 同时用于内存历史、待重试的历史保存队列、磁盘 JSON 和返回的 `ChatReply`。最终气泡、当前/历史消息的复制、md/txt 导出及后续请求中的会话历史均读取该结果，不在各出口重复清理。

- 只删除两侧均为 CJK 码点的连续半角空格、全角空格及 tab。覆盖汉字及其 BMP/补充平面扩展、兼容汉字、平假名、片假名和指定 CJK 标点/全角字符区间，不将韩文音节或 Jamo 纳入匹配。
- 保留 CRLF/LF/CR、首尾空白、NBSP/细空格、拉丁文/数字/emoji 边界和用户消息。
- 保留反引号和波浪线围栏（围栏长短匹配、最多三格缩进、未闭合围栏到文末）；保留单/多反引号行内代码及其换行、tab。处理转义反引号与不匹配的分隔符，不重写 Markdown。
- 流式快照及原有 50ms 合批不变；仅已成功完成的回答进入清理和历史提交。取消/失败沿用原有行为。
- 新 helper 不写日志，也不增加任何聊天正文、Prompt 或剪贴板内容日志。

## 修改文件（14 个）

| 文件 | 修改 |
| --- | --- |
| `src/OfficeTranslate.Core/ChatCjkSpacing.cs`（新增） | Chat 专用确定性清理，保护代码与 Unicode 边界 |
| `src/OfficeTranslate.Core/DocumentChat.cs` | 完成回答后、写入历史前接入，共用清理后的 answer |
| `tests/OfficeTranslate.Tests/ChatCjkSpacingTests.cs`（新增） | 4 个测试方法、40 个实际用例 |
| `tests/OfficeTranslate.Tests/DocumentChatHostTests.cs` | 三宿主程序集版本断言更新为 2.3.6.0 |
| `VERSION.txt` | 统一版本源 2.3.6.0；Core/package/MSI 版本由此派生 |
| `src/OfficeTranslate.WordAddIn/Properties/AssemblyInfo.cs` | AssemblyVersion/FileVersion 2.3.6.0 |
| `src/OfficeTranslate.ExcelAddIn/Properties/AssemblyInfo.cs` | AssemblyVersion/FileVersion 2.3.6.0 |
| `src/OfficeTranslate.PowerPointAddIn/Properties/AssemblyInfo.cs` | AssemblyVersion/FileVersion 2.3.6.0 |
| `src/OfficeTranslate.WordAddIn/UiText.cs` | 五语言 2.3.6 更新记录 |
| `src/OfficeTranslate.WordAddIn/AboutForm.cs` | About 更新记录顶部接入 2.3.6 |
| `README.md` | 当前验收版本、功能说明与验证链接 |
| `APPLY.md` | 当前版本、构建安装说明与验收边界 |
| `CHANGELOG.md` | 2.3.6 更新记录 |
| `tests/VALIDATION-2.3.6-CHAT-CJK-SPACING.md`（新增） | 本报告 |

已扫描版本引用；2.3.5 等历史更新记录、旧验证文档及旧复制测试的证据目录保留原版本号。没有创建新 Git tag 或 release。

## 测试与构建

- 先加测试再实现：初始 39 项在旧逻辑上 **29 失败、10 通过**；10 项为应保持原样的边界用例，证据 `red.trx`。
- 实现后初始专项 **39/39 通过**（`focused.trx`）；随后增加 1 项 Unicode 范围/幂等测试，并增强既有 6 项三宿主用例的历史气泡与复制断言。
- 最终全量 Release 回归 **843/843 通过，0 失败、0 跳过**，测试报告耗时 **2 分 22 秒**，退出码 **0**。相对 main 的 803 项新增 **40** 项，原有 803 项保留并通过。最终 40 个新增用例全部包含在 `all.trx` 中。
- 全量测试命令：`dotnet test tests/OfficeTranslate.Tests/OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=all.trx' --results-directory TestResults/cjk-space-cleanup-2.3.6`。
- 全量运行在已登录 Windows 桌面 **Session 1 / Interactive=True**，使用实际编译的三宿主 WinForms 控件、STA 消息循环与已有真实系统剪贴板回归；文档、网络模型和历史目录使用隔离夹具及合成文本。
- 执行 `dotnet clean OfficeTranslate.Word.sln -c Release` 后，运行 `./build.ps1 -Configuration Release`：**0 错误、7 条既有 nullable 警告**，WiX MSI 打包成功。后续仅增强测试断言并重建测试项目；MSI 四个 DLL 的哈希与最终 Release 输出逐一相同。
- 最终测试期间核对 120 个源代码/测试/项目文件的 SHA-256 清单一致；`git diff --check` 通过。

## 边界场景清单

| 场景 | 结果 |
| --- | --- |
| `韩 国 翻 译`、混合半角/全角/tab | 清为 `韩国翻译` |
| `我在用 iPhone 16`、`2024 年 10 月` | 保留中英/中数空格 |
| 韩文回答、韩文 Jamo、纯英文回答 | 空格不变 |
| 汉字扩展与兼容区、补充平面、假名、CJK 标点 | 正确按码点匹配，无代理对拆分 |
| emoji、异常代理字符、NBSP/细空格 | 保持边界，无异常 |
| CRLF/LF/CR、首尾空白 | 原样保留 |
| 单/多反引号行内代码、跨行代码 | 内部完全保留 |
| 三/四反引号、缩进、不同长度、未闭合围栏、波浪线围栏 | 内部完全保留；围栏外清理 |
| 用户含中文词间空格 | 内存、JSON、用户气泡复制与导出均保留 |
| Question / Summarize / Rewrite，历史保存失败再重试 | 保存的是清理后的回答；下一次请求历史一致 |
| Word / Excel / PowerPoint × 翻译设置开/关（6 项） | 流式原样；最终气泡、磁盘 JSON、历史气泡、当前/历史复制、md/txt 文件一致 |
| 既有完整回归 | 气泡布局、Markdown、流式合批、复制按钮、历史会话、导出、Enter/Shift+Enter、多语言、翻译设置开关及翻译/OCR 回归通过 |

以上是自动化控件与隔离夹具验证；没有实际调用远端模型、安装本包或完成 Office 真机验收。

## MSI

- 路径：`D:\pythonProject\office-translate\native-word\artifacts\OfficeTranslate.Office.x64.msi`
- 字节数：**958464**。
- SHA-256：`E958E210A64FE28BC09DE5E9FFCB128ED175C580C36DF6CD2A301D29F7D6602D`。
- 打包前实测 2.3.5 MSI 与用户提供的 `C43658B34CB02948C39DE1A1E92A7905EB67E84519F1ED6CF77459A4EE09CAF0` 一致；新包哈希不同。
- 直接读取 MSI 数据库：ProductVersion **2.3.6.0**，6 个 COM Assembly 注册项均为 **2.3.6.0**。
- 使用 WiX dark 解出 Core / Word / Excel / PowerPoint 四个 DLL，其 AssemblyVersion/FileVersion 均为 **2.3.6.0**，SHA-256 与最终 Release 输出逐一相同。

证据目录：`TestResults/cjk-space-cleanup-2.3.6/`，包括 `red.trx`、`focused.trx`、`all.trx`、`all.exit`、`session.txt`、`clean.log`、`release-build.log`、`test-build.log`、`source-manifest.json`、`msi.json`、`versions.json` 和 MSI 解包证据。
