# 2.3.3 Chat 输入快捷键验证

分支：`feat/chat-history`。版本按本次要求保持 **2.3.3.0**。仅本地修改和构建，未安装、push、合并、发布或创建 tag；保留工作区原有历史会话功能修改。

## 改动

- `src/OfficeTranslate.WordAddIn/DocumentChatPane.cs`：Enter 且未按 Shift 时发送并抑制换行；Shift+Enter 交给原生多行文本框插入换行，显式启用 `AcceptsReturn`。Excel、PowerPoint 已通过项目链接编译同一份控件，无需另建处理逻辑。
- `src/OfficeTranslate.WordAddIn/UiText.cs`：输入提示和空白会话说明统一改为 Enter 发送、Shift+Enter 换行，覆盖简体/繁体中文、韩文、日文、英文；原有标签、ToolTip、AccessibleName 继续通过 UiText 更新。
- `tests/OfficeTranslate.Tests/ChatUiRegressionTests.cs`：新增三宿主中英文共 6 项回归，向真实编译控件发送窗口键盘消息，验证 Shift+Enter 实际产生 CRLF 且不发送、Enter 发送且不插入换行、多行提问进入会话、等待回复时不重复发送、空输入不发送、标签/ToolTip/AccessibleName 的双语文本。
- `README.md`、`CHANGELOG.md`：同步快捷键说明；本文件记录验证结果。

## 结果

- Release 构建及 WiX MSI：`./build.ps1 -Configuration Release`，退出码 0，**0 错误**。首次产品源码重编有 6 条既有 nullable 警告；调整测试输入方式后的最终增量构建有 4 条既有测试 nullable 警告。
- 交互桌面全量测试：`dotnet test OfficeTranslate.Word.sln -c Release --no-build --no-restore`，**719/719 通过，0 失败、0 跳过**，约 1 分 37 秒，退出码 0。`SessionId=1; UserInteractive=True`。包含全部 6 项新增回归。
- 首次测试使用全局 SendKeys，发生按键未送达和 Windows `Access is denied`，结果为原有 713 项通过、6 项新增测试失败。最终改为仅针对测试控件的窗口消息，并在测试线程临时设置/恢复修饰键状态；不再注入全局键盘输入。初次日志保留在证据目录的 `initial-desktop.*`。
- MSI `ProductVersion`、Core 及三宿主程序集/文件版本均为 **2.3.3.0**。使用 WiX dark 提取 MSI 后，Core 及三个宿主 DLL 的 SHA-256 均与构建输出一致。
- `git diff --check` 通过。测试使用隔离文档、模型和临时历史目录；没有安装候选版本或在真实 Office 中进行人工操作。临时交互桌面计划任务已删除。

最终安装包：`artifacts/OfficeTranslate.Office.x64.msi`

- 字节数：**937984**
- SHA-256：`F883964FE5FF3EEC6B63976172023A56EFA92AA7C056064F041C0CF1E608AB54`

证据目录（git 忽略）：`TestResults/chat-input-2.3.3/`，含 `release-build.log`、`full-desktop.log/.trx/.exit`、`desktop-context.txt`、`artifact-metadata.json` 和 MSI 提取文件。
