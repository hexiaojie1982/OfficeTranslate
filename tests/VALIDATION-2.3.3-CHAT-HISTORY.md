# 2.3.3 本地历史会话验证

分支：`feat/chat-history`，基于本地 `main`（2.3.2）。开发版本：`2.3.3.0`。仅本地开发、构建和测试，不安装、不 push、不合并、不发布、不创建 tag。

## 行为与存储格式

- 三宿主共用功能组「总结 / 改写 / Chat 设置」；操作组「清空 / 更新上下文 / 导出 / 历史会话」。
- 每次完整问答成功后保存；失败、取消、清空前的过期请求不进入历史。清空保留已存文件，下一轮使用新会话 ID。退出宿主不删除已存文件。
- 路径：`%APPDATA%\OfficeTranslate\ChatHistory\<SessionId>.json`。SessionId 为 GUID 的 N 格式（32 位十六进制）。UTF-8 无 BOM；一个会话一个文件。
- 每轮读入已有 JSON 并追加问答，写同目录临时文件，再用 `File.Replace` 原子替换；初次保存用 `File.Move`。没有独立索引文件。
- JSON 中时间使用 ISO 8601 `O` 格式，保留时区。`CreatedAt` 为首轮完整问答完成时间，`DocumentName` 为该轮捕获的文档名。读取名称失败时为空，不阻断问答。
- `SchemaVersion` 为 2，`Host` 由宿主显式传入，值为 Word/Excel/PowerPoint，未保存、无扩展名的文档也能识别。首轮成功问答的来源随保存失败重试保留，清空后重新捕获。读取兼容版本 1；缺少 Host 时按 `.doc/.docx`、`.xls/.xlsx`、`.ppt/.pptx` 推断（忽略大小写），无法推断时留空。只读打开不改写旧 JSON，追加时升级为版本 2。不支持的格式版本按无法读取的文件跳过并提示。
- 历史列表首列显示「来源应用」：蓝色 Word、绿色 Excel、橙色 PowerPoint 文字，其他列保留默认文字颜色；不引入外部图片。列标题通过 UiText 覆盖简体/繁体中文、韩文、日文、英文。
- `Question` 为实际用户提问或总结/改写请求，`Answer` 保留 AI 回答原文，包括 Markdown 与原始换行。详情视图仅为 WinForms 显示转换 CRLF，不修改 JSON。
- 活跃会话仍仅保留最近 10 轮供展示、导出及模型上下文使用；完整历史从磁盘读取，不自动恢复到模型上下文。历史窗口按创建时间倒序显示时间、文档名及首问最多 100 个字符预览，可只读查看全部轮次。
- 单条删除与清空全部均需确认，默认选「否」。清空全部只删除专用目录顶层 JSON 文件，包含损坏的历史文件；不递归删除目录或其他类型文件。删除活跃会话的历史后，下一轮只保存此后新增轮次，不从内存恢复之前已删除轮次；活跃内存仍受「清空」按钮控制。
- 保存失败时保留待保存轮次并显示本地化错误，下一轮和「清空」重试；重试失败则拒绝清空，避免静默丢失。关闭前应解决存储故障；进程退出或崩溃时尚未保存的轮次无法保证恢复。
- 列表跳过无法读取的文件并提示。读写、删除错误不显示异常原文，不将历史正文写日志。历史文件只由本地文件系统接口读取；不上传，不保存文档上下文、系统提示词、API 凭据。
- 新增文字全部使用 UiText，覆盖简体/繁体中文、韩文、日文、英文。

示例（格式化展示，实际为紧凑 JSON）：

```json
{
  "SchemaVersion": 2,
  "SessionId": "1cd052d6268b48e9a0d5ebdd3026c246",
  "CreatedAt": "2026-10-16T12:30:00.0000000+08:00",
  "DocumentName": "报告.docx",
  "Host": "Word",
  "Turns": [
    {
      "Timestamp": "2026-10-16T12:30:00.0000000+08:00",
      "Question": "请总结当前文档上下文。",
      "Answer": "## 总结\n回答原文"
    }
  ]
}
```

## 改动文件

- `src/OfficeTranslate.Core/ChatHistoryStore.cs`：JSON 格式、本地追加/原子写入、读取、摘要列表、单条/全部删除。
- `src/OfficeTranslate.Core/DocumentChat.cs`：文档名、可选历史存储、完整轮次保存、失败重试、清空轮换 ID；保留 10 轮限制。
- `src/OfficeTranslate.WordAddIn/ChatHistoryForm.cs`：共用历史列表、只读原文、删除确认、安全错误提示。
- `src/OfficeTranslate.WordAddIn/DocumentChatPane.cs`：调整按钮、接入存储及历史窗口、保存失败提示。
- `src/OfficeTranslate.WordAddIn/DocumentChatHost.cs`：三宿主捕获文档名。
- `src/OfficeTranslate.WordAddIn/UiText.cs`：全部历史窗口文案及五语言更新记录。
- Excel/PowerPoint 的 `.csproj`：链接共用 `ChatHistoryForm.cs`。
- `tests/OfficeTranslate.Tests/ChatHistoryTests.cs`：新增存储及三宿主中英文历史窗体回归、隔离的临时历史目录。
- `tests/OfficeTranslate.Tests/ChatUiRegressionTests.cs`：新按钮分组、保存失败 UI、清空重试、历史按钮打开正确存储；现有测试使用临时目录。
- `tests/OfficeTranslate.Tests/DocumentChatHostTests.cs`：三宿主文档名捕获、临时历史目录、版本断言。
- `VERSION.txt`、三宿主 `Properties/AssemblyInfo.cs`：同步 2.3.3.0。Core/包版本由已有 `Directory.Build.props` 读取版本文件；MSI 继续由已有构建脚本注入版本。
- `README.md`、`CHANGELOG.md`、`APPLY.md` 及本文件：更新本地持久化规格、版本、应用与验证说明。

## 验证结果

1. 最终 Release 构建（含 WiX MSI）成功：**0 错误**，6 条既有 Word/旧测试 nullable 警告；新增代码无编译警告。全量重编 Core 时还会出现其既有 `OfficeUiDispatcher.cs` nullable 警告。
2. 最终交互桌面全量测试：**685/685 通过，0 失败、0 跳过**，1 分 33 秒，退出码 0；Session 1 / `UserInteractive=True`。相比 main 的 670 项新增 15 项（6 项 Core 历史行为、三宿主中英文历史窗口 6 项、三宿主保存失败/清空重试/历史按钮 3 项）；扩展既有分组及文档捕获测试。
3. 实测覆盖：Unicode/Markdown/原始换行及时区往返；同一文件追加；13 轮完整落盘与活跃 10 轮限制；清空新建 ID；失败/取消/过期请求不落盘；按创建时间排序及首问预览；单条/全部删除及取消确认；保存目标被占用时原文件不变、临时文件清理、重试不重复；删除活跃历史后不恢复旧轮次；三宿主 UI 保存失败提示及历史按钮接线。
4. 已检查中英文历史详情/空状态截图及 320/440/600 宽度侧栏布局；详情换行正常，按钮和只读原文可用。历史窗口截图保存在本证据目录；侧栏布局截图沿用既有测试目录 `TestResults/chat-bubbles-2.3.2/screenshots/`。
5. 版本核对：Core、三宿主程序集版本/文件版本及最终 MSI `ProductVersion` 均为 **2.3.3.0**；版本证据和 MSI SHA-256 保存在本证据目录。MSI 未安装或发送。
6. `git diff --check` 通过。翻译/OCR、模型传输、配置/凭据存储实现未修改。交互桌面测试临时计划任务完成后删除。

首次在 Session 0 运行的全量测试未通过，原因包括既有窗体测试要求交互桌面及旧版本断言；保留该次 `full-suite.log/.trx` 便于区分。上述最终结果来自修正后候选程序集的交互桌面全量复测，结果为 `full-desktop.log/.trx/.exit`。

证据目录（git 忽略）：`TestResults/chat-history-2.3.3/`。构建命令：`./build.ps1 -Configuration Release`（含 WiX MSI）。全量测试：`dotnet test OfficeTranslate.Word.sln -c Release --no-build --no-restore`。使用项目本地 `DOTNET_CLI_HOME=dotnet-home`、`NUGET_PACKAGES=packages`。

验证边界：真实编译的三宿主控件与窗体在 STA 消息循环中运行，文档和模型使用隔离夹具。没有安装候选版本、调用真实模型或修改真实用户历史；真实 Office 功能区打开侧栏及其他设备/DPI 的人工验收仍需安装后进行。

## 来源应用标识追加验证（2026-10-16）

基于 `feat/chat-history` 的 `c7fb415`，版本按要求保持 **2.3.3.0**。

- Release 构建：`./build.ps1 -Configuration Release -SkipInstaller`，**0 错误、7 条既有 nullable 警告**，新增代码无编译警告。本次未生成安装包、未安装。
- 交互桌面全量测试：**713/713 通过，0 失败、0 跳过**，约 1 分 38 秒，退出码 0，`SessionId=1; UserInteractive=True`。
- 在原有 685 项基础上新增 28 项：三宿主显式 Host 保存/读回、失败重试和清空后重新捕获（3）；六种扩展名、大小写、空名及未知扩展名的版本 1 兼容、只读不改写和追加升级（9）；版本 2 缺少 Host 兼容及未知格式版本跳过（1）；三宿主、五种界面的混合来源列表、彩色首列及旧记录推断（15）。另扩展既有三宿主适配器测试，断言选中内容和全文路径均传入正确 Host。
- 新增未知格式版本测试首次发现 `InvalidDataException` 未被存储错误筛选捕获，导致列表不能跳过该文件；已修正并通过最终全量复测。首次结果为 712/713，保留于 `initial-desktop.log/.trx/.exit`。
- 三宿主五语言列表截图已生成，已人工检查简体中文、繁体中文与英文截图，来源标签与后续列显示正常。未知来源保留空白，三宿主共用原存储目录，无外部图片资源。
- Core、Word、Excel、PowerPoint 的程序集版本及文件版本均为 **2.3.3.0**。`git diff --check` 通过。仅本地修改，未 push、未合并、未发布。

本次证据目录（git 忽略）：`TestResults/chat-history-host-2.3.3/`，包含 `release-build.log`、`full-desktop.log/.trx/.exit`、`desktop-context.txt`、`dll-versions.json` 及三宿主五语言 `*-hosts.png`。测试沿用隔离文档/模型夹具及临时历史目录；未使用真实用户历史。临时交互桌面计划任务完成后删除。
