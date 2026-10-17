# 2.3.5 Chat 逐条复制验证

日期：2026-10-17。基于干净的 `main`（`62de3da`，merge #27）创建本地分支 `feat/copy-button`。产品、Core、三个宿主程序集、MSI 和 COM 注册版本统一为 **2.3.5.0**。未合并、未创建 tag、未 push、未发布、未安装或发送文件，等待用户真机验收后明确说“合并发布”。

## 实现

- 三宿主共享 `ChatCopyButton`，使用 XAML `Geometry.Parse` 定义复制/对勾矢量路径，再由现有 WinForms/GDI+ 界面绘制。仅增加 .NET Framework 自带的 PresentationCore / WindowsBase 引用，无图片文件、浏览器或新增 NuGet 包。
- 用户和 AI 气泡右下角独立预留 24px（随 DPI 缩放）按钮区域，正文和时间戳不重叠；保留短气泡收缩和 80% 宽度上限。浅灰常显、悬停加深，键盘 Enter/Space 也可触发。
- 从消息的原始 source 复制，不读取 RichEdit 的渲染文本；保留 Markdown、CRLF/LF、尾部空格和 Unicode，不附加角色、时间戳、头像或错误诊断。
- AI 流式期间隐藏按钮，输出结束后显示；空串或纯空白消息隐藏。取消/失败停止流式后，可复制已显示的部分正文，不复制诊断后缀。
- 剪贴板操作限定所属 UI STA 线程；后台调用通过已存在的控件句柄调回 UI，未创建/已销毁句柄时不在后台创建句柄或操作剪贴板。失败只显示本地化提示，不写正文、Prompt、剪贴板或异常原文日志。
- 点击成功显示对勾与“已复制”，失败显示“复制失败”；WinForms Timer 1500ms 后恢复。重复点击重新计时，销毁时释放 Timer/ToolTip。复制、已复制、复制失败通过 UiText 覆盖简中、繁中、韩语、日语、英语；语言刷新也更新当前反馈。
- 历史详情复用只读气泡，逐条复制与当前会话相同；保留文档名、创建时间、每轮时间戳、宿主列、删除确认和原始 JSON。选中历史会话后从首条消息开始阅读。

## 修改文件（20 个）

| 文件 | 内容 |
| --- | --- |
| `src/OfficeTranslate.WordAddIn/ChatCopyButton.cs`（新增） | 矢量按钮、STA 调度、剪贴板失败隔离、反馈计时及释放 |
| `src/OfficeTranslate.WordAddIn/ChatConversationView.cs` | 气泡按钮、流式/空白状态、底部布局及历史回到首条 |
| `src/OfficeTranslate.WordAddIn/DocumentChatPane.cs` | UI 语言、流式开始/完成/取消/失败状态接入 |
| `src/OfficeTranslate.WordAddIn/ChatHistoryForm.cs` | 历史消息只读气泡与元数据 |
| `src/OfficeTranslate.WordAddIn/UiText.cs` | 五语言复制提示及 2.3.5 更新记录 |
| `src/OfficeTranslate.WordAddIn/AboutForm.cs` | 更新记录显示 2.3.5 |
| `src/OfficeTranslate.WordAddIn/OfficeTranslate.WordAddIn.csproj` | Framework Geometry 引用 |
| `src/OfficeTranslate.ExcelAddIn/OfficeTranslate.ExcelAddIn.csproj` | 链接共享按钮及 Geometry 引用 |
| `src/OfficeTranslate.PowerPointAddIn/OfficeTranslate.PowerPointAddIn.csproj` | 链接共享按钮及 Geometry 引用 |
| `src/OfficeTranslate.WordAddIn/Properties/AssemblyInfo.cs` | 2.3.5.0 |
| `src/OfficeTranslate.ExcelAddIn/Properties/AssemblyInfo.cs` | 2.3.5.0 |
| `src/OfficeTranslate.PowerPointAddIn/Properties/AssemblyInfo.cs` | 2.3.5.0 |
| `VERSION.txt` | 统一版本源 2.3.5.0，Core/package/MSI 从此派生 |
| `tests/OfficeTranslate.Tests/ChatCopyButtonTests.cs`（新增） | 11 个测试方法，35 个数据化用例 |
| `tests/OfficeTranslate.Tests/ChatHistoryTests.cs` | 历史详情改为验证只读气泡、原始 source、元数据；保留原删除和本地化断言 |
| `tests/OfficeTranslate.Tests/DocumentChatHostTests.cs` | 同步宿主版本断言 |
| `README.md` | 复制功能和本地验收状态 |
| `APPLY.md` | 2.3.5 构建、安装及等待验收边界 |
| `CHANGELOG.md` | 2.3.5 更新记录 |
| `tests/VALIDATION-2.3.5-CHAT-COPY.md`（新增） | 本报告 |

版本已用 `git grep` 全量扫描。历史 CHANGELOG、About 历史更新项和旧验证报告仍保留各自原始版本；没有创建 2.3.5 Git tag。

## 测试与构建

- 先添加针对性测试再实现：初始 26 个数据化用例在原实现上 **26 失败、0 通过**，结果 `red.trx`。实现过程中继续补齐真实剪贴板占用、取消/失败、矢量状态等 9 个用例；最终新增 **35** 项。
- 专项回归：**211/211 通过，0 失败、0 跳过**，约 1 分 15 秒（`focused.trx`）。
- 完整 Release 清理重建：`dotnet clean OfficeTranslate.Word.sln -c Release`，随后 `./build.ps1 -Configuration Release`，**0 错误、7 条既有 nullable 警告**，WiX MSI 打包成功。
- 最终完整 Release 构建的全量测试：**803/803 通过，0 失败、0 跳过**，约 **2 分 24 秒**，退出码 0（`all.trx` / `all.exit`）。相对 main 的 768 项新增 **35** 项，原有 768 项全部保留。最终构建前后及测试后核对 118 个源代码/测试/项目文件 SHA-256，一致。
- 在已登录 Windows 桌面 Session 1（`Interactive=True`）执行真实 Word / Excel / PowerPoint 编译控件和 STA 消息循环；文档上下文、网络模型和历史目录使用隔离夹具。真实系统剪贴板测试使用合成数据，无真实模型调用或用户聊天正文。

## 针对性回归清单

1. 三宿主用户原文和 AI Markdown 原文逐字一致；100000 个中文字符、emoji、代码围栏、混合换行及尾部空格不截断（长 payload 与昂贵的 RichEdit 自然高度测量分开验证）。
2. 三宿主 AI 非空流式隐藏、完成显示；用户/AI 空串和纯空白隐藏，隐藏状态不执行复制。
3. 三宿主后台触发，实际 ClipboardWriter 在线程 ID 与拥有控件的 UI 相同、ApartmentState 为 STA 的线程运行。
4. 注入 ExternalException 失败不抛出；真实后台 OpenClipboard 占用导致“复制失败”，释放后可重试成功。
5. 成功/失败 tooltip、对勾、1500ms 恢复、重复点击重新计时、窗口销毁停止计时；中英韩日及繁体中文、已有消息切换语言后的当前提示。
6. 真实 UnicodeText 剪贴板往返超过 70000 个韩文字符及 Markdown 原文完全一致。
7. 取消和传输失败停止流式后复制部分原文，不附带错误诊断，不产生已完成历史。
8. 两种角色悬停图标实际像素变暗；复制后图标像素变为对勾，保存正常/悬停/已复制截图。
9. 三宿主历史四条只读消息逐条复制原文，读取/复制前后 JSON 字节不变，打开时从首条开始阅读。
10. 三宿主 260/320/440/600px 宽度：按钮位于气泡右下角，正文不重叠，最大 80%、短气泡收缩及下方时间戳不变。
11. 原有全量用例继续验证 Markdown、50ms 流式合批、时间戳、圆形头像、Enter/Shift+Enter、持久化/宿主标识、md/txt 导出、五语言切换和翻译设置拼接 system prompt，以及翻译/OCR 其他功能。

## MSI

- 路径：`D:\pythonProject\office-translate\native-word\artifacts\OfficeTranslate.Office.x64.msi`
- 字节数：**954368**。
- SHA-256：`3528BB5870E380FB91BA504DADDE414ED50F286FA390072CCC706A8F35F6B675`。
- 重建前 2.3.4 MSI 实测为 `3FEC42A39CA3B87EF43186915DE8885339765836F56D7EF1792EB9E53698BD34`，新哈希不同。
- 直接读取 MSI 数据库：ProductVersion **2.3.5.0**；6 个 COM Assembly 注册项版本均为 **2.3.5.0**。
- 用 WiX dark 解出 MSI 中的 Core、Word、Excel、PowerPoint DLL；四个 DLL 的 AssemblyVersion/FileVersion 均为 **2.3.5.0**，SHA-256 与最终 Release 输出逐一相同（`versions.json`）。没有执行安装。

证据目录：`TestResults/copy-button-2.3.5/`，含红灯/专项/全量 TRX、Release 构建、源码哈希清单、MSI 信息及控件截图。自动控件回归和截图不等于真实 Office 端到端验收，仍需用户在实际宿主中验收交互与视觉效果。
