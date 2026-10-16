# 2.3.3 Chat 复用翻译设置验证

分支 `feat/chat-history`，版本按本次明确要求保持 **2.3.3.0**。保留工作区已有历史会话及输入框改动。本次仅本地修改、构建和测试，不安装、push、合并、发布或创建 tag。

## 实现

- `src/OfficeTranslate.Core/TranslationSettings.cs`：新增 `ChatUseTranslationSettings`，缺省为 `false`，使用现有 `SettingsStore` 与 Chat 设置一起持久化，不复制风格、额外要求或术语表。
- `src/OfficeTranslate.WordAddIn/DocumentChatPane.cs`：在 Chat 设置功能组下方增加勾选框；初始化、重新打开及刷新时恢复保存的状态，切换立即保存。保存失败恢复原状态，仅显示本地化错误。沿用宿主共享控件，覆盖 Word、Excel、PowerPoint。
- `src/OfficeTranslate.WordAddIn/UiText.cs`：增加开关文案，中英文及原有繁体、韩文、日文均走 UiText。
- `src/OfficeTranslate.WordAddIn/SettingsForm.cs`：翻译服务设置重建配置对象时保留开关；Chat 设置保存也保留开关。
- `src/OfficeTranslate.Core/ChatPrompts.cs`、`DocumentChat.cs`：按所选模式的原 system prompt 和原安全指令构建基础文本。关闭开关直接返回基础文本，逐字符不变；开启后顺序追加 `Translation style:`、`Additional requirements:`、`Glossary:` 段。风格通过 `TranslationStyleCatalog.GetPrompt(settings.TranslationStyle)` 解析，额外要求读 `settings.CustomInstructions`，术语表读 `settings.Glossary`；空值跳过，额外要求与风格正文完全相同时只拼一次。Custom 风格不凭空补默认风格。每次发送读取共享配置的最新值，不写回 Chat prompt 字段。长文中间摘要同样应用开关；拼接在预算计算之前，超预算在调用模型前拒绝。
- `src/OfficeTranslate.Core/ChatErrors.cs`：扩展错误脱敏，过滤风格正文、额外要求和术语表及其段落；未新增 prompt 正文日志。历史会话仍仅保存用户问题和模型回答。
- `tests/OfficeTranslate.Tests/ChatTranslationSettingsTests.cs`：新增 21 项回归，覆盖中英文三模式开/关差异、最新配置读取、空值/Custom/风格去重、长文 map/final 请求及预算、超预算拒绝、错误脱敏、旧配置及开关存取、三宿主中英文控件恢复/保存失败/语言切换/真实请求路径，以及翻译服务和 Chat 设置保存后保留开关。
- `README.md`、`CHANGELOG.md`：同步使用说明与本地开发变更记录。

## 验证

- Release：`./build.ps1 -Configuration Release -SkipInstaller`，成功，**0 错误，7 条既有 nullable 警告**。此次没有生成或安装新 MSI。
- 新增回归：`dotnet test tests/OfficeTranslate.Tests/OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --filter FullyQualifiedName~ChatTranslationSettingsTests`，**21/21 通过，0 失败、0 跳过**。
- 交互桌面全量：`dotnet test OfficeTranslate.Word.sln -c Release --no-build --no-restore`，**740/740 通过，0 失败、0 跳过**，约 1 分 42 秒，退出码 0。包含新增 21 项回归；运行环境 `SessionId=1; UserInteractive=True`。临时计划任务已删除。
- `VERSION.txt`、Core 及三个宿主 DLL 的程序集/文件版本均为 **2.3.3.0**，版本及 SHA-256 记录在 `binary-metadata.json`。
- `git diff --check` 通过。
- 构建及测试证据目录（git 忽略）：`TestResults/chat-translation-settings-2.3.3/`。
- 验证使用隔离配置文件、临时历史目录、模型替身和编译后的三宿主真实 WinForms 控件；没有调用真实模型或在真实 Office 中进行人工验收。
