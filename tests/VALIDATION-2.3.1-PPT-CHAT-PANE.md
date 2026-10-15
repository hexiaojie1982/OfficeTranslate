# 2.3.1 PowerPoint Chat 侧栏修复

日期：2026-10-15。直接修改本地 `feat/doc-chat-2.3.0`，按本次要求保持 **2.3.1.0**。保留工作区已有的 DeepSeek 修复；未提交、push、合并或发布，未安装候选 MSI。

## 根因与复现

三宿主共用 `AddIn.Chat.cs`，通过 `ICustomTaskPaneConsumer.CTPFactoryAvailable` 获取 Office 原生工厂，再以各自 ProgID 调用 `ICTPFactory.CreateCTP`，不是 VSTO 的 CustomTaskPanes 集合。PowerPoint 的 `_host` 和工厂初始化没有发现遗漏。

旧代码将 `ActiveWindow` 保存为 `dynamic`，随后读取 `window.HWND`。在 NUC 的真实 PowerPoint、Session 1 中，动态 COM 调用抛出：

```text
System.MissingMemberException: Error while invoking HWND.
  at System.Dynamic.ComRuntimeHelpers.CheckThrowException(...)
  at CallSite.Target(...)
  at WindowProbe.Late(Object w)
```

对同一个真实窗口，`((Microsoft.Office.Interop.PowerPoint.DocumentWindow)window).HWND` 成功返回非零句柄。也就是说，PowerPoint 的强类型 COM 接口可读取该属性，但这里的 IDispatch/dynamic 路径失败；异常发生在 CreateCTP 前。旧 catch 吞掉异常，统一提示检查 Chat 注册，因此提示不能证明注册缺失。

MSI 的三个控件均有独立 ProgID/CLSID、Control、Implemented Categories、mscoree.dll、Class、Assembly、RuntimeVersion 和各自 CodeBase。PowerPoint 中 `Class=OfficeTranslate.WordAddIn.DocumentChatPane` 是共享源码编译后的实际类型名，Assembly/CodeBase 正确指向 PowerPoint DLL。回归测试按编译后的类型核对注册，没有修改 MSI 注册内容。

## 修改文件

- `src/OfficeTranslate.AddInShared/AddIn.Chat.cs`：三宿主统一使用强类型窗口和句柄访问；继续以原文档窗口绑定工厂、缓存及清理，未改用 Application.HWND。拆出同一个 Ribbon 入口使用的创建核心，便于验证真实异常。工厂未提供时报告明确原因；新侧栏成功显示后才入缓存，失败时清理控件/窗格，清理异常不覆盖原异常。
- `src/OfficeTranslate.WordAddIn/UiText.cs`：替换误导性的注册/固定版本兜底提示，补充 2.3.1 更新记录。打开失败显示阶段、真实异常类型、脱敏消息、HRESULT 和最多六个关键堆栈方法；标准消息框可按 Ctrl+C 复制。复用 ChatErrors 脱敏配置中的密钥和 Prompt；不附带文档正文、参数、局部变量或源码路径，不持久化聊天数据。
- `tests/OfficeTranslate.Tests/ChatPaneActivationTests.cs`：新增 15 项回归，使用实际编译宿主、控件和 COM 接口代理，覆盖创建/初始化/隐藏后复用、初始化或显示失败后的清理和重试、保留原异常、安全诊断及三宿主 MSI 注册匹配。使用不存在的隔离设置路径，不读取用户模型密钥。
- `tests/OfficeTranslate.OfficeValidation/PowerPointChatProbeAddIn.cs`、`RunPowerPointChatProbe.ps1`、`OfficeTranslate.OfficeValidation.csproj`：增加不进入产品解决方案/MSI 的真实 Office 验证夹具，将真实工厂交给实际产品 AddIn，在 Office UI STA 上执行。脚本核对候选 DLL 哈希，临时隔离已安装同名程序集，备份并恢复 COM/加载项注册，恢复后导出比对哈希。NUC 本次使用 `-RegistrationHive HKLM`；HKCU 测试注册在此环境未能通过 COM 激活。该夹具不是交付产品的一部分。
- `README.md` 和本文：补充故障说明、验证及边界。

## 最终验证

1. `build.ps1 -Configuration Release` 成功，**0 错误**，6 条原有 Word/旧测试 nullable 警告；WiX 编译/链接成功。本次新增代码没有编译警告。
2. 交互桌面全量测试 **583/583 通过，0 失败、0 跳过**，包括新增的 15 项侧栏回归。Session 1 / UserInteractive=True，退出码 0。
3. 真实 PowerPoint 原生工厂验证通过：动态 HWND 反证失败、强类型 HWND 成功、COM 激活实际产品 UserControl、右停靠 440 宽度、上下文读取、Ribbon 入口隐藏后重开且复用、两个演示文稿独立侧栏、切换文档后仍写回原演示文稿、关闭窗口清理及新建文档重开。使用临时演示文稿和固定文本，不连接真实模型。验证结束 PowerPoint 退出，临时注册删除，原注册及 LoadBehavior=3 恢复并核对。
4. Core/三宿主 AssemblyVersion、FileVersion、About、MSI ProductVersion、包内四个 DLL、六项 COM Assembly 注册值均为 **2.3.1.0**。`git diff --check` 通过。
5. 本地 MSI：`artifacts/OfficeTranslate.Office.x64.msi`，SHA-256：`77ADCEB2FF70103FE1B123439EF9FBEFBD936EC4729D92D415277294F1E59746`。未安装、发送、发布此 MSI。

证据目录（Git 忽略）：`TestResults/ppt-chat-2.3.1/`。包含 `window-probe.txt`、`activation.trx`、`build.log`、`all-desktop.trx`、`desktop-test.log`、`desktop-test.exit`、`desktop-context.txt`、`version-verification.log`；最终真实宿主证据位于 `real-final/real-powerpoint-chat.txt`、`real-final/real-powerpoint-run.log` 和该目录中的注册备份/恢复比对文件。临时计划任务已删除。

## 复跑

```powershell
.\build.ps1 -Configuration Release
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore
dotnet build tests\OfficeTranslate.OfficeValidation\OfficeTranslate.OfficeValidation.csproj -c Release
# 先关闭 PowerPoint；在已登录的交互桌面运行，证据目录需存在且不得含上次结果。
.\tests\OfficeTranslate.OfficeValidation\RunPowerPointChatProbe.ps1 -EvidenceDirectory <新证据目录> -RegistrationHive HKLM
```

真实验证覆盖 NUC 当前 Office；未代替用户其他 Office 版本/DPI 环境的验收，也未发送真实模型请求。Word/Excel 本轮通过共享侧栏激活回归和原有全量测试，未额外执行两者的真实 Office 侧栏自动化。

原生工厂的父窗口语义参考：[Microsoft ICTPFactory.CreateCTP](https://learn.microsoft.com/en-us/office/vba/api/office.ictpfactory.createctp)。本次 PowerPoint 动态访问问题的结论来自上述本机真实 COM 复现。
