# 2.3.1 DeepSeek Chat 修复

分支：`feat/doc-chat-2.3.0`。版本：`2.3.1.0`。只修改 Chat 业务、对应测试及版本元数据；不修改翻译/OCR 实现，不安装、不提交、不 push、不合并、不发布。

## 定位与修复

- Chat 和翻译均通过 `CreateRequest` 构造 UTF-8 JSON POST、Bearer 认证及同一个模型名。均使用配置的 Base URL 追加 `/chat/completions`（Ollama 仍用 `/api/chat`）。未发现将云端强制路由到 vLLM 默认端口/路径的代码。
- **可复现的请求差异**：翻译已对官方 `api.deepseek.com` 设置 `thinking: { type: "disabled" }`，Chat 的 SSE/普通恢复请求均未设置。DeepSeek 默认思考模式可能把生成预算花在 `reasoning_content`，以 `finish_reason=length` 结束，未产生完整 `content`；短输入不能排除这种输出耗尽。修复在 Chat 两条请求中复用翻译的官方主机判断及同一非思考模式策略，vLLM/其他兼容服务不发送该扩展参数。660 字测试用受控服务模拟默认思考耗尽的响应；不将这个模拟结果当作用户家用电脑实际异常的证据。
- **可复现的 SSE 边界问题**：旧解析器只在空行时处理事件，`[DONE]` 后直接 EOF 会误判中断。现在处理 EOF 未派发事件、首行 BOM、CRLF、多 data 行、无空格前缀；忽略 keep-alive、role/null content、reasoning_content 和 usage-only 事件。JSON 语法异常转换为不含响应原文的诊断。
- **超时差异**：Chat 未调用 `Configure(settings)`，因而使用 HttpClient 默认 100 秒而非保存的 TimeoutSeconds；现在沿用翻译的配置方法。每次尝试仍有独立总时限，取消会处置阻塞中的响应流；主动取消不降级。
- **错误被吞**：旧代码将 HTTP/网络/解析等异常全部触发普通请求，再统一显示 ChatError。现在区分网络/超时、401/403 认证、模型/服务、输入/输出预算、限流、余额和配置错误。仅协议中断、网络失败或明确不支持 stream 的错误允许恢复；认证、确定的模型错误与预算错误不盲目重复请求。完整 JSON 响应直接使用，避免额外计费请求。
- 聊天框展示可复制的异常类型、消息、HTTP 状态或 finish_reason；网络和本地错误保留底层异常链，移除已知密钥、文档段落、问题、系统提示词和 URL。恢复也失败时同时保留两次安全诊断。服务错误体仅用于分类，JSON 解析异常改为安全说明，不显示/记录服务原始正文、Prompt、密钥。失败不提交历史，不产生写回按钮。

DeepSeek 参数依据：[官方思考模式文档](https://api-docs.deepseek.com/guides/thinking_mode/)。本文没有断言用户此次真实失败必然由思考模式造成；原版本没有留下可用异常，需要用户真机重试或新错误详情完成最后确认。

## 回归范围

新增 `ChatCompatibilityTests.cs` 覆盖 DeepSeek 根路径、`/v1`、`/beta`、大小写主机、伪后缀主机、Bearer/JSON/messages/model/stream/max_tokens、恢复请求策略、660 字输入、vLLM 无 Key、SSE reasoning/keep-alive/usage/BOM/CRLF/EOF/多行、JSON 响应避免重复请求、JSON/纯文本不支持流、HTTP 分类、SSE 错误/长度/过滤/服务资源不足、隐私信息、网络异常链、请求超时与读取中取消。

三宿主控件回归验证分类提示、可复制详情、无失败历史/写回入口。既有翻译/OCR 测试仍在全量套件中运行。

## 验证结果

1. `build.ps1 -Configuration Release` 成功：**0 错误**，7 条原有代码/测试的 nullable 警告，Chat 新代码无警告；WiX 编译/链接成功。
2. 最终交互桌面全量测试 **568/568 通过，0 失败、0 跳过**，退出码 0；Session 1 / UserInteractive=True。Chat 相关共 82 项，其中新加 38 项（34 项服务兼容与诊断、4 项三宿主分类错误 UI）。独立 Core Chat 回归 63/63 通过。
3. 将旧 `HEAD` 的 `TranslationClient.Chat.cs` 临时替换到相同测试环境进行反证，660 字默认思考场景和两个 EOF/SSE 场景 **3/3 失败**，分别为 `InvalidDataException: Incomplete chat response.` 和 `Interrupted or empty SSE response.`。已恢复修复代码并重新完成上述最终构建、全量测试。
4. Core 与三宿主 AssemblyVersion / FileVersion、三个 About.CurrentVersion、MSI ProductVersion、MSI 中四个产品 DLL、六个 COM Assembly 注册值全部为 **2.3.1.0**。
5. `git diff --check` 通过；翻译客户端主文件、Recognition、TranslationSettings、SettingsStore、三宿主翻译服务与 HEAD 无差异。其他改动仅为 Chat、测试和版本/说明。
6. 本地构建产物：`artifacts/OfficeTranslate.Office.x64.msi`，SHA-256：`B3160AA99949615EF072B9638A47AD6BD84F455055C28907147E52E252103414`。未安装、发送、发布该 MSI。

证据目录（git 忽略）：`TestResults/doc-chat-2.3.1/`，包含 `build.log`、`chat.trx`、`baseline-repro.trx`、`baseline-repro.log`、`all-desktop.trx`、`desktop-test.log`、`desktop-test.exit`、`desktop-context.txt`、`version-verification.log`、`frozen-paths.log`。临时交互测试计划任务已删除。

复现命令：

```powershell
.\build.ps1 -Configuration Release
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=all-desktop.trx' --results-directory TestResults\doc-chat-2.3.1
```

全量测试在已登录交互桌面运行（避免 Session 0 对窗体/剪贴板测试的影响）。本次未连接真实 DeepSeek 或 vLLM，未使用真实文档或读取用户 API 密钥；服务兼容性通过 mock HTTP/SSE 验证，真实云端请求及用户真机仍需重试验收。
