# OfficeTranslate for Microsoft Office（原生 COM 加载项）

面向完全禁止 Office Web Add-in 的 Windows 企业环境。该版本不使用 `manifest.xml`，通过 MSI 注册原生 Word COM 加载项。

## 第一版功能

- 翻译并直接替换当前选区。
- 可切换双语模式：启用后，翻译选区或全文时保留原文，并在每段原文下方插入译文。
- 全文翻译按段实时写回：每完成一段立即更新文档，无需等待全部段落翻译完成。
- Excel：翻译选中单元格或当前工作表，跳过公式、数字与空白，逐单元格实时写回；双语模式在同一单元格换行追加译文。
- PowerPoint：翻译选中文字、形状、表格、组合形状、选中幻灯片或整份演示文稿，逐文本框实时写回。
- 逐段翻译正文并直接替换；倒序写回，避免位置漂移。
- Word 功能区快速选择源语言和目标语言，源语言支持自动检测，并可一键反转。
- 紧凑型语言工具栏和统一的自绘线性图标，不依赖不同 Office 版本的内置图标映射。
- 单次 Word 撤销记录，可用一次“撤销”恢复翻译前内容。
- Ollama `/api/chat`。
- OpenAI-compatible `/chat/completions`，支持 OpenAI、DeepSeek、LM Studio、vLLM、LocalAI 等。
- 根据 Base URL 和 API Key 从 `/api/tags` 或 `/models` 更新模型下拉列表。
- 目标语言、提示词、术语表、分块大小配置。
- API Key 使用当前 Windows 用户的 DPAPI 加密。
- 进度显示在 Word 状态栏，支持取消。

## 构建环境

构建机需要：

1. Visual Studio 2022 Build Tools，安装“.NET 桌面生成工具”。
2. .NET Framework 4.8 Developer Pack。
3. .NET SDK 8.x。
4. WiX Toolset 3.14；脚本会自动检查 PATH、`WIX` 环境变量和默认安装目录。
5. NuGet 访问权限（首次恢复 `Microsoft.Office.Interop.Word`）。

注意：安装 `.NET Runtime` 不够，`dotnet --list-sdks` 必须至少显示一个 8.x SDK。安装 WiX 后应能在新 PowerShell 窗口中运行 `candle.exe -?` 和 `light.exe -?`。

在 x64 PowerShell 中执行：

```powershell
.\build.ps1 -Configuration Release
```

输出：`artifacts\OfficeTranslate.Office.x64.msi`，统一安装 Word、Excel 和 PowerPoint 加载项。

## 安装与卸载

当前 MSI 面向 **64 位 Office**：

```powershell
msiexec /i OfficeTranslate.Office.x64.msi /qn /l*v install.log
msiexec /x OfficeTranslate.Office.x64.msi /qn /l*v uninstall.log
```

安装后重新启动 Word、Excel 或 PowerPoint，功能区会出现“OfficeTranslate”。首次使用先打开“翻译设置”。

> Windows 位数不等于 Office 位数。若组织仍使用 32 位 Office，需要另行构建 x86 MSI，不能把 x64 MSI 部署到 32 位 Office。

## 企业部署注意事项

- MSI 应使用组织信任的代码签名证书签名，再通过 Intune、Configuration Manager 或组策略分发。
- 可通过防火墙仅允许 `WINWORD.EXE` 访问批准的模型地址。
- 使用 Ollama 时，模型服务应由 IT 统一安装、配置和更新。
- 若 Word 将加载项放入“禁用项目”，请先检查事件日志和安装日志，不建议通过策略强制忽略崩溃。

## 当前边界

- 仅处理主文档正文；页眉页脚、脚注、尾注、文本框和批注暂不翻译。
- 全文按 Word 段落翻译。段落内部复杂混合格式可能被统一为该段落首个字符的格式；段落样式和表格结构通常保留。
- 密码保护、只读或受信息权限管理保护的文档无法写回。
- 当前安装定义仅覆盖 64 位 Office，尚未加入签名和自动更新。
