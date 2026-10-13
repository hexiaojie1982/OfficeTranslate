# OfficeTranslate 2.2.0：纯 OCR 识别按钮

日期：2026-10-14。基线 `main / 05622be`；开发分支 `feature/ocr-button-2.2.0`。
按项目四段版本惯例，DLL/MSI 使用 `2.2.0.0`。本记录覆盖发布前构建和验证，未执行安装。

## 行为

- Word、Excel、PowerPoint 的共享 Ribbon 新增「OCR」按钮，使用 32×32 文档加放大镜图标。
- 只处理选中的图片；无图片时提示「请先选中一张图片」。PowerPoint 选中幻灯片缩略图不触发整页图片扫描。
- 不依赖图片翻译开关、双语模式、源/目标语言、术语表或翻译风格。请求仅包含区域检测与原文转录指令；意外返回的 translation 内容被忽略。
- `TranslationClient.RecognizeImageAsync` 直接调用现有 `RequestImageRegionsAsync`，沿用图片模型、8192 输出预算、独立图片超时、重试、结构化输出回退和截断拒绝策略。无额外翻译调用。
- 复用原图/剪贴板捕获基础设施、Word 预捕获预算、`ImageOverlayPlanner`、宿主真实排版检查、固定几何验证和旁注拆分/恢复。新编排独立保存在 partial 文件中，避免修改旧翻译方法。
- 新覆盖框使用 `OfficeTranslateOCR:Recognition:`，新旁注使用 `OfficeTranslateOCR-Note:Recognition:`，后接原图片 owner。旧翻译清理按完整 owner 匹配，两种结果互不清理。沿用父前缀使旧文本收集器继续排除插件生成的框。
- PowerPoint 同样写 AlternativeText 独立前缀，并保留不同值的 `OfficeTranslateOCR` tag 供旧文本收集器排除；OCR 清理依据 AlternativeText，避免 Tags 大写规范化干扰匹配。
- 原文旁注标头为「图片原文待检查」。独立 OCR 摘要显示识别统计，失败时保留完整未放置旁注，不声称发生了翻译。
- OCR 结果窗体的尺寸、字体、颜色、布局和按钮风格与翻译结果窗体统一，保留可复制的完整原文旁注。

## 验证结果

1. Release 构建 **0 错误**，MSI 生成成功。最新 UI 风格构建记录为 `TestResults/ocr-result-style/build.log`，含 6 个既有可空性警告。
2. 最新桌面**全量测试 409/409 通过，0 失败，0 跳过**。相对 2.1.38 新增 34 项：19 项纯 OCR 逻辑/请求用例，15 项实际编译宿主 DLL 的 Ribbon、OCR 摘要及窗体风格用例。
3. 新增覆盖：两种 provider 的单请求与模型选择、忽略翻译设置/意外译文、无文字、原文排版、坏坐标旁注、超长原文与旁注全文恢复、双向标记隔离、结构化输出回退、429/503 重试、图片超时、用户取消、截断及无效 JSON 拒绝。
4. 当前命令会话处于非交互 Session 0，首次全量测试 22 个已有窗体用例因 `UserInteractive=False` 失败。随后使用本机 Interactive 计划任务在 Session 1 执行全量测试，全部通过。临时计划任务已删除。
5. 对 `TranslationClient.cs` 及三个 `*TranslationService.cs` 做基线全文比较：仅类声明增加 `partial`，去掉这个声明差异后与 `05622be` 完全相等。旧 `TranslateImageAsync`、图片请求通道及各宿主旧捕获、规划、写回方法均未改变。
6. 四个产品 DLL 的 AssemblyVersion/FileVersion、MSI ProductVersion、File 表中四个产品 DLL 版本及三条 COM Assembly 注册值全部为 `2.2.0.0`。
7. 全仓库版本搜索已执行；历史验证文档、历史变更日志保留原版本。`Product.wxs` 已通过 `$(var.ProductVersion)` 使用 VERSION.txt 注入值，无硬编码旧版本，因此未修改模板。
8. `git diff --check` 通过。

本轮未安装到 Office，也未连接真实模型执行三宿主 Ribbon 人工验收；自动测试验证了请求、排版决策、标记隔离、编译后的 Ribbon 回调与结果展示，不等同于真实 Office 文档上的端到端验收。

## 改动文件

新增：

- `src/OfficeTranslate.Core/TranslationClient.Recognition.cs`
- `src/OfficeTranslate.Core/ImageRecognitionText.cs`
- `src/OfficeTranslate.WordAddIn/WordRecognitionService.cs`
- `src/OfficeTranslate.ExcelAddIn/ExcelRecognitionService.cs`
- `src/OfficeTranslate.PowerPointAddIn/PowerPointRecognitionService.cs`
- `src/OfficeTranslate.WordAddIn/OcrResultForm.cs`
- `tests/OfficeTranslate.Tests/PureOcrTests.cs`
- `tests/OfficeTranslate.Tests/PureOcrHostTests.cs`
- `tests/VALIDATION-2.2.0-OCR.md`

修改：

- `src/OfficeTranslate.Core/TranslationClient.cs`：仅声明 partial。
- Word、Excel、PowerPoint 各自的 `*TranslationService.cs`：仅声明 partial。
- `src/OfficeTranslate.WordAddIn/AddIn.cs`、`src/OfficeTranslate.AddInShared/AddIn.cs`：新增 OCR 回调与运行入口，旧入口保留。
- `src/OfficeTranslate.WordAddIn/Ribbon.xml`、`TranslationProgressForm.cs`、`AboutForm.cs`。
- Excel、PowerPoint 各自的 `.csproj`：链接共享 OCR 摘要窗体。
- Word、Excel、PowerPoint 各自的 `Properties/AssemblyInfo.cs`。
- `README.md`、`VERSION.txt`、`build.ps1`。

## 复现和本机证据

```powershell
.\build.ps1 -Configuration Release
# 以下全量测试须在交互桌面运行，包含原有窗体用例。
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=all-desktop.trx' --results-directory TestResults\ocr-2.2.0
```

本机证据位于忽略目录 `TestResults/ocr-2.2.0/`：

- `build.log`、`all-desktop.trx`、`desktop-test.log`、`desktop-test.exit`、`desktop-context.txt`。
- `frozen-translation-path.txt`：旧路径全文比对及 SHA-256。
- `version-scan.log`、`verify-version.ps1`：DLL/MSI/注册版本核验。
- `all.trx`、`all-tests.log`：非交互首轮的环境失败记录。

最终 UI 风格验证证据位于 `TestResults/ocr-result-style/`：`build.log`、`all-desktop.trx`、`desktop-test.log`、`desktop-test.exit`、`desktop-context.txt`。测试于 2026-10-14 06:34 完成，退出码 0，Session 1 / UserInteractive=True。

本地构建产物：`artifacts/OfficeTranslate.Office.x64.msi`。
SHA-256：`0EE4E239ECFECC482AE8F4FC1CE3691758D8334650FC463CC8FE7074D3809E90`。
