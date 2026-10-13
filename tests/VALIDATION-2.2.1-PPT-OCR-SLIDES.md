# OfficeTranslate 2.2.1：PowerPoint OCR 支持选中幻灯片

日期：2026-10-14。基线 `main / 6b3c029`；本地分支 `fix/ppt-ocr-slides-2.2.1`。
项目四段版本为 `2.2.1.0`。本轮未 push、合并、发布或安装。

## 行为与范围

- `PowerPointRecognitionService.cs` 的 OCR 图片收集新增 `ppSelectionSlides`：遍历选中的 `SlideRange`，逐页遍历 `Shapes`，复用旧图片翻译的 `AddImageShape` 过滤图片、链接图片以及组合内图片。
- 单页和多页选区均收集为任务开始前的图片列表；忽略空页及非图片形状，不读取未选中的幻灯片。
- 每张图片继续走原有 `RecognizeImageAsync` → 原文覆盖框 → 旁注降级；这部分方法与基线全文相同。
- 原图形选区及无图片提示「请先选中一张图片」保持不变。
- Word/Excel OCR、三个宿主旧翻译服务、共享 Ribbon 回调、Core 翻译和识别请求代码均与基线相同。
- README、VERSION.txt、三个宿主 AssemblyInfo、构建版本示例及共享 About 更新日志同步版本；About 新增 2.2.1 条目并保留 2.2.0 历史。

## 验证结果

1. `build.ps1 -Configuration Release` 成功，**0 错误**，7 个已有可空性警告；MSI 生成成功。
2. 新增 **8/8** 回归用例通过。测试用严格的 Office 接口替身调用实际编译后的 PowerPoint OCR 收集器与预检查：单页、含多张图片的多页、空页、非图片过滤、原图形选区、多页无图片、无选区、文字选区、无活动窗口及 COM 失败。
3. 交互桌面全量测试 **417/417 通过，0 失败，0 跳过**；Session 1 / `UserInteractive=True`，退出码 0。临时计划任务已删除。
4. Core、Word、Excel、PowerPoint 四个 DLL 的 AssemblyVersion/FileVersion，MSI ProductVersion、File 表中的四个产品 DLL 版本及三项 COM Assembly 注册版本均为 `2.2.1.0`。
5. 全仓库版本搜索已执行，旧版本仅保留于历史验证记录及历史更新日志。Core 版本来自 `src/Directory.Build.props`；安装器从 VERSION.txt 注入版本，无需修改模板。
6. 冻结路径全文比较通过；`git diff --check` 通过。

未安装候选加载项，也未连接真实模型在 PowerPoint 中执行人工端到端验收。新增自动测试验证实际编译的选区遍历和图片过滤，现有纯 OCR 测试验证请求、原文输出和降级逻辑。

## 复现与证据

```powershell
.\build.ps1 -Configuration Release
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --filter FullyQualifiedName~PowerPointOcrSelectionTests --logger 'trx;LogFileName=selection.trx' --results-directory TestResults\ppt-ocr-slides-2.2.1
# 全量测试在交互桌面运行（已有窗体测试要求交互会话）
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=all-desktop.trx' --results-directory TestResults\ppt-ocr-slides-2.2.1
```

本机证据目录 `TestResults/ppt-ocr-slides-2.2.1/`（git 忽略）：

- `build.log`、`selection-test.log`、`selection.trx`。
- `desktop-test.log`、`all-desktop.trx`、`desktop-test.exit`、`desktop-context.txt`。
- `version-verification.log`、`version-scan.log`、`frozen-paths.log`。

本地 MSI：`artifacts/OfficeTranslate.Office.x64.msi`。
SHA-256：`4529FCD9C3EEBD9DFE71B1B6B760F54881709C84052C3BFF32BEA42B881995B0`。
