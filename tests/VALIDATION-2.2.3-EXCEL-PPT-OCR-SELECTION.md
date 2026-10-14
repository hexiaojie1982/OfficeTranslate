# OfficeTranslate 2.2.3：Excel / PowerPoint OCR 文本框选区翻译

日期：2026-10-14。基线 `main / 925cee4`；本地分支 `fix/excel-ppt-ocr-shape-select-2.2.3`。
版本 `2.2.3.0`。未 push、合并、发布或安装。Word 分支的 `f5a5986` 未带入；两个分支最终合并时需整合 README 和共享 About 更新日志。

## 修改文件与行为

- `src/OfficeTranslate.ExcelAddIn/ExcelTranslationService.cs`：`AddShape` 新增可选参数 `includeOcrOverlays = false`。显式图形选区传 `true`；整表扫描沿用默认值，继续按 AlternativeText 跳过 OCR 覆盖框。组合图形递归透传参数。
- `src/OfficeTranslate.PowerPointAddIn/PowerPointTranslationService.cs`：同样新增参数；仅 `ppSelectionShapes` 传 `true`，组合图形递归透传。选中幻灯片和整份演示文稿仍按 Tags 跳过 OCR 覆盖框。
- `README.md`、`VERSION.txt`、`build.ps1`、三个宿主的 `Properties/AssemblyInfo.cs`、共享 `src/OfficeTranslate.WordAddIn/AboutForm.cs`：统一版本并新增本分支更新说明，保留历史记录。Core 版本通过现有 `src/Directory.Build.props` 从 VERSION.txt 读取，MSI 版本继续由 build.ps1 注入。
- `tests/OfficeTranslate.Tests/OcrShapeTextSelectionTests.cs`：新增 21 个数据驱动回归用例，加载实际编译的 Excel / PowerPoint DLL，通过严格 Office 接口替身调用真实收集器和写回回调。
- 本验证记录。

Word 业务代码与 main 相同。未修改 OCR 识别、图片收集、图片翻译或覆盖框写回链路。两个宿主业务代码改动仅涉及 AddShape 参数、OCR 跳过条件、显式图形选择调用点和递归透传。

## 验证结果

1. `build.ps1 -Configuration Release` 成功：**0 错误，6 个已有可空性警告**；本地 MSI 生成成功。
2. 新增回归 **21/21 通过**：显式选中 OCR 图形可收集原文并写回译文；直接图形、两层组合、带标记的组合均覆盖。整表/全文及选中幻灯片跳过覆盖框且不读取其正文，普通图形照常收集和写回。另覆盖空白文本及默认参数继续排除覆盖框。
3. 登录桌面全量测试 **475/475 通过，0 失败，0 跳过**，退出码 0，耗时约 1 分 17 秒；Session 1 / `UserInteractive=True`。
4. Core、Word、Excel、PowerPoint DLL 的 AssemblyVersion/FileVersion，三个 About 的 CurrentVersion，MSI ProductVersion、四个产品 DLL 的 File 表版本及三项 COM Assembly 注册版本均为 **2.2.3.0**。
5. 全仓库版本扫描通过；2.2.2 只保留于历史验证记录、历史链接和历史更新日志。`git diff --check` 通过。

本轮测试验证编译后的宿主目标收集和写回回调；未安装候选加载项，未连接真实模型执行 Office 人工端到端验收。

## 复现与证据

```powershell
.\build.ps1 -Configuration Release
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --filter FullyQualifiedName~OcrShapeTextSelectionTests --logger 'trx;LogFileName=regression.trx' --results-directory TestResults\excel-ppt-ocr-shape-select-2.2.3
# 现有窗体和剪贴板测试要求在登录桌面执行：
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=all-desktop.trx' --results-directory TestResults\excel-ppt-ocr-shape-select-2.2.3
```

本机证据目录 `TestResults/excel-ppt-ocr-shape-select-2.2.3/`（git 忽略）：`build.log`、`regression.log`、`regression.trx`、`desktop-test.log`、`all-desktop.trx`、`desktop-test.exit`、`desktop-context.txt`、`version-verification.log`、`version-scan.log`、`frozen-paths.log`。测试临时计划任务在完成后删除。

本地 MSI：`artifacts/OfficeTranslate.Office.x64.msi`，仅构建，未安装或分发。
