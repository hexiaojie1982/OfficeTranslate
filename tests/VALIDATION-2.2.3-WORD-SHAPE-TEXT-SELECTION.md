# OfficeTranslate 2.2.3：Word 选中图形文本翻译

日期：2026-10-14。基线 `main / 925cee4 / v2.2.2`；本地分支 `fix/word-shape-text-selection-2.2.3`。
项目四段版本为 `2.2.3.0`。未 push、合并、发布或安装候选加载项。

## 修改文件与行为

- `src/OfficeTranslate.WordAddIn/WordTranslationService.cs`：选区收集正文及 `Selection.ShapeRange` 中有可翻译文字的 Shape，跳过空框、不支持 TextFrame 的图形及纯对象标记。Shape 选中模式不把锚点正文误当成选中文字；正文与图形混选时两边都收集。框内局部文字选中时避免重复翻译整个框。OCR 标记不参与文字筛选。
- 同一文件：Target 保存对应 Shape；现有 UI 线程、取消检查和撤销记录内调用统一写回方法。Shape 替换写入 `TextFrame.TextRange.Text`，恢复首字符的基础 Font 和 ParagraphFormat；双语模式在框内追加译文，保留原文。正文及全文的 Range 写回行为保持一致。
- `README.md`、`VERSION.txt`、`build.ps1`、Word/Excel/PowerPoint 的三个 `Properties/AssemblyInfo.cs`、共享 `AboutForm.cs`：同步版本和更新日志，保留历史版本记录。Core 版本由已有 `src/Directory.Build.props` 读取 VERSION；About 由各宿主程序集读取版本；MSI 及 COM 注册版本由已有构建参数派生。
- `tests/OfficeTranslate.Tests/WordShapeTextSelectionTests.cs`：新增 11 个回归用例，使用严格 Office 接口替身直接执行实际 Word DLL 的收集和写回方法。覆盖空 Selection.Range 的文本框、独立 Shape 写回和格式恢复、多图形与不支持文本框的图形、空白/对象标记跳过、混选正文及文本框、普通正文、框内局部选中去重，以及正文/文本框双语追加。
- `tests/RunWordShapeTextSelectionProbe.ps1`：真实 Word COM 探针，仅创建并关闭临时未保存文档，使用本机 HTTP 模型替身调用候选 DLL 的完整 `TranslateAsync`；无需安装候选加载项。
- 本验证记录。

冻结路径比较通过：Word 图片捕获、翻译及覆盖写回链路与 main 完全相同；OCR、Core、Excel、PowerPoint 业务代码未改动，仅同步宿主程序集版本。

## 验证结果

1. `build.ps1 -Configuration Release` 成功，**0 错误**，本地 MSI 已生成。
2. 交互桌面全量测试 **465/465 通过，0 失败，0 跳过**，退出码 0；Session 1 / `UserInteractive=True`。新增回归 **11/11** 通过。
3. 真实 Word COM 验证三项通过：
   - 选中带真实 OCR 标记的浮动文本框 → 完整选区翻译 → 本机 HTTP → 对应框写回；摘要 Total=1、Translated=1。保留 19pt、粗体、东西文字体配置及段后距，正文和 OCR 标记保持原值。
   - 同时选中文本框、带字 AutoShape、空文本框：只产生两个文字目标，分别写回各自文本框，正文未改动。
   - 双语译文追加在选中的文本框内，19pt 保留。
4. Core、Word、Excel、PowerPoint 四个 DLL 的 AssemblyVersion/FileVersion，三个 About 的 CurrentVersion，MSI ProductVersion、四个产品 DLL 的 File 表版本及三项 COM Assembly 注册版本均为 `2.2.3.0`。
5. 版本全仓扫描及 `git diff --check` 通过；旧版本仅保留在历史记录/链接和历史更新日志。测试临时计划任务已删除，临时 Word 文档关闭且未保存。

格式保留边界：替换整框时保留基础字体、字号和段落格式；复杂混合文字样式无法与译文逐字符对应，采用首字符的基础格式。双语追加保留已有原文格式。本轮使用真实 Word 和本机 HTTP 替身，未连接用户的真实模型，也未执行功能区人工点击验收。

## 复现命令与证据

```powershell
.\build.ps1 -Configuration Release
# 以下测试在已登录的交互桌面执行：
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=all-desktop.trx' --results-directory TestResults\word-shape-text-2.2.3
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File tests\RunWordShapeTextSelectionProbe.ps1
```

本机证据目录 `TestResults/word-shape-text-2.2.3/`（git 忽略）：`build.log`、`shape-regression.log`、`shape-regression.trx`、`desktop-test.log`、`all-desktop.trx`、`desktop-test.exit`、`desktop-context.txt`、`word-com-probe.log`、`word-com-probe.exit`、`version-verification.log`、`version-scan.log`、`frozen-paths.log`。

本地 MSI：`artifacts/OfficeTranslate.Office.x64.msi`。
SHA-256：`262A3BA1D826D2DDC4BBA1B617482E64B7C270D22BECD1469484FA2A8CB5848C`。
