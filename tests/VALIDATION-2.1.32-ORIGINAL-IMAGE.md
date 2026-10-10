# OfficeTranslate 2.1.32 原图捕获候选验证

## 范围与结论

- 基线：`8b8b252f8b061e289a397ffdfa3808dc35cff2fe`（2.1.31）。候选分支：`fix/ocr-original-image-all-hosts`。
- VERSION.txt = 2.1.32，三宿主 AssemblyVersion/AssemblyFileVersion = 2.1.32.0；共用 About 增加独立 2.1.32 条目，保留 2.1.31 与更早历史。
- 完整 Release 重建通过：0 错误，7 个既有 nullable 警告，新增代码无警告。
- 最终全量单测：238/238 通过，失败/跳过 0；基线 189 项全部保留，新增 20 个测试方法、49 个参数化执行用例。
- 仅本地候选提交。不 push、不合并、不打 tag、不安装、不生成 MSI、不发布；没有改动 main、2.1.30/2.1.31 分支或 review-image-overlay 工作区。
- 没有启动/操纵真实 Office、没有公司服务器/Qwen/公司照片测试。以下包与像素结果不能等同于公司环境 OCR 召回率或速度验收。

## 根因复核及实现

原路径确认：Word 2.1.31 仅行内图原图提取；浮动图 CopyAsPicture 回退。Excel 用屏幕 Clipboard Copy/CopyPicture；PowerPoint 仅无尺寸 Shape.Export。后两者不能保证发送源文件像素，而放大屏幕捕获不是稳定的原图证明。

共享 `OriginalImageCapture` 在 Core 内完成图片变换守卫、只读 ZIP 关系定位、栅格解码、EXIF 校正、透明度检查、长宽比及 PNG 编码。Word 的 Flat OPC 定位保留，复用共享守卫/编码器；EXIF/透明度函数机械迁至共享类，没有另写三套算法。Core 不依赖 Office interop。

### Word

- 行内图：仍只读 `Range.WordOpenXML`；要求唯一 inline/pic。
- 浮动图：只读 Anchor.WordOpenXML；要求唯一 `wp:anchor/wp:docPr@name == Shape.Name` 且该 anchor 恰好一个 pic。锚点相对页/段落的定位不改变图片内部像素，因此模型 0–1000 bbox 仍相对当前图片矩形；覆盖框定位与参考系逻辑未修改。
- 同名歧义、名称未命中、Anchor XML 未包含目标、分组/AlternateContent、非零 effectExtent 或相对尺寸，均回退已有 CopyAsPicture。只读 XML 路径不改图片尺寸、文档、选区或视图。
- 保留全图预捕获与剩余/累计 PNG 预算，不把写回插入捕获阶段。
- 限制：Flat OPC 仍要求 `/word/document.xml` 与对应关系；不是该结构时回退，不猜其他 Story 部件。

### Excel

- 以图片自身 Parent 工作表和工作簿为起点；必须有保存路径且 `Workbook.Saved == true`。已有磁盘文件但有未保存修改也回退，避免读取旧图。不会触发保存或设置 Saved。
- 支持 `.xlsx/.xlsm/.xltx/.xltm`；用 `FileStream(FileMode.Open, FileAccess.Read, FileShare.ReadWrite)`，不是默认独占打开。
- 通过 workbook.xml 中工作表名称 → worksheet 关系 → 唯一 drawing 关系 → 该 drawing 内唯一图片名称 → embed 图像关系 → media 部件。绝不按 sheet/drawing 序号猜。
- 找不到、歧义、多个 drawing、旧 `.xls/.xlsb`、外链、加密/损坏包、未知格式/变换等，回退原有 Copy/CopyPicture 捕获。
- 图片写回绑定捕获时的工作表，避免网络等待期间切表后在另一工作表创建译文框。正文单元格翻译未修改。
- 第一次图片写回通常会使工作簿变为未保存状态；同一任务后续图片可能走回退。这是避免磁盘旧状态的保守限制，未通过保存、修改 Saved 或磁盘缓存绕过。

### PowerPoint

- 已保存且 `Presentation.Saved == msoTrue` 的 `.pptx/.pptm/.potx/.potm/.ppsx/.ppsm`：通过 presentation.xml 的 SlideID → slide 关系 → 本 slide 内唯一 Shape.Name → embed/media。不是根据 SlideIndex 猜文件名。
- 原图不可证明时，优先按图片尺寸、2560 长边目标与 12 倍显示倍率上限计算均匀缩放，调用带尺寸的 Export；失败才无尺寸 Export。裁剪/翻转/旋转等回退仍由 Office 渲染最终外观，不对模型 bbox 二次镜像。
- 本机 .NET Framework 只读反射确认签名：`Void Export(String, PpShapeFormat, Int32, Int32, PpExportMode)`，ScaleWidth/ScaleHeight 默认 0，ExportMode 默认 ppRelativeToSlide。
- 参数实际是相对整张幻灯片的缩放尺寸（点），不是直接的图片像素。换算示例：180×135pt 图片、960×540pt 幻灯片，参数 10240×5760pt；按 96DPI 理论目标 PNG 为 2560×1920。这是纯计算与文档语义核验，未实测 COM 导出像素。
- 依据：[Microsoft Shape.Export 文档](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.export)。不采用文档标为保留用途的 ppScaleXY/ppScaleToFit。
- 导出文件使用唯一临时路径，重试先删本次不完整文件，finally 清理。读取前检查字节数及完整 33 字节 IHDR 的像素限制；容量违规明确停止，不靠低清重试绕过限制。
- 占位图片可能继承 layout/master 的几何或效果，原图分支无法独立证明，因此回退渲染。写回后变 dirty 的后续图片也保守回退。

### 三宿主共用守则与诊断

- 只接受内嵌 PNG/JPEG/BMP；关系外链、未知扩展、分组、裁剪 srcRect/fillRect 非零、翻转/旋转、效果/3D/样式/描边及非矩形均回退。
- EXIF 缺失/1 不变，2–8 在解码像素上校正一次；非法/不可解析值回退，输出新 PNG 不继承旧方向标记。
- 校正后的源像素与 Office 矩形长宽比偏差 >1% 回退。缩放复用 ImageCaptureResolution.Plan：源像素不无意义上采样，长边 2560 与显示倍率上限 12。
- 源图及导出像素限制 8000/边、3200 万像素；单 PNG 128MiB，Word 还保留累计 512MiB。它们不是 Office 总进程内存或 OOM 保证。
- ZIP XML 部件限制 8MiB、禁 DTD/外部解析器；重复 ZIP 部件或关系不猜测；media 解压读取也限额。未写 OOXML 源文件。
- 每次图片请求的 `ocr_summary` 增加 `capture=original/fallback` 和 `capture_reason`，与原区域统计、PNG 尺寸、finish_reason/truncated 一起记录；0 区域与失败仍记录。原因限定短元数据 token，不记录图片、正文、Prompt、密钥或文件路径。
- 正文翻译请求、OCR 请求 Token 预算、语言规则、覆盖框/旁注字号和安全原则未修改。

## 测试证据

最终 TRX：`TestResults/ocr-original-image/2.1.32-original-final3.trx`（测试产物已被仓库忽略）。238 executed/passed；failed/error/timeout/aborted/notExecuted 均 0。最终运行耗时约 32 秒。

合成 ZIP 使用工作表 Data → sheet17 → drawing23，以及 SlideID 701 → slide9，刻意不匹配序号。内嵌不对称 JPEG 方向 1–8 的输出与已验证 Word 路径逐字节一致。3200×2400 JPEG 在 Excel/PPT Core 路径都编码为 2560×1920；Word 浮动 EXIF 6 合成 anchor 输出 80×120 且原 XML 未变。它们是 Core 层测试，不是实时菜单测试。

`.xlsx/.xlsm/.pptx/.pptm` 合成文件保持 ReadWrite 句柄占用时仍可提取，且文件内容前后逐字节不变。另测原图预算超限不会被 fallback 绕过、透明/损坏/DTD/歧义等明确回退，以及导出文件字节/IHDR/像素限制。

### 新增测试方法（49 个执行用例）

`OriginalImageAllHostsTests`（16 方法，43 用例）：

- PackageIdentityAndPictureName_SelectRelationshipNotPartIndex
- UnsafeOrAmbiguousPackagePicture_RequiresRenderedFallback
- PackageExif_AllOrientationsReuseWordCorrectedPixels
- PackageInvalidExifAndAspectMismatch_AreExplicitFallbacks
- PackageMissingSheetOrSlide_DoesNotGuessAnotherPart
- MultipleWorksheetDrawings_AreNotGuessed
- PackagePngBudgetFailure_IsNotBypassedByFallback
- DirtyUnsavedAndLegacyDocuments_DoNotReadDisk
- SavedPackageFile_IsReadableWhileHostKeepsWriteSharing
- FloatingWordAnchor_UniqueNamedPictureUsesOriginalWithoutMutation
- FloatingWordAnchor_MismatchAmbiguityAndTransformsFallBack
- PowerPointExportScale_UsesSlideRelativePointDimensions
- OcrSummary_CaptureMetadataIsExplicitAndRejectsContent
- PackageLargeOriginal_UsesShared2560ResolutionBudget
- PackageXmlDtdCorruptRasterAndTransparency_AreExplicitFallbacks
- KnownDpiHint_StillUsesOriginalPixels

`ExportPngBudgetTests`（4 方法，6 用例）：

- ExportPng_CompleteIhdrIsAcceptedWithoutChangingBytes
- ExportPng_ByteLimitAndZeroBudgetAreExplicitFailures
- ExportPng_TruncatedHeaderAndInvalidSignatureAreRejected
- ExportPng_PixelBudgetIsCheckedBeforeReadingPayload

`WordOriginalImageCaptureTests.PatternJpeg` 仅由 private 改 internal 复用夹具；既有 EXIF 不对称像素位置用例没有改期待值。

## 实际执行命令

目录：`D:\pythonProject\office-translate\native-word`，PowerShell，.NET SDK 8.0.423 / net48。

```powershell
$env:DOTNET_CLI_HOME = Join-Path (Get-Location) 'dotnet-home'
$env:NUGET_PACKAGES = Join-Path (Get-Location) 'packages'
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1'
$env:DOTNET_CLI_TELEMETRY_OPTOUT='1'
& 'C:\Program Files\dotnet\dotnet.exe' restore OfficeTranslate.Word.sln --ignore-failed-sources
& 'C:\Program Files\dotnet\dotnet.exe' build OfficeTranslate.Word.sln -c Release --no-restore -t:Rebuild
& 'C:\Program Files\dotnet\dotnet.exe' test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=2.1.32-original-final3.trx' --results-directory TestResults\ocr-original-image
```

Restore：全部最新，退出 0。最后 Rebuild：退出 0，1.31 秒，0 错误/7 警告；Tests：退出 0，238/238。没有执行 build.ps1，因为该脚本会生成 MSI。

既有警告：OfficeUiDispatcher.cs:134 CS8601；WordTranslationService.cs:77 CS8600；WordSelectionProbe.cs:71 CS8602；ImageOverlayTextTests.cs:14、ReviewFixesTests.cs:49、StagingNameContractTests.cs:52/94 CS8625。

最后本地 Release DLL SHA-256：

| DLL | SHA-256 |
| --- | --- |
| Core | 257781E0CF14DAEBBCDF078F576E76CF9D3ACBD8A0895443B5F2318558C820EA |
| Word | 886A6B4CEB0BC6554E90D6F78142FA611881BE140379DA002C7E3D102BB2CE3A |
| Excel | B9BB58826522B74BF135D6ECFA8AF215A3761233C1B3B93A9206D428DA6B7952 |
| PowerPoint | 39140E3BF3549AA64320E5DEA9CB595D63D997DFFE0B30FDB6230D4186AC0039 |

## 改动文件

- VERSION.txt
- src/OfficeTranslate.Core/ImageCaptureBudget.cs
- src/OfficeTranslate.Core/ImageCaptureResolution.cs
- src/OfficeTranslate.Core/ImageOverlayDiagnostics.cs
- src/OfficeTranslate.Core/OfficeTranslate.Core.csproj
- src/OfficeTranslate.Core/OriginalImageCapture.cs（新增）
- src/OfficeTranslate.Core/TranslationClient.cs
- src/OfficeTranslate.Core/WordOriginalImageCapture.cs
- src/OfficeTranslate.ExcelAddIn/ExcelTranslationService.cs
- src/OfficeTranslate.ExcelAddIn/Properties/AssemblyInfo.cs
- src/OfficeTranslate.PowerPointAddIn/PowerPointTranslationService.cs
- src/OfficeTranslate.PowerPointAddIn/Properties/AssemblyInfo.cs
- src/OfficeTranslate.WordAddIn/AboutForm.cs
- src/OfficeTranslate.WordAddIn/Properties/AssemblyInfo.cs
- src/OfficeTranslate.WordAddIn/WordTranslationService.cs
- tests/OfficeTranslate.Tests/OfficeTranslate.Tests.csproj
- tests/OfficeTranslate.Tests/WordOriginalImageCaptureTests.cs
- tests/OfficeTranslate.Tests/OriginalImageAllHostsTests.cs（新增）
- tests/OfficeTranslate.Tests/ExportPngBudgetTests.cs（新增）
- tests/VALIDATION-2.1.32-ORIGINAL-IMAGE.md（本报告）

## 未覆盖项与后续真实验收

- Word 浮动图的真实 Anchor.WordOpenXML 是否含正确 docPr 名称、Office 各版本 EXIF 显示方向、相对定位/页眉等 Story：未实测；未命中即回退。
- Excel/PPT 真实已保存包的 COM Parent/Name/Saved 对应及高分辨率接线：只完成编译与 Core 测试，真实菜单未验。
- PowerPoint 带尺寸 Export 的实际像素、裁剪/翻转/旋转/效果图外观与坐标：纯换算及 PIA 签名已验，真实渲染未验。
- 公司文档加密、云端/未完全下载文件、并发磁盘保存和第三方修改：未实测；无法读取/证明时回退，不触发保存。
- 同任务多图在首图提交后 dirty 的质量退化尚存在保守回退；未实现全任务磁盘原图快照，因为需要额外证明每张图在其提交时仍对应当前包。
- 真实公司 Qwen3.8-27b/服务器/密集照片的 OCR 完整率、延迟、模型输入视觉 Token 和费用：均未验。提高像素不保证服务端不再缩图，也可能增加视觉 Token/延迟。
- 三宿主实际图片 OCR、重翻/取消、菜单 MVID、MSI 升级安装：本轮未做；不能从 238 个单测宣称这些通过。
