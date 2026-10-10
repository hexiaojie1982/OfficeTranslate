# OfficeTranslate 2.1.32 原图捕获候选验证

## 范围与结论

- 基线：`8b8b252f8b061e289a397ffdfa3808dc35cff2fe`（2.1.31）。候选分支：`fix/ocr-original-image-all-hosts`。
- VERSION.txt = 2.1.32，三宿主 AssemblyVersion/AssemblyFileVersion = 2.1.32.0；共用 About 增加独立 2.1.32 条目，保留 2.1.31 与更早历史。
- 完整 Release 重建通过：0 错误，7 个既有 nullable 警告，新增代码无警告。
- 原图阶段全量单测：238/238 通过；本轮超时补充后的最新全量结果为 258/258（详见文末新增章节）。基线 189 项全部保留；原图阶段新增 20 个测试方法、49 个参数化执行用例。
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

原图提交 `05deaf7` 的 TRX：`TestResults/ocr-original-image/2.1.32-original-final3.trx`（测试产物已被仓库忽略）。238 executed/passed；failed/error/timeout/aborted/notExecuted 均 0。该阶段运行耗时约 32 秒；最新超时补充结果见文末。

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

原图提交 `05deaf7` 的本地 Release DLL SHA-256（历史证据，不代表超时补充后的 DLL；最新哈希见文末）：

| DLL | SHA-256 |
| --- | --- |
| Core | 257781E0CF14DAEBBCDF078F576E76CF9D3ACBD8A0895443B5F2318558C820EA |
| Word | 886A6B4CEB0BC6554E90D6F78142FA611881BE140379DA002C7E3D102BB2CE3A |
| Excel | B9BB58826522B74BF135D6ECFA8AF215A3761233C1B3B93A9206D428DA6B7952 |
| PowerPoint | 39140E3BF3549AA64320E5DEA9CB595D63D997DFFE0B30FDB6230D4186AC0039 |

## 原图提交 05deaf7 的改动文件（超时补充清单见文末）

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

## 2.1.32 补充：图片请求长超时与可见失败（独立本地提交）

### 问题与现行源码复核

用户在 PPT 上观测到密集照片任务等待约 364 秒后消失；`ocr_summary` 为 `parsed=False`、计数 -1、`finish_reason=unknown`、`status=cancelled`。计数 -1 表示尚未取得/解析响应，不是实际识别为零。

任务背景中“裸 HttpClient 默认 100 秒、未配置 Timeout”与本分支现行源码不完全一致：`Configure` 已设置 `settings.TimeoutSeconds`，`TranslationSettings` 默认 120 秒。三次 120 秒加 1/3 秒退避约 364 秒，与观测相符；未做公司现场 HTTP 抓包，不能把这一时间吻合写成逐次请求实测。确认的代码缺口是图片和文本共用短超时、非用户 TaskCanceledException 被重试，最终 OCE 被无条件按用户取消静默处理。

### 最终数值、机制与策略

- 图片：内部常量 `ImageRequestTimeoutSeconds = 600`。每次图片 HTTP 发送创建独立 deadline CTS（CancelAfter 600 秒），与用户 token 通过 linked CTS 合并；覆盖完整 `ResponseContentRead` 的非流式 JSON 响应体，而非只等响应头。
- 使用单独 `_imageHttp`，仅其底层 HttpClient.Timeout 为 Infinite；实际请求仍有上述 600 秒期限。没有将文本 `_http` 的超时改为 600 秒或无限。
- 文本/模型列表：仍用 `_http` 与原 `Configure(settings)`；默认 120 秒/次，实际服从既有 `TimeoutSeconds` 配置。文本超时仍最多三次尝试，原 1/3 秒退避未改（默认累计约 364 秒，不含额外处理）。
- 图片本地超时：转换为 `ImageOcrTimeoutException : TimeoutException`，立刻退出，不会进入 TaskCanceledException 的重试分支，也不会再次启动整图生成。
- 用户主动取消优先：用户 token 已取消时保留 OperationCanceledException，既不重试也不弹窗。Deadline 与用户取消同时被观测时，用户取消判定优先。
- HTTP 429、500/502/503/504 仍沿用原重试、最多三次及 Retry-After（最多 30 秒）策略；原 HttpRequestException 瞬时网络失败重试也保留。任一尝试发生本地图片超时即终止，后续不重试。600 秒是单次 HTTP 等待期限，不是含 HTTP 重试/结构化回退的整项总时限。
- 400/422 结构化参数不支持仍可回退一次；回退请求同样有独立图片期限，若它超时不会继续重新生成。
- 服务端自发取消但用户 token/deadline 都未触发时，显式失败而非伪造“600 秒后超时”或“用户取消”。包裹在 HttpRequestException 内的取消也按真实 token 分类。
- 两个 HttpClient 均释放；单测共享 fake handler 时仅有一方持有其释放责任。CTS、失败后已取得的 response 也释放。

### 日志与三宿主提示

| 情况 | ocr_summary status | UI |
| --- | --- | --- |
| 本地图片等待期限耗尽 | timeout | 共用 TranslationErrorDialog，警告图标、确定按钮，明确说明等待秒数与下一步 |
| 用户 token 主动取消 | cancelled | 不弹窗，更新“已取消”状态并按既有 finally 关闭进度 |
| 非用户/非 deadline 的运输取消或其他错误 | failed | 可见错误，不伪装为取消 |

图片超时提示：

> 图片识别请求等待 600 秒后超时，服务器/模型未返回结果，未写入当前图片的任何内容；此前已完成的内容保留。可重试、减小图片或改用更快的图片模型。

只承诺当前失败图片未写入，不谎称整项任务此前也没有写回。Word 和 Excel/PPT 两处宿主入口共用 `TranslationFailurePolicy.IsUserCancellation`，只有 OCE 且用户 token 已取消才静默；最终文本超时也不再误判为用户取消（网络超时/重试数值不变）。三宿主共用同一链接编译的错误提示组件，保留各自 UI 调度及活窗口 owner 校验。About 的当前 2.1.32 条目补充该变更；版本号不变。

日志沿用既有元数据摘要，没有新增图片/正文/Prompt/密钥记录；没有把内部 timeout 异常的 InnerException 正文写入诊断。

### 新增单测与真实执行结果

`ImageRequestTimeoutTests`：新增 12 个测试方法、20 个执行用例；假 handler + 注入 80ms/5s 小期限，不访问真实服务。

- ImageDeadline_IsVisibleTimeoutNotCancellation_AndNeverRetries（OpenAI-compatible/Ollama）
- InFlightUserCancel_IsSilentCancellationAndNeverRetries（两 Provider）
- SlowImageResponse_IsNotLimitedByConfiguredTextTimeout（两 Provider，文本设 1 秒，图片 1.25 秒仍成功）
- ImageClient_DoesNotExtendTextTimeoutOrChangeItsThreeAttemptRetry（先图片成功，再三次文本 1 秒超时）
- TransientImageHttpStatus_StillRetriesAndCanSucceed（429/500/502/503/504）
- TimeoutAfterTransientHttpResponse_StopsWithoutThirdAttempt
- StructuredFallbackTimeout_DoesNotRestartImageGeneration
- NonUserTransportCancellation_IsVisibleFailure_NotSilentCancelOrRetry
- UserCancelDuringHttpBackoff_DoesNotStartAnotherAttempt
- WrappedTransportCancellation_PreservesTimeoutOrUserCancelWithoutRetry（超时/主动取消）
- DeadlineCoversResponseBody_NotOnlyHttpHeaders（延迟 Content）
- SharedFailurePolicy_RequiresUserTokenAndKeepsDefaultBudgets

实际执行（目录仍为 native-word，.NET SDK 8.0.423；不调用 build.ps1）：

```powershell
$env:DOTNET_CLI_HOME = Join-Path (Get-Location) 'dotnet-home'
$env:NUGET_PACKAGES = Join-Path (Get-Location) 'packages'
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1'
$env:DOTNET_CLI_TELEMETRY_OPTOUT='1'
& 'C:\Program Files\dotnet\dotnet.exe' build OfficeTranslate.Word.sln -c Release --no-restore -t:Rebuild
& 'C:\Program Files\dotnet\dotnet.exe' test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=2.1.32-timeout-final.trx' --results-directory TestResults\ocr-original-image
```

最新完整 Rebuild：退出 0，1.87 秒，0 错误/7 个原有警告，新增代码无警告。全量单测：退出 0，258/258，失败/跳过 0，约 43 秒。TRX：`TestResults/ocr-original-image/2.1.32-timeout-final.trx`；此前专项初轮 18/18 通过，最终新增 20 项全部纳入全量。既有 238 项（含原图提取与 EXIF）全部通过。

当前补充后 Release DLL SHA-256：

| DLL | SHA-256 |
| --- | --- |
| Core | 0544C837038497A72B274551EE12C1015BCCF17DBDBE3BDA490F9C8B7BD349C2 |
| Word | 1F9A3643A1A683E52B6E559E057C52DBED2E59A30AC65771B69741001DE9F41B |
| Excel | 974C6298207A938FD63384E6412E019BE5DC0EB74151BC896AD4DC372886B3C4 |
| PowerPoint | 347364E7C09F3C371050E61939F0C13A52DB22F3F44EC1657D35C0BEE5D37310 |

本补充相对 `05deaf7` 的改动文件：

- src/OfficeTranslate.Core/TranslationClient.cs
- src/OfficeTranslate.Core/ImageOcrTimeoutException.cs（新增，含共用取消分类）
- src/OfficeTranslate.AddInShared/AddIn.cs
- src/OfficeTranslate.WordAddIn/AddIn.cs
- src/OfficeTranslate.WordAddIn/TranslationErrorDialog.cs（新增，三宿主链接共用）
- src/OfficeTranslate.WordAddIn/AboutForm.cs
- src/OfficeTranslate.ExcelAddIn/OfficeTranslate.ExcelAddIn.csproj
- src/OfficeTranslate.PowerPointAddIn/OfficeTranslate.PowerPointAddIn.csproj
- tests/OfficeTranslate.Tests/ImageRequestTimeoutTests.cs（新增）
- tests/VALIDATION-2.1.32-ORIGINAL-IMAGE.md（本节）

### 补充未覆盖项与交付边界

- 没有等待真实 600 秒；单测验证同一生产 deadline/分类机制，注入小期限缩短耗时。
- 三宿主实际菜单的 warning 弹窗、owner/焦点及取消按钮交互没有现场验证；完成的是三宿主构建、共享 UI 接线静态核对及 Core 离线测试，不能称真机弹窗已通过。
- 公司 Qwen3.8-27B/密集照片完整结果、模型服务/代理自身更短超时、网络故障和取消后服务端是否停止计算未实测。600 秒不保证该模型必能完成，也不保证视觉 Token/服务端算力节约。
- 普通文本 120 秒默认保持，用户自定义 TimeoutSeconds 的实际值可能不同；本轮没有修改 settings.json。
- 原图捕获、覆盖框/旁注、已保存/dirty 回退边界不改；上节列出的实际 Office/公司验收缺口仍然存在。
- 同一 fix/ocr-original-image-all-hosts 分支追加独立本地提交，仍为 2.1.32；不 push/合并/tag/安装/发布/生成 MSI。
