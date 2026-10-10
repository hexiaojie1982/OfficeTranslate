# 2.1.31 图片 OCR 召回修复候选：本地验证报告

日期：2026-10-10。仓库：`D:\pythonProject\office-translate\native-word`。
基线：`d23d68e6c32e8a50b37ca19775644476c83e300a`（2.1.30 候选）。
分支：`fix/ocr-recognition-recall`。本地提交哈希见交付回复及 `git log -1`。
范围：仅本地候选；不 push、合并、打 tag、生成 MSI、安装或发布。
EXIF 补修基线：`3d5ce6c9763891ed2c0ef75a73b42f129ae3f3a5`；仍为未定稿的 2.1.31 候选，在同一分支追加一个本地提交，不变更其他分支。

## 根因与证据边界

源码复核确认三个可改善点：

1. Word 现有路径使用显示尺寸的 CopyAsPicture/剪贴板 Bitmap，而非直接使用内嵌位图原始像素。用户提供的此前 COM 对照是 3200×2400 原图、180×135pt 显示时捕获 240×180；放大到 1440×1080pt 时捕获 1920×1440。这些是任务给出的既有测量，**不是本轮重新实测**。
2. 图片请求一次完成 OCR、翻译、bbox 和 JSON。OpenAI-compatible 未显式设置输出上限；Ollama 图片请求没有 options。比 Open WebUI 的“仅 OCR”任务承担更多工作，不能把全部延迟归因于低分辨率，更不能据此承诺速度会更快。
3. 旧解析器忽略非字典、非四元 bbox 和空译文项，没有返回原始/保留/丢弃统计。非空 regions 与“已翻译 1”只说明一张图有输出，**不证明完整识别**。格式正确但几何异常的旧区域原本会进入 Planner 降级，不能说旧版所有坏 bbox 都被丢弃。

本轮未访问公司服务器、Qwen3.8-27b 或公司照片；服务器默认输出长度、视觉输入缩放、推理模式及 Open WebUI 参数均未知。因此这是针对已证实输入链路和源码缺口的候选，不是公司环境识别率、时延或 Token 改善的实测结论。

## 改动与安全原则

- Word 对单个行内图片 Range.WordOpenXML 的 Flat OPC 做只读提取。只接受可证明的内嵌 PNG/JPEG/BMP、单一行内矩形，无分组/替代渲染、Office 几何裁剪/旋转/翻转、效果、外链、未知图像扩展或实际透明像素。EXIF orientation 1–8 均支持：缺失及 1 无变换，2–8 先在解码位图上校正方向，再进行透明度、长宽比、缩放及编码；其他守卫通过时走高清原图路径。非法或不可解析属性才以 `exif-orientation-invalid` 回退（包括非单个 SHORT、短字节或数值不在 1–8）；GDI+ 校正执行失败另以 `exif-orientation-transform-failed` 回退，不猜测方向。校正后的源宽高供 sourceWidth/sourceHeight 和长宽比校验使用；校正后长宽比与 Office 矩形相差超过 1% 仍回退（比 Planner 的 5% 更保守）。
- 不改原图片尺寸、内容、书签或视图，不创建草稿文档。不能证明安全的情况（含浮动图）回退既有 CopyAsPicture，并记录固定原因及捕获尺寸；容量限制异常不伪装成普通回退。
- `ImageCaptureResolution.Plan` 纯计算：不超过原始像素；长边目标 2560px，显示尺寸 96DPI 像素的最大倍率 12；8000px 单边、32MP 总像素上限。小图不插值放大，小图标不发送无意义的巨大图。
- 原图解码也拒绝超 8000px/32MP 的来源；编码继续服从单图 128MiB、Word 512MiB 累计及剩余 PNG 字节预算。**这些不是总进程内存或 OOM 保证**：WordOpenXML 字符串、base64、原始解码和临时位图有额外分配。
- bbox 继续相对最终 PNG 的 0–1000 坐标；未修改 Planner、覆盖字号 8–18pt、覆盖区域几何、1400pt 旁注上限、真实排版校验或 2.1.30 无损旁注分段原则。
- Excel/PPT **没有改捕获分辨率**：不为形式一致扩展未证实的宿主捕获风险；它们共用此次图片请求、解析、待检查旁注及摘要改善。
- OpenAI-compatible 图片请求 `temperature=0.1, max_tokens=8192`；Ollama 图片请求 `options.temperature=0.1, options.num_predict=8192`。兼容服务拒绝结构化参数时，原有一次纯提示词回退仍带同一预算。普通成功图片路径仍只有一次请求，不增加二阶段模型调用。
- 提示词要求整图逐行逐块完整扫描，包含密集小字、表格、注释、边缘文字；保留源语言、混合语言、术语表、额外要求及 JSON 字段。相邻段落可合区，但不允许摘要或遗漏行。
- `ImageRecognitionResult` 返回 Regions、Statistics、FinishReason、Truncated；纯 `ImageRegionParser` 返回 raw/kept/dropped、bbox 无效/转旁注、空译文/无可用文字丢弃、非字典和 review 数。子原因计数可重叠；总项数满足 raw=kept+dropped。
- 原文非空、译文空：保留完整原文，明确标注“未提供译文”，强制待检查旁注；Translation 保持空，不把原文当已翻译或画定位覆盖。bbox 损坏但文本可用同样强制旁注；无可用文本或非字典才丢弃，并在统计中显式显示。
- 每次进入图片模型请求的操作均在 finally 写 `event=ocr_summary`，包括 0 区域、取消、失败和截断：PNG 尺寸、区域数、已解析标志、finish_reason、truncated、elapsed_ms。不写正文、图像、Prompt 或密钥；任意服务器 finish 字符串只记录为 `other`。不可解析时计数为 -1（未知），不伪装成 0。捕获阶段失败仍走既有 `capture_failed`，未发送的图没有模型统计。
- 长度截断不将部分结果冒充完整译文写回；解析失败也明确报错。不会为提高“成功数”跳过安全检查。
- 任务摘要额外显示图片及区域统计；文本严格重试提示仅使用 TextNeedsReview，图片排版/缺译文不套用文本重试口径。来源仅有原文的图不计为已翻译，也不误计 OCR 未识别。
- VERSION.txt、三宿主 AssemblyVersion/AssemblyFileVersion、共享 About 更新日志同步到 2.1.31(.0)，2.1.30 历史条目保留。

## 实际构建、测试命令

工作目录为 native-word。实际使用的 PowerShell 命令：

```powershell
$env:DOTNET_CLI_HOME = Join-Path (Get-Location) 'dotnet-home'
$env:NUGET_PACKAGES = Join-Path (Get-Location) 'packages'
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = '1'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
& 'C:\Program Files\dotnet\dotnet.exe' restore OfficeTranslate.Word.sln --ignore-failed-sources
& 'C:\Program Files\dotnet\dotnet.exe' build OfficeTranslate.Word.sln -c Release --no-restore -t:Rebuild
& 'C:\Program Files\dotnet\dotnet.exe' test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=2.1.31-exif-final.trx' --results-directory TestResults\ocr-recognition-recall
```

- Restore：退出码 0，所有项目已还原/缓存最新，无还原错误。
- 完整 Release Rebuild（Core、Word、Excel、PPT、Tests）：退出码 0；**0 Error，7 Warning**。7 项 nullable 警告位于既有语句：OfficeUiDispatcher(134)、WordTranslationService(77)、WordSelectionProbe(71)、ImageOverlayTextTests(14)、StagingNameContractTests(52/94)、ReviewFixesTests(49)；本轮新增源文件无编译警告。
- 全量 MSTest（EXIF 补修后）：**189/189 通过，失败 0，跳过 0**。既有 141 项 + 首轮 OCR 召回新增 28 项 + EXIF 补修新增 20 项；测试约 31 秒。首轮 169/169 的记录保留在原 TRX 中。
- 最终 TRX：`TestResults/ocr-recognition-recall/2.1.31-exif-final.trx`；预验证映射 TRX：`2.1.31-exif-mapping.trx`（忽略的本地结果，不提交生成物）。补修只重建/测试，未重复执行依赖已就绪的 restore；上述 restore 是候选首轮的实际命令。
- 7 个 HTTP 测试使用内存 HttpMessageHandler，不连接真实模型服务；取消和结构化回退均有隔离测试。
- `build.ps1` 未执行（会生成 MSI）；未生成任何新安装包。

## 像素与解析测试数据

| 场景 | 输入 | 本轮结果/证据类型 |
| --- | --- | --- |
| 密集图尺寸规则 | 原始 3200×2400，显示 180×135pt | Plan 2560×1920，纯计算 |
| 原图提取和真实 PNG 编码 | **合成** Flat OPC 中 3200×2400 位图，显示 180×135pt | 生成 PNG 2560×1920，MSTest/GDI+；不是实际 Word 捕获 |
| 小源图 | 400×300，显示 180×135pt | 400×300，无插值放大 |
| 图标 | 原始 6000×6000，显示 12×12pt | 192×192，倍率限幅 |
| 大图预算 | 原始 50000×30000 / 50000×1000 / 30000×30000，测试放宽目标 | 每边≤8000，总像素≤32MP，纯计算 |
| 正常响应 | 一条完整 source/translation/bbox | raw=1、kept=1、dropped=0、review=0，全文/换行保留 |
| 原文有而译文空 | 一条有效 bbox + source | kept=1，空译文=1，强制待检查，不计已翻译 |
| 坏 bbox | 三元、越界、反向、NaN、bool、null | 可用文字完整保留为旁注；不静默丢弃 |
| 混合损坏条目 | null、1、数组、空字典、source-only、translation-only | raw=6、kept=2、dropped=4、non_dictionary=3 |
| 零区域响应 | regions=[] | raw=0、kept=0，PNG=16×12 的零统计诊断被实读验证 |
| 截断响应 | 两种服务 finish=length；有效/损坏 JSON | 明确失败；日志保留 finish/truncated，损坏 JSON 的计数未知 |

真实 Word COM 尺寸对照：**本轮未完成**。独立自测探针的初始编译发现 Word Application PIA 无 Hwnd；使用 ActiveWindow 证明实例归属的修订执行又被执行保护拒绝，要求同时证明无窗口时的新进程清理，避免误操作现有 Word。没有执行创建文档/捕获/退出命令，不将其称为 COM 通过。已有 WINWORD PID 14152 前后不变。未运行探针源文件移入忽略的 `artifacts/ocr-recall-probe/unexecuted-source.cs` 保存，不进候选提交。

## 新增单测名（合计 35 个方法/48 项，以下先列首轮 28 项）

`ImageRecognitionRecallTests`（14）：

- SmallOriginal_IsNeverUpsampled
- DenseOriginalAtSmallDisplay_Reaches2560WithSameAspect
- TinyIcon_IsBoundedByDisplayMultiplier
- HugeInput_IsBoundedBySideAndTotalPixels
- PortraitAndLandscape_PreserveAspectWithinPixelRounding
- InvalidResolutionOrShapeDimensions_AreRejected
- NormalJson_PreservesFullTextsAndCoordinates
- SourceOnly_IsKeptForReviewNeverAsTranslatedOverlay
- BrokenBbox_WithUsableTextIsDowngradedNotDropped
- NonFiniteOutOfRangeReversedAndBooleanBoxes_AreDowngraded
- NonDictionaryAndEmptyText_AreCountedWithoutLosingUsableRegions
- ZeroRegions_StillReturnsStatisticsAndSummary
- MalformedRegions_IsExplicitFailure
- OcrMetadata_DoesNotLogServerSuppliedContent

`ImageRequestRecallTests`（7）：

- OpenAiImageRequest_Has8192BudgetFullScanAndExistingLanguageRules
- OllamaImageRequest_HasMatchingOptionsAndSingleRequest
- StructuredFallback_PreservesBudgetForBothProviders
- TruncatedValidJson_IsRejectedInsteadOfWritingPartialRegions
- MalformedTruncatedJson_LogsFinishAndExplicitUnknownCounts
- EmptyImageResponse_LogsRealZeroCountsAndPixelSize
- CancelledImageRequest_DoesNotRetryOrReturnRegions

`WordOriginalImageCaptureTests`（7）：

- OriginalRasterExtraction_IsReadOnlyAndDoesNotInventPixels
- TransformsEffectsExternalAndUnknownExtensions_RequireRenderedFallback
- AspectMismatch_RequiresRenderedFallback
- ExtractionPngBudget_IsNotSilentlyBypassedByFallback
- DtdAndInvalidPackage_AreRejectedSafely
- Original3200x2400_InSyntheticFlatOpc_Produces2560x1920
- OriginalPixelsOverSafetyLimit_RequireRenderedFallback

### EXIF 补修新增（7 个方法/20 项）

- GdiRotateFlip_AllExifMappingsMatchIndependentPixelCoordinates（1–8，共 8 项）
- Exif6_CorrectsStoredPixelsAndSwapsCaptureDimensions
- Exif2_CorrectsLeftRightMirrorPixels
- OtherValidExifOrientations_PreserveCorrectViewedPixels（1/3/4/5/7/8，共 6 项）
- ExifBigEndian6_IsNormalizedByGdiAndCorrectedOnce
- InvalidExifOrientation_HasDistinctFallbackReason（0/9，共 2 项）
- OneByteExifOrientation_HasDistinctFallbackWithoutThrowing

先不改产品逻辑，单独运行 `--filter 'FullyQualifiedName~GdiRotateFlip_AllExifMappingsMatchIndependentPixelCoordinates'`（日志名 `2.1.31-exif-mapping.trx`），**8/8 通过**后才实现校正。fixture 为 GDI+ 实际编码的 120×80 JPEG，加合法 APP1 Exif/TIFF IFD，不修改产品公开接口；六个 40×40 色块构成不对称的 3×2 像素图案。坐标期望独立于 RotateFlipType：6 将存储 (x,y) 放到 (79-y,x)，2 放到 (119-x,y)。JPEG 通道允许 ±12 压缩误差。

- orientation=6：FromStream 实测仍为未校正的 120×80；校正和最终 PNG 均为 80×120。红块中心 (20,20)→(59,20)，蓝块 (100,20)→(59,100)，黄块 (20,60)→(19,20)，其余三个色块也逐一核对。这证明不是只改宽高，且 DrawImage 没有二次旋转。
- orientation=2：最终 120×80；红块 (20,20)→(99,20)，蓝块 (100,20)→(19,20)，下排同样左右反转。1/3/4/5/7/8 的最终 PNG 也按六个独立坐标逐项验证。
- big-endian TIFF 的 orientation=6 也通过：GDI+ 属性值读取与本机字节序一致，校正一次；所有成功输出 PNG 均实测无 0x112 属性。
- orientation=0/9 和确实由 GDI+ 暴露的单字节 BYTE 属性：返回 false、空输出和 `exif-orientation-invalid`，没有抛到外层。
- 原有无 EXIF PNG、透明度、几何回退、容量及解析/请求测试继续通过。using 资源生命周期保留；没有改宿主 COM、源文档、覆盖框或旁注逻辑。

这些是实际 JPEG 解码、GDI+ 变换和编码后的像素测试，**不是实际 Word 显示测试**。本轮未操作 Word；尚未证明各 Word 版本对具体手机照片的 EXIF 显示行为，未发现实际宿主不一致的证据，也不宣称完成该验收。方向校正后若与 Office 矩形长宽比不符仍安全回退；镜像方向不能仅凭长宽比证明 Word 的显示处理，须现场用不对称照片核对。

本次补修仅改三个文件：`src/OfficeTranslate.Core/WordOriginalImageCapture.cs`、`tests/OfficeTranslate.Tests/WordOriginalImageCaptureTests.cs`、本报告。下方 22 文件清单描述完整 2.1.31 候选相对 d23d68e 的累积改动。

## 改动文件清单（22）

1. VERSION.txt
2. src/OfficeTranslate.Core/ImageCaptureResolution.cs（新增）
3. src/OfficeTranslate.Core/ImageRecognitionResult.cs（新增）
4. src/OfficeTranslate.Core/WordOriginalImageCapture.cs（新增）
5. src/OfficeTranslate.Core/ImageOverlayDiagnostics.cs
6. src/OfficeTranslate.Core/TranslationClient.cs
7. src/OfficeTranslate.Core/TranslationResults.cs
8. src/OfficeTranslate.Core/OfficeTranslate.Core.csproj
9. src/OfficeTranslate.WordAddIn/WordTranslationService.cs
10. src/OfficeTranslate.WordAddIn/TranslationResultForm.cs
11. src/OfficeTranslate.WordAddIn/UiText.cs
12. src/OfficeTranslate.WordAddIn/AboutForm.cs
13. src/OfficeTranslate.WordAddIn/Properties/AssemblyInfo.cs
14. src/OfficeTranslate.ExcelAddIn/ExcelTranslationService.cs
15. src/OfficeTranslate.ExcelAddIn/Properties/AssemblyInfo.cs
16. src/OfficeTranslate.PowerPointAddIn/PowerPointTranslationService.cs
17. src/OfficeTranslate.PowerPointAddIn/Properties/AssemblyInfo.cs
18. tests/OfficeTranslate.Tests/OfficeTranslate.Tests.csproj
19. tests/OfficeTranslate.Tests/ImageRecognitionRecallTests.cs（新增）
20. tests/OfficeTranslate.Tests/ImageRequestRecallTests.cs（新增）
21. tests/OfficeTranslate.Tests/WordOriginalImageCaptureTests.cs（新增）
22. tests/VALIDATION-2.1.31-OCR-RECALL.md（本报告）

## 未覆盖项与公司验收口径

- 公司服务器、Qwen3.8-27b、公司照片均不可在 nuc 复现；未验证识别完整度、译文质量、请求实际时延、input/output/视觉 Token 消耗或与 Open WebUI 的同条件速度对比。
- 实际 Word Range.WordOpenXML 的结构/源像素读取、真实捕获像素、同图多轮、页眉/页脚、浮动/裁剪/翻转/旋转回退，以及 Word 显示与 EXIF 1–8 的对应关系均待验。纯 XML/JPEG/GDI+ 像素回归不能替代这些宿主验证。
- 三宿主真实 Ribbon 菜单、OCR 覆盖/旁注排版、结果窗口显示均未做本轮验收；Excel/PPT 捕获提高清晰度未实施。
- 8192 是最大生成预算，不是强制生成量，也不保证服务接受该长度。输入视觉 Token 增加可能拖慢模型；服务端上下文、视觉缩放或超时配置不足时需要现场确认，不自动降低预算并掩盖漏识别。
- 无“完整识别”质量判定：模型没有返回的文字无法单靠 regions 数推知；新增日志用于区分低输入像素、模型少返回、解析损失和截断。格式有效而漏识别仍需要与原图人工对照。
- 内存规则不是进程内存封顶；公司加密/安全软件是否允许 WordOpenXML 内嵌图读取、只读提取是否频繁降级待现场验证。
- 未 push、合并、发布、打 tag、做 MSI 或变更正式安装；main 和 2.1.30 候选分支不动。

公司测试建议：同一图先确认诊断 `ocr_capture path=original-raster` 的 PNG 像素（若回退，看 reason）；再按 `ocr_summary` raw/kept/dropped、empty_translation、bbox_note、finish_reason、truncated 比对全文。分别记录插件和 Open WebUI 的实际任务类型、图片尺寸、输出上限、时延及服务端 usage，不能用“仅 OCR”和“OCR+翻译+坐标”的时延直接宣称加速或退化。
