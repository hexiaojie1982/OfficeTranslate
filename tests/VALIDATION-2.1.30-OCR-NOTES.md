# 2.1.30 图片合并旁注溢出修复：本地验证报告

日期：2026-10-10。工作仓库：`D:\pythonProject\office-translate\native-word`。
基线：main / v2.1.29，`a674cd1d2230faf15f5b5e3c4c1dac4558ee2555`。
分支：`fix/ocr-combined-note-overflow`。提交哈希见本轮最终回复或本地 `git log -1`。

## 根因复核

确认用户提供的报错路径与 main 源码一致：

1. `ImageOverlayPlanner` 调用 `ImageOverlayLayout.FitFontSize`；8pt 下仍不适配的区域返回 SideNote。
2. 三宿主把降级区域的原因和译文收集到 noteEntries，经 `ImageOverlayNotes.Combine` 合为一条旁注；宽度 160–420pt、字号 9pt。
3. 三个 `GrowNoteToFit` 均在 1400pt 后仍溢出时删除该条旁注并抛 InvalidOperationException。
4. 此操作位于该图片的旧结果清理之后。**当整张图的区域都降级为旁注时**，唯一的新旁注删除后，整图确实无译文；若部分区域已成功覆盖，原实现没有在此路径删除那些新覆盖框，不能泛化为任何溢出都删除全部新覆盖框。

以上是源码路径确认；本轮未打开 Office 复现真实排版报错。

## 修复策略和不变原则

- Core 统一切分、估高和放置调度。先按完整条目装箱；单条过长则优先按行切分，超长单行也可切分。
- 保留条目索引、原始字符偏移与完整原文片段。后续段明确显示“续接”及字符范围；不丢失或重复有效载荷，不以截断文本冒充译文。
- 按实际旁注 8pt 总横向边距估高；预切分使用 1400pt 的 70% 估高预算，给真实字体/排版留余量。估高不是真实排版的替代。
- 仍使用各宿主原有真实排版校验及有界增高。达到上限或校验不收敛时，该次形状删除并发出专用溢出异常；Core 对该块继续无损细分。
- 最小片段仍无法放置，或出现非排版创建/COM 错误，只记录这一块的完整内容和原因，继续其余块。取消异常不作为旁注失败吞掉。
- 每个成功块沿用同一图片 owner：Word/Excel 的 AlternativeText 为 `ImageOverlayTags.NoteFor(ownerId)`；PowerPoint Tag 为 `Note:` + ownerId。重译的旧结果清理仍按 owner 一次执行，新增块之间不互删。
- 用成功块的实际最终 Height + 6pt 堆叠下一块；失败块不占用空间。
- 保持覆盖框逻辑、8–18pt 字号规则、旁注 9pt 和 1400pt 高度上限不变。
- 失败摘要显示未放置块数量、图片 owner、错误原因及完整内容；滚动文本框可复制，存在未放置块时取消倒计时，直到用户关闭。无此类失败的摘要仍 5 秒自动关闭。
- VERSION.txt、三宿主 AssemblyVersion/AssemblyFileVersion、关于窗口及更新日志同步为 2.1.30(.0)。MSI 版本仍来自 VERSION.txt；本轮不生成 MSI、不创建 tag。

## 改动文件清单（仓库相对路径）

1. `VERSION.txt`
2. `src/OfficeTranslate.Core/ImageOverlayNotes.cs`
3. `src/OfficeTranslate.Core/ImageOverlayNoteSplitter.cs`（新增）
4. `src/OfficeTranslate.Core/TranslationResults.cs`
5. `src/OfficeTranslate.ExcelAddIn/ExcelTranslationService.cs`
6. `src/OfficeTranslate.ExcelAddIn/Properties/AssemblyInfo.cs`
7. `src/OfficeTranslate.PowerPointAddIn/PowerPointTranslationService.cs`
8. `src/OfficeTranslate.PowerPointAddIn/Properties/AssemblyInfo.cs`
9. `src/OfficeTranslate.WordAddIn/WordTranslationService.cs`
10. `src/OfficeTranslate.WordAddIn/TranslationResultForm.cs`
11. `src/OfficeTranslate.WordAddIn/UiText.cs`
12. `src/OfficeTranslate.WordAddIn/AboutForm.cs`
13. `src/OfficeTranslate.WordAddIn/Properties/AssemblyInfo.cs`
14. `tests/OfficeTranslate.Tests/ImageOverlayNoteSplitterTests.cs`（新增）
15. `tests/VALIDATION-2.1.30-OCR-NOTES.md`（本报告）

## 新增单测

新增 16 个测试方法，因一项有 3 组数据，共 18 个执行用例。既有单测未修改。

- `Split_ManyEntries_KeepsWholeEntriesAndOrder`
- `Split_SingleHugeMultilineEntry_PreservesLinesAndContinuation`
- `Split_SingleHugeLine_SplitsWithoutDroppingText`
- `Split_EveryChunkEstimate_IsUnderCap`（160 / 300 / 420pt 三组）
- `Split_UnicodeAndMixedLineEndings_KeepExactTextElements`
- `Split_EmptyInput_HasNoNotes`
- `Split_EmptyEntries_AreRepresentedExactlyOnce`
- `Split_InvalidWidth_RejectsNonFiniteGeometry`
- `SplitForRetry_MultipleEntries_PreservesOrder`
- `SplitForRetry_LongEntry_PreservesOffsetsAndText`
- `Place_RealOverflow_SplitsAgainAndStacksActualHeights`
- `Place_OneFailure_PreservesSuccessAndContinues`
- `Place_UnsplittableOverflow_ReportsOnlyThatFragment`
- `Place_Cancellation_IsNotSwallowedAsNoteFailure`
- `Summary_UnplacedNotes_KeepFullTextAndOwner`
- `ChunkMarkers_AllBelongToSameImage_NotOtherImage`

用委托模拟真实校验的返回高度、溢出与错误；测试的是生产共用的调度逻辑，不冒充 Office COM 实测。归属测试验证 marker 判定契约，不冒充真实重译清理验收。

## 实际构建和测试命令

在上述 native-word 仓库中使用 PowerShell、.NET SDK 8.0.423：

```powershell
$env:DOTNET_CLI_HOME = Join-Path (Get-Location) 'dotnet-home'
$env:NUGET_PACKAGES = Join-Path (Get-Location) 'packages'
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = '1'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
dotnet restore OfficeTranslate.Word.sln --ignore-failed-sources
dotnet build OfficeTranslate.Word.sln -c Release --no-restore -t:Rebuild
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=2.1.30-final.trx' --results-directory TestResults\ocr-combined-note-overflow
git diff --check
```

- 首次 restore 因沙箱无法读取用户 NuGet.Config 失败；解除该执行限制后重跑成功，依赖均已恢复。不是源码构建失败。
- 最终完整 solution Rebuild：**退出码 0，0 错误，7 警告**，用时 1.81s。覆盖 Core、Word、Excel、PowerPoint、Tests 五个项目。
- 7 条警告均为既有 nullable 问题：OfficeUiDispatcher、WordTranslationService、WordSelectionProbe、StagingNameContractTests（2 条）、ReviewFixesTests、ImageOverlayTextTests。新代码无新增编译警告。
- 最终全部 MSTest：**141 总数，141 通过，0 失败，0 跳过**，测试用时 31s；TRX Counters 独立读取核对一致。123 既有用例 + 18 新增用例。
- 最终 TRX：`TestResults/ocr-combined-note-overflow/2.1.30-final.trx`（本机生成文件，按现有规则忽略，不提交）。
- `git diff --check` 通过；Git 的 LF→CRLF 提醒仅为本机换行策略提示。
- 没有调用 `build.ps1`，因为该脚本在编译之后还会生成 MSI；仅执行其 dotnet 构建步骤。

## 未覆盖项与用户后续验收

本轮**未做真实 Word/Excel/PowerPoint 菜单、OCR 请求或 COM 排版验收**，未安装或注册候选 DLL。需用户后续验证：

- 小字紧凑图的实际识别与双语/单语结果，多条旁注完整显示、不互删。
- 不同默认字体/段落样式的真实 Overflowing / BoundHeight 校验、触发二次切分。
- 单张图多轮重译：其所有段被一并清理，相邻其他图片结果保持不变。
- 成功覆盖框与失败旁注混合时，成功结果仍在；失败摘要的完整文字可复制且不会自动关闭。
- 超多块向下堆叠在 Word 页面、Excel 工作表及 PPT 幻灯片中的实际位置与宿主坐标限制；本修复不新增分页或移到下一张幻灯片的策略。若宿主拒绝某块坐标，应在摘要明确报告，而不是称全部放置成功。
- Office 对话框的五种界面语言实际显示与滚动/复制行为。

本轮仅本地分支提交；main/v2.1.29 不动。未 push、未合并、未打 tag、未生成 MSI、未发布、未发邮件；未修改项目根目录的旧 TS 工程或 review-image-overlay 工作区。
