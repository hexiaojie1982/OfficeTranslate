# OfficeTranslate 2.1.34 候选：独立宽高 OCR 覆盖框放大

## 范围与结论

- 基线：`c66ca5a28e81460360ede283f5938e34c8065a5e`（2.1.33）；新分支 `fix/ocr-overlay-expansion-2`，只在新分支本地提交。
- VERSION.txt 为 2.1.34；Word/Excel/PowerPoint AssemblyVersion、AssemblyFileVersion 为 2.1.34.0。共用 AboutForm 新增 2.1.34，保留独立 2.1.33 和更早日志。
- 放大算法仅实现于 Core 的 ImageOverlayPlanner；宿主同步宽高倍数日志，并将真实排版检查的字号下限同步到 Core 的共享常量，没有建立宿主专用放大算法。
- 本轮三个改动点：**宽高独立放大，每维上限 5x；信任门 3 单位；最小字号 8pt → 6pt**。最大字号仍为 18pt。
- Release 构建 0 错误；最终全量单测 313/313，失败/跳过 0。相对基线新增 8 个单测方法/执行用例（初始 6 项 + 本补充 2 项）；基线 305 项全部保留（冲突旧意图按下表更新）。
- 不合并、不 push、不打 tag、不生成 MSI、不安装、不发布；不动 main、其他候选分支或 review-image-overlay 工作区。

## 最终数值与 2.1.33 差异

| 项目 | 2.1.33 | 2.1.34 |
| --- | --- | --- |
| 宽高放大 | 公共倍数，沿该均匀增长路径求最小倍数 | 独立宽高，遍历拟合前沿，求覆盖面积最小的组合 |
| 每维上限 | 3 倍 | 5 倍，且不超过图片该维尺寸 |
| 信任门 | 每维至少 5 个 bbox 单位 | 每维至少 3 单位（0.3%），严格小于 3 仍旁注 |
| 最小/最大字号 | 8pt/18pt | 6pt/18pt（第三改动点） |
| grow 日志 | 单值、最多三位小数 | 宽×高，各一位小数，例如 2.0x3.5 |
| 未放大/旁注日志 | 1.0 / 0.0 | 1.0x1.0 / 0.0 |
| 宿主真实排版失败 | 转旁注 | 不变，仍转旁注 |

所有有效性检查仍先于放大：0–1000 范围/非有限 bbox → 正向框 → 信任门 → PNG 与 Office 尺寸有效性 → 旋转 → 长宽比 → 原映射框非有限值/边界 → 原 FitFontSize。除了用户明确要求修改的小框阈值，原坐标校验顺序、结果和文案不动。旋转 0.5°、长宽比 5%、边界容差 0.05pt 不变。

原矩形能容纳译文时，不缩小原框也不放大。否则分别取 maxWidth=min(图片宽, 原宽×5)、maxHeight=min(图片高, 原高×5)，在最大候选仍不满足原 FitFontSize 时直接旁注。不会因为某一维先封顶，就拒绝另一维仍可增长的合法组合。

### 最小面积选择依据

FitFontSize 的算法不改，仅将实际位于 ImageOverlayLayout 的 MinFontSize 常量由 8F 改为 6F；ImageOverlayPlanner 继续调用这一共享判定。最小字号下的所需高度随宽度是由换行数形成的下降阶梯。每个拟合阶梯的最小宽/最小高角点才可能是面积最优，阶梯内继续增加宽高不会改善面积。

1. 用最大宽找到最低可拟合高度，作为后续面积下界。
2. 用最大高找到第一个可拟合宽度。
3. 对当前宽度求最小可拟合高度，比较 width×height（double 计算，避免 float 面积乘积溢出）。
4. 以该高度的前一个 float 为目标，找到下一次可降低换行高度的最小宽度，继续比较；已到最低高度/最大宽度，或后续宽×最低高度不可能改善当前面积时停止。

单调搜索使用正有限 float 的整数位序寻找第一个可拟合值，避免固定采样网格漏掉最优阶梯，也避免近似高度导致重复停留在同一阶梯。不复制字符测量常数、不修改原排版判定。结论限于原 FitFontSize 估算器和可表示 float 范围，不等于真实宿主排版的最小面积。

选中宽高后，仍围绕原 bbox 中心定位，贴边时向内平移保持完整尺寸；最终重新拟合并检查边界。记录 WidthGrowthFactor、HeightGrowthFactor；兼容的 GrowthFactor 只作为两者最大值的派生属性，不参与新面积选择。SideNote 两因子均为 0。

实例：

- 6×6pt 小框 + A：原估算器 MeasureUnits 有一单位下限，6pt 的最小框为约 10×9.8pt，倍数约 1.666667×1.633333，日志 grow=1.7x1.6；不再强制两维相等。
- 10×18pt + ABCDE：最小组合约 12.7×18pt，只增长宽度，grow≈1.27×1.0。
- 保留旧输入 10×24pt + ABCDE：6pt 下最小组合变为三行约 10×25.4pt，只增长高度；另加回归用例，不能只改输入藏掉这一方向变化。
- 10×20pt + ABCDEFGHIJ：6pt 下最小组合为三行约 15.6×25.4pt，面积约 396.24pt²，两行需 21.4×20pt（面积 428pt²）。单测与其他所有可行换行数逐项比较。
- 原宽占满图片的 1000×20pt 框 + 400 个汉字：宽保持 1000，高约 25.4pt，仍可独立增高，证明不会因宽被图片封顶而直接旁注。

最多每维 5 倍意味着面积理论上可达 25 倍。增长可能遮住邻近原文或其他覆盖框；沿用用户接受的取舍，不做碰撞避让。旁注 splitter、1400pt 上限、归属/清理、摘要、OCR 捕获/网络/语言保护都不改。

## 初始新增单测（IndependentOverlayExpansionTests，6/6，最终均保留）

- Plan_WideShortText_GrowsOnlyWidthWhenThatIsMinimumArea
- Plan_IndependentCombination_MinimizesAreaAcrossAllWrapCounts
- Plan_MultilineMinimumArea_MatchesExhaustiveGridOracle（使用原 FitFontSize 对 10–50pt 宽、20–100pt 高、0.5pt 间隔的所有网格对作独立比较）
- Plan_WidthAtImageCap_StillGrowsHeightIndependently
- Plan_SaturatedHeightCap_CannotFit_ReturnsSideNote
- Plan_BetweenOldAndNewHeightCaps_NowPlaces（高度需超过 3 倍但不超过 5 倍的新成功路径）

## 初始 2.1.34 修改的旧测试：逐项说明（8pt 阶段历史记录）

| 测试/辅助 | 变化与原因 |
| --- | --- |
| InBounds 共享辅助 | 单一 grow≤3 改成分别断言宽/高因子都在 [1,5]；影响调用它的中心、四角、宽度封顶、已拟合用例 |
| Plan_SmallTrustedBox_GrowsMinimallyAroundCenter | 不再断言 width=height；改为 12×12.4pt、分别对应因子，并分别缩减宽或高验证不拟合；上限注释从 18pt 改为 30pt |
| Plan_ThreeTimesStillTooSmall_UsesExistingSideNoteChannel → Plan_FiveTimesStillTooSmall_UsesExistingSideNoteChannel | 保留同样的 20 汉字/6×6pt 输入，验证新的 5 倍仍失败，而不是只验证旧 3 倍 |
| Plan_TrustGateIsFiveUnits_NotTen → Plan_TrustGateIsThreeUnits | (4.999,6)/(6,4.999) 改为 (2.9,6)/(6,2.9) 仍旁注；(5,5) 改为 (3,3) 可进入放大；原 (9.999,9.999) 仍保留成功 |
| Plan_ValidationOrder_RemainsRangeThenBoxThenTrustThenDimensionsThenRotationThenAspect | 同时非法输入中的小框从 4×4 改为 2.9×2.9，确保真正低于新门槛；其余范围/框/尺寸/旋转先后意图不变 |
| Diagnostics_GrowFieldIsConsistentInvariantAndMetadataOnly（Word/Excel/PowerPoint 三例） | 单值 grow=2.067 改为 grow=2.0x2.1；未放大改为 1.0x1.0，SideNote 保持 0.0；仍验证法语文化下的小数点格式 |
| Plan_TinyBbox_GoesToSideNote | bbox 从 4×4 调为约 2.9×2.9，继续测试信任门以下噪声 |
| Plan_EdgeTinyRegion_GrowsButStaysInsideImage | 2pt 原宽上限由 6 调为 10pt，20pt 原高由 60 调为 100pt，grow 上限改为 5；图片边界断言保留 |
| Plan_TextTooLong_GoesToSideNote | 500 汉字在 5 倍下已可能放下，因此改为 2000 汉字，保留“最大候选仍不合适必须旁注”的真实意图 |

其余范围/倒置/NaN/Infinity/长宽比/旋转等既有用例不改，保留全量验证。信任门降为 3 是明确需求，不声称有公司样本统计证明它是最佳数值。

首轮专项 43/48：5 个预期值失败（中心框最小宽的字符单位下限、三个诊断字符串、短文本的一行/两行最小面积）。均按未改的估算器修正测试预期，未修改原 FitFontSize；其后全量 311/311。失败专项 TRX 保留为过程证据。

## 本补充：第三改动点 6pt，新增/修改断言单独列出

Word 中显示的图片通常受页面/栏宽限制，绝对尺寸比 PowerPoint 幻灯片上的图片小；相同的相对放大规则下，Word 可用的绝对高度更少，更容易撞到最小字号门槛。将下限降到 6pt 是为补偿这一差异，并不承诺所有区域都能覆盖，也没有公司照片的定量验证。

任务文件写下限在 ImageOverlayPlanner，实际源码常量在 ImageOverlayLayout（Planner 调用它）。三宿主真实排版检查此前又各自写死 8F，故一并改用 ImageOverlayLayout.MinFontSize。原循环仅 12 次，从 18pt 逐点下降时来不及检查 6pt；现在根据共享最大/最小字号计算迭代预算，确保检查到下限。只同步字号边界和所需检查次数，未改 Overflowing/BoundHeight 判定、几何、失败转旁注或旁注字号/切分/限高逻辑。

| 本补充影响的测试方法 | 新断言/变化 |
| --- | --- |
| Plan_SmallTrustedBox_GrowsMinimallyAroundCenter | 12×12.4pt → 10×9.8pt；仍逐维缩减验证不拟合，中心/边界/独立倍数断言保留 |
| Plan_EdgeBox_TranslatesGrownFrameInsideImage（四角 4 项） | 高 12.4 → 9.8pt；四角仍验证平移不裁框，宽高在 5x 内 |
| Diagnostics_GrowFieldIsConsistentInvariantAndMetadataOnly（三宿主 3 项） | grow=2.0x2.1 → grow=1.7x1.6；原框 1.0x1.0、旁注 0.0、文化及隐私断言不变 |
| Plan_WideShortText_GrowsOnlyWidthWhenThatIsMinimumArea | 高 24 的旧输入在 6pt 下已不再仅增宽；改为高 18，预期 12.7×18、倍数 1.27×1.0、中心 y=409，保持真实的“只增宽最优”意图 |
| Plan_IndependentCombination_MinimizesAreaAcrossAllWrapCounts | 27.2×22.8 → 15.6×25.4（三行）；独立换行数 oracle 从文字宽 46.4/行高 10.4 更新为 34.8/7.8 |
| FitFontSize_ShortText_UsesLargeSizeAndFits | 下限断言改用共享 MinFontSize，仍 ≤18pt |
| FitFontSize_MediumText_ShrinksToFit | 下限断言改用共享 MinFontSize |
| FitFontSize_SmallButUsableRegion_Fits | 下限断言改用共享 MinFontSize |
| FitFontSize_InvalidDimensions_FallsBackToSafeDefaults | 默认 24×12 在 6pt 现在能放下 x：Fits=false/8pt 改为 Fits=true/约 7.685697pt，按既有八轮二分计算 |
| **新增** FitFontSize_SixPointFloor_FitsWhereEightPointWouldOverflow | 10×9.8pt 的 A 能以 6pt 完整放下；降到 9.79pt 则失败；显式断言共享下限=6 |
| **新增** Plan_PreviousWideFixture_NowGrowsOnlyHeightAtSixPoints | 保留旧 10×24pt/ABCDE 输入，断言 10×25.4pt、宽倍数=1、高倍数约 1.058333、原中心及真实估算 Fits |

FitFontSize_LongText_ShrinksInsteadOfGrowing 只同步注释为“6pt 仍放不下”，失败断言不变。其他安全门、最大 5x 仍失败、宽封顶、超过旧 3x 高度、未放大的断言保持并全部重跑。独立宽高算法、3 单位信任门、检查顺序及旁注路径相对 e03a8d8 未改。

## 实际命令与数字

目录 `D:\pythonProject\office-translate\native-word`，SDK 8.0.423、net48；沿用上一轮已有 -SkipInstaller，不修改 build.ps1：

```powershell
$env:PATH='C:\Program Files\dotnet;'+$env:PATH
.\build.ps1 -Configuration Release -SkipInstaller
$env:DOTNET_CLI_HOME=Join-Path (Get-Location) 'dotnet-home'
$env:NUGET_PACKAGES=Join-Path (Get-Location) 'packages'
& 'C:\Program Files\dotnet\dotnet.exe' test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=2.1.34-overlay-final3.trx' --results-directory TestResults\ocr-overlay-expansion-2
```

- 本补充全解决方案 Release 编译：退出 0，0 错误/7 个既有 nullable 警告，2.73 秒。版本校验通过，日志明确 installer skipped。
- 加入保留旧输入的回归用例后，再次完整执行同一 Release 构建命令：退出 0，0 错误/4 个既有测试 nullable 警告，1.04 秒；3 个产品既有警告未在增量重报。新增代码无警告。
- 本补充首次全量单测：312/312，failed/skipped 0，退出 0；随后加入保留旧输入的回归用例重跑，**最终 313/313，failed/skipped 0，退出 0，43 秒**。最终 TRX `TestResults/ocr-overlay-expansion-2/2.1.34-overlay-final3.trx` 覆盖了同名的 312 项中间结果，以其中 counters 为准。原 8pt 阶段 final/final2 均 311/311，不能当作 6pt 的验证。
- 专项命令使用 `--filter 'FullyQualifiedName~OverlayExpansionTests|FullyQualifiedName~ImageOverlayPlannerTests'`，结果 43/48，TRX `2.1.34-overlay-focused.trx`；该失败轮不当作最终通过。
- TestResults 本地测试产物按既有规则忽略；git diff --check 通过。

Release DLL SHA-256：

| DLL | SHA-256 |
| --- | --- |
| Core | 8AB4E6597A4239ACFF857AB198F37D9B4A65164B3D8C47F59B7B0AD5891E4EA9 |
| Word | 5A94D80D45CB413481D81B4E353D1EE338807DFF5FDAB96D4B3A89FFDB303C72 |
| Excel | 4116783A345C2327D9A32FE4E9ACA8F82232FAEF92182F3FFE675CD2FA66E0B0 |
| PowerPoint | 4E464A32718EE0ED0649160A75F480489BB589DBA083789C3CF0D9BD29264E9B |

## 未覆盖项与硬边界

- 没有真实 Word/Excel/PowerPoint 宿主内的排版/视觉/遮挡结果验证，没有公司照片复测、公司模型请求或菜单端到端验收。
- **宿主写回时**的真实 Overflowing/BoundHeight 检查下限同步为 6pt，若仍不合适，仍会转旁注；回退行为未关闭、未重写，不承诺新条件让所有区域都成为覆盖框。6pt 的实际可读性和各宿主实际布局仍待用户菜单验收。
- 最小面积只对原 FitFontSize 估算成立，不等于真实宿主的最优排版；没有现场速度或 Token 消耗数据。
- 没有修改覆盖/旁注透明度、真实布局、所有权、清理、取消、UI 恢复票据、捕获预算、网络超时等行为。日志只增加宽高数值，不能把规划 Place 当作真实宿主成功的证明。
- 默认 MSI 路径未执行、没有任何 MSI 发布/安装。只执行已有 -SkipInstaller 的 DLL 构建路径。
- main、fix/ocr-overlay-expansion、fix/ocr-original-image-all-hosts、fix/ocr-combined-note-overflow 的 ref 不动；只在 fix/ocr-overlay-expansion-2 本地提交，不 merge/push/tag/release，不改 review-image-overlay。
