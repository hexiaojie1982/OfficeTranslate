# OfficeTranslate 2.1.33 候选：图片 OCR 覆盖框适度放大

## 范围与结论

- 基线：`b2fcabae45cd85e7fd4eb9349d70aba2a03d08ff`（2.1.32 混合段落修复）；新分支 `fix/ocr-overlay-expansion`。
- VERSION.txt 为 2.1.33，Word/Excel/PowerPoint 的 AssemblyVersion/AssemblyFileVersion 同步为 2.1.33.0。三宿主链接同一 AboutForm，新增 2.1.33 日志并保留独立 2.1.32 历史。
- 使用 build.ps1 的 Release 路径构建成功，0 错误；全量测试 305/305 通过，失败/跳过 0。新增 9 个方法、27 个参数化执行用例；基线 278 项全部保留（含两处旧意图更新）。
- 仅新分支本地提交；不 push、不合并、不打 tag、不安装、不发布、不生成 MSI。main、fix/ocr-original-image-all-hosts、fix/ocr-combined-note-overflow 分支和 review-image-overlay 工作区不改。

## 规则与实现

全部几何规则仅实现于 Core `ImageOverlayPlanner`，三宿主共用，调用接口 Plan 不变：

| 参数 | 最终值 | 说明 |
| --- | --- | --- |
| bbox 信任门 | 每维至少 5 个 0–1000 单位 | 严格小于 5（整图 0.5%）仍旁注，沿用“坐标框过小，不可信”文案；恰好 5 允许进入布局验证 |
| 放大上限 | 每维最多原 bbox 尺寸的 3 倍 | 同时不超过图片该维尺寸，最大面积可能达到原区域 9 倍 |
| 字号 | 最小 8pt、最大 18pt | 原 ImageOverlayLayout.FitFontSize 完全不改，不用更小字号冒充能放下 |
| 搜索 | 公共倍数 [1,3]，20 次二分 | 返回可放下端点；倍数区间误差至多 2/2^20，约 0.00000191，不一律取 3 |
| 边界容差 | 0.05pt | 沿用 BoundsEpsilon；原框及最终框都检查 |
| 其他信任门 | 旋转 0.5°、长宽比偏差 5% | 不改原数值或检查顺序 |

判定顺序仍为：bbox 0–1000 范围/非有限值 → 正向有效框 → 小框信任门 → PNG/Office 尺寸有效性 → 旋转 → 长宽比 → 映射矩形非有限值/越界 → FitFontSize。仅在最后布局不合适时新增放大路径，绝不用放大挽救无效几何。信任门按用户指定的 5 单位执行；它是本轮明确接受的召回取舍，不宣称已由公司样本统计证明最佳阈值。

先用原 bbox 矩形及原 FitFontSize；能放下时输出完全不变、grow=1。不能放下时先验证“每维 3 倍且封顶图片尺寸”的最大候选；仍不能放下则旁注。最大候选可放下时，围绕原中心按公共倍数增长宽高，并对宽高独立封顶图片尺寸；在此增长路径上二分找最小可放下倍数。并非分别寻找全局最小面积的任意宽高组合。

贴边时保留已可放下的宽高，只将 Left/Top 平移进 `[0, 图片尺寸-框尺寸]`，避免裁掉所需空间。最终重新 FitFontSize 并检查边界，失败仍转原旁注流程。

纯计算实例：1000×1000pt 图片上 bbox [500,500,506,506] 的 6×6pt 原框，文字 A，在约 2.066667 倍时达到 12.4×12.4pt，中心保持 503/503；高度达到 8×1.3+2=12.4pt 的原估算条件。缩小倍数 0.0001 后不再通过估算，证明不是直接顶满 3。贴右下角 [994,994,1000,1000] 同样增至约 12.4pt，Left/Top 移到约 987.6pt，不越图片边界。

用户已明确接受放大遮住邻近原文。此实现不做邻近区域避让/碰撞检测，增长后的不同覆盖框可能重叠；这是取舍，不声称智能排版。几何规划仅采用既有字符宽度估算；宿主真实 Overflowing/BoundHeight 检查、缩字号到 8pt、删不适合的框并降级旁注、AutoSize=None、MatchesPlan 等安全逻辑均保留。不会将真实裁切当成成功。

旁注分段、1400pt 上限、续接关系、归属/清理、未放置摘要等完全不改；OCR 请求、捕获、超时、语言保护和正文翻译管线完全不改。

## 诊断约定

PlannedOverlay 新增 GrowthFactor，表示输出矩形实际 `max(输出宽/原宽, 输出高/原高)`。所有 SideNote（包括宿主直接构造的 SideNote）统一为 0；Place 未放大为 1。

三宿主的 LogOverlay 都传该值到 ImageOverlayDiagnosticEntry，沿用原日志字段并在末尾追加 `grow=...`。格式使用 invariant culture `0.0##`：例如 grow=1.0、grow=2.067、SideNote grow=0.0。只记几何元数据，不传正文/Prompt/密钥。该逐区日志记录规划结果，不能用规划 Place 推断宿主真实布局最终一定成功。

## 单测清单

新增 `ImageOverlayExpansionTests`（9 个方法、27 个执行用例）：

1. Plan_SmallTrustedBox_GrowsMinimallyAroundCenter（1）：5–10 单位内的小框可放下、保中心、不顶满、比最终倍数略小即不合适。
2. Plan_EdgeBox_TranslatesGrownFrameInsideImage（4）：四角向内平移、尺寸不裁掉、每维不越界。
3. Plan_ImageWidthCap_PreservesFitAndPerDimensionLimit（1）：宽度已满图片时封顶宽度、增长高度、仍满足最小拟合。
4. Plan_ThreeTimesStillTooSmall_UsesExistingSideNoteChannel（1）：3 倍仍放不下，旁注/grow=0，不输出假几何。
5. Plan_TrustGateIsFiveUnits_NotTen（4）：任一维 4.999 仍不信任，5 和 9.999 可放大验证。
6. Plan_AlreadyFits_KeepsExactRectAndGrowthOne（1）：原来可放下的框坐标/尺寸不动。
7. Plan_InvalidInput_IsRejectedBeforeGrowth（11）：范围、倒置、NaN/Infinity bbox/尺寸、旋转、长宽比等先拒绝。
8. Plan_ValidationOrder_RemainsRangeThenBoxThenTrustThenDimensionsThenRotationThenAspect（1）：多种同时非法的输入仍返回原先最早的拒绝理由。
9. Diagnostics_GrowFieldIsConsistentInvariantAndMetadataOnly（3）：三宿主 grow 格式一致，法语 culture 仍用小数点，SideNote 为 0。

修改原 `ImageOverlayPlannerTests`：

- Plan_TinyBbox_GoesToSideNote：原 5×5 改成 4×4，继续验证真正低于信任门的噪声框。
- Plan_EdgeTinyRegion_DoesNotExpandBeyondImage 重命名为 Plan_EdgeTinyRegion_GrowsButStaysInsideImage：旧规则“必旁注”改成验证新规则下增长、每维最多 3 倍且仍在图片内。依据现有 FitFontSize 估算可放下，不意味着真实宿主必然能容纳极窄字符。
- 原 Plan_TextTooLong_GoesToSideNote 等测试不改，通过全量回归确认。

过程记录：首次专项 41/42，唯一失败是上述贴边旧断言仍要求 SideNote；按新意图更新后，全量 305/305。保留首次 TRX，不将失败轮当成验收通过。

## 实际构建和测试命令

目录：`D:\pythonProject\office-translate\native-word`；.NET SDK 8.0.423 / net48：

```powershell
$env:PATH='C:\Program Files\dotnet;'+$env:PATH
.\build.ps1 -Configuration Release -SkipInstaller
& 'C:\Program Files\dotnet\dotnet.exe' test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=2.1.33-overlay-final.trx' --results-directory TestResults\ocr-overlay-expansion
```

为同时满足“用 build.ps1”及“禁止 MSI”，脚本新增可选 -SkipInstaller，默认不带参数的原 MSI 路径不变。SkipInstaller 时不用 WiX 作为先决条件，仍执行 SDK/Framework/版本校验、restore、solution build；build 成功后 return，位置在 artifact/WiX 调用之前。没有调用 candle/light，也没有新的 MSI/WiX 对象。

- 初次全解决方案编译：0 错误/7 个既有 nullable 警告，3.51 秒，脚本明确提示 installer skipped。
- 更新旧测试后最终 Release 构建：0 错误/4 个既有测试 nullable 警告，1.33 秒（产品项目增量复用，初次产品编译的 3 个既有警告未重报），退出 0。新增代码无警告。
- 全量单测：305 executed/passed，failed/skipped 0，约 43 秒，退出 0。TRX `TestResults/ocr-overlay-expansion/2.1.33-overlay-final.trx`。
- 初次专项：`--filter 'FullyQualifiedName~ImageOverlayExpansionTests|FullyQualifiedName~ImageOverlayPlannerTests'`，TRX `2.1.33-overlay-focused.trx`，41/42；该失败轮仅为过程证据。
- git diff --check 通过。以上本地测试产物被仓库忽略，没有作为产品交付提交。

当前 Release DLL SHA-256：

| DLL | SHA-256 |
| --- | --- |
| Core | 5EC4E50648B50CC383AC5AEF6239300DF2149EB597E4B16735F121143146C4AB |
| Word | 4C496172E5CB8B4B04747100F3859DCEFDBBF6750402E332625C1D7EE5CC09DB |
| Excel | 2415FA865949BBA92B0BFF129F5C5890B8BBD248FEF3003FB1FA1765A9A4DFDD |
| PowerPoint | B4ADB46355C66C62B41CB43A59F80243F415FF304331879AA78A970FF914DF06 |

三宿主目录中复制的 Core.dll 哈希与 Core 输出一致。VERSION.txt/三 AssemblyInfo 静态读回已同步；About 界面未真实打开。

## 未覆盖项与交付边界

- 真实 Word/Excel/PowerPoint 内的排版、覆盖/遮挡效果、重译清理、字体/缩放/DPI 和菜单端到端未验；没有启动或操纵 Office。
- 公司照片未复测，公司模型/服务器没有在线请求；不宣称所有小字都会得到覆盖框或公司效果已通过。
- 原 FitFontSize 是估算，尤其极窄框的字符宽度估算可与真实宿主不同；本轮按要求不改该算法，真实失败仍走旁注。没有以降低字号或关闭真实排版验证绕过安全性。
- 可能覆盖邻近文字/其他覆盖框，不做碰撞排版。该行为是用户已接受的明确取舍。
- 默认 MSI 构建路径未执行，也不声称该路径重新验收；本次仅验证 -SkipInstaller 的 DLL 路径。已有 artifacts MSI 保留不动，其历史版本不是新 2.1.33 安装包。
- UI 队列/恢复票据、捕获预算、语言保护和网络超时等沿用原实现；没有扩大这次修改范围。

最终只有本分支代码、测试、版本/关于文本及安全 DLL 构建参数和本报告；不动其他分支/工作区，不 push/merge/tag/MSI/install/release。
