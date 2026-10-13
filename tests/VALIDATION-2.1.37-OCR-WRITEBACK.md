# OfficeTranslate 2.1.37.0：Word OCR 覆盖框写回溢出修复

日期：2026-10-12。基线：`6a5225d`（2.1.36.0）。
分支：`fix/ocr-writeback-overflow-2.1.37`。本地候选；不合并 main、不安装、不发邮件。

## 检测机制与根因

`WordTranslationService.PlaceOverlay` 创建固定尺寸、不透明的覆盖框，填入译文后调用
`ShrinkFontToFit`。后者调用任务文档 `Repaginate()`，读取 `TextFrame.Overflowing`。
它是 **整个带格式文本流能否全部放入文本框的布尔结果**，不是字形包围盒高度，也不是代码
自行比较某个 `BoundHeight`。Word 的换行、字体、段落行距、段距、文档网格和框内边距都会影响它。
因此不能把该结果解释成“1pt 字形本身仍太高”。

plan 的 `ImageOverlayLayout` 则是纯几何估算：扣除水平 4pt / 垂直 2pt 边距，按字符宽度估计
换行数，用 `行数 × 字号 × 1.3` 估计高度。它没有 Word 文档的段距、自动段距或网格量化。
2.1.36 的写回只设置框的边距、换行、AutoSize 和字号，段落格式全部继承。这是可复现的系统性偏差。

本机 Word 16.0 / build 16.0.14332，30×20pt、同一合成文本、1pt 对照：

| 设置 | 真实 Overflowing |
| --- | --- |
| 默认：段前/后 0、自动段距关、单倍行距、DisableLineHeightGrid=0 | True |
| 仅改 DisableLineHeightGrid=-1 | False |
| 网格关，但三段文字段前/后各 12pt | True |
| 网格关，显式段距 0，但自动段前/后打开 | True |
| 网格关，但固定 24pt 或多倍行距（LineSpacing=120） | True |
| 网格关、自动段距关、段前/后 0、单倍行距 | False |

**已证明段落格式能独立造成该写回降级，本机默认样例的直接原因是对齐网格。**
该 COM 属性名称和 VBA 文档文字容易误读；实测 `DisableLineHeightGrid=-1` 输出
`<w:snapToGrid w:val="0"/>`，确认是取消网格对齐。单倍行距下 `LineSpacing` 仍可能返回 12，
此属性值不能直接当作 1pt 字体的实际行高。没有采用固定 1.3pt 行距，避免裁剪实际字体的字形。

API 依据：[Overflowing](https://learn.microsoft.com/en-us/office/vba/api/word.textframe.overflowing)、
[自动段距会覆盖显式段距](https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.spaceafterauto)、
[ContainingRange](https://learn.microsoft.com/en-us/office/vba/api/word.textframe.containingrange)。
网格取值以 Word 对照及 OOXML 断言为依据。

## 修复范围

- `WordTranslationService.cs`：仅在 OCR 覆盖框真实排版判定前，关闭自动段前/后，设置段前/后 0、
  单倍行距、取消网格对齐。覆盖全部段落和末尾段落标记。
- 使用新建、未链接覆盖框的 `ContainingRange` 设置段落格式及逐步缩小的字号。
  长文本实测 `TextRange` 曾只返回 602 字符中的 406 字符，不能把它当作完整文本范围。
  301 段超限回归从 18pt 开始，候选逐字符验证全部 602 字符达到 1pt，随后仍正确拒绝放置。
  Word 对长范围的聚合 `Font.Size` 可能返回 `wdUndefined=9999999`，即使逐字符都是 1pt；
  该回归使用逐字符断言，不把聚合哨兵误报为真实字号。
- 设置段落格式遇到 COM 失败时返回固定中文原因，进入既有删除覆盖框/旁注降级路径。
  不记录异常 Message 或任何图片、OCR 文本、正文、Prompt、剪贴板、密钥。
- 保留 Word 实际排版缩字号和最终溢出门、1pt 下限、不透明白底、AutoSize=None、固定框几何、
  每维最多 5 倍扩展及图片边界。没有更改共享 plan 算法、普通文字翻译、旁注排版或其他宿主逻辑。
  紧凑段落消除已确认的偏差；plan 仍是字体宽度/换行估算，不声称与 Word 完全一致。
- `AboutForm.cs`：新增 2.1.37.0 中文日志段，保留 2.1.36.0 历史段。
- `VERSION.txt`、三宿主 `AssemblyInfo.cs`、build 版本示例、README 同步；Core、包版本、
  MSI ProductVersion 和 COM 注册版本从既有版本机制同步。

## 验证结果

最终 `build.ps1 -Configuration Release` 成功，**0 错误**；最终增量构建 2 个既有 Word nullable 警告
（首次完整相关项目构建 7 个既有 nullable 警告）。WiX candle/light 成功。

全量 MSTest：**359/359 通过，失败 0、跳过 0，58 秒**。
在已登录交互桌面 Session 1 运行，包含现有窗体/剪贴板测试，没有跳过 UI 测试。

`RunWord237LayoutProbe.ps1` 在同一交互桌面创建独立 Word 空白文档，分别加载保留的实际
2.1.36.0 DLL 与最终 2.1.37.0 DLL。无需切换 COM 注册，无模型请求、不读取用户文档。

| 回归 | 2.1.36.0 | 2.1.37.0 |
| --- | --- | --- |
| 网格、段距、自动段距、固定行距、多倍行距、混合段落（含末尾空段）、中英双语：7 组 | plan=Place；1pt 仍溢出 | 7 组全部 1pt 放入，框仍为 30×20pt |
| 真正超限的 301 段 | 拒绝放置 | 完整范围逐字符均为 1pt 后仍拒绝 |
| 调用真实 `PlaceOverlay`，30 个 W | 返回 False，删除失败框 | 返回 True，实际 3pt；白底不透明、无边线、前置覆盖、固定位置和宽高 |
| 强制 writeback 真超限 | — | 返回 False，失败框已删除 |
| 已删除形状引发段落格式 COM 失败 | — | 返回 False，原因分类正确 |

候选共 **11 组 Word 检查通过**，另核对正文文本、正文段落格式和 Normal 段落格式保持不变。
每段 OOXML 都断言 `snapToGrid=0`。未把长文本的聚合字号哨兵当作失败。
测试文档不保存，探针创建的 Word 已正常退出。

这里验证的是实际 Word COM 排版和生产写回方法；**未取得用户原始密集图片，未宣称完成该原图
或功能区 OCR 全链路验收**。字体差异、实际容量或其他几何信任门仍可合法触发旁注。

## 复现与交付

```powershell
.\build.ps1 -Configuration Release
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=all-final.trx' --results-directory TestResults\ocr-writeback-2.1.37
# 以下在交互桌面、无其他 Word 实例时执行
.\tests\RunWord237LayoutProbe.ps1 -EvidenceDirectory .\TestResults\ocr-writeback-2.1.37
.\tests\RunWord237LayoutProbe.ps1 -EvidenceDirectory .\TestResults\ocr-writeback-2.1.37 -AssemblyDirectory .\TestResults\ocr-writeback-2.1.37\baseline-dlls -ExpectBaseline
```

本机忽略目录 `TestResults/ocr-writeback-2.1.37/` 保存：
`build-release.txt`、`all-final.trx` / `.txt` / `.exit`、`baseline-word.txt` / `.exit`、
`candidate-word.txt` / `.exit`，以及初始单因素调查 `investigate.txt`。
以最终 `.exit=PASS` 和 TRX 为准；调查期间的失败诊断保留，不能把过时 `.error` 当作最终状态。

候选 MSI：`artifacts/candidate-2.1.37.0/OfficeTranslate.Office.x64.msi`。
已直接查询 MSI：ProductVersion、File 表四个产品 DLL、Registry 表三条 COM Assembly 均为
`2.1.37.0`。版本扫描保留历史验证报告、旧版探针名称和 About 历史日志的原始版本。

MSI SHA-256：`DD499B0202213134926301EE540361C9A383508754EC2AC458229B4F5F8FB019`。
