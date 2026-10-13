# OfficeTranslate 2.1.36.0：设置窗口 owner 与 OCR 1pt 候选验证

日期：2026-10-12。基线：`6e0173c`（main / 2.1.35，PR #16）。
开发分支：`fix/summary-settings-handle-2.1.36`。仅本地候选提交和 MSI；不合并 main。

## 1. 设置窗口与摘要

### 已确认的缺陷与修复

Word `AddIn.OpenSettings` 与 Excel/PowerPoint 共享入口原来调用无参 `ShowDialog()`。
在同一 STA 中，摘要为活动窗口时，WinForms 把摘要的 HWND 当成设置窗口的 owner，
而非当前 Office 文档窗口。摘要的关闭/销毁生命周期因此会影响设置窗口。

在 Windows 交互桌面加载三宿主 **2.1.35 实际 DLL**，分别显示摘要再打开设置，
通过 Win32 `GetWindow(GW_OWNER)` 验证：Word、Excel、PowerPoint 的 owner 断言均失败；
活动窗口改为模拟宿主后通过。证据：`baseline-desktop.trx`（3 失败、2 通过）。

修复：`SettingsForm.ShowFor` 使用明确的 HWND 包装对象；调用方在构造设置后重新读取当前宿主窗口，
只接受仍存在且属于当前原生 UI 线程的 HWND。无可用宿主时显式使用空 HWND，
配合 `ShowInTaskbar=false` 使用 WinForms 隐藏 owner，避免再次采用活动摘要。
不读取摘要/已释放控件的 `Handle`，不增加吞掉句柄异常的 catch。
摘要全文、关闭行为及普通文字翻译流程未修改。

**证据边界：尚未在本机直接重现用户报告的 `Error creating window handle` 异常。**
常规窗口组合及在设置 Load 时关闭摘要均未抛出该异常；已复现并修正的是错误的 owner 依赖。
尚无现场异常堆栈，不能认定已排除其他 Win32 句柄创建原因，也不声称已完成真实功能区点击验收。
需要用户在原始“需处理摘要未关闭”场景验收候选。

## 2. OCR 1pt 与旁注调查

`ImageOverlayLayout.MinFontSize` 从 6F 改为 1F；三宿主共用规划器和真实排版缩字号循环均引用此常量。
仍保留 18pt 上限、bbox 信任门、每维最多 5 倍且不超出图片的扩展限制，以及实际排版校验。
既有测试以 1pt 下仍需扩展/仍超限的新样例保持边界覆盖，并保留原 6pt 失败样例验证其现在可以放置。

实际决策顺序如下：

| 阶段 | 触发旁注的条件 | 降至 1pt 能否解决 |
| --- | --- | --- |
| 识别结果 | bbox 损坏/无效、保留原文但缺少译文 | 不能 |
| Word 定位 | 图片 Left/Top 为居中、靠右等对齐哨兵值 | 不能 |
| 几何信任 | bbox 超界/倒置/非有限；任一边小于 3/1000；尺寸无效；旋转超过 0.5°；长宽比偏差超过 5% | 不能 |
| 估算排版 | 在 5 倍扩展与图片边界内仍放不下 | 部分样例可以；极长内容仍失败 |
| Word 写回 | 定位参考系失败、字号设置失败、真实排版读取失败、1pt 仍溢出、宿主改变框几何 | 取决于实际原因 |

当前解析器没有独立的数值 `confidence` 门；所谓 bbox 信任主要是有效性与最小边长检查。
本机既有日志未包含 2.1.35 现场翻译，因此不能统计用户图片究竟在哪个门被拒绝。

### 真实 Word 排版证据

`RunWord236LayoutProbe.ps1` 在独立空白文档中，加载候选 DLL，通过 COM 调用真实 Word 形状与
候选 `ShrinkFontToFit`。不读取用户文档、不请求模型、不安装或切换 add-in 注册。
Word 版本：16.0，build 16.0.14332；候选程序集：2.1.36.0。

- 固定 30×20pt 框、合成可换行文本：规划为 Place；从 18pt 开始真实缩字号，在 2pt 放入，框尺寸不变。
- 同框从 1pt 开始：真实返回 fit=True，Word 保留 1pt，无溢出。
- 同框的另一合成文本：规划为 Place（估算字号约 3.46pt），真实缩至 1pt 仍 overflow=True。
- 对后一个样例只做实验性对照：`DisableLineHeightGrid=0`、单倍行距（报告 12pt）时仍溢出；
  改为 `DisableLineHeightGrid=-1` 后不溢出；使用精确 1.3pt 行距也不溢出；只加宽到 60pt 同样不溢出。

**已证明还有 Word 段落排版导致的降级，不能把“改成 1pt”描述成所有旁注问题都已解决。**
以上是受控样例的因果对照，不是对用户原图主因的断言。本候选不强改覆盖框继承的段落网格或行距。
Microsoft 的 [ParagraphFormat 属性文档](https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.disablelineheightgrid)
说明该属性涉及行网格；此处结论以本机实测数值为准。

### 诊断补齐

旧版只记录计划 `Place/SideNote`，实际 `PlaceOverlay` 返回失败后转旁注没有对应日志，
且字号设置/排版读取异常与真溢出都被标作“最小字号仍溢出”。
候选补充 `stage=plan/writeback`、`min_font=1`，并拆分上述固定原因。
到达 1pt 仍溢出时写入 `event=overlay_layout_failure`，仅记录字号、行距规则、行距数值、网格属性。
`stage=writeback` 行中的 font 是规划字号，真实失败字号见 `overlay_layout_failure`。
无图片、OCR 文字、正文、Prompt、剪贴板、密钥或异常 Message 新增到日志。

## 3. 构建、测试与版本

最终 `build.ps1 -Configuration Release`：成功，**0 错误**；7 个既有 nullable 警告
（Core 1、Word 2、旧测试 4）；WiX candle/light 成功生成 MSI。

最终全量 MSTest：**359/359 通过，失败 0、跳过 0，约 61 秒**。
基线 347 项，新增设置窗体 10 项、1pt 超限样例 1 项、诊断字段 1 项。
窗体检查直接加载三宿主 DLL，覆盖活动摘要/活动宿主、重复打开、设置创建过程中摘要关闭、
零/无效/已销毁 owner。全量包含既有普通文字、部分结果、取消、OCR、所有宿主回归。

初次 Session 0 运行窗体测试被 UserInteractive 检查拦截；改在已登录的 Session 1 通过一次性计划任务执行。
没有跳过 UI 测试，也没有把环境失败记作通过。Word 布局探针单独通过；新建文档未保存，测试 Word 已退出。
它验证真实 Office 布局，不等同于功能区端到端验收。

版本扫描范围包括 README、VERSION.txt、全部源码 AssemblyInfo、csproj、WiX、build.ps1 与 APPLY 文件名。
仓库原先没有 APPLY 文件；历史 About 更新日志和历史验证报告保留原始版本。
VERSION.txt 改用完整四段 `2.1.36.0`，构建/安装器不再重复追加 `.0`；
新增 `src/Directory.Build.props` 使 Core 自动生成的程序集版本也一致。
已核对 MSI ProductVersion、File 表中 Core/三宿主四个 DLL 版本，以及三处 COM Assembly 注册字符串，均为 2.1.36.0。
About 顶部版本为 2.1.36.0，并有独立中文更新日志段。改动 PS1 已以 UTF-8 BOM 保存，中文源码通过严格 UTF-8 检查。

执行命令（Release 完成后）：

```powershell
$env:PATH='C:\Program Files\dotnet;'+$env:PATH
.\build.ps1 -Configuration Release
$env:DOTNET_CLI_HOME=Join-Path (Get-Location) 'dotnet-home'
$env:NUGET_PACKAGES=Join-Path (Get-Location) 'packages'
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=all-final.trx' --results-directory TestResults\summary-settings-2.1.36
.\tests\RunWord236LayoutProbe.ps1 -EvidenceDirectory .\TestResults\summary-settings-2.1.36
```

后两条含 UI/Office 的命令在交互桌面会话执行。证据保留于本机忽略目录 `TestResults/summary-settings-2.1.36/`：
`baseline-desktop.trx`、`candidate-focused.trx`（110 项定向通过）、`build-final.txt`、`all-final.trx`、
`all-final.txt`、`word-layout.txt`、`word-layout.exit`（PASS）。早期探针调试错误文件保留，最终以对应 exit 和文件时间核对。

## 4. 本机交付与验收

候选 MSI：`artifacts/candidate-2.1.36.0/OfficeTranslate.Office.x64.msi`。
SHA-256：`B65D643BE1992EC624059304F4A3314BCB7FB0DC8C61DE5D145B6496BBAA8A59`。

在原始 Windows 环境安装候选后，保留“需处理”摘要，点击设置，关闭设置后确认摘要可继续复制；
重复打开，并检查另一文档窗口的设置归属。使用原始图片重试 OCR，按新增元数据区分 plan 和 writeback。
若仍出现句柄异常，需要异常堆栈/Win32 原生错误码；若仍为旁注，需要旁注原因或对应元数据日志，勿提供敏感原文。
此交付不含邮件、合并、标签、发布或本机安装。
