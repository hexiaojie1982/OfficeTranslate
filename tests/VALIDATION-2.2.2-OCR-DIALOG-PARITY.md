# OfficeTranslate 2.2.2：OCR 结果窗行为对齐

日期：2026-10-14。基线 `main / 6a7b281`；本地分支 `fix/ocr-dialog-parity-2.2.2`。
项目四段版本为 `2.2.2.0`。本轮未 push、合并、发布或安装。

## 修改文件与范围

- `src/OfficeTranslate.WordAddIn/OcrResultForm.cs`：沿用翻译结果窗的 5 秒、每秒更新关闭按钮及归零关闭逻辑；窗口显示后才启动计时。`error != null` 或有未放置旁注时不启动计时，保持完整、可复制的恢复内容。保留 OCR 窗原有模态显示、owner 和布局，关闭时停止并释放计时器。
- `src/OfficeTranslate.WordAddIn/UiText.cs`：只新增 OCR 结果文案键，包含标题、完成/取消/失败、统计、旁注提示及倒计时；关闭按钮复用 `Close`。简体/繁体中文、英文有对应文案，韩语/日语新键暂用英文。`OcrResultCloseCountdown` 使用“关闭（5s）”，已有 `ResultCloseCountdown` 与翻译结果窗不变。
- `src/OfficeTranslate.WordAddIn/AddIn.cs`、`src/OfficeTranslate.AddInShared/AddIn.cs`：仅 OCR 结果展示入口捕获并传递本次任务的界面语言，覆盖成功、取消及异常分支。
- `README.md`、`VERSION.txt`、`build.ps1`、三个宿主的 `Properties/AssemblyInfo.cs`、共享 `AboutForm.cs`：同步 2.2.2 版本及更新日志，保留历史记录。
- `tests/OfficeTranslate.Tests/OcrResultFormTests.cs`：新增 20 个数据驱动回归用例，直接加载三个实际宿主 DLL；覆盖倒计时启动、5→1 秒按钮更新、实际自动关闭、计时器释放、旁注/错误/取消/空摘要错误不启动、五种界面的文案与统计。
- `tests/OfficeTranslate.Tests/PureOcrHostTests.cs`：适配新增语言参数和成功时倒计时按钮，保留长文本及样式回归。
- 本验证记录。

冻结路径比较通过：Core、识别服务、翻译服务、`TranslationResultForm`、`TranslationProgressForm` 共 39 个文件与 main 相同；两个 AddIn 文件在 `RunOcrAsync` 之外全文相同。未改 OCR 请求、识别、写回及 Word/Excel/PowerPoint 业务逻辑。

## 验证结果

1. `build.ps1 -Configuration Release` 成功，**0 错误**，7 个已有可空性警告；本地 MSI 生成成功。
2. 交互桌面全量测试 **437/437 通过，0 失败，0 跳过**，退出码 0；Session 1 / `UserInteractive=True`。新增 OCR 回归 **20/20** 通过。
3. Word / Excel / PowerPoint 的实际 `ShowResult` 模态入口验证成功场景自动关闭，测得测试时长约 5.38 / 5.04 / 5.04 秒（包含创建和观察开销）；按钮显示完整 5、4、3、2、1 秒序列。恢复场景明确验证计时器未启用，按钮无倒计时。
4. Core、Word、Excel、PowerPoint 四个 DLL 的 AssemblyVersion/FileVersion，三个 About 的 CurrentVersion，MSI ProductVersion、四个产品 DLL 的 File 表版本及三项 COM Assembly 注册版本均为 `2.2.2.0`。
5. 全仓库版本扫描通过，旧版本仅保留于历史验证记录、历史链接及历史更新日志；`git diff --check` 通过。测试临时计划任务已删除。

未安装候选加载项，也未连接真实模型执行 Office 人工端到端验收；本轮变更由实际宿主 DLL 的窗体测试验证。

## 复现与证据

```powershell
.\build.ps1 -Configuration Release
# 已有窗体及剪贴板测试要求交互桌面，在登录会话执行：
dotnet test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=all-desktop.trx' --results-directory TestResults\ocr-dialog-parity-2.2.2
```

本机证据目录 `TestResults/ocr-dialog-parity-2.2.2/`（git 忽略）：`build.log`、`desktop-test.log`、`all-desktop.trx`、`desktop-test.exit`、`desktop-context.txt`、`regression-summary.log`、`version-verification.log`、`version-scan.log`、`frozen-paths.log`。

本地 MSI：`artifacts/OfficeTranslate.Office.x64.msi`。
SHA-256：`B7294C7AF32361C008EB00C2375FBD58BA58AB2303507D7BF47AB6CB605A1B08`。
