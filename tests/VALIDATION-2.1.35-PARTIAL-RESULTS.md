# OfficeTranslate 2.1.35：部分结果与 OCR 失败回退

复核日期：2026-10-11。基线 `b81139268e062c4578d232de504a8b2a30908e7f`（main / v2.1.34）。本地分支 `fix/translation-partial-results`。

## 修复与边界

### R1：逐段回退遇到连续空回复

- 引入专用内部异常 `EmptyTranslationResponseException`，保留原错误文案及 InvalidOperationException 继承关系。
- 只在混合文本逐段回退边界捕获该异常：该段保留原文并计入检查，继续后续段落。原分隔符、首尾空白不变，成功段保留，部分结果不进入成功缓存。
- 取消在恢复前再次检查。鉴权错误、网络失败、解析错误等不吞掉；独立单段连续空回复仍明确失败，不伪装成功。

### R2：后续错误或取消吞掉先前未放置旁注

- 三宿主 service 在任务初始阶段暴露本次 Summary；入口先清空上一次引用。AddIn 在任务正常返回前也能取得已记录的旁注全文。
- 共用 `TranslationTaskFailure` 和 `TranslationErrorDialog.ShowTaskFailure`：有未放置内容时显示可复制、不会自动关闭的恢复摘要，分别标注“已取消”或“已中止”，错误路径保留原错误信息；不标作全部完成。
- 没有未放置内容时沿用旧行为：主动取消不弹错误，非主动取消仍显示原错误。
- 三宿主正文统计移到写回成功之后；非空图片的 OCR 统计在图片提交流程返回后记录，空识别仍单独计数。避免等待写回时取消却提前增加成功数。
- Word 恢复摘要为非模态窗口，诊断使用 `recovery_summary_show/returned`，不把 Show 返回误记成用户已关闭窗口。未变更原选区恢复票据、所有权清理、取消的单项提交边界。

### R3：迭代复查发现 Excel/PPT 覆盖框初始化 COM 异常绕过旁注

- Excel 原来创建失败直接 throw；PowerPoint 的 AddTextbox/属性初始化无对应清理保护。这会在旧结果清理之后中断后续区域，PPT 还可能留下未初始化完的框。
- 共用 `ImageOverlayWriteback.TryCreate` 包含创建、初始化及失败清理。COM 错误只把当前区域降级到原旁注通道，其余区域继续；只删除当前失败对象。
- 删除失败会明确说明“未能删除失败的框，请检查原图区域”，不谎报清理成功。取消与非 COM 异常在清理后原样抛出，不 blanket retry。
- 不改变原 5x/6pt 几何、真实排版校验或旁注 1400pt 上限。Word 本来已有覆盖框失败转旁注逻辑，不重复重写。

## 测试清单

基线 313 项，新增 34 个执行用例，共 347 项。

`MixedParagraphFallbackTests` 新增 5 方法 × OpenAI-compatible/Ollama = 10 项：

- `EmptyPureMiddleTwice_PreservesThatParagraphContinuesAndDoesNotCache`
- `EmptyStandaloneParagraphTwice_StillFailsRatherThanClaimingSuccess`
- `Fallback_UnauthorizedMiddleIsNotConvertedToReview`
- `Fallback_NetworkFailureMiddlePropagatesAfterExistingRetries`
- `Fallback_CancelDuringEmptyRetryDoesNotContinue`

`TranslationTaskFailureTests`：6 项，覆盖真实 splitter/placement 协调逻辑的“首图一段失败 + 后续超时/取消”、全文与已成功内容保留、无旁注取消、摘要建立前失败、非用户 OCE，以及取消同时出现其他错误。

`InterruptedResultFormTests`：13 项，使用真实 STA WinForms：

- 五种界面语言下保留 40000 字以上内容及换行，正确显示错误/取消标题，不显示“全部完成”，关闭按钮无倒计时且 Timer 未开启。
- 普通成功仍开启原五秒倒计时；无旁注取消不新开窗口。
- **直接加载编译后的 Word/Excel/PowerPoint DLL，各测失败和取消，共 6 项**。通过实际共享 presenter 打开各 DLL 自己的恢复窗口并核验完整文本、标题及 Timer，不只是测试项目链接的窗体副本。
- 测试项目以 `ReferenceOutputAssembly=false` 声明宿主构建依赖，避免只运行 dotnet test 时误测旧 DLL；不向单测编译引用引入 Office COM 类型。路径按当前 Debug/Release 配置解析。

`ImageOverlayWritebackTests`：5 项，覆盖创建失败、初始化失败只删除该框且前后兄弟区域继续、取消清理后原异常透传、非 COM 错误透传、删除失败的明确提示。

## 实际命令

目录 `D:\pythonProject\office-translate\native-word`，.NET SDK 8.0.423 / net48：

```powershell
$env:PATH='C:\Program Files\dotnet;'+$env:PATH
.\build.ps1 -Configuration Release -SkipInstaller
$env:DOTNET_CLI_HOME=Join-Path (Get-Location) 'dotnet-home'
$env:NUGET_PACKAGES=Join-Path (Get-Location) 'packages'
& 'C:\Program Files\dotnet\dotnet.exe' test tests\OfficeTranslate.Tests\OfficeTranslate.Tests.csproj -c Release --no-build --no-restore --logger 'trx;LogFileName=2.1.35-final.trx' --results-directory TestResults\partial-results-2.1.35
git diff --check
```

过程证据（均位于本地忽略目录 `TestResults/partial-results-2.1.35`）：

- `2.1.35-focused-1.trx`：37/43，6 项失败。原因是新负面用例的模拟整块回复仍保留源文中实际存在的英文术语，意外通过整块校验，未进入预期回退分支。修正为确定丢失术语的响应，未删除断言、未削弱产品校验。
- `2.1.35-focused-2.trx`：43/43，失败/跳过 0。
- `2.1.35-all-1.trx` 和 `2.1.35-all-2.trx`：各 342/342，失败/跳过 0，含三宿主实际 DLL 的恢复窗口检查。
- `2.1.35-all-3.trx`：347/347，失败/跳过 0，增加 R3 的 5 项。
- 最终 build 日志 `build-final.txt`：退出 0，0 错误、2 个既有 Word nullable 警告；此前完整重编译为 7 个既有 nullable 警告（Core 1、Word 2、旧测试 4），增量构建不会每次重报。没有新增代码警告，版本检查通过，明确跳过 installer。
- 最终 `2.1.35-final.trx` / `test-final.txt`：**347/347，失败 0、跳过 0，56 秒，退出 0**。最后一次对格式及更新日志整理后的同一产物复跑全量，而非复用旧一轮通过数。

最终 Release DLL SHA-256：

| DLL | SHA-256 |
| --- | --- |
| Core | `05270AA0A6ECE13B2DD7D20FC2A542046F2E897D81B080F79108DAF04F874656` |
| Word | `CA4D439E5BDAB5B0341382D445285DE3965B52CF39EDCE6B5FF8795FA01FA769` |
| Excel | `2C8572D580BC8A1B21D6A3F3B9FEB2798450160CA37575DE696BA3195ACBF65D` |
| PowerPoint | `13821E273319C2372B6FDE04EBA833A3DB694233ABED630C419D86C9486551D4` |

## 交付限制

- VERSION.txt、三个宿主 AssemblyVersion/AssemblyFileVersion 同步为 2.1.35 / 2.1.35.0；About 新增 2.1.35，保留独立 2.1.34 历史日志。
- 本地候选分支提交；没有 push、merge、tag、MSI 构建、安装、发布或邮件发送。main/v2.1.34 保持基线，未触碰 review-image-overlay。
- 模型调用为受控 HTTP handler，没有访问公司服务器或读取真实公司照片。
- 宿主 DLL 窗体集成不是 Office 菜单端到端验收；COM 创建失败使用故障注入代理动作，不声称真实 Excel/PPT 人为故障已复现。真实菜单、Office 文档布局和公司模型仍需现场验收。
- 结论限于已确认问题和本报告覆盖的回归：不能由单测全绿推断任意文档、任意模型或真实公司环境绝无缺陷。
