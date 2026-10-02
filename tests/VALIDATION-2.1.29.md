# 2.1.29 三项验收缺口处理报告

日期：2026-10-02。基于已发布 main `2de89f2`，独立分支
`fix/validation-boundaries`。本次没有合并 main、发布 Release 或发邮件。

## 构建与测试

- VERSION.txt、三个宿主 AssemblyVersion/AssemblyFileVersion 均为 2.1.29(.0)；
  MSI 与关于窗口跟随版本源，更新日志含独立 2.1.29 条目。
- 完整 Release 构建成功，0 错误、7 条既有 nullable 警告。
- 最终 MSTest：123/123 通过，新增 13 项。
- 最终 MSI SHA-256：
  `6715B5FC7291F2400AE59022A38D87D90FF699CFF60909BCF753A8BA8E164F52`。
- 最终 Core SHA-256：
  `D042522DF10C6B2C08C446633EF009A3E34B82D2A75A23AE83031D54DA9BD8C2`；
  MVID：`6bcac013-c1fe-44b5-840d-1d56b4cda604`。

## 1. 真正的 Word UI 队列超时与任务代次

把原来 Word 私有票据逻辑提取为 Core 的 `UiRestoreGate`，Word 宿主仍在
进度/结果/错误窗口收尾后调用；超时等待采用 ConfigureAwait(false)，票据
失效不依赖被阻塞的 UI 队列。已开始的 COM 调用不承诺中途可撤销。

独立测试辅助组件**不进入 solution/MSI**，通过已有 Word 加载入口的临时
HKCU CLSID 覆盖载入。正常启动 Word；OnConnection 捕获 UI 调度器；调用时
核对 HWND 原生线程 ID 等于当前原生线程 ID。不是另开 WinForms STA 线程来
冒充 Office 队列，也不是只用代理 Range。使用真实合成 Word Range。

最终实测（native UI thread=14292，managed UI thread=1，Core MVID 如上）：

| 情况 | 结果 |
| --- | --- |
| 实际 Word UI 线程无消息泵阻塞 2.4s，放行前选择新范围 7–12 | expired；旧恢复调用 0 次；最终仍 7–12 |
| 排队期间提升任务代次，选择新范围 7–12 | superseded；旧恢复调用 0 次；最终仍 7–12 |
| 正常队列 | 恢复调用 1 次；最终 0–5 |

另有单测覆盖已开始回调与超时竞争、正常恢复、回调异常不掩盖任务结果。
临时 HKCU 注册已恢复并核验；没有执行机器级辅助注册或修改安全策略。

**边界**：实测的是生产共用票据在真正 Office 队列上的行为；没有据此声称
所有翻译菜单、页眉/页脚视图切换或取消路径在本版本重新做过全量回归。

## 2. 512MiB 预捕获容量：安全测试入口及提前保护

Word 使用 `ImagePreCapture.CaptureAllAsync` 完成全部捕获后才进入 OCR/写回。
捕获委托与预算按调用传入，不新增全局静态注入后门。测试和宿主使用同一
`ImageCaptureBudget` 算法，而非复制一个模拟实现。

- 真正的 512MiB 数值边界：恰好达到允许，多 1 字节拒绝；不分配巨型数组。
- 超大 long 计数不会溢出；拒绝后预算不变化。
- 小容量注入到**生产批量流程**：超限不返回部分批次，不进入下游处理；
  余额为零时不再捕获下一张；捕获后取消也不接受批次。
- Word/Excel 剪贴板 PNG 编码时检查单图 128MiB 及 Word 剩余预算，避免等完整
  byte[] 生成后才发现超限。像素保护 8000px/边、3200万总像素。
- 真实 GDI+ 编码发现：它可能吞掉 IStream 超限异常并返回“成功”的截断输出。
  已加入编码后的粘性超限标志校验；测试证实超限显式失败、正常 PNG 可解码。

**边界**：这是安全的容量/流程边界验证，未向 Word 灌入超过 512MiB 的真实
图片数据，也不等同于内存压力/OOM 测试。预算限制保留的 PNG 数据，不限制
Office 自身或 Clipboard.GetImage 的首次解码分配、GDI/托管临时副本、进程 RSS。
原来的“512MB 防止 OOM”不能作为硬保证。

## 3. MSI 升级、加载与人工回退流程

- 起点：正式安装 2.1.23.0，ProductCode
  `{6B793222-771A-4E72-85D1-46DD4F9E576A}`。
- 回退介质取自官方 v2.1.23 Release，SHA-256
  `AC0B70D88C8A4CEAEBDA9CEFF144249571494634E8FC7AD521DA445FF1A057C9`。
- 初次无管理员令牌的静默升级失败（1603/1730）；旧安装完整未变。
  正常 RunAs/UAC 提升后最终包升级成功（MSI 返回 0）。
- 最终 2.1.29.0 ProductCode：`{079558FA-BA61-4F29-B806-2ED8FA503D45}`，
  与旧版 UpgradeCode 相同：`{5D99BB0D-887E-4AA7-A504-B763AC97F899}`。
- 升级后仅有一个安装项，三个宿主 DLL 版本正确，Core.dll 哈希与候选一致，
  原设置未改。
- 正常 UI 启动三个宿主：Word/Excel/PowerPoint 的 OfficeTranslate 菜单均出现，
  COMAddIns.Connect 均为 True。这里只记加载烟测，不冒充翻译/OCR 端到端。
- 卸载候选 + 重新安装旧包均返回 0。独立 Verify 校验原 ProductCode/版本、
  全部安装文件、六个原 HKLM COM/Office 注册项及 settings.json 均与快照一致。

**边界**：验证的是成功升级以及卸载重装的恢复步骤；不是在安装中途注入故障
后 Windows Installer 自动事务回滚的保证。

## 证据与本机最终状态

本机原始证据：`validation-fixtures/review-2.1.29-evidence/final/`，包含 TRX、
三份真实队列结果、MSI 日志和原/恢复注册导出。MSI 验证脚本为
`validation-fixtures/review-2.1.29-msi.ps1`；队列辅助项目和使用说明在本目录。

测试完恢复原 2.1.23 安装与设置；Office 测试实例退出；临时辅助注册不保留。
2.1.29 是本地候选，MSI 在 `artifacts/OfficeTranslate.Office.x64.msi`。
