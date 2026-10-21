# OfficeTranslate 应用说明

当前开发版本：**2.3.9.0**；分支 `fix/security-audit-findings`，基于 main `3fba5c2`（v2.3.8）。本次只提交 PR，未合并、打 tag、发布、安装或发送文件。

## 安全修复

- Excel 替换与双语翻译通过共享字面文本 helper 写回，避免模型输出成为公式。
- Chat 响应最多 8 MiB 实际字节/累计字符，图片响应最多 16 MiB 实际字节；错误响应同样受限。
- 翻译设置和 Chat 设置保存时，公网 HTTP 携带 API Key 会明确警告明文传输风险，默认选择不保存。localhost、私有网段、链路本地和仅解析到本地地址的局域网域名不弹该警告。
- 图片最多处理前 512 个区域；超出部分显示丢弃数量及待检查提示，结果窗保持打开供确认。Word / Excel / PowerPoint 的 OCR 和图片翻译写回逐条检查取消。

## 构建与测试

在此分支根目录执行：

```powershell
./build.ps1 -Configuration Release -SkipInstaller
dotnet test OfficeTranslate.Word.sln -c Release --no-build --no-restore --logger 'trx;LogFileName=security-after.trx' --results-directory TestResults
```

全量测试包含真实窗口、键盘、剪贴板及 Excel COM 集成用例，须在已登录的交互桌面运行。不要通过过滤或跳过测试来声称全量通过。基线、最终结果及环境限制见 [2.3.9 验证记录](tests/VALIDATION-2.3.9-SECURITY-AUDIT.md)。

`VERSION.txt`、三宿主 AssemblyVersion / FileVersion、Core Version 均为 **2.3.9.0**。WiX 的 ProductVersion 和 COM 注册版本由构建脚本读取同一版本文件。本次使用 `-SkipInstaller`，未生成或验证 2.3.9 MSI；现存 artifacts 下的旧包不能代表本次代码。

## 功能使用

Word / Excel / PowerPoint 共用 Chat 面板，功能组为「翻译 / 总结 / 改写 / Chat 设置」。翻译直接使用翻译服务当前目标语言，选中优先，否则读取全文。Chat 提示词可编辑并持久化；「使用翻译设置」控制是否追加风格、额外要求和术语表。

Question / Rewrite / Translate 完成后可确认替换选中或插入下方；Summarize 仅插入开头。所有模式在发送开始时清空输入框，完成回调保留下一条草稿。CJK 清理保留韩文间距、中英/中数边界、换行及代码。

合并、正式包构建、安装、发布及邮件发送均由后续独立核验流程执行。
