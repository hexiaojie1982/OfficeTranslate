# OfficeTranslate 应用说明

当前验收版本：**2.3.6.0**，分支 `feat/cjk-space-cleanup`，基于 `main` merge #28。未合并、未创建 tag、未 push。

本次修复 Chat 韩译中词间多余空格：AI 回答完成后、保存历史前执行确定性清理，最终气泡、历史、导出与复制一致。保留中英、中数和韩文空格、换行、围栏代码块及行内代码；用户消息和普通翻译路径不变。保留逐条复制、并行翻译、输入快捷键及翻译设置开关。

## 构建

在仓库根目录执行 `./build.ps1 -Configuration Release`，完整构建 Word、Excel、PowerPoint 加载项和 WiX MSI。

最终安装包：`artifacts/OfficeTranslate.Office.x64.msi`，仅适用于 64 位 Office。

## 安装或升级

1. 关闭 Word、Excel、PowerPoint，保存未完成的文档。
2. 运行 MSI，按安装向导完成安装或升级；不要仅替换 DLL，侧栏依赖 MSI 中的 COM 注册。
3. 重新启动 Office，在 OfficeTranslate 功能区打开「工具 → 文档内 Chat」，检查按钮分组、历史会话持久化、只读查看及删除确认。

静默部署命令：`msiexec /i OfficeTranslate.Office.x64.msi /qn /l*v install.log`。

功能测试及验证边界见 `tests/VALIDATION-2.3.6-CHAT-CJK-SPACING.md`。程序集、MSI 和 COM 注册版本统一为 **2.3.6.0**；本次自动化验证不代表实际 Office 宿主真机验收。
