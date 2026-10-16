# OfficeTranslate 应用说明

当前正式版本：**2.3.2.0**；Git tag：`v2.3.2`。

## 构建

在仓库根目录执行 `./build.ps1 -Configuration Release`，完整构建 Word、Excel、PowerPoint 加载项和 WiX MSI。

最终安装包：`artifacts/OfficeTranslate.Office.x64.msi`，仅适用于 64 位 Office。

## 安装或升级

1. 关闭 Word、Excel、PowerPoint，保存未完成的文档。
2. 运行 MSI，按安装向导完成安装或升级；不要仅替换 DLL，侧栏依赖 MSI 中的 COM 注册。
3. 重新启动 Office，在 OfficeTranslate 功能区打开「工具 → 文档内 Chat」，检查左右聊天气泡、流式回答、滚动和导出。

静默部署命令：`msiexec /i OfficeTranslate.Office.x64.msi /qn /l*v install.log`。

本次交付只生成安装包，不自动安装。验证范围及真机验收项目见 `tests/VALIDATION-2.3.2-CHAT-BUBBLES.md`。
