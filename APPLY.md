# OfficeTranslate 应用说明

当前正式版本：**2.3.4.0**（`v2.3.4`），从 `main` 构建。

本次新增 Word 正文有界并行翻译及按序增量写回，翻译设置可调并发数（默认 4）。保留已发布的历史会话、输入快捷键、翻译设置开关；图片 OCR 翻译仍串行。

## 构建

在仓库根目录执行 `./build.ps1 -Configuration Release`，完整构建 Word、Excel、PowerPoint 加载项和 WiX MSI。

最终安装包：`artifacts/OfficeTranslate.Office.x64.msi`，仅适用于 64 位 Office。

## 安装或升级

1. 关闭 Word、Excel、PowerPoint，保存未完成的文档。
2. 运行 MSI，按安装向导完成安装或升级；不要仅替换 DLL，侧栏依赖 MSI 中的 COM 注册。
3. 重新启动 Office，在 OfficeTranslate 功能区打开「工具 → 文档内 Chat」，检查按钮分组、历史会话持久化、只读查看及删除确认。

静默部署命令：`msiexec /i OfficeTranslate.Office.x64.msi /qn /l*v install.log`。

功能测试及验证边界见 `tests/VALIDATION-2.3.4-PARALLEL-TRANSLATE.md`。正式版构建后，安装和实际模型端到端验收仍需单独执行。
