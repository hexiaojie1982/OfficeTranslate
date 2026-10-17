# OfficeTranslate 应用说明

当前正式版本：**2.3.5.0**，发布分支 `main`，Git tag：`v2.3.5`。

本次新增聊天和历史会话逐条复制按钮：用户复制原文，AI 复制 Markdown 原文；流式与空消息隐藏，成功/失败提示约 1.5 秒。保留并行翻译、历史持久化、输入快捷键及翻译设置开关。

## 构建

在仓库根目录执行 `./build.ps1 -Configuration Release`，完整构建 Word、Excel、PowerPoint 加载项和 WiX MSI。

最终安装包：`artifacts/OfficeTranslate.Office.x64.msi`，仅适用于 64 位 Office。

## 安装或升级

1. 关闭 Word、Excel、PowerPoint，保存未完成的文档。
2. 运行 MSI，按安装向导完成安装或升级；不要仅替换 DLL，侧栏依赖 MSI 中的 COM 注册。
3. 重新启动 Office，在 OfficeTranslate 功能区打开「工具 → 文档内 Chat」，检查按钮分组、历史会话持久化、只读查看及删除确认。

静默部署命令：`msiexec /i OfficeTranslate.Office.x64.msi /qn /l*v install.log`。

功能测试及验证边界见 `tests/VALIDATION-2.3.5-CHAT-COPY.md`。全量测试 **803/803 通过，0 失败、0 跳过**；用户已完成真机验收并授权合并发布。正式 MSI 从合并后的 `main` 执行 Release 清理及全量重建，程序集、MSI 和 COM 注册版本统一为 **2.3.5.0**。
