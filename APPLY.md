# OfficeTranslate 应用说明

当前开发版本：**2.3.7.0**，分支 `feat/chat-translate-button`，起点为 `main` 的 `a6bbdfa`（2.3.5）；尚未合并、打 tag、push 或发布。

三宿主共用 Chat 面板：功能组为「翻译 / 总结 / 改写 / Chat 设置」。翻译直接读取当前翻译服务的目标语言，按选中优先、否则全文获取上下文；结果走现有 Chat SSE 流式显示。翻译系统提示词可编辑并持久化；风格、额外要求和术语表沿用「使用翻译设置」开关。译文提供「替换选中」（目标可替换时启用）和「插入到下方」，不提供插入到开头。

## 构建

在本分支工作树根目录执行 `./build.ps1 -Configuration Release`，构建 Word、Excel、PowerPoint 加载项和 WiX MSI。`VERSION.txt`、三宿主程序集、Core 包版本及 MSI/COM 注册统一为 **2.3.7.0**。

安装包：`artifacts/OfficeTranslate.Office.x64.msi`，仅适用于 64 位 Office。

## 安装或升级

1. 关闭 Word、Excel、PowerPoint，保存未完成的文档。
2. 运行 MSI，按安装向导完成安装或升级；不要仅替换 DLL，侧栏依赖 MSI 中的 COM 注册。
3. 重新启动 Office，打开「工具 → 文档内 Chat」，检查翻译按钮位于总结之前、Chat 设置位于功能组最后。
4. 切换翻译服务的目标语言，分别以选中文字和无选中全文翻译；检查流式、复制、历史、导出与写回确认。
5. 编辑翻译系统提示词，关闭重开设置窗确认保存；保存翻译服务设置后再次确认提示词仍保留。

静默部署命令：`msiexec /i OfficeTranslate.Office.x64.msi /qn /l*v install.log`。

本次不包含 `feat/cjk-space-cleanup`（2.3.6），原分支和工作树保持原样。超出上下文预算时沿用现有分段摘要机制，译文会基于摘要，界面有提示。测试、安装包哈希和验证边界见 [2.3.7 验证报告](tests/VALIDATION-2.3.7-CHAT-TRANSLATE.md)。本次未执行真实 Office 安装或远程模型人工验收。

本次全量回归 **858/858 通过，0 失败、0 跳过**，比 main 新增 55 项用例；Release 构建 **0 错误、6 条既有警告**，MSI 数据库与解包载荷版本均已核对。
