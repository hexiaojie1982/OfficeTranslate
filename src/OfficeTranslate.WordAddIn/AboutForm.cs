using System;
using System.Drawing;
using System.Windows.Forms;

namespace OfficeTranslate.WordAddIn
{
    internal sealed class AboutForm : Form
    {
        // Version is read from the assembly so it can never drift from the
        // AssemblyVersion declared in AssemblyInfo.cs (checked by build.ps1
        // against VERSION.txt).
        public static string CurrentVersion
        {
            get
            {
                var version = typeof(AboutForm).Assembly.GetName().Version;
                return version == null ? string.Empty : version.Major + "." + version.Minor + "." + version.Build;
            }
        }
        private readonly string _language;
        private readonly TextBox _content = new TextBox();

        public AboutForm(string uiLanguage)
        {
            _language = uiLanguage;
            Text = UiText.Get(_language, "AboutTitle");
            Width = 620; Height = 520; MinimumSize = new Size(520, 420);
            StartPosition = FormStartPosition.CenterParent; Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.FromArgb(247, 249, 252);

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), RowCount = 3, ColumnCount = 1 };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var header = new Label { Text = "OfficeTranslate  " + CurrentVersion, AutoSize = true, Font = new Font(Font.FontFamily, 16F, FontStyle.Bold), ForeColor = Color.FromArgb(32, 55, 88), Margin = new Padding(0, 0, 0, 14) };
            root.Controls.Add(header, 0, 0);
            _content.Dock = DockStyle.Fill; _content.Multiline = true; _content.ReadOnly = true; _content.TabStop = false; _content.ScrollBars = ScrollBars.Vertical; _content.BackColor = Color.White; _content.BorderStyle = BorderStyle.FixedSingle; _content.Font = new Font(Font.FontFamily, 10F); _content.Text = AboutText();
            root.Controls.Add(_content, 0, 1);

            var actions = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2, Padding = new Padding(0, 14, 0, 0) };
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); actions.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var left = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
            var changelog = Button(UiText.Get(_language, "Changelog"), false); changelog.Click += (s, e) => _content.Text = ChangelogText();
            var guide = Button(UiText.Get(_language, "Guide"), false); guide.Click += (s, e) => _content.Text = GuideText();
            left.Controls.Add(changelog); left.Controls.Add(guide); actions.Controls.Add(left, 0, 0);
            var close = Button(UiText.Get(_language, "Close"), true); close.DialogResult = DialogResult.OK; actions.Controls.Add(close, 1, 0);
            root.Controls.Add(actions, 0, 2); Controls.Add(root); AcceptButton = close; CancelButton = close; ActiveControl = close;
        }

        private static Button Button(string text, bool primary)
        {
            var button = new Button { Text = text, AutoSize = true, Height = 34, Padding = new Padding(12, 3, 12, 3), Margin = new Padding(0, 0, 8, 0), FlatStyle = FlatStyle.Flat };
            button.FlatAppearance.BorderSize = primary ? 0 : 1;
            button.FlatAppearance.BorderColor = Color.FromArgb(190, 200, 214);
            button.BackColor = primary ? Color.FromArgb(45, 108, 223) : Color.White;
            button.ForeColor = primary ? Color.White : Color.FromArgb(42, 58, 78);
            button.UseVisualStyleBackColor = false;
            return button;
        }

        private string AboutText()
        {
            if (_language == "en") return "Native translation add-ins for Microsoft Word, Excel, and PowerPoint.\r\n\r\nVersion: " + CurrentVersion + "\r\nArchitecture: x64 native COM add-ins\r\nConfiguration: shared by all three Office applications\r\nTranslation services: Ollama API and OpenAI-compatible API";
            if (_language == "ja") return "Microsoft Word、Excel、PowerPoint 用ネイティブ翻訳アドイン。\r\n\r\nバージョン：" + CurrentVersion + "\r\nアーキテクチャ：x64 ネイティブ COM アドイン\r\n設定：3 つの Office アプリで共有\r\n翻訳サービス：Ollama API / OpenAI 互換 API";
            if (_language == "ko") return "Microsoft Word, Excel, PowerPoint용 네이티브 번역 추가 기능입니다.\r\n\r\n버전: " + CurrentVersion + "\r\n아키텍처: x64 네이티브 COM 추가 기능\r\n설정: 세 Office 앱에서 공유\r\n번역 서비스: Ollama API / OpenAI 호환 API";
            if (_language == "zh-HK") return "適用於 Microsoft Word、Excel 及 PowerPoint 的原生翻譯增益集。\r\n\r\n版本：" + CurrentVersion + "\r\n架構：x64 原生 COM 增益集\r\n設定：三個 Office 應用程式共用\r\n翻譯服務：Ollama API、OpenAI 相容 API";
            return "适用于 Microsoft Word、Excel 和 PowerPoint 的原生翻译加载项。\r\n\r\n版本：" + CurrentVersion + "\r\n架构：x64 原生 COM 加载项\r\n配置：三个 Office 应用共享\r\n翻译服务：Ollama API、OpenAI-compatible API";
        }

        private string ChangelogText()
        {
            // 3b2a252 review: the title already shows the current candidate
            // version dynamically (CurrentVersion); every section below
            // carries its own version header so the entry text stays
            // consistent across the three hosts sharing this window.
            return "OfficeTranslate " + CurrentVersion + "\r\n• Word 图片预捕获使用可验证的容量预算，累计上限 512MiB；Word/Excel 剪贴板 PNG 编码过程中限制单图 128MiB，并检查 Word 的剩余预算。该限制不代表 Office 总进程内存上限。\r\n• 最终选区恢复采用共用一次性 UI 票据，超时失效不依赖 Office UI 消息队列，过期及被新任务取代的回调不触碰选区。\r\n\r\n" + ChangelogVersion128Text();
        }

        private string ChangelogVersion128Text()
        {
            return "OfficeTranslate 2.1.28\r\n• 修复错误框 owner 校验逻辑：已关闭源文档的缓存窗口句柄可能被新窗口复用（IsWindow 仍返回真），现改为校验缓存 HWND 是否仍属于当前 Word 窗口集合；不属于则改用当前存活窗口或无 owner 显示，避免错误框关闭后留下阻塞 Word 退出的空白窗口。\r\n\r\n" + ChangelogVersion127Text();
        }

        private string ChangelogVersion127Text()
        {
            return "OfficeTranslate 2.1.27\r\n• 修复等待网络时关闭源文档后错误提示的收尾问题：错误框不再使用已销毁的源文档窗口作为 owner，关闭前校验窗口存活，失效时改用当前 Word 实例的存活窗口或无 owner 显示；错误框关闭与最终恢复均有日志，不再留下阻塞 Word 操作与退出的空白窗口。Excel/PowerPoint 错误框做同类 owner 校验。\r\n• 多轮取图证据截断严格限定 2000 字符（含截断标记长度），截断不拆散代理对。\r\n• Word 图片预捕获增加 512MB 总量上限：超限时在写回任何内容前明确停止并提示分批翻译。\r\n• 原生剪贴板格式枚举达到 64 项上限时标注“仅列出前64项”，不再把部分列表当作完整枚举。\r\n\r\n" + ChangelogVersion126Text();
        }

        private string ChangelogVersion126Text()
        {
            return "OfficeTranslate 2.1.26\r\n• Word 图片任务改为先捕获全部图片、再逐图识别与写回：多图全文翻译时，前一张图的书签与覆盖框写回不再干扰后一张图的剪贴板捕获，修复第二张图报 0x800A11FD 的失败；捕获失败则在写回任何内容之前停止。\r\n• 原生剪贴板格式枚举读取错误码，区分“枚举结束”与“枚举失败”，避免把枚举失败误报为“无格式”。\r\n• 多轮取图证据的长度截断改为拼接完成后校验，避免超长单轮证据突破上限。\r\n\r\n" + ChangelogVersion125Text();
        }

        private string ChangelogVersion125Text()
        {
            return "OfficeTranslate 2.1.25\r\n• Word 翻译任务绑定入口文档：网络等待期间切换到其他文档，结果仍写回原文档；源文档关闭则明确停止，不再回退到活动文档。\r\n• 剪贴板诊断改用 Win32 原生格式枚举（仅记录格式 ID 与名称，不读内容），区分“剪贴板无数据”与“数据不可读”。\r\n• EMF 获取改走 Win32 句柄复制，不再依赖托管 DataObject 对 metafile 的识别；取消正常穿透，暂不可用数据有界重试并记录阶段。\r\n• 光栅化失败路径释放位图并设像素上限；多轮取图证据聚合保留最早轮次；格式读取失败与空格式区分记录。\r\n\r\n" + ChangelogVersion124Text();
        }

        private string ChangelogVersion124Text()
        {
            return "OfficeTranslate 2.1.24\r\n• 图片复制诊断记录目标选区与实际选区的行内图状态，定位复制与剪贴板环节。\r\n• 视图修复绑定任务入口窗口，不再误改无关窗口；入口窗口未知时直接跳过。\r\n• 位图轮询失败后尝试读取剪贴板 metafile 并按原生分辨率光栅化为 PNG 兜底。\r\n• 复制硬失败时链入此前各轮取图的剪贴板序号、格式与轮询证据。\r\n\r\n" + ChangelogImageOverlayText();
        }

        private string ChangelogImageOverlayText()
        {
            return "图片 OCR 译文框定位\r\n• 图片 OCR 译文框定位重构：统一 bbox 校验与坐标换算，Word 浮动图片继承自身相对定位基准，修复译文框错位。\r\n• 遮盖层改为不透明并严格限定在原文区域内；译文在区域内换行并自动缩小字号，不再向下扩高；放不下或坐标不可信时生成“图片译文待检查”旁注。\r\n• 重复翻译同一图片时更新已有译文框，不再叠加。\r\n• 新增图片几何诊断日志：记录捕获像素尺寸、模型原始 bbox、Office 图片矩形与最终文本框矩形，便于区分模型定位误差与坐标换算误差（不记录图片内容与密钥）。\r\n• 模型返回的原始 bbox 不再预先截断，异常坐标交由校验器判定并降级为旁注。\r\n\r\n" + ChangelogVersion223Text();
        }

        private string ChangelogVersion223Text()
        {
            return "OfficeTranslate 2.1.23\r\n• 增强翻译请求的超时与重试处理，并对异常或截断的模型结果停止写回。\r\n• Excel 按区域批量读取单元格，提高大范围翻译时的读取效率并修复数组边界问题。\r\n• 图片 OCR 增加兼容服务结构化输出回退，并改进剪贴板捕获失败时的处理。\r\n• 修复 MSI 安装目录及版本一致性检查，补充回归测试。\r\n\r\n" + ChangelogVersion222Text();
        }

        private string ChangelogVersion222Text()
        {
            return "OfficeTranslate 2.1.22\r\n• 对 Word、Excel 和 PowerPoint 的换行符进行程序级保护，避免多行文本翻译后合并成一行。\r\n• 支持保留 CRLF、CR、LF、Word 手动换行、分页符及 Unicode 行分隔符的原始类型和顺序。\r\n• 正常翻译仍保持整段一次请求，仅在模型破坏换行时进行一次占位符安全重试。\r\n\r\n" + ChangelogVersion221Text();
        }

        private string ChangelogVersion221Text()
        {
            return "OfficeTranslate 2.1.21\r\n• 新增目标语言文字校验，避免中英混合段落翻译为韩语时仍返回中文却被误判为成功。\r\n• 中文源语言纳入残留文字检测，整段未翻译或严重漏译会自动安全重试。\r\n• 自动检测模式会检查非目标文字残留，且提示词明确要求韩文或日文的目标书写系统。\r\n\r\n" + ChangelogVersion220Text();
        }

        private string ChangelogVersion220Text()
        {
            return "OfficeTranslate 2.1.20\r\n• DeepSeek 官方接口的文本翻译明确关闭思考模式，避免思考 Token 挤占译文空间。\r\n• 空正文或校验失败时的唯一一次安全重试使用扩展输出预算，提高复杂模型兼容性。\r\n• 正常请求继续使用较低输出上限，并保持最多两次 API 请求的限制。\r\n\r\n" + ChangelogVersion219Text();
        }

        private string ChangelogVersion219Text()
        {
            return "OfficeTranslate 2.1.19\r\n• 精简指定源语言的翻译提示词，降低混合语言长文的模型推理负担。\r\n• 混合语言保护失败时最多进行一次占位符兜底，API 请求由最多三次降为最多两次。\r\n• 按目标语言自适应限制输出 Token，避免异常重复内容跑满输出上限。\r\n• 新增长度与重复内容校验，防止超长或循环译文写回文档。\r\n• 翻译温度调整为 0.1，提高术语与保护内容输出的稳定性。\r\n\r\n" + ChangelogVersion218Text();
        }

        private string ChangelogVersion218Text()
        {
            return "OfficeTranslate 2.1.18\r\n• 韩语、日语等混合语言改为整段翻译优先，避免拆成大量无上下文碎片。\r\n• 校验英文、缩写、型号等非源语言内容的原文与顺序，保护失败时自动整段重试。\r\n• 新增源语言残留比例校验，严重漏译或半译结果不再写回文档。\r\n• 减少数字类保护片段，解决长文翻译缓慢及只翻译中间少量内容的问题。\r\n\r\n" + ChangelogVersion217Text();
        }

        private string ChangelogVersion217Text()
        {
            return "OfficeTranslate 2.1.17\r\n• DeepSeek 等兼容服务不再发送非官方 prompt_cache_key 参数，继续使用服务自身的自动前缀缓存。\r\n• 为文本翻译设置合理的最大输出长度，避免模型异常长输出造成长时间等待。\r\n• 混合语言安全降级取消重复严格请求；远程 API 最多四路并发，本地 Ollama 保持单路。\r\n\r\n" + ChangelogVersion216Text();
        }

        private string ChangelogVersion216Text()
        {
            return "OfficeTranslate 2.1.16\r\n• 修复翻译结果窗口关闭后访问已释放对象的异常。\r\n• 结果窗口改由 Office UI 线程创建和管理，解决窗口未响应的问题。\r\n• 结果窗口不再阻塞翻译任务收尾，并继续支持 5 秒自动关闭。\r\n\r\n" + ChangelogVersion215Text();
        }

        private string ChangelogVersion215Text()
        {
            return "OfficeTranslate 2.1.15\r\n• 修复以英文缩写开头的混合句被强制逐片请求、导致单段翻译耗时过长的问题。\r\n• 正常情况下恢复为每段一次模型请求；仅在模型实际破坏保护内容时才进入安全降级。\r\n• 降级翻译连续返回空内容时仍会保留原文并标记需要检查，不会中止整个任务。\r\n\r\n" + ChangelogVersion214Text();
        }

        private string ChangelogVersion214Text()
        {
            return "OfficeTranslate 2.1.14\r\n• 增强韩英、日英混合句句首全大写英文缩写的保护恢复。\r\n• 分片模型连续返回空内容时保留原文并标记需要检查，不再中止整个翻译任务。\r\n\r\n" + ChangelogVersion213Text();
        }

        private string ChangelogVersion213Text()
        {
            return "OfficeTranslate 2.1.13\r\n• 翻译结束后关闭进度窗口，并显示独立、醒目的结果摘要窗口。\r\n• 结果窗口支持手动确认，并在 5 秒倒计时结束后自动关闭。\r\n• 韩语和日语翻译中的英文术语与相邻字母、数字自动补充边界空格，避免回填粘连。\r\n• 韩语和日语中的数字片段纳入原位保护，增强术语表英文结果的稳定性。\r\n\r\n" + ChangelogVersion212Text();
        }

        private string ChangelogVersion212Text()
        {
            return "OfficeTranslate 2.1.12\r\n• 新增翻译结果校验：空译文或未变化结果会自动严格重试，仍无有效结果时保留原文并提示检查。\r\n• 新增翻译任务摘要，显示成功翻译、会话缓存命中、跳过、需检查及 OCR 无文本数量。\r\n• 新增 Office 进程内的会话级翻译结果缓存，相同内容和配置可直接复用译文。\r\n• 稳定模型请求的提示词前缀，并为兼容服务提供前缀缓存键及自动降级。\r\n\r\n" + ChangelogVersion211Text();
        }

        private string ChangelogVersion211Text()
        {
            return "OfficeTranslate 2.1.11\r\n• 保护校验失败时自动切换为源语言分片翻译，并由程序按原位置恢复英文、缩写和型号。\r\n• 避免指令遵循能力较弱的模型丢失中间保护标记而导致整段翻译失败。\r\n• 分片降级翻译支持取消操作，并保留非源语言内容和原始空白。\r\n\r\n" + ChangelogVersion210Text();
        }

        private string ChangelogVersion210Text()
        {
            return "OfficeTranslate 2.1.10\r\n• 修复韩英混合文本中句首或句尾英文保护标记被模型省略时翻译失败的问题。\r\n• 增强残缺、全角及含零宽字符的保护标记恢复，防止 OT_KEEP 内部标记残留在译文中。\r\n• 整段均为非源语言时直接保留原文，不再发送翻译请求。\r\n\r\n" + ChangelogHistoryText();
        }

        private string ChangelogHistoryText()
        {
            return "OfficeTranslate 2.1.9\r\n• 明确选择源语言时，仅翻译该语言片段，其他语言内容保持原样。\r\n• 新增非源语言占位符保护与完整性校验，防止英文、型号、网址、邮箱和代码被误译。\r\n• 图片 OCR 同步采用选择性翻译规则；自动检测继续使用整段翻译。\r\n\r\n2.1.8\r\n• 图片译文框根据文字长度、宽度和双语行数自动增加高度，改善文字显示不全的问题。\r\n• 重绘语言菜单国旗图标，增强五星、紫荆花、星条、太极与国徽等辨识细节。\r\n\r\n2.1.7\r\n• 新增基于本地视觉模型的图片 OCR 翻译。\r\n• Word、Excel、PowerPoint 均支持图片选区与全文处理。\r\n• 译文以可编辑文本框覆盖原文字区域，并支持图片双语显示。\r\n• 新增独立图片模型设置和会话级图片开关。\r\n\r\n2.1.6\r\n• Excel 支持仅翻译组合流程图中选中的子图形。\r\n• 修复关闭翻译错误提示后 Excel 工作簿被最小化的问题。\r\n\r\n2.1.5\r\n• PowerPoint 表格支持仅翻译选中的单元格，选中整个表格时仍翻译全表。\r\n• Excel 支持翻译文本框、流程图及组合图形中的可编辑文字。\r\n\r\n2.1.4\r\n• 统一三个 Office 应用的翻译进度窗口，支持独立取消并保留已译内容。\r\n• 修复 Word 与 PowerPoint 表格翻译兼容问题。\r\n• 为不同服务类型分别保存连接配置。\r\n• 语言与双语模式改为会话设置，不再写入配置文件。\r\n\r\n2.1.3\r\n• 新增翻译风格预设与可编辑 Prompt。\r\n• 优化安装界面、升级与同版本重装行为。\r\n\r\n2.1.2\r\n• 统一 Word、Excel、PowerPoint 关于窗口的按钮风格。\r\n• 修复打开窗口时正文被自动选中的问题。\r\n\r\n2.1.1\r\n• 修复设置窗口取消按钮被裁切的问题。\r\n\r\n2.1.0\r\n• 新增关于、更新日志和使用说明。\r\n\r\n2.0.0\r\n• 新增 Excel 与 PowerPoint 原生翻译加载项。\r\n• 三个 Office 应用共享设置和界面语言。\r\n\r\n1.5.x\r\n• 新增五种界面语言、语言旗帜及逐段实时写回。";
        }

        private string GuideText()
        {
            if (_language == "en") return "WORD\r\nSelection translates selected text; Document translates the main document paragraph by paragraph.\r\n\r\nEXCEL\r\nSelection translates selected text cells; Document translates the active worksheet. Formulas and numbers are skipped.\r\n\r\nPOWERPOINT\r\nSelection translates selected text, shapes, tables, groups, or selected slides; Document translates the presentation.\r\n\r\nBILINGUAL\r\nEnable Bilingual before translating to keep the original and append the translation below it.";
            return "WORD\r\n“选区”翻译当前选中文字；“全文”逐段翻译文档正文。\r\n\r\nEXCEL\r\n“选区”翻译选中的文本单元格；“全文”翻译当前工作表。公式、数字和空白会自动跳过。\r\n\r\nPOWERPOINT\r\n“选区”可翻译选中文字、形状、表格、组合形状或左侧缩略图中选中的幻灯片；“全文”翻译整份演示文稿。\r\n\r\n双语模式\r\n翻译前启用“双语”，即可保留原文并在下方追加译文。翻译过程中可点击“取消”。";
        }
    }
}
