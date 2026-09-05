using System.Drawing;
using System.Windows.Forms;

namespace OfficeTranslate.WordAddIn
{
    internal sealed class AboutForm : Form
    {
        public const string CurrentVersion = "2.1.8";
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
            return "OfficeTranslate 2.1.8\r\n• 图片译文框根据文字长度、宽度和双语行数自动增加高度，改善文字显示不全的问题。\r\n• 重绘语言菜单国旗图标，增强五星、紫荆花、星条、太极与国徽等辨识细节。\r\n\r\n2.1.7\r\n• 新增基于本地视觉模型的图片 OCR 翻译。\r\n• Word、Excel、PowerPoint 均支持图片选区与全文处理。\r\n• 译文以可编辑文本框覆盖原文字区域，并支持图片双语显示。\r\n• 新增独立图片模型设置和会话级图片开关。\r\n\r\n2.1.6\r\n• Excel 支持仅翻译组合流程图中选中的子图形。\r\n• 修复关闭翻译错误提示后 Excel 工作簿被最小化的问题。\r\n\r\n2.1.5\r\n• PowerPoint 表格支持仅翻译选中的单元格，选中整个表格时仍翻译全表。\r\n• Excel 支持翻译文本框、流程图及组合图形中的可编辑文字。\r\n\r\n2.1.4\r\n• 统一三个 Office 应用的翻译进度窗口，支持独立取消并保留已译内容。\r\n• 修复 Word 与 PowerPoint 表格翻译兼容问题。\r\n• 为不同服务类型分别保存连接配置。\r\n• 语言与双语模式改为会话设置，不再写入配置文件。\r\n\r\n2.1.3\r\n• 新增翻译风格预设与可编辑 Prompt。\r\n• 优化安装界面、升级与同版本重装行为。\r\n\r\n2.1.2\r\n• 统一 Word、Excel、PowerPoint 关于窗口的按钮风格。\r\n• 修复打开窗口时正文被自动选中的问题。\r\n\r\n2.1.1\r\n• 修复设置窗口取消按钮被裁切的问题。\r\n\r\n2.1.0\r\n• 新增关于、更新日志和使用说明。\r\n\r\n2.0.0\r\n• 新增 Excel 与 PowerPoint 原生翻译加载项。\r\n• 三个 Office 应用共享设置和界面语言。\r\n\r\n1.5.x\r\n• 新增五种界面语言、语言旗帜及逐段实时写回。";
        }

        private string GuideText()
        {
            if (_language == "en") return "WORD\r\nSelection translates selected text; Document translates the main document paragraph by paragraph.\r\n\r\nEXCEL\r\nSelection translates selected text cells; Document translates the active worksheet. Formulas and numbers are skipped.\r\n\r\nPOWERPOINT\r\nSelection translates selected text, shapes, tables, groups, or selected slides; Document translates the presentation.\r\n\r\nBILINGUAL\r\nEnable Bilingual before translating to keep the original and append the translation below it.";
            return "WORD\r\n“选区”翻译当前选中文字；“全文”逐段翻译文档正文。\r\n\r\nEXCEL\r\n“选区”翻译选中的文本单元格；“全文”翻译当前工作表。公式、数字和空白会自动跳过。\r\n\r\nPOWERPOINT\r\n“选区”可翻译选中文字、形状、表格、组合形状或左侧缩略图中选中的幻灯片；“全文”翻译整份演示文稿。\r\n\r\n双语模式\r\n翻译前启用“双语”，即可保留原文并在下方追加译文。翻译过程中可点击“取消”。";
        }
    }
}
