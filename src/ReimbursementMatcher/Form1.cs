using System.ComponentModel;
using System.Diagnostics;

namespace ReimbursementMatcher;

public partial class Form1 : Form
{
    private readonly WorkspaceService _workspace = new();
    private readonly MaterialScanner _scanner;
    private readonly MatchingService _matcher = new();
    private readonly ConfirmationService _confirmations;
    private readonly ReportService _reporter;
    private readonly InvoiceFileCleaner _invoiceCleaner;
    private readonly EmailAuditReportService _emailAuditReporter;
    private readonly EmailAuditService _emailAuditService;

    private readonly BindingList<EvidenceItem> _items = new();
    private readonly BindingList<EvidenceItem> _invoiceItems = new();
    private readonly BindingList<MatchCandidate> _matches = new();
    private readonly BindingList<DownloadRecord> _downloadRecords = new();
    private readonly BindingList<EmailAuditItem> _emailAuditItems = new();
    private readonly List<EvidenceItem> _allItems = new();
    private readonly List<EvidenceItem> _allInvoiceItems = new();
    private readonly List<MatchCandidate> _allMatches = new();
    private readonly List<DownloadRecord> _allDownloadRecords = new();
    private readonly List<EmailAuditItem> _allEmailAuditItems = new();

    private readonly DataGridView _itemGrid = Grid();
    private readonly DataGridView _invoiceGrid = Grid();
    private readonly DataGridView _matchGrid = Grid();
    private readonly DataGridView _downloadGrid = Grid();
    private readonly DataGridView _emailAuditGrid = Grid();
    private readonly TextBox _detailBox = new() { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
    private readonly TextBox _invoiceDetailBox = new() { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
    private readonly TextBox _emailAuditDetailBox = new() { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
    private readonly TextBox _logBox = new() { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, BackColor = Color.FromArgb(16, 24, 40), ForeColor = Color.White, Font = new Font("Consolas", 9) };

    private readonly TextBox _sourceRoot = new();
    private readonly TextBox _invoiceDir = new();
    private readonly TextBox _templatePath = new();
    private readonly TextBox _outputDir = new();
    private readonly TextBox _ruleDir = new();
    private readonly TextBox _operator = new();
    private readonly TextBox _dateStart = new();
    private readonly TextBox _dateEnd = new();

    private readonly CheckBox _emailEnabled = new() { Text = "启用邮箱发票下载", AutoSize = true, Checked = true };
    private readonly TextBox _emailHost = new();
    private readonly TextBox _emailPort = new();
    private readonly TextBox _emailFolder = new();
    private readonly TextBox _emailUser = new();
    private readonly TextBox _emailPassword = new() { UseSystemPasswordChar = true };
    private readonly TextBox _emailOutput = new();
    private readonly TextBox _emailStart = new();
    private readonly TextBox _emailEnd = new();
    private readonly TextBox _emailInterval = new();

    private readonly TextBox _noteBox = new() { Dock = DockStyle.Fill };
    private readonly TextBox _invoiceNoteBox = new() { Dock = DockStyle.Fill };
    private readonly TextBox _emailDecisionNoteBox = new() { Dock = DockStyle.Fill };
    private readonly Label _summary = new() { AutoSize = true, ForeColor = Color.FromArgb(52, 64, 84) };
    private readonly CheckBox _hideDoneItems = new() { Text = "只看待处理/异常", Checked = true, AutoSize = true };
    private readonly CheckBox _hideDoneInvoices = new() { Text = "只看待处理/异常", Checked = true, AutoSize = true };
    private readonly CheckBox _hideDoneMatches = new() { Text = "只看待处理", Checked = true, AutoSize = true };
    private readonly CheckBox _hideDoneDownloads = new() { Text = "只看待处理/异常", Checked = true, AutoSize = true };
    private readonly CheckBox _hideDoneEmailAudit = new() { Text = "只看需要人工确认", Checked = true, AutoSize = true };

    private AppConfig _config = new();
    private ConfirmationStore _store = new();
    private EmailDecisionStore _emailDecisionStore = new();
    private DataGridView? _lastEvidenceGrid;

    public Form1()
    {
        InitializeComponent();
        _scanner = new MaterialScanner(_workspace);
        _confirmations = new ConfirmationService(_workspace);
        _reporter = new ReportService(_workspace);
        _invoiceCleaner = new InvoiceFileCleaner(_workspace);
        _emailAuditReporter = new EmailAuditReportService(_workspace);
        _emailAuditService = new EmailAuditService(_workspace);
        BuildLayout();
        LoadConfig();
    }

    private void BuildLayout()
    {
        Text = "自动化报销匹配工作台";
        MinimumSize = new Size(980, 700);
        Size = new Size(1400, 880);
        Font = new Font("Microsoft YaHei UI", 10);
        BackColor = Color.FromArgb(246, 247, 249);

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, Padding = new Padding(14), BackColor = BackColor };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 150));
        Controls.Add(root);

        root.Controls.Add(BuildTopPanel(), 0, 0);
        root.Controls.Add(BuildMainTabs(), 0, 1);
        root.Controls.Add(_logBox, 0, 2);
    }

    private Control BuildTopPanel()
    {
        var panel = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 4, RowCount = 4, AutoSize = true, BackColor = BackColor };
        for (var i = 0; i < 4; i++) panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));

        AddLabeled(panel, "素材根目录", _sourceRoot, 0, 0, 2);
        AddLabeled(panel, "报销模板", _templatePath, 2, 0, 2);
        AddLabeled(panel, "输出目录", _outputDir, 0, 1, 1);
        AddLabeled(panel, "规则目录", _ruleDir, 1, 1, 1);
        AddLabeled(panel, "经办人", _operator, 2, 1, 1);
        panel.Controls.Add(Button("保存配置", (_, _) => SaveConfig()), 3, 1);

        AddLabeled(panel, "开始日期", _dateStart, 0, 2, 1);
        AddLabeled(panel, "结束日期", _dateEnd, 1, 2, 1);
        AddLabeled(panel, "发票目录", _invoiceDir, 2, 2, 2);

        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
        buttons.Controls.Add(Button("扫描原始素材", (_, _) => ScanMaterials(), primary: true));
        buttons.Controls.Add(Button("保存确认记录", (_, _) => SaveConfirmations()));
        buttons.Controls.Add(Button("生成报销Excel", (_, _) => GenerateReport()));
        buttons.Controls.Add(Button("打开选中文件", (_, _) => OpenSelectedFile()));
        buttons.Controls.Add(Button("打开规则目录", (_, _) => OpenDir(_ruleDir.Text)));
        buttons.Controls.Add(Button("打开输出目录", (_, _) => OpenDir(_outputDir.Text)));
        buttons.Controls.Add(Button("打开发票目录", (_, _) => OpenDir(_invoiceDir.Text)));
        buttons.Controls.Add(_summary);
        panel.Controls.Add(buttons, 0, 3);
        panel.SetColumnSpan(buttons, 4);
        return panel;
    }

    private Control BuildMainTabs()
    {
        ConfigureItemGrid();
        ConfigureInvoiceGrid();
        ConfigureMatchGrid();
        ConfigureDownloadGrid();
        ConfigureEmailAuditGrid();

        var tabs = new TabControl { Dock = DockStyle.Fill };
        tabs.TabPages.Add(BuildItemsPage());
        tabs.TabPages.Add(BuildInvoicePage());
        tabs.TabPages.Add(BuildEmailAuditPage());
        tabs.TabPages.Add(BuildDownloadAuditPage());
        tabs.TabPages.Add(BuildMatchPage());
        return tabs;
    }

    private TabPage BuildItemsPage()
    {
        var page = new TabPage("逐项确认") { BackColor = Color.White };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 72));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 28));
        page.Controls.Add(layout);
        var left = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
        left.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        left.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        _hideDoneItems.CheckedChanged += (_, _) => RefreshEvidenceViews();
        left.Controls.Add(_hideDoneItems, 0, 0);
        left.Controls.Add(_itemGrid, 0, 1);
        layout.Controls.Add(left, 0, 0);
        layout.Controls.Add(BuildDecisionPanel(includeInvoiceButtons: true), 1, 0);
        return page;
    }

    private TabPage BuildInvoicePage()
    {
        var page = new TabPage("发票下载核验") { BackColor = Color.White, AutoScroll = true };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1 };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        page.Controls.Add(layout);

        var hint = new Label
        {
            Dock = DockStyle.Top,
            Height = 34,
            Text = "在这里配置邮箱并下载发票。下载后的 PDF/OFD/XML/图片会进入下方列表，直接在软件内确认有效、无效、重复或日期不符。",
            TextAlign = ContentAlignment.MiddleLeft,
            ForeColor = Color.FromArgb(52, 64, 84)
        };
        layout.Controls.Add(hint, 0, 0);

        layout.Controls.Add(BuildEmailPanel(), 0, 1);

        var middle = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
        middle.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 72));
        middle.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 28));
        var left = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
        left.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        left.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        _hideDoneInvoices.CheckedChanged += (_, _) => RefreshEvidenceViews();
        left.Controls.Add(_hideDoneInvoices, 0, 0);
        left.Controls.Add(_invoiceGrid, 0, 1);
        middle.Controls.Add(left, 0, 0);
        middle.Controls.Add(BuildInvoiceSidePanel(), 1, 0);
        layout.Controls.Add(middle, 0, 2);
        return page;
    }

    private TabPage BuildMatchPage()
    {
        var page = new TabPage("发票/付款匹配") { BackColor = Color.White };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1 };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        _hideDoneMatches.CheckedChanged += (_, _) => RefreshMatchView();
        layout.Controls.Add(_hideDoneMatches, 0, 0);
        layout.Controls.Add(_matchGrid, 0, 1);
        layout.Controls.Add(BuildMatchButtons(), 0, 2);
        page.Controls.Add(layout);
        return page;
    }

    private TabPage BuildDownloadAuditPage()
    {
        var page = new TabPage("下载完整性核验") { BackColor = Color.White };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        page.Controls.Add(layout);
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
        _hideDoneDownloads.CheckedChanged += (_, _) => RefreshDownloadView();
        buttons.Controls.Add(_hideDoneDownloads);
        buttons.Controls.Add(Button("读取最新下载清单", (_, _) => LoadDownloadRecords()));
        buttons.Controls.Add(Button("只重处理异常邮件", async (s, _) => await DownloadEmailInvoicesAsync(s as Button, onlyAbnormal: true)));
        buttons.Controls.Add(Button("打开选中链接/文件", (_, _) => OpenSelectedDownloadTarget()));
        buttons.Controls.Add(Button("生成邮件统计Excel", (_, _) => GenerateEmailAuditReport()));
        buttons.Controls.Add(Button("打开清单目录", (_, _) => OpenDir(_config.Email.OutputDir)));
        buttons.Controls.Add(Button("归档非PDF重复格式", (_, _) => ArchiveNonPdfInvoiceFormats()));
        buttons.Controls.Add(new Label
        {
            Text = "重点看状态为“未下载到文件”“失败”“需人工确认”的行；这些代表邮件里可能有发票，但程序没能直接下载成文件。",
            AutoSize = true,
            ForeColor = Color.FromArgb(102, 112, 133),
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(8)
        });
        layout.Controls.Add(buttons, 0, 0);
        layout.Controls.Add(_downloadGrid, 0, 1);
        return page;
    }

    private TabPage BuildEmailAuditPage()
    {
        var page = new TabPage("邮件发票判断") { BackColor = Color.White };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 74));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 26));
        page.Controls.Add(layout);

        var left = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
        left.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        left.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
        _hideDoneEmailAudit.CheckedChanged += (_, _) => RefreshEmailAuditView();
        buttons.Controls.Add(_hideDoneEmailAudit);
        buttons.Controls.Add(Button("读取邮件判断清单", (_, _) => LoadEmailAuditItems(), primary: true));
        buttons.Controls.Add(Button("生成邮件判断Excel", (_, _) => GenerateEmailChecklistExcel()));
        buttons.Controls.Add(Button("保存下载规则", (_, _) => SaveEmailDownloadRules()));
        buttons.Controls.Add(Button("选中标有发票", (_, _) => ConfirmSelectedEmails(EmailAuditService.HasInvoice)));
        buttons.Controls.Add(Button("选中标无发票", (_, _) => ConfirmSelectedEmails(EmailAuditService.NoInvoice)));
        buttons.Controls.Add(Button("当前显示全部无发票", (_, _) => ConfirmVisibleEmailsNoInvoice()));
        buttons.Controls.Add(Button("打开规则目录", (_, _) => OpenDir(_config.RuleDir)));
        left.Controls.Add(buttons, 0, 0);
        left.Controls.Add(_emailAuditGrid, 0, 1);
        layout.Controls.Add(left, 0, 0);

        var side = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 8, ColumnCount = 1, Padding = new Padding(10) };
        side.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        for (var i = 1; i < side.RowCount; i++) side.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        side.Controls.Add(_emailAuditDetailBox, 0, 0);
        side.Controls.Add(new Label { Text = "人工备注", AutoSize = true }, 0, 1);
        side.Controls.Add(_emailDecisionNoteBox, 0, 2);
        side.Controls.Add(Button("明确有发票", (_, _) => ConfirmSelectedEmails(EmailAuditService.HasInvoice)), 0, 3);
        side.Controls.Add(Button("明确无发票", (_, _) => ConfirmSelectedEmails(EmailAuditService.NoInvoice)), 0, 4);
        side.Controls.Add(Button("需要人工确认", (_, _) => ConfirmSelectedEmails(EmailAuditService.NeedsReview)), 0, 5);
        side.Controls.Add(Button("打开下载清单目录", (_, _) => OpenDir(_config.Email.OutputDir)), 0, 6);
        side.Controls.Add(new Label
        {
            Text = "明确有/无发票的邮件后续会复用判断；只有不确定的邮件继续拉出来人工核对。",
            AutoSize = true,
            ForeColor = Color.FromArgb(102, 112, 133)
        }, 0, 7);
        layout.Controls.Add(side, 1, 0);
        return page;
    }

    private Control BuildDecisionPanel(bool includeInvoiceButtons)
    {
        var panel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = includeInvoiceButtons ? 9 : 7, ColumnCount = 1, Padding = new Padding(10) };
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        for (var i = 1; i < panel.RowCount; i++) panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.Controls.Add(_detailBox, 0, 0);
        panel.Controls.Add(new Label { Text = "备注/确认说明", AutoSize = true }, 0, 1);
        panel.Controls.Add(_noteBox, 0, 2);
        panel.Controls.Add(Button("标记报销", (_, _) => ConfirmSelectedItem("报销", null)), 0, 3);
        panel.Controls.Add(Button("标记不报销", (_, _) => ConfirmSelectedItem("不报销", null)), 0, 4);
        panel.Controls.Add(Button("标记待确认", (_, _) => ConfirmSelectedItem("待确认", null)), 0, 5);
        if (includeInvoiceButtons)
        {
            panel.Controls.Add(Button("发票/文件有效", (_, _) => ConfirmSelectedItem(null, "有效")), 0, 6);
            panel.Controls.Add(Button("发票/文件无效", (_, _) => ConfirmSelectedItem(null, "无效")), 0, 7);
            panel.Controls.Add(Button("打开文件", (_, _) => OpenSelectedFile()), 0, 8);
        }
        else
        {
            panel.Controls.Add(Button("打开文件", (_, _) => OpenSelectedFile()), 0, 6);
        }
        return panel;
    }

    private Control BuildInvoiceSidePanel()
    {
        var panel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 8, ColumnCount = 1, Padding = new Padding(10) };
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        for (var i = 1; i < panel.RowCount; i++) panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.Controls.Add(_invoiceDetailBox, 0, 0);
        panel.Controls.Add(Button("发票有效", (_, _) => ConfirmSelectedItem(null, "有效")), 0, 1);
        panel.Controls.Add(Button("发票无效", (_, _) => ConfirmSelectedItem(null, "无效")), 0, 2);
        panel.Controls.Add(Button("待确认", (_, _) => ConfirmSelectedItem(null, "待确认")), 0, 3);
        panel.Controls.Add(Button("打开选中发票", (_, _) => OpenSelectedFile()), 0, 4);
        panel.Controls.Add(Button("打开发票目录", (_, _) => OpenDir(_invoiceDir.Text)), 0, 5);
        panel.Controls.Add(new Label { Text = "备注/确认说明", AutoSize = true }, 0, 6);
        panel.Controls.Add(_invoiceNoteBox, 0, 7);
        return panel;
    }

    private Control BuildEmailPanel()
    {
        var group = new GroupBox { Text = "邮箱发票下载", Dock = DockStyle.Top, Padding = new Padding(8), AutoSize = true };
        var panel = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 4, RowCount = 4, AutoSize = true };
        for (var i = 0; i < 4; i++) panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
        group.Controls.Add(panel);

        panel.Controls.Add(_emailEnabled, 0, 0);
        var passwordHint = new Label
        {
            Text = "QQ邮箱通常需要使用 IMAP/SMTP 授权码；点击保存配置后会保存到本机配置文件。",
            AutoSize = true,
            ForeColor = Color.FromArgb(102, 112, 133),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft
        };
        panel.Controls.Add(passwordHint, 1, 0);
        panel.SetColumnSpan(passwordHint, 2);
        panel.Controls.Add(Button("下载邮箱发票", async (s, _) => await DownloadEmailInvoicesAsync(s as Button), primary: true), 3, 0);

        AddLabeled(panel, "IMAP服务器", _emailHost, 0, 1, 1);
        AddLabeled(panel, "端口", _emailPort, 1, 1, 1);
        AddLabeled(panel, "邮箱文件夹", _emailFolder, 2, 1, 1);
        AddLabeled(panel, "邮箱账号", _emailUser, 3, 1, 1);

        AddLabeled(panel, "授权码/密码", _emailPassword, 0, 2, 1);
        AddLabeled(panel, "下载开始日期", _emailStart, 1, 2, 1);
        AddLabeled(panel, "下载结束日期", _emailEnd, 2, 2, 1);
        AddLabeled(panel, "请求间隔秒", _emailInterval, 3, 2, 1);

        AddLabeled(panel, "保存到目录", _emailOutput, 0, 3, 4);
        return group;
    }

    private Control BuildMatchButtons()
    {
        var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        panel.Controls.Add(Button("确认匹配", (_, _) => ConfirmSelectedMatch("确认匹配")));
        panel.Controls.Add(Button("否决匹配", (_, _) => ConfirmSelectedMatch("否决匹配")));
        panel.Controls.Add(Button("待确认", (_, _) => ConfirmSelectedMatch("待确认")));
        return panel;
    }

    private void ConfigureItemGrid()
    {
        _itemGrid.DataSource = _items;
        AddEvidenceColumns(_itemGrid, includeReimburse: true);
        _itemGrid.SelectionChanged += (_, _) =>
        {
            _lastEvidenceGrid = _itemGrid;
            UpdateDetail(SelectedItem());
        };
        _itemGrid.DataBindingComplete += (_, _) => ColorRows(_itemGrid);
    }

    private void ConfigureInvoiceGrid()
    {
        _invoiceGrid.DataSource = _invoiceItems;
        AddEvidenceColumns(_invoiceGrid, includeReimburse: false);
        _invoiceGrid.SelectionChanged += (_, _) =>
        {
            _lastEvidenceGrid = _invoiceGrid;
            UpdateDetail(SelectedItem());
        };
        _invoiceGrid.DataBindingComplete += (_, _) => ColorRows(_invoiceGrid);
    }

    private void ConfigureMatchGrid()
    {
        _matchGrid.DataSource = _matches;
        AddColumn(_matchGrid, "ExpenseTitle", "消费/订单", 320);
        AddColumn(_matchGrid, "InvoiceTitle", "发票", 320);
        AddColumn(_matchGrid, "ExpenseAmount", "消费金额", 90);
        AddColumn(_matchGrid, "InvoiceAmount", "发票金额", 90);
        AddColumn(_matchGrid, "ExpenseDate", "消费日期", 100);
        AddColumn(_matchGrid, "InvoiceDate", "发票日期", 100);
        AddColumn(_matchGrid, "Score", "评分", 70);
        AddColumn(_matchGrid, "Reason", "理由", 260);
        AddColumn(_matchGrid, "Decision", "确认", 100);
        _matchGrid.DataBindingComplete += (_, _) => ColorRows(_matchGrid);
    }

    private void ConfigureDownloadGrid()
    {
        _downloadGrid.DataSource = _downloadRecords;
        AddColumn(_downloadGrid, "Status", "状态", 100);
        AddColumn(_downloadGrid, "DownloadStatus", "下载状态", 100);
        AddColumn(_downloadGrid, "Kind", "类型", 90);
        AddColumn(_downloadGrid, "Date", "日期", 100);
        AddColumn(_downloadGrid, "Subject", "邮件主题", 360);
        AddColumn(_downloadGrid, "Sender", "发件人", 180);
        AddColumn(_downloadGrid, "SavedFileCount", "下载数", 70);
        AddColumn(_downloadGrid, "AttachmentTotal", "附件数", 70);
        AddColumn(_downloadGrid, "LinkCandidateTotal", "链接数", 70);
        AddColumn(_downloadGrid, "InvoiceNumber", "发票号", 150);
        AddColumn(_downloadGrid, "InvoiceAmount", "发票金额", 90);
        AddColumn(_downloadGrid, "InvoiceDate", "开票日期", 100);
        AddColumn(_downloadGrid, "File", "文件", 420);
        AddColumn(_downloadGrid, "Url", "链接", 360);
        AddColumn(_downloadGrid, "Error", "错误/说明", 260);
        AddColumn(_downloadGrid, "DuplicateOf", "重复于", 360);
        _downloadGrid.DataBindingComplete += (_, _) => ColorRows(_downloadGrid);
    }

    private void ConfigureEmailAuditGrid()
    {
        _emailAuditGrid.DataSource = _emailAuditItems;
        _emailAuditGrid.MultiSelect = true;
        AddColumn(_emailAuditGrid, "FinalDecision", "最终判断", 120);
        AddColumn(_emailAuditGrid, "RuleDecision", "规则判断", 120);
        AddColumn(_emailAuditGrid, "ManualDecision", "人工判断", 120);
        AddColumn(_emailAuditGrid, "Date", "日期", 100);
        AddColumn(_emailAuditGrid, "Subject", "邮件主题", 420);
        AddColumn(_emailAuditGrid, "Status", "下载状态", 120);
        AddColumn(_emailAuditGrid, "HasInvoiceKeyword", "发票关键字", 90);
        AddColumn(_emailAuditGrid, "HasPlatformKeyword", "平台关键字", 90);
        AddColumn(_emailAuditGrid, "HasAttachment", "有附件", 70);
        AddColumn(_emailAuditGrid, "AttachmentCount", "附件数", 70);
        AddColumn(_emailAuditGrid, "HasLikelyInvoiceAttachment", "疑似附件", 90);
        AddColumn(_emailAuditGrid, "HasLikelyInvoiceLink", "疑似链接", 90);
        AddColumn(_emailAuditGrid, "DownloadedFileCount", "下载数", 70);
        AddColumn(_emailAuditGrid, "SkippedOrDuplicateCount", "跳过数", 70);
        AddColumn(_emailAuditGrid, "Reason", "判断原因", 280);
        AddColumn(_emailAuditGrid, "Note", "人工备注", 180);
        _emailAuditGrid.SelectionChanged += (_, _) => UpdateEmailAuditDetail(SelectedEmailAuditItem());
        _emailAuditGrid.DataBindingComplete += (_, _) => ColorRows(_emailAuditGrid);
    }

    private static void AddEvidenceColumns(DataGridView grid, bool includeReimburse)
    {
        AddColumn(grid, "Kind", "类型", 90);
        AddColumn(grid, "Platform", "来源", 80);
        AddColumn(grid, "Date", "日期", 100);
        AddColumn(grid, "Amount", "金额", 80);
        AddColumn(grid, "Title", "标题/文件", 300);
        AddColumn(grid, "InvoiceNumber", "发票号", 150);
        if (includeReimburse) AddColumn(grid, "ReimburseDecision", "报销确认", 110);
        AddColumn(grid, "FileDecision", "文件确认", 100);
        AddColumn(grid, "MatchStatus", "匹配", 90);
        AddColumn(grid, "Project", "项目", 220);
        AddColumn(grid, "Weight", "权重", 60);
        AddColumn(grid, "Suggestion", "建议", 100);
        AddColumn(grid, "DuplicateInfo", "重复/合并", 160);
        AddColumn(grid, "RelativePath", "路径", 360);
    }

    private void LoadConfig()
    {
        _config = _workspace.LoadConfig();
        _sourceRoot.Text = _config.SourceRoot;
        _invoiceDir.Text = _config.InvoiceDir;
        _templatePath.Text = _config.TemplatePath;
        _outputDir.Text = _config.OutputDir;
        _ruleDir.Text = _config.RuleDir;
        _operator.Text = _config.Operator;
        _dateStart.Text = _config.DateStart;
        _dateEnd.Text = _config.DateEnd;

        _emailEnabled.Checked = _config.Email.Enabled;
        _emailHost.Text = _config.Email.Host;
        _emailPort.Text = _config.Email.Port.ToString();
        _emailFolder.Text = _config.Email.Folder;
        _emailUser.Text = _config.Email.User;
        _emailPassword.Text = _config.Email.Password;
        _emailOutput.Text = _config.Email.OutputDir;
        _emailStart.Text = _config.Email.Start;
        _emailEnd.Text = _config.Email.End;
        _emailInterval.Text = _config.Email.RequestIntervalSec.ToString("0.###");

        _store = _confirmations.Load(_config);
        _emailDecisionStore = _emailAuditService.LoadDecisions(_config);
        Log("配置已读取。");
    }

    private void SaveConfig()
    {
        _config.SourceRoot = _sourceRoot.Text.Trim();
        _config.InvoiceDir = _invoiceDir.Text.Trim();
        _config.TemplatePath = _templatePath.Text.Trim();
        _config.OutputDir = _outputDir.Text.Trim();
        _config.RuleDir = _ruleDir.Text.Trim();
        _config.Operator = _operator.Text.Trim();
        _config.DateStart = _dateStart.Text.Trim();
        _config.DateEnd = _dateEnd.Text.Trim();

        _config.Email.Enabled = _emailEnabled.Checked;
        _config.Email.Host = _emailHost.Text.Trim();
        _config.Email.Port = int.TryParse(_emailPort.Text.Trim(), out var port) ? port : 993;
        _config.Email.Folder = string.IsNullOrWhiteSpace(_emailFolder.Text) ? "inbox" : _emailFolder.Text.Trim();
        _config.Email.User = _emailUser.Text.Trim();
        _config.Email.Password = _emailPassword.Text;
        _config.Email.OutputDir = _emailOutput.Text.Trim();
        _config.Email.Start = string.IsNullOrWhiteSpace(_emailStart.Text) ? _config.DateStart : _emailStart.Text.Trim();
        _config.Email.End = string.IsNullOrWhiteSpace(_emailEnd.Text) ? _config.DateEnd : _emailEnd.Text.Trim();
        _config.Email.RequestIntervalSec = double.TryParse(_emailInterval.Text.Trim(), out var interval) ? interval : 0.1;

        _workspace.SaveConfig(_config);
        Log("配置已保存。");
    }

    private void ScanMaterials()
    {
        SaveConfig();
        _store = _confirmations.Load(_config);
        var scanned = _scanner.Scan(_config).Where(InDateRangeOrUndated).ToList();
        var candidates = _matcher.BuildCandidates(scanned);
        _confirmations.Apply(_store, scanned, candidates);

        _allItems.Clear();
        _allItems.AddRange(scanned);
        _allInvoiceItems.Clear();
        _allInvoiceItems.AddRange(scanned.Where(i => i.Kind == EvidenceKinds.Invoice));
        _allMatches.Clear();
        _allMatches.AddRange(candidates);

        RefreshEvidenceViews();
        RefreshMatchView();

        LoadDownloadRecords();
        LoadEmailAuditItems(showMessageWhenMissing: false);
        _lastEvidenceGrid = _itemGrid;
        UpdateSummary();
        Log($"扫描完成：素材 {_items.Count} 个，发票 {_invoiceItems.Count} 个，匹配候选 {_matches.Count} 个。");
    }

    private async Task DownloadEmailInvoicesAsync(Button? button, bool onlyAbnormal = false)
    {
        SaveConfig();
        if (string.IsNullOrWhiteSpace(_emailPassword.Text))
        {
            MessageBox.Show("请输入邮箱授权码/密码。QQ邮箱通常需要 IMAP/SMTP 授权码。", "缺少授权码", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        _invoiceDir.Text = _config.Email.OutputDir;
        _config.InvoiceDir = _config.Email.OutputDir;
        _workspace.SaveConfig(_config);

        if (button != null) button.Enabled = false;
        try
        {
            var downloader = new EmailDownloader(_workspace, Log);
            var manifest = await downloader.DownloadAsync(_config, _emailPassword.Text, onlyAbnormal);
            Log($"邮箱发票下载清单：{manifest}");
            ScanMaterials();
            LoadEmailAuditItems(showMessageWhenMissing: false);
            MessageBox.Show("邮箱发票下载完成，已刷新发票下载核验列表。", "下载完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            Log("邮箱发票下载失败：" + ex.Message);
            MessageBox.Show(ex.Message, "邮箱发票下载失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (button != null) button.Enabled = true;
        }
    }

    private bool InDateRangeOrUndated(EvidenceItem item)
    {
        if (string.IsNullOrWhiteSpace(item.Date)) return true;
        if (!DateTime.TryParse(item.Date, out var date)) return true;
        if (DateTime.TryParse(_config.DateStart, out var start) && date < start) return false;
        if (DateTime.TryParse(_config.DateEnd, out var end) && date > end) return false;
        return true;
    }

    private void ConfirmSelectedItem(string? reimburseDecision, string? fileDecision)
    {
        if (SelectedItem() is not { } item) return;
        var note = item.Kind == EvidenceKinds.Invoice ? _invoiceNoteBox.Text.Trim() : _noteBox.Text.Trim();
        _confirmations.ConfirmItem(_config, _store, item, reimburseDecision, fileDecision, note);
        RefreshEvidenceViews();
        UpdateDetail(item);
        UpdateSummary();
        Log($"已确认素材：{item.Title} -> {item.ReimburseDecision}/{item.FileDecision}");
    }

    private void ConfirmSelectedMatch(string decision)
    {
        if (SelectedMatch() is not { } match) return;
        _confirmations.ConfirmMatch(_config, _store, match, decision, _noteBox.Text.Trim());
        if (decision == "确认匹配")
        {
            foreach (var item in _allItems.Where(i => i.Id == match.ExpenseId || i.Id == match.InvoiceId))
            {
                item.MatchStatus = "已匹配";
            }
        }
        RefreshMatchView();
        RefreshEvidenceViews();
        UpdateSummary();
        Log($"已确认匹配：{match.ExpenseTitle} <-> {match.InvoiceTitle}，{decision}");
    }

    private void SaveConfirmations()
    {
        _confirmations.Save(_config, _store);
        Log("确认记录和判断权重已保存。");
    }

    private void GenerateReport()
    {
        SaveConfirmations();
        var output = _reporter.Generate(_config, _allItems.ToList(), _allMatches.ToList());
        Log($"已生成报销Excel：{output}");
        Process.Start(new ProcessStartInfo(output) { UseShellExecute = true });
    }

    private void OpenSelectedFile()
    {
        if (SelectedItem() is not { } item || !File.Exists(item.FilePath))
        {
            return;
        }
        Process.Start(new ProcessStartInfo(item.FilePath) { UseShellExecute = true });
    }

    private void OpenDir(string path)
    {
        var resolved = _workspace.Resolve(path);
        Directory.CreateDirectory(resolved);
        Process.Start(new ProcessStartInfo(resolved) { UseShellExecute = true });
    }

    private EvidenceItem? SelectedItem()
    {
        if (_lastEvidenceGrid?.CurrentRow?.DataBoundItem is EvidenceItem activeItem) return activeItem;
        if (_itemGrid.CurrentRow?.DataBoundItem is EvidenceItem item) return item;
        return _invoiceGrid.CurrentRow?.DataBoundItem as EvidenceItem;
    }

    private MatchCandidate? SelectedMatch() => _matchGrid.CurrentRow?.DataBoundItem as MatchCandidate;

    private DownloadRecord? SelectedDownloadRecord() => _downloadGrid.CurrentRow?.DataBoundItem as DownloadRecord;

    private void OpenSelectedDownloadTarget()
    {
        if (SelectedDownloadRecord() is not { } record)
        {
            return;
        }

        if (!string.IsNullOrWhiteSpace(record.File) && File.Exists(record.File))
        {
            Process.Start(new ProcessStartInfo(record.File) { UseShellExecute = true });
            return;
        }

        if (!string.IsNullOrWhiteSpace(record.Url))
        {
            Process.Start(new ProcessStartInfo(record.Url) { UseShellExecute = true });
        }
    }

    private void UpdateDetail(EvidenceItem? item)
    {
        if (item == null)
        {
            _detailBox.Clear();
            _invoiceDetailBox.Clear();
            return;
        }

        var text =
            $"类型：{item.Kind}{Environment.NewLine}" +
            $"来源：{item.Platform}{Environment.NewLine}" +
            $"日期：{item.Date}{Environment.NewLine}" +
            $"金额：{item.Amount}{Environment.NewLine}" +
            $"标题：{item.Title}{Environment.NewLine}" +
            $"报销确认：{item.ReimburseDecision}{Environment.NewLine}" +
            $"文件确认：{item.FileDecision}{Environment.NewLine}" +
            $"项目：{item.Project}{Environment.NewLine}" +
            $"权重：{item.Weight} {item.Suggestion}{Environment.NewLine}" +
            $"重复/合并：{item.DuplicateInfo}{Environment.NewLine}" +
            $"路径：{item.RelativePath}{Environment.NewLine}" +
            $"备注：{item.Note}";
        _detailBox.Text = text;
        _invoiceDetailBox.Text = text;
        _noteBox.Text = item.Note;
        _invoiceNoteBox.Text = item.Note;
    }

    private void UpdateSummary()
    {
        var invoices = _allItems.Count(i => i.Kind == EvidenceKinds.Invoice);
        var confirmed = _allItems.Count(i => i.ReimburseDecision == "报销" || i.FileDecision == "有效");
        var pending = _allItems.Count(i => i.ReimburseDecision.Contains("待") || i.FileDecision.Contains("待"));
        var excluded = _allItems.Count(i => i.ReimburseDecision == "不报销" || i.FileDecision == "无效");
        _summary.Text = $"素材 {_allItems.Count} 个；当前显示 {_items.Count} 个；发票 {invoices} 个；已确认 {confirmed}；待确认 {pending}；已排除/无效 {excluded}；匹配候选 {_allMatches.Count}";
    }

    private void RefreshEvidenceViews()
    {
        _items.Clear();
        foreach (var item in _allItems.Where(i => !_hideDoneItems.Checked || NeedsItemAttention(i)))
        {
            _items.Add(item);
        }

        _invoiceItems.Clear();
        foreach (var item in _allInvoiceItems.Where(i => !_hideDoneInvoices.Checked || NeedsInvoiceAttention(i)))
        {
            _invoiceItems.Add(item);
        }
    }

    private void RefreshMatchView()
    {
        _matches.Clear();
        foreach (var match in _allMatches.Where(m => !_hideDoneMatches.Checked || m.Decision == "待确认"))
        {
            _matches.Add(match);
        }
    }

    private void RefreshDownloadView()
    {
        _downloadRecords.Clear();
        foreach (var record in _allDownloadRecords.Where(r => !_hideDoneDownloads.Checked || NeedsDownloadAttention(r)))
        {
            _downloadRecords.Add(record);
        }
    }

    private void RefreshEmailAuditView(string? preferredMessageKey = null, int preferredIndex = -1, int firstDisplayedRow = -1)
    {
        _emailAuditItems.Clear();
        foreach (var item in _allEmailAuditItems.Where(i => !_hideDoneEmailAudit.Checked || i.NeedsHumanReview))
        {
            _emailAuditItems.Add(item);
        }

        RestoreEmailAuditGridPosition(preferredMessageKey, preferredIndex, firstDisplayedRow);
    }

    private void RestoreEmailAuditGridPosition(string? preferredMessageKey, int preferredIndex, int firstDisplayedRow)
    {
        if (_emailAuditItems.Count == 0 || _emailAuditGrid.Rows.Count == 0)
        {
            UpdateEmailAuditDetail(null);
            return;
        }

        var rowIndex = -1;
        if (!string.IsNullOrWhiteSpace(preferredMessageKey))
        {
            for (var i = 0; i < _emailAuditGrid.Rows.Count; i++)
            {
                if (_emailAuditGrid.Rows[i].DataBoundItem is EmailAuditItem item
                    && item.MessageKey.Equals(preferredMessageKey, StringComparison.OrdinalIgnoreCase))
                {
                    rowIndex = i;
                    break;
                }
            }
        }
        if (rowIndex < 0 && preferredIndex >= 0)
        {
            rowIndex = Math.Min(preferredIndex, _emailAuditGrid.Rows.Count - 1);
        }
        if (rowIndex < 0)
        {
            return;
        }

        _emailAuditGrid.ClearSelection();
        _emailAuditGrid.Rows[rowIndex].Selected = true;
        _emailAuditGrid.CurrentCell = _emailAuditGrid.Rows[rowIndex].Cells[0];

        if (firstDisplayedRow >= 0 && firstDisplayedRow < _emailAuditGrid.Rows.Count)
        {
            try
            {
                _emailAuditGrid.FirstDisplayedScrollingRowIndex = firstDisplayedRow;
            }
            catch
            {
                // Grid can reject the value while layout is still recalculating.
            }
        }
    }

    private static bool NeedsItemAttention(EvidenceItem item)
    {
        if (item.Kind == EvidenceKinds.Template) return false;
        if (item.Kind == EvidenceKinds.Invoice) return NeedsInvoiceAttention(item);
        if (item.ReimburseDecision is "报销" or "不报销") return false;
        return item.ReimburseDecision.Contains("待") || item.ReimburseDecision.Contains("建议") || item.FileDecision.Contains("待");
    }

    private static bool NeedsInvoiceAttention(EvidenceItem item)
    {
        if (item.FileDecision is "有效" or "无效") return false;
        return item.FileDecision.Contains("待")
            || item.Suggestion.Contains("未识别")
            || item.DuplicateInfo.Length > 0;
    }

    private static bool NeedsDownloadAttention(DownloadRecord record)
    {
        return record.Status is "未下载到文件" or "失败" or "需人工确认" or "待核验" or "异常" or "链接取票待处理"
            || !string.IsNullOrWhiteSpace(record.Error) && record.Status is not "PDF已存在跳过" and not "已处理跳过" and not "人工确认无发票跳过" and not "附件已取得发票，链接跳过";
    }

    private EmailAuditItem? SelectedEmailAuditItem() => _emailAuditGrid.CurrentRow?.DataBoundItem as EmailAuditItem;

    private void UpdateEmailAuditDetail(EmailAuditItem? item)
    {
        if (item == null)
        {
            _emailAuditDetailBox.Clear();
            _emailDecisionNoteBox.Clear();
            return;
        }

        _emailAuditDetailBox.Text =
            $"最终判断：{item.FinalDecision}{Environment.NewLine}" +
            $"规则判断：{item.RuleDecision}{Environment.NewLine}" +
            $"人工判断：{item.ManualDecision}{Environment.NewLine}" +
            $"日期：{item.Date}{Environment.NewLine}" +
            $"主题：{item.Subject}{Environment.NewLine}" +
            $"状态：{item.Status}{Environment.NewLine}" +
            $"原因：{item.Reason}{Environment.NewLine}" +
            $"附件数：{item.AttachmentCount}{Environment.NewLine}" +
            $"下载数：{item.DownloadedFileCount}{Environment.NewLine}" +
            $"重复/跳过数：{item.SkippedOrDuplicateCount}{Environment.NewLine}" +
            $"文件：{Environment.NewLine}{item.Files}{Environment.NewLine}{Environment.NewLine}" +
            $"链接：{Environment.NewLine}{item.Urls}{Environment.NewLine}{Environment.NewLine}" +
            $"备注：{item.Note}";
        _emailDecisionNoteBox.Text = item.Note;
    }

    private void ConfirmSelectedEmail(string decision)
    {
        if (SelectedEmailAuditItem() is not { } item) return;
        var currentIndex = _emailAuditGrid.CurrentRow?.Index ?? -1;
        var firstDisplayed = _emailAuditGrid.FirstDisplayedScrollingRowIndex;
        var preferredKey = _hideDoneEmailAudit.Checked && decision != EmailAuditService.NeedsReview ? null : item.MessageKey;
        _emailAuditService.Confirm(_config, _emailDecisionStore, item, decision, _emailDecisionNoteBox.Text.Trim());
        RefreshEmailAuditView(preferredKey, currentIndex, firstDisplayed);
        UpdateEmailAuditDetail(SelectedEmailAuditItem());
        Log($"已确认邮件：{item.Subject} -> {decision}");
    }

    private void ConfirmSelectedEmails(string decision)
    {
        var selectedRowIndexes = _emailAuditGrid.SelectedRows
            .Cast<DataGridViewRow>()
            .Select(r => r.Index)
            .Concat(_emailAuditGrid.SelectedCells.Cast<DataGridViewCell>().Select(c => c.RowIndex))
            .Where(i => i >= 0)
            .Distinct()
            .OrderBy(i => i)
            .ToList();

        var targets = selectedRowIndexes
            .Select(i => _emailAuditGrid.Rows[i].DataBoundItem)
            .OfType<EmailAuditItem>()
            .GroupBy(i => i.MessageKey, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .ToList();
        if (targets.Count == 0 && SelectedEmailAuditItem() is { } current)
        {
            targets.Add(current);
        }
        if (targets.Count == 0)
        {
            return;
        }

        var firstDisplayed = _emailAuditGrid.FirstDisplayedScrollingRowIndex;
        var currentIndex = _emailAuditGrid.CurrentRow?.Index ?? -1;
        foreach (var item in targets)
        {
            _emailAuditService.Confirm(_config, _emailDecisionStore, item, decision, $"批量确认：{decision}");
        }

        RefreshEmailAuditView(null, currentIndex, firstDisplayed);
        SaveEmailDownloadRules();
        Log($"已批量标记邮件：{targets.Count} 封 -> {decision}");
    }

    private void ConfirmVisibleEmailsNoInvoice()
    {
        if (_emailAuditItems.Count == 0)
        {
            return;
        }

        var count = _emailAuditItems.Count;
        var result = MessageBox.Show(
            $"确认将当前列表显示的 {count} 封邮件全部标记为“明确无发票”吗？这个动作会写入规则记录，后续不再要求人工确认。",
            "批量确认无发票",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);
        if (result != DialogResult.Yes)
        {
            return;
        }

        var targets = _emailAuditItems.ToList();
        foreach (var item in targets)
        {
            _emailAuditService.Confirm(_config, _emailDecisionStore, item, EmailAuditService.NoInvoice, "批量确认：当前显示列表均无发票");
        }

        RefreshEmailAuditView();
        SaveEmailDownloadRules();
        Log($"已批量标记明确无发票：{targets.Count} 封。");
    }

    private void LoadEmailAuditItems(bool showMessageWhenMissing = true)
    {
        SaveConfig();
        _emailDecisionStore = _emailAuditService.LoadDecisions(_config);
        _allEmailAuditItems.Clear();
        _allEmailAuditItems.AddRange(_emailAuditService.LoadLatestAudit(_config, _emailDecisionStore));
        RefreshEmailAuditView();

        if (_allEmailAuditItems.Count == 0)
        {
            Log("没有找到邮箱发票下载清单，请先执行一次邮箱下载。");
            if (showMessageWhenMissing)
            {
                MessageBox.Show("没有找到邮箱发票下载清单，请先在“发票下载核验”里执行一次邮箱下载。", "没有邮件清单", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return;
        }

        var review = _allEmailAuditItems.Count(i => i.NeedsHumanReview);
        var hasInvoice = _allEmailAuditItems.Count(i => i.FinalDecision == EmailAuditService.HasInvoice);
        var noInvoice = _allEmailAuditItems.Count(i => i.FinalDecision == EmailAuditService.NoInvoice);
        Log($"已读取邮件判断清单：共 {_allEmailAuditItems.Count} 封；明确有发票 {hasInvoice}；明确无发票 {noInvoice}；需要人工确认 {review}。");
    }

    private void GenerateEmailChecklistExcel()
    {
        SaveConfig();
        if (_allEmailAuditItems.Count == 0) LoadEmailAuditItems(showMessageWhenMissing: false);
        if (_allEmailAuditItems.Count == 0) return;
        var output = _emailAuditService.GenerateChecklistExcel(_config, _allEmailAuditItems);
        Log($"已生成邮件发票判断清单：{output}");
        Process.Start(new ProcessStartInfo(output) { UseShellExecute = true });
    }

    private void SaveEmailDownloadRules()
    {
        SaveConfig();
        if (_allEmailAuditItems.Count == 0) LoadEmailAuditItems(showMessageWhenMissing: false);
        var path = _emailAuditService.SaveDownloadRules(_config, _allEmailAuditItems);
        Log($"已保存邮件下载规则：{path}");
        MessageBox.Show("邮件下载规则已保存。后续“明确无发票”的邮件会直接跳过，不再重复拉出来确认。", "规则已保存", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void LoadDownloadRecords()
    {
        _allDownloadRecords.Clear();
        var dir = _workspace.Resolve(_config.Email.OutputDir);
        if (!Directory.Exists(dir))
        {
            RefreshDownloadView();
            return;
        }

        var latest = Directory.EnumerateFiles(dir, "邮箱发票下载清单_*.csv", SearchOption.TopDirectoryOnly)
            .OrderByDescending(File.GetLastWriteTime)
            .FirstOrDefault();
        if (latest == null)
        {
            RefreshDownloadView();
            return;
        }

        var lines = File.ReadAllLines(latest, System.Text.Encoding.UTF8);
        if (lines.Length == 0)
        {
            RefreshDownloadView();
            return;
        }
        var headers = ParseCsvLine(lines[0]);
        foreach (var line in lines.Skip(1))
        {
            var values = ParseCsvLine(line);
            string Get(string name)
            {
                var index = headers.FindIndex(h => h.Equals(name, StringComparison.OrdinalIgnoreCase));
                return index >= 0 && index < values.Count ? values[index] : "";
            }
            var status = Get("status");
            var file = Get("file");
            var error = Get("error");
            if (string.IsNullOrWhiteSpace(status))
            {
                status = !string.IsNullOrWhiteSpace(error)
                    ? "失败"
                    : !string.IsNullOrWhiteSpace(file) && File.Exists(file)
                        ? "已下载"
                        : "待核验";
            }
            _allDownloadRecords.Add(new DownloadRecord
            {
                Status = status,
                DownloadStatus = Get("download_status"),
                Kind = Get("kind"),
                Date = Get("date"),
                Subject = TextEncodingFixer.Fix(Get("subject")),
                Sender = TextEncodingFixer.Fix(Get("sender")),
                MessageId = Get("msg_id"),
                MessageKey = Get("message_key"),
                File = file,
                Url = Get("url"),
                Error = error,
                InvoiceNumber = Get("invoice_number"),
                InvoiceAmount = Get("invoice_amount"),
                InvoiceDate = Get("invoice_date"),
                Md5 = Get("md5"),
                Sha256 = Get("sha256"),
                SavedFileCount = Get("saved_file_count"),
                AttachmentTotal = Get("attachment_total"),
                LinkCandidateTotal = Get("link_candidate_total"),
                DuplicateOf = Get("duplicate_of")
            });
        }

        RefreshDownloadView();
        var risky = _allDownloadRecords.Count(NeedsDownloadAttention);
        Log($"已读取下载清单：{Path.GetFileName(latest)}，记录 {_allDownloadRecords.Count} 条，当前显示 {_downloadRecords.Count} 条，需要核验 {risky} 条。");
    }

    private void ArchiveNonPdfInvoiceFormats()
    {
        SaveConfig();
        var moved = _invoiceCleaner.ArchiveNonPdfFormats(_config);
        Log($"已归档非PDF重复格式文件：{moved} 个。");
        ScanMaterials();
        MessageBox.Show($"已归档 {moved} 个同一发票已有 PDF 的非 PDF 格式文件。", "归档完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void GenerateEmailAuditReport()
    {
        SaveConfig();
        try
        {
            var output = _emailAuditReporter.Generate(_config);
            Log($"已生成逐封邮件统计Excel：{output}");
            Process.Start(new ProcessStartInfo(output) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            Log("生成逐封邮件统计失败：" + ex.Message);
            MessageBox.Show(ex.Message, "生成逐封邮件统计失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static List<string> ParseCsvLine(string line)
    {
        var result = new List<string>();
        var current = new System.Text.StringBuilder();
        var quoted = false;
        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];
            if (quoted)
            {
                if (ch == '"' && i + 1 < line.Length && line[i + 1] == '"')
                {
                    current.Append('"');
                    i++;
                }
                else if (ch == '"')
                {
                    quoted = false;
                }
                else
                {
                    current.Append(ch);
                }
            }
            else if (ch == ',')
            {
                result.Add(current.ToString());
                current.Clear();
            }
            else if (ch == '"')
            {
                quoted = true;
            }
            else
            {
                current.Append(ch);
            }
        }
        result.Add(current.ToString());
        return result;
    }

    private static DataGridView Grid() => new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        MultiSelect = false,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        AutoGenerateColumns = false,
        AllowUserToAddRows = false,
        AllowUserToDeleteRows = false
    };

    private static void AddColumn(DataGridView grid, string property, string header, int width)
    {
        grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = property, HeaderText = header, Width = width });
    }

    private static void ColorRows(DataGridView grid)
    {
        foreach (DataGridViewRow row in grid.Rows)
        {
            if (row.DataBoundItem is EvidenceItem item)
            {
                if (item.ReimburseDecision == "报销" || item.FileDecision == "有效") row.DefaultCellStyle.BackColor = Color.FromArgb(217, 234, 211);
                else if (item.ReimburseDecision.Contains("待") || item.FileDecision.Contains("待") || item.ReimburseDecision.Contains("建议")) row.DefaultCellStyle.BackColor = Color.FromArgb(255, 242, 204);
                else if (item.ReimburseDecision == "不报销" || item.FileDecision == "无效") row.DefaultCellStyle.BackColor = Color.FromArgb(244, 204, 204);
            }
            if (row.DataBoundItem is MatchCandidate match)
            {
                if (match.Decision == "确认匹配") row.DefaultCellStyle.BackColor = Color.FromArgb(217, 234, 211);
                else if (match.Decision == "待确认") row.DefaultCellStyle.BackColor = Color.FromArgb(255, 242, 204);
                else if (match.Decision == "否决匹配") row.DefaultCellStyle.BackColor = Color.FromArgb(244, 204, 204);
            }
            if (row.DataBoundItem is DownloadRecord record)
            {
                if (record.Status is "已下载" or "正常已下载" or "页面解析已下载" or "已处理跳过" or "无发票内容已跳过") row.DefaultCellStyle.BackColor = Color.FromArgb(217, 234, 211);
                else if (record.Status is "人工确认无发票跳过") row.DefaultCellStyle.BackColor = Color.FromArgb(229, 231, 235);
                else if (record.Status is "重复跳过" or "重复已存在" or "重复发票" or "PDF已存在跳过" or "非发票邮件" or "附件已取得发票，链接跳过") row.DefaultCellStyle.BackColor = Color.FromArgb(229, 231, 235);
                else if (record.Status is "未下载到文件" or "失败" or "需人工确认" or "待核验" or "异常" or "链接取票待处理" || !string.IsNullOrWhiteSpace(record.Error)) row.DefaultCellStyle.BackColor = Color.FromArgb(255, 242, 204);
            }
            if (row.DataBoundItem is EmailAuditItem email)
            {
                if (email.FinalDecision == EmailAuditService.HasInvoice) row.DefaultCellStyle.BackColor = Color.FromArgb(217, 234, 211);
                else if (email.FinalDecision == EmailAuditService.NoInvoice) row.DefaultCellStyle.BackColor = Color.FromArgb(229, 231, 235);
                else row.DefaultCellStyle.BackColor = Color.FromArgb(255, 242, 204);
            }
        }
    }

    private void AddLabeled(TableLayoutPanel panel, string label, TextBox textBox, int column, int row, int span)
    {
        var box = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Margin = new Padding(3), AutoSize = true };
        box.Controls.Add(new Label { Text = label, AutoSize = true, ForeColor = Color.FromArgb(52, 64, 84) }, 0, 0);
        textBox.Dock = DockStyle.Fill;
        box.Controls.Add(textBox, 0, 1);
        panel.Controls.Add(box, column, row);
        panel.SetColumnSpan(box, span);
    }

    private static Button Button(string text, EventHandler handler, bool primary = false)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = true,
            Height = 34,
            Dock = DockStyle.Fill,
            Margin = new Padding(4),
            BackColor = primary ? Color.FromArgb(37, 99, 235) : Color.White,
            ForeColor = primary ? Color.White : Color.FromArgb(24, 34, 48),
            FlatStyle = FlatStyle.Flat
        };
        button.FlatAppearance.BorderColor = Color.FromArgb(208, 213, 221);
        button.Click += handler;
        return button;
    }

    private void Log(string text)
    {
        _logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {text}{Environment.NewLine}");
    }
}
