namespace ReimbursementMatcher;

static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        if (Environment.GetCommandLineArgs().Skip(1).FirstOrDefault() is { } arg && arg.StartsWith("--"))
        {
            RunCommand(arg);
            return;
        }

        // To customize application configuration such as set high DPI settings or default font,
        // see https://aka.ms/applicationconfiguration.
        ApplicationConfiguration.Initialize();
        Application.Run(new Form1());
    }

    private static void RunCommand(string arg)
    {
        var workspace = new WorkspaceService();
        var config = workspace.LoadConfig();
        if (arg == "--scan")
        {
            var items = new MaterialScanner(workspace).Scan(config);
            var matches = new MatchingService().BuildCandidates(items);
            Console.WriteLine($"素材：{items.Count}");
            Console.WriteLine($"发票：{items.Count(i => i.Kind == EvidenceKinds.Invoice)}");
            Console.WriteLine($"付款/截图/订单：{items.Count(i => i.Kind != EvidenceKinds.Invoice && i.Kind != EvidenceKinds.Template)}");
            Console.WriteLine($"匹配候选：{matches.Count}");
            return;
        }
        if (arg == "--export")
        {
            var confirmations = new ConfirmationService(workspace);
            var store = confirmations.Load(config);
            var items = new MaterialScanner(workspace).Scan(config);
            var matches = new MatchingService().BuildCandidates(items);
            confirmations.Apply(store, items, matches);
            var output = new ReportService(workspace).Generate(config, items, matches);
            Console.WriteLine(output);
            return;
        }
        if (arg == "--email-audit")
        {
            var output = new EmailAuditReportService(workspace).Generate(config);
            Console.WriteLine(output);
            return;
        }
        if (arg == "--email-checklist")
        {
            var service = new EmailAuditService(workspace);
            var store = service.LoadDecisions(config);
            var items = service.LoadLatestAudit(config, store);
            var output = service.GenerateChecklistExcel(config, items);
            Console.WriteLine(output);
            return;
        }
        if (arg == "--email-presence-check")
        {
            var service = new EmailAuditService(workspace);
            var store = service.LoadDecisions(config);
            var items = service.LoadLatestAudit(config, store);
            var output = service.GenerateInvoicePresenceExcel(config, items);
            Console.WriteLine(output);
            return;
        }
        if (arg == "--previous-index")
        {
            var output = new PreviousReimbursementIndexService(workspace).Generate(config);
            Console.WriteLine(output);
            return;
        }
        if (arg == "--email-download" || arg == "--email-download-abnormal")
        {
            var password = config.Email.Password;
            if (string.IsNullOrWhiteSpace(password))
            {
                throw new InvalidOperationException("matcher_config.json 里没有保存邮箱授权码，无法命令行下载。");
            }
            var downloader = new EmailDownloader(workspace, Console.WriteLine);
            var output = downloader.DownloadAsync(config, password, arg == "--email-download-abnormal").GetAwaiter().GetResult();
            Console.WriteLine(output);
            return;
        }
        if (arg == "--archive-nonpdf-invoices")
        {
            var moved = new InvoiceFileCleaner(workspace).ArchiveNonPdfFormats(config);
            Console.WriteLine($"已归档非PDF重复格式文件：{moved}");
            return;
        }
        if (arg == "--archive-date-rejected-invoices")
        {
            var moved = new InvoiceFileCleaner(workspace).ArchiveDateRejectedInvoices(config);
            Console.WriteLine($"已归档日期不符发票：{moved}");
            return;
        }
        if (arg == "--consolidate-pdf-invoices")
        {
            var output = new PdfInvoiceConsolidationService(workspace).Generate(config);
            Console.WriteLine(output);
            return;
        }
        if (arg == "--process-nonpdf-qr")
        {
            var output = new NonPdfInvoiceMaintenanceService(workspace, Console.WriteLine)
                .ProcessAsync(config)
                .GetAwaiter()
                .GetResult();
            Console.WriteLine(output);
            return;
        }
        throw new ArgumentException("未知参数：" + arg);
    }
}
