using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Diagnostics;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using UglyToad.PdfPig;

namespace ReimbursementMatcher;

public sealed class PreviousReimbursementIndexService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private readonly WorkspaceService _workspace;

    public PreviousReimbursementIndexService(WorkspaceService workspace)
    {
        _workspace = workspace;
    }

    public string Generate(AppConfig config)
    {
        var roots = FindPreviousRoots(config).ToList();
        if (roots.Count == 0)
        {
            throw new DirectoryNotFoundException("没有找到上期报销目录。可以在“上期发票目录”填写，或把材料放到 报销准备的资料/上一期报销。");
        }

        var archiveRoots = roots.SelectMany(root => ExtractArchivesForIndex(config, root)).ToList();
        var scanRoots = roots.Concat(archiveRoots).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

        var documents = scanRoots
            .SelectMany(root => EnumerateFilesSafe(root, "*.*")
                .Where(IsPreviousDocument)
                .Select(file => BuildDocument(root, file)))
            .OrderBy(d => d.SourcePath, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var invoices = scanRoots
            .SelectMany(root => EnumerateFilesSafe(root, "*.*")
                .Where(IsInvoiceFileForIndex)
                .Select(file => BuildInvoice(config, root, file, documents)))
            .GroupBy(InvoiceDedupeKey, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.OrderByDescending(i => i.InvoiceNumber.Length)
                .ThenByDescending(i => i.Amount)
                .ThenBy(i => i.SourcePath, StringComparer.OrdinalIgnoreCase)
                .First())
            .OrderBy(i => i.InvoiceDate)
            .ThenBy(i => i.Category)
            .ThenBy(i => i.Amount)
            .ToList();

        var index = new PreviousReimbursementIndex
        {
            UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
            SourceRoots = scanRoots,
            Invoices = invoices,
            Documents = documents
        };

        var jsonPath = IndexPath(config);
        Directory.CreateDirectory(Path.GetDirectoryName(jsonPath)!);
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(index, JsonOptions) + Environment.NewLine, Encoding.UTF8);

        var outputDir = _workspace.Resolve(config.OutputDir);
        Directory.CreateDirectory(outputDir);
        var workbookPath = Path.Combine(outputDir, $"上期报销查询表_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        WriteWorkbook(workbookPath, index, jsonPath);
        return workbookPath;
    }

    public PreviousReimbursementIndex Load(AppConfig config)
    {
        var path = IndexPath(config);
        if (!File.Exists(path))
        {
            return new PreviousReimbursementIndex();
        }

        try
        {
            return JsonSerializer.Deserialize<PreviousReimbursementIndex>(File.ReadAllText(path, Encoding.UTF8), JsonOptions)
                ?? new PreviousReimbursementIndex();
        }
        catch
        {
            return new PreviousReimbursementIndex();
        }
    }

    public string IndexPath(AppConfig config)
    {
        return Path.Combine(_workspace.Resolve(config.RuleDir), "上期报销索引.json");
    }

    public string? LatestWorkbook(AppConfig config)
    {
        var outputDir = _workspace.Resolve(config.OutputDir);
        return Directory.Exists(outputDir)
            ? Directory.EnumerateFiles(outputDir, "上期报销查询表_*.xlsx", SearchOption.TopDirectoryOnly)
                .OrderByDescending(File.GetLastWriteTime)
                .FirstOrDefault()
            : null;
    }

    public PreviousReimbursementMatch MatchMail(PreviousReimbursementIndex index, string subject, string sender, string date, IEnumerable<string> attachmentNames, IEnumerable<string> urls)
    {
        if (index.Invoices.Count == 0)
        {
            return new PreviousReimbursementMatch();
        }

        var combined = TextEncodingFixer.Fix(string.Join(" ", new[]
        {
            subject,
            sender,
            date,
            string.Join(" ", attachmentNames),
            string.Join(" ", urls)
        }));

        var invoiceNumbers = ExtractInvoiceNumbers(combined).ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (invoiceNumbers.Count > 0)
        {
            var byNo = index.Invoices.FirstOrDefault(i =>
                !string.IsNullOrWhiteSpace(i.InvoiceNumber)
                && invoiceNumbers.Contains(i.InvoiceNumber));
            if (byNo != null)
            {
                return BuildMatch("发票号命中上期索引", 100, byNo);
            }
        }

        var amount = ExtractAmount(combined);
        if (amount > 0)
        {
            var byVendorAndAmount = index.Invoices
                .Where(i => i.Amount == amount)
                .Where(i => !string.IsNullOrWhiteSpace(i.Vendor) && combined.Contains(i.Vendor, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(i => i.InvoiceDate)
                .FirstOrDefault();
            if (byVendorAndAmount != null)
            {
                return BuildMatch("金额和销售方命中上期索引", 88, byVendorAndAmount);
            }
        }

        return new PreviousReimbursementMatch();
    }

    public PreviousReimbursementMatch MatchInvoiceNumber(PreviousReimbursementIndex index, string invoiceNumber)
    {
        if (index.Invoices.Count == 0 || string.IsNullOrWhiteSpace(invoiceNumber))
        {
            return new PreviousReimbursementMatch();
        }

        var normalized = NormalizeInvoiceNo(invoiceNumber);
        var invoice = index.Invoices.FirstOrDefault(i => i.InvoiceNumber.Equals(normalized, StringComparison.OrdinalIgnoreCase));
        return invoice == null
            ? new PreviousReimbursementMatch()
            : BuildMatch("发票号命中上期索引", 100, invoice);
    }

    public PreviousReimbursementMatch MatchFile(PreviousReimbursementIndex index, string file)
    {
        if (index.Invoices.Count == 0 || string.IsNullOrWhiteSpace(file) || !File.Exists(file))
        {
            return new PreviousReimbursementMatch();
        }

        var invoiceNumber = ExtractInvoiceNo($"{Path.GetFileName(file)} {ExtractPdfText(file)}");
        return MatchInvoiceNumber(index, invoiceNumber);
    }

    public IEnumerable<string> FindPreviousRoots(AppConfig config)
    {
        var explicitRoots = config.PreviousInvoiceDirs
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Select(_workspace.Resolve);

        var sourceRoot = _workspace.Resolve(config.SourceRoot);
        var autoRoots = Directory.Exists(sourceRoot)
            ? Directory.EnumerateDirectories(sourceRoot, "*", SearchOption.AllDirectories)
                .Where(dir => IsPreviousMaterialDir(Path.GetFileName(dir)))
                .Where(HasAnyUsefulFile)
                .Select(Path.GetFullPath)
            : [];

        var matches = explicitRoots
            .Concat(autoRoots)
            .Where(Directory.Exists)
            .Select(Path.GetFullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(dir => dir.Length)
            .ToList();

        var roots = new List<string>();
        foreach (var dir in matches)
        {
            if (!roots.Any(root => IsSubPath(root, dir)))
            {
                roots.Add(dir);
            }
        }
        return roots;
    }

    private static PreviousReimbursementInvoice BuildInvoice(AppConfig config, string root, string path, List<PreviousReimbursementDocument> documents)
    {
        var text = ExtractInvoiceText(path);
        var combined = $"{Path.GetFileName(path)} {path} {text}";
        var qr = InvoiceQrCodeService.ExtractFromPdfImages(path);
        var invoice = new PreviousReimbursementInvoice
        {
            SourceRoot = root,
            SourcePath = path,
            FileName = Path.GetFileName(path),
            Sha256 = Sha256(path),
            InvoiceNumber = FirstNonEmpty(ExtractInvoiceNo(combined), qr.InvoiceNumber),
            Amount = ExtractAmount(combined),
            InvoiceDate = FirstNonEmpty(ExtractDate(combined), qr.InvoiceDate),
            Vendor = ExtractVendor(combined, config),
            Category = DetectCategory(combined),
            RelatedDocument = FindRelatedDocument(path, documents),
            TextStatus = BuildTextStatus(text, qr)
        };
        if (invoice.Amount <= 0 && qr.Amount > 0)
        {
            invoice.Amount = qr.Amount;
        }
        FillRecognitionStatus(invoice, text);
        return invoice;
    }

    private IEnumerable<string> ExtractArchivesForIndex(AppConfig config, string root)
    {
        var archiveFiles = EnumerateFilesSafe(root, "*.*")
            .Where(file => Path.GetExtension(file).Equals(".zip", StringComparison.OrdinalIgnoreCase)
                || Path.GetExtension(file).Equals(".rar", StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (archiveFiles.Count == 0)
        {
            return [];
        }

        var cacheRoot = _workspace.Resolve(Path.Combine(config.RuleDir, "上期压缩包解压缓存"));
        Directory.CreateDirectory(cacheRoot);
        var extracted = new List<string>();
        foreach (var archive in archiveFiles)
        {
            var hash = Sha256(archive);
            if (string.IsNullOrWhiteSpace(hash))
            {
                continue;
            }

            var target = Path.Combine(cacheRoot, hash[..Math.Min(16, hash.Length)]);
            if (!Directory.Exists(target) || !Directory.EnumerateFiles(target, "*.pdf", SearchOption.AllDirectories).Any())
            {
                Directory.CreateDirectory(target);
                TryExtractArchive(archive, target);
            }
            if (Directory.Exists(target) && Directory.EnumerateFiles(target, "*.pdf", SearchOption.AllDirectories).Any())
            {
                extracted.Add(target);
            }
        }
        return extracted;
    }

    private static void TryExtractArchive(string archive, string target)
    {
        try
        {
            var start = new ProcessStartInfo
            {
                FileName = "tar",
                Arguments = $"-xf \"{archive}\" -C \"{target}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            using var process = Process.Start(start);
            process?.WaitForExit(15000);
        }
        catch
        {
            // Archive extraction is best-effort. Already extracted folders are still scanned.
        }
    }

    private static PreviousReimbursementDocument BuildDocument(string root, string path)
    {
        var info = new FileInfo(path);
        return new PreviousReimbursementDocument
        {
            SourceRoot = root,
            SourcePath = path,
            FileName = Path.GetFileName(path),
            Kind = Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ? "报销Excel" : "上期资料",
            SizeBytes = info.Length,
            UpdatedAt = info.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
        };
    }

    private static string FindRelatedDocument(string invoicePath, List<PreviousReimbursementDocument> documents)
    {
        if (documents.Count == 0) return "";
        var dir = Path.GetDirectoryName(invoicePath) ?? "";
        return documents
            .OrderBy(d => CommonPrefixLength(dir, Path.GetDirectoryName(d.SourcePath) ?? ""))
            .LastOrDefault()?.SourcePath ?? "";
    }

    private static void WriteWorkbook(string path, PreviousReimbursementIndex index, string jsonPath)
    {
        using var wb = new XLWorkbook();
        var summary = wb.Worksheets.Add("总览");
        var metrics = new (string, object)[]
        {
            ("索引更新时间", index.UpdatedAt),
            ("索引JSON", jsonPath),
            ("上期来源目录", string.Join(Environment.NewLine, index.SourceRoots)),
            ("上期PDF发票数量", index.Invoices.Count),
            ("上期PDF发票金额合计", index.Invoices.Sum(i => i.Amount)),
            ("上期报销文档数量", index.Documents.Count)
        };
        summary.Cell(1, 1).Value = "指标";
        summary.Cell(1, 2).Value = "值";
        summary.Range(1, 1, 1, 2).Style.Font.Bold = true;
        for (var i = 0; i < metrics.Length; i++)
        {
            summary.Cell(i + 2, 1).Value = metrics[i].Item1;
            if (metrics[i].Item2 is decimal amount)
            {
                summary.Cell(i + 2, 2).Value = amount;
                summary.Cell(i + 2, 2).Style.NumberFormat.Format = "¥#,##0.00";
            }
            else
            {
                summary.Cell(i + 2, 2).Value = metrics[i].Item2.ToString();
            }
        }
        summary.Columns().AdjustToContents(8, 80);

        var invoices = wb.Worksheets.Add("上期发票索引");
        var invoiceHeaders = new[] { "类别", "开票日期", "金额", "发票号", "销售方", "识别状态", "缺失项/原因", "文本状态", "发票文件", "关联报销文档", "SHA256" };
        WriteHeaders(invoices, invoiceHeaders);
        var r = 2;
        foreach (var item in index.Invoices)
        {
            invoices.Cell(r, 1).Value = item.Category;
            invoices.Cell(r, 2).Value = item.InvoiceDate;
            invoices.Cell(r, 3).Value = item.Amount;
            invoices.Cell(r, 3).Style.NumberFormat.Format = "¥#,##0.00";
            invoices.Cell(r, 4).Value = item.InvoiceNumber;
            invoices.Cell(r, 5).Value = item.Vendor;
            invoices.Cell(r, 6).Value = item.RecognitionStatus;
            invoices.Cell(r, 7).Value = item.RecognitionIssues;
            invoices.Cell(r, 8).Value = item.TextStatus;
            invoices.Cell(r, 9).Value = item.SourcePath;
            invoices.Cell(r, 10).Value = item.RelatedDocument;
            invoices.Cell(r, 11).Value = item.Sha256;
            r++;
        }
        FormatSheet(invoices, invoiceHeaders.Length);

        var documents = wb.Worksheets.Add("上期报销文档");
        var docHeaders = new[] { "类型", "文件名", "更新时间", "大小KB", "路径" };
        WriteHeaders(documents, docHeaders);
        r = 2;
        foreach (var doc in index.Documents)
        {
            documents.Cell(r, 1).Value = doc.Kind;
            documents.Cell(r, 2).Value = doc.FileName;
            documents.Cell(r, 3).Value = doc.UpdatedAt;
            documents.Cell(r, 4).Value = Math.Round(doc.SizeBytes / 1024m, 1);
            documents.Cell(r, 5).Value = doc.SourcePath;
            r++;
        }
        FormatSheet(documents, docHeaders.Length);

        wb.SaveAs(path);
    }

    private static void WriteHeaders(IXLWorksheet ws, IReadOnlyList<string> headers)
    {
        for (var c = 0; c < headers.Count; c++)
        {
            ws.Cell(1, c + 1).Value = headers[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
            ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E5E7EB");
        }
    }

    private static void FormatSheet(IXLWorksheet ws, int columnCount)
    {
        ws.Columns().AdjustToContents(8, 70);
        ws.Range(1, 1, Math.Max(1, ws.LastRowUsed()?.RowNumber() ?? 1), columnCount).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        ws.SheetView.FreezeRows(1);
        ws.Range(1, 1, Math.Max(1, ws.LastRowUsed()?.RowNumber() ?? 1), columnCount).SetAutoFilter();
    }

    private static IEnumerable<string> EnumerateFilesSafe(string root, string pattern)
    {
        try
        {
            return Directory.EnumerateFiles(root, pattern, SearchOption.AllDirectories)
                .Where(file => !file.Contains($"{Path.DirectorySeparatorChar}~$", StringComparison.OrdinalIgnoreCase))
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    private static PreviousReimbursementMatch BuildMatch(string type, int score, PreviousReimbursementInvoice invoice) => new()
    {
        IsMatch = true,
        Score = score,
        MatchType = type,
        Summary = $"{type}：{invoice.InvoiceDate} {invoice.Amount:0.##} {invoice.InvoiceNumber} {invoice.Vendor}",
        SourcePath = invoice.SourcePath,
        InvoiceNumber = invoice.InvoiceNumber,
        Amount = invoice.Amount,
        InvoiceDate = invoice.InvoiceDate,
        Vendor = invoice.Vendor
    };

    private static string InvoiceDedupeKey(PreviousReimbursementInvoice row)
    {
        return !string.IsNullOrWhiteSpace(row.InvoiceNumber)
            ? $"no:{row.InvoiceNumber}"
            : $"sha:{row.Sha256}";
    }

    private static bool IsPreviousDocument(string file)
    {
        var ext = Path.GetExtension(file);
        return ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".xls", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".docx", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".doc", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsInvoiceFileForIndex(string file)
    {
        var ext = Path.GetExtension(file);
        return ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".ofd", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".xml", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsPreviousMaterialDir(string name)
    {
        return ContainsAny(name, "上一期", "上期", "历史", "已报销", "往期", "上次报销", "前期报销");
    }

    private static bool HasAnyUsefulFile(string dir)
    {
        try
        {
            return Directory.EnumerateFiles(dir, "*.pdf", SearchOption.AllDirectories).Any()
                || Directory.EnumerateFiles(dir, "*.ofd", SearchOption.AllDirectories).Any()
                || Directory.EnumerateFiles(dir, "*.xml", SearchOption.AllDirectories).Any()
                || Directory.EnumerateFiles(dir, "*.zip", SearchOption.AllDirectories).Any()
                || Directory.EnumerateFiles(dir, "*.rar", SearchOption.AllDirectories).Any()
                || Directory.EnumerateFiles(dir, "*.xlsx", SearchOption.AllDirectories).Any();
        }
        catch
        {
            return false;
        }
    }

    private static bool IsSubPath(string parent, string child)
    {
        var parentFull = Path.GetFullPath(parent).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        var childFull = Path.GetFullPath(child).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        return childFull.StartsWith(parentFull, StringComparison.OrdinalIgnoreCase);
    }

    private static string ExtractPdfText(string path)
    {
        try
        {
            using var document = PdfDocument.Open(path);
            var builder = new StringBuilder();
            foreach (var page in document.GetPages().Take(3))
            {
                builder.AppendLine(page.Text);
            }
            return Regex.Replace(builder.ToString(), @"\s+", " ");
        }
        catch
        {
            return "";
        }
    }

    private static string ExtractInvoiceText(string path)
    {
        var ext = Path.GetExtension(path);
        if (ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            return ExtractPdfText(path);
        }

        if (ext.Equals(".xml", StringComparison.OrdinalIgnoreCase))
        {
            return ReadTextFile(path);
        }

        if (ext.Equals(".ofd", StringComparison.OrdinalIgnoreCase))
        {
            return ExtractOfdText(path);
        }

        return "";
    }

    private static string ReadTextFile(string path)
    {
        try
        {
            var bytes = File.ReadAllBytes(path);
            var utf8 = Encoding.UTF8.GetString(bytes);
            return TextEncodingFixer.Fix(Regex.Replace(utf8, @"\s+", " "));
        }
        catch
        {
            return "";
        }
    }

    private static string ExtractOfdText(string path)
    {
        try
        {
            using var stream = File.OpenRead(path);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            var builder = new StringBuilder();
            foreach (var entry in archive.Entries
                .Where(e => e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .Take(30))
            {
                using var entryStream = entry.Open();
                using var reader = new StreamReader(entryStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                builder.Append(' ').Append(reader.ReadToEnd());
            }
            return TextEncodingFixer.Fix(Regex.Replace(builder.ToString(), "<[^>]+>", " "));
        }
        catch
        {
            return ReadTextFile(path);
        }
    }

    private static List<string> ExtractInvoiceNumbers(string text)
    {
        var numbers = new List<string>();

        var oldStyle = Regex.Matches(text, @"发票代码\s*[:：]?\s*([0-9]{10,12}).{0,80}?发票号码\s*[:：]?\s*([0-9]{8})", RegexOptions.Singleline);
        numbers.AddRange(oldStyle.Select(m => NormalizeInvoiceNo(m.Groups[1].Value + m.Groups[2].Value)));

        var labeled = Regex.Matches(text, @"(?:发票号码|发票号|电子发票号码|EIid|InvoiceNumber)[_：:\s-]*([0-9]{8,30})", RegexOptions.IgnoreCase);
        numbers.AddRange(labeled
            .Select(m => NormalizeInvoiceNo(m.Groups[1].Value))
            .Where(v => v.Length >= 18));

        numbers.AddRange(Regex.Matches(text, @"(?<!\d)([0-9]{10,12})\s+([0-9]{8})(?!\d)")
            .Select(m => NormalizeInvoiceNo(m.Groups[1].Value + m.Groups[2].Value)));

        var dzfp = Regex.Matches(text, @"dzfp[_-]([0-9]{12,30})", RegexOptions.IgnoreCase);
        numbers.AddRange(dzfp.Select(m => NormalizeInvoiceNo(m.Groups[1].Value)));

        numbers.AddRange(Regex.Matches(text, @"(?<!\d)([0-9]{18,24})(?!\d)")
            .Select(m => NormalizeInvoiceNo(m.Groups[1].Value)));

        return numbers
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Where(v => !LooksLikeNumericTaxId(v))
            .Where(v => !LooksLikeDateNoise(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string ExtractInvoiceNo(string text)
    {
        return ExtractInvoiceNumbers(text).FirstOrDefault() ?? "";
    }

    private static string NormalizeInvoiceNo(string value)
    {
        value = Regex.Replace(value, @"\D", "");
        if (value.Length > 20 && !value.StartsWith("20", StringComparison.Ordinal))
        {
            value = value[..20];
        }
        return value;
    }

    private static bool LooksLikeNumericTaxId(string value)
    {
        return value.Length == 18 && (value.StartsWith("91", StringComparison.Ordinal) || value.StartsWith("92", StringComparison.Ordinal));
    }

    private static bool LooksLikeDateNoise(string value)
    {
        if (Regex.IsMatch(value, @"^20\d{6}20\d{6}\d*$"))
        {
            return true;
        }

        var dateTokens = Regex.Matches(value, @"20\d{6}")
            .Select(m => m.Value)
            .Count(v => DateTime.TryParseExact(v, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _));
        return dateTokens >= 2;
    }

    private static string ExtractDate(string text)
    {
        var match = Regex.Match(text, @"(20\d{2})\s*[-_.年/]\s*(\d{1,2})\s*[-_.月/]\s*(\d{1,2})");
        if (match.Success)
        {
            var value = $"{int.Parse(match.Groups[1].Value):0000}-{int.Parse(match.Groups[2].Value):00}-{int.Parse(match.Groups[3].Value):00}";
            return DateTime.TryParse(value, out var parsed) ? parsed.ToString("yyyy-MM-dd") : "";
        }

        match = Regex.Match(text, @"(?<!\d)(20\d{6})(?!\d)");
        if (match.Success)
        {
            var value = match.Groups[1].Value;
            var normalized = $"{value[..4]}-{value.Substring(4, 2)}-{value.Substring(6, 2)}";
            return DateTime.TryParse(normalized, out var parsed) ? parsed.ToString("yyyy-MM-dd") : "";
        }
        return "";
    }

    private static decimal ExtractAmount(string text)
    {
        var patterns = new[]
        {
            @"([0-9]{1,6}(?:\.[0-9]{1,2})?)元-合并发票文件",
            @"发票金额\s*[：:】\]\s]*[¥￥]?\s*([0-9]{1,6}(?:\.[0-9]{1,2})?)",
            @"共\s*\d+\s*笔.{0,30}?合计\s*([0-9]{1,6}(?:\.[0-9]{1,2})?)\s*元",
            @"合计\s*([0-9]{1,6}(?:\.[0-9]{1,2})?)\s*元",
            @"金额\s*[：:】\]\s]*[¥￥]\s*([0-9]{1,6}(?:\.[0-9]{1,2})?)"
        };

        foreach (var pattern in patterns)
        {
            var match = Regex.Match(text, pattern, RegexOptions.Singleline);
            if (match.Success && TryParseReasonableAmount(match.Groups[1].Value, out var amount))
            {
                return amount;
            }
        }
        var currencyAmounts = Regex.Matches(text, @"[¥￥]\s*([0-9]{1,6}(?:\.[0-9]{1,2})?)")
            .Select(m => TryParseReasonableAmount(m.Groups[1].Value, out var amount) ? amount : 0m)
            .Where(amount => amount > 0)
            .ToList();
        if (currencyAmounts.Count > 0)
        {
            return currencyAmounts.Max();
        }
        return 0m;
    }

    private static bool TryParseReasonableAmount(string value, out decimal amount)
    {
        if (decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out amount)
            && amount > 0
            && amount <= 1_000_000m)
        {
            return true;
        }

        amount = 0m;
        return false;
    }

    private static void FillRecognitionStatus(PreviousReimbursementInvoice invoice, string extractedText)
    {
        var issues = new List<string>();
        if (string.IsNullOrWhiteSpace(extractedText))
        {
            issues.Add(invoice.TextStatus.Contains("已识别发票二维码", StringComparison.OrdinalIgnoreCase)
                ? "图片型PDF，已通过二维码补核心字段"
                : "PDF无可提取文本/可能是扫描件");
        }

        if (string.IsNullOrWhiteSpace(invoice.InvoiceNumber))
        {
            issues.Add("缺发票号");
        }

        if (string.IsNullOrWhiteSpace(invoice.InvoiceDate))
        {
            issues.Add("缺开票日期");
        }

        if (invoice.Amount <= 0)
        {
            issues.Add("缺金额");
        }

        if (string.IsNullOrWhiteSpace(invoice.Vendor))
        {
            issues.Add("缺销售方");
        }

        if (invoice.FileName.Contains("合并发票", StringComparison.OrdinalIgnoreCase)
            || Regex.IsMatch(invoice.FileName, @"\d+张-\d", RegexOptions.IgnoreCase))
        {
            issues.Add("合并发票文件，建议按单张明细核验");
        }

        invoice.RecognitionIssues = string.Join("；", issues.Distinct());
        invoice.RecognitionStatus = issues.Count == 0 ? "完整" : "需复核";
    }

    private static string BuildTextStatus(string text, InvoiceQrMetadata qr)
    {
        var parts = new List<string>();
        parts.Add(string.IsNullOrWhiteSpace(text) ? "无可提取文本" : $"已提取文本({text.Length})");
        if (!string.IsNullOrWhiteSpace(qr.Raw))
        {
            parts.Add("已识别发票二维码");
        }
        return string.Join("；", parts);
    }

    private static string FirstNonEmpty(params string[] values)
    {
        return values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "";
    }

    private static string ExtractVendor(string text, AppConfig config)
    {
        var patterns = new[]
        {
            @"销售方信息.{0,120}?名\s*称[:：]?\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,60})",
            @"销售方\s*名\s*称[:：]\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,60})",
            @"销方\s*名\s*称[:：]\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,60})",
            @"销\s*货\s*方.{0,40}?名\s*称[:：]\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,60})",
            @"<[^>]*(?:SellerName|Seller|Xsfmc|XSFMC)[^>]*>\s*([^<]{4,80})\s*</"
        };

        foreach (var pattern in patterns)
        {
            var match = Regex.Match(text, pattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            if (match.Success)
            {
                var value = NormalizeVendorCandidate(match.Groups[1].Value, config);
                if (LooksLikeValidVendor(value))
                {
                    return value;
                }
            }
        }

        var valueFromCompanyName = ExtractNeighborVendorByCompanyName(text, config);
        if (LooksLikeValidVendor(valueFromCompanyName))
        {
            return valueFromCompanyName;
        }

        var valueFromTaxId = ExtractNeighborVendorByCompanyTaxId(text, config);
        if (LooksLikeValidVendor(valueFromTaxId))
        {
            return valueFromTaxId;
        }

        var fileNameVendor = Regex.Match(text, @"[0-9]+(?:\.[0-9]{1,2})?元-([^-\\/:]{4,60}?)-20\d{2}[._-]?\d{1,2}", RegexOptions.Singleline);
        if (fileNameVendor.Success)
        {
            var value = NormalizeVendorCandidate(fileNameVendor.Groups[1].Value, config);
            if (LooksLikeValidVendor(value))
            {
                return value;
            }
        }

        var name = Path.GetFileNameWithoutExtension(text);
        var parts = Regex.Split(name, @"[_\s]+").Where(p => p.Length >= 4).ToList();
        return parts
            .Select(p => NormalizeVendorCandidate(p, config))
            .FirstOrDefault(p =>
                p.Any(ch => ch >= '\u4e00' && ch <= '\u9fff')
                && LooksLikeValidVendor(p)
                && p.Length <= 40) ?? "";
    }

    private static string ExtractNeighborVendorByCompanyName(string text, AppConfig config)
    {
        var companyName = CleanVendor(config.CompanyName);
        if (string.IsNullOrWhiteSpace(companyName) || !text.Contains(companyName, StringComparison.OrdinalIgnoreCase))
        {
            return "";
        }

        var escaped = Regex.Escape(companyName);
        var match = Regex.Match(text, escaped + @"(?<vendor>[\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,80}?)(?:9[0-9A-Z]{14,19}|[¥￥]|[*＊]|合计|价税合计|下载次数)", RegexOptions.Singleline);
        if (match.Success)
        {
            var value = NormalizeVendorCandidate(match.Groups["vendor"].Value, config);
            if (LooksLikeValidVendor(value))
            {
                return value;
            }
        }

        match = Regex.Match(text, escaped + @"(?<tail>.{0,160})", RegexOptions.Singleline);
        if (match.Success)
        {
            foreach (var value in ExtractCompanyLikeNames(match.Groups["tail"].Value, config))
            {
                if (!value.Equals(companyName, StringComparison.OrdinalIgnoreCase) && LooksLikeValidVendor(value))
                {
                    return value;
                }
            }
        }

        return "";
    }

    private static string ExtractNeighborVendorByCompanyTaxId(string text, AppConfig config)
    {
        var companyTaxId = Regex.Replace(config.CompanyTaxId ?? "", @"[^0-9A-Z]", "", RegexOptions.IgnoreCase).ToUpperInvariant();
        if (string.IsNullOrWhiteSpace(companyTaxId) || !text.Contains(companyTaxId, StringComparison.OrdinalIgnoreCase))
        {
            return "";
        }

        var escaped = Regex.Escape(companyTaxId);
        var match = Regex.Match(text, escaped + @"(?<vendor>[\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,80}?)(?:9[0-9A-Z]{14,19}|[¥￥]|[*＊]|合计|价税合计|下载次数)", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        if (match.Success)
        {
            var value = NormalizeVendorCandidate(match.Groups["vendor"].Value, config);
            if (LooksLikeValidVendor(value))
            {
                return value;
            }
        }

        match = Regex.Match(text, @"(?<head>.{0,120})" + escaped, RegexOptions.Singleline | RegexOptions.IgnoreCase);
        if (match.Success)
        {
            var candidates = ExtractCompanyLikeNames(match.Groups["head"].Value, config).ToList();
            if (candidates.Count > 0)
            {
                var value = candidates.Last();
                if (LooksLikeValidVendor(value))
                {
                    return value;
                }
            }
        }

        return "";
    }

    private static IEnumerable<string> ExtractCompanyLikeNames(string text, AppConfig config)
    {
        return Regex.Matches(text, @"[\u4e00-\u9fff][\u4e00-\u9fffA-Za-z0-9（）()·\-]{1,59}(?:有限公司|分公司|公司|餐饮服务有限公司|餐饮店|旅馆|酒店|宾馆|烧烤店|商贸有限公司|科技有限公司|高速公路有限公司|路桥投资建设有限公司|服务区)")
            .Select(m => NormalizeVendorCandidate(m.Value, config))
            .Where(LooksLikeValidVendor)
            .Distinct(StringComparer.OrdinalIgnoreCase);
    }

    private static string NormalizeVendorCandidate(string value, AppConfig config)
    {
        value = CleanVendor(value);
        var companyName = CleanVendor(config.CompanyName);
        var companyTaxId = Regex.Replace(config.CompanyTaxId ?? "", @"[^0-9A-Z]", "", RegexOptions.IgnoreCase).ToUpperInvariant();

        if (!string.IsNullOrWhiteSpace(companyTaxId))
        {
            value = Regex.Replace(value, Regex.Escape(companyTaxId), "", RegexOptions.IgnoreCase);
        }

        if (!string.IsNullOrWhiteSpace(companyName) && value.Contains(companyName, StringComparison.OrdinalIgnoreCase))
        {
            value = value.Replace(companyName, "", StringComparison.OrdinalIgnoreCase);
        }

        value = Regex.Replace(value, @"^(?:20\d{2})?年?\d{1,2}月\d{1,2}日", "");
        value = Regex.Replace(value, @"^[0-9A-Z]{15,}", "", RegexOptions.IgnoreCase);
        value = Regex.Replace(value, @"^\d{8,}", "");
        return CleanVendor(value);
    }

    private static string CleanVendor(string value)
    {
        return Regex.Replace(value ?? "", @"\s+", "").Trim();
    }

    private static bool LooksLikeValidVendor(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        value = CleanVendor(value);
        if (value.Length < 4 || value.Length > 80)
        {
            return false;
        }

        if (LooksLikeInvoiceTableHeader(value) || LooksLikeVendorNoise(value))
        {
            return false;
        }

        return ContainsAny(value,
            "公司", "分公司", "有限公司", "商贸", "科技", "餐", "饭", "酒店", "宾馆", "烧烤",
            "旅馆", "住宿", "餐饮", "高速", "路桥", "服务区", "商行", "商店", "店");
    }

    private static bool LooksLikeVendorNoise(string value)
    {
        if (ContainsAny(value, ":", "：", ";", "；", "项目名称", "规格型号", "统一社会信用代码", "纳税人识别号",
            "下载次数", "税率", "征收率", "单价", "金额", "发票代码", "发票号码", "开票日期", "银行账号", "账号",
            "复核人", "开票人", "收款人", "机器编号", "校验码", "电子支付标识", "密码区", "电话", "地址", "车牌号"))
        {
            return true;
        }

        if (Regex.IsMatch(value, @"[0-9A-Z]{15,}|\d{8,}|20\d{2}年", RegexOptions.IgnoreCase))
        {
            return true;
        }

        var chineseCount = value.Count(ch => ch >= '\u4e00' && ch <= '\u9fff');
        if (chineseCount < 2)
        {
            return true;
        }

        return Regex.IsMatch(value, @"^[0-9A-Za-z%+\-*<>.]+$");
    }

    private static bool LooksLikeInvoiceTableHeader(string value)
    {
        return ContainsAny(value, "项目名称", "规格型号", "统一社会信用代码", "纳税人识别号", "下载次数", "税率", "征收率", "单价", "金额", "名称名称",
            "发票代码", "发票号码", "开票日期", "银行账号", "复核人", "开票人", "收款人", "机器编号", "校验码");
    }

    private static string DetectCategory(string text)
    {
        if (ContainsAny(text, "通行费", "高速", "ETC", "收费公路", "过路费")) return "高速/通行费";
        if (ContainsAny(text, "滴滴", "DIDI", "网约车", "出行科技", "小桔")) return "滴滴/网约车";
        if (ContainsAny(text, "铁路", "火车票", "12306", "客票", "行程单")) return "火车/铁路";
        if (ContainsAny(text, "航空", "机票", "航旅", "机场")) return "飞机/机票";
        if (ContainsAny(text, "美团", "三快", "餐饮", "饭店", "餐厅", "饮食")) return "餐饮";
        if (ContainsAny(text, "京东", "淘宝", "天猫", "拼多多", "采购", "五金", "电子", "科技")) return "采购";
        if (ContainsAny(text, "联通", "移动", "电信", "话费", "通信")) return "通信";
        return "其他发票";
    }

    private static string Sha256(string path)
    {
        try
        {
            using var sha = SHA256.Create();
            using var stream = File.OpenRead(path);
            return Convert.ToHexString(sha.ComputeHash(stream)).ToLowerInvariant();
        }
        catch
        {
            return "";
        }
    }

    private static int CommonPrefixLength(string left, string right)
    {
        var max = Math.Min(left.Length, right.Length);
        var i = 0;
        while (i < max && char.ToUpperInvariant(left[i]) == char.ToUpperInvariant(right[i]))
        {
            i++;
        }
        return i;
    }

    private static bool ContainsAny(string value, params string[] keywords)
    {
        return keywords.Any(k => value.Contains(k, StringComparison.OrdinalIgnoreCase));
    }
}
