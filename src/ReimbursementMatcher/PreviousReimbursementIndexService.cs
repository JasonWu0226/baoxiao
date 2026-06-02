using System.Globalization;
using System.Security.Cryptography;
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

        var documents = roots
            .SelectMany(root => EnumerateFilesSafe(root, "*.*")
                .Where(IsPreviousDocument)
                .Select(file => BuildDocument(root, file)))
            .OrderBy(d => d.SourcePath, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var invoices = roots
            .SelectMany(root => EnumerateFilesSafe(root, "*.pdf")
                .Select(file => BuildInvoice(root, file, documents)))
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
            SourceRoots = roots,
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

    private static PreviousReimbursementInvoice BuildInvoice(string root, string path, List<PreviousReimbursementDocument> documents)
    {
        var text = ExtractPdfText(path);
        var combined = $"{Path.GetFileName(path)} {path} {text}";
        return new PreviousReimbursementInvoice
        {
            SourceRoot = root,
            SourcePath = path,
            FileName = Path.GetFileName(path),
            Sha256 = Sha256(path),
            InvoiceNumber = ExtractInvoiceNo(combined),
            Amount = ExtractAmount(combined),
            InvoiceDate = ExtractDate(combined),
            Vendor = ExtractVendor(combined),
            Category = DetectCategory(combined),
            RelatedDocument = FindRelatedDocument(path, documents)
        };
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
        var invoiceHeaders = new[] { "类别", "开票日期", "金额", "发票号", "销售方", "发票文件", "关联报销文档", "SHA256" };
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
            invoices.Cell(r, 6).Value = item.SourcePath;
            invoices.Cell(r, 7).Value = item.RelatedDocument;
            invoices.Cell(r, 8).Value = item.Sha256;
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

    private static bool IsPreviousMaterialDir(string name)
    {
        return ContainsAny(name, "上一期", "上期", "历史", "已报销", "往期", "上次报销", "前期报销");
    }

    private static bool HasAnyUsefulFile(string dir)
    {
        try
        {
            return Directory.EnumerateFiles(dir, "*.pdf", SearchOption.AllDirectories).Any()
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

    private static List<string> ExtractInvoiceNumbers(string text)
    {
        var numbers = new List<string>();
        var labeled = Regex.Matches(text, @"(?:发票号码|发票号|电子发票号码|EIid|InvoiceNumber)[_：:\s-]*([0-9]{8,30})", RegexOptions.IgnoreCase);
        numbers.AddRange(labeled.Select(m => NormalizeInvoiceNo(m.Groups[1].Value)));
        var dzfp = Regex.Matches(text, @"dzfp[_-]([0-9]{12,30})", RegexOptions.IgnoreCase);
        numbers.AddRange(dzfp.Select(m => NormalizeInvoiceNo(m.Groups[1].Value)));
        numbers.AddRange(Regex.Matches(text, @"(?<!\d)([0-9]{18,24})(?!\d)")
            .Select(m => NormalizeInvoiceNo(m.Groups[1].Value)));
        return numbers.Where(v => !string.IsNullOrWhiteSpace(v)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static string ExtractInvoiceNo(string text)
    {
        return ExtractInvoiceNumbers(text).FirstOrDefault() ?? "";
    }

    private static string NormalizeInvoiceNo(string value)
    {
        if (value.Length > 20 && !value.StartsWith("20", StringComparison.Ordinal))
        {
            value = value[..20];
        }
        return value;
    }

    private static string ExtractDate(string text)
    {
        var match = Regex.Match(text, @"(20\d{2})[-_.年/](\d{1,2})[-_.月/](\d{1,2})");
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
            @"价税合计.{0,80}?[小写）\)]\s*[¥￥]?\s*([0-9]+(?:\.[0-9]{1,2})?)",
            @"小写\s*[）\)]?\s*[¥￥]?\s*([0-9]+(?:\.[0-9]{1,2})?)",
            @"发票金额\s*[：:】\]\s]*[¥￥]?\s*([0-9]+(?:\.[0-9]{1,2})?)",
            @"金额\s*[：:】\]\s]*[¥￥]?\s*([0-9]+(?:\.[0-9]{1,2})?)",
            @"[¥￥]\s*([0-9]+(?:\.[0-9]{1,2})?)"
        };

        foreach (var pattern in patterns)
        {
            var match = Regex.Match(text, pattern, RegexOptions.Singleline);
            if (match.Success && decimal.TryParse(match.Groups[1].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var amount))
            {
                return amount;
            }
        }
        return 0m;
    }

    private static string ExtractVendor(string text)
    {
        var seller = Regex.Match(text, @"销售方信息.{0,80}?名称[:：]\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,40})");
        if (seller.Success) return seller.Groups[1].Value.Trim();

        var name = Path.GetFileNameWithoutExtension(text);
        var parts = Regex.Split(name, @"[_\s]+").Where(p => p.Length >= 4).ToList();
        return parts.FirstOrDefault(p => p.Any(ch => ch >= '\u4e00' && ch <= '\u9fff')) ?? "";
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
