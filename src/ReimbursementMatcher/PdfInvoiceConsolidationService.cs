using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using UglyToad.PdfPig;

namespace ReimbursementMatcher;

public sealed class PdfInvoiceConsolidationService
{
    private readonly WorkspaceService _workspace;

    public PdfInvoiceConsolidationService(WorkspaceService workspace)
    {
        _workspace = workspace;
    }

    public string Generate(AppConfig config)
    {
        var invoiceDir = _workspace.Resolve(config.InvoiceDir);
        if (!Directory.Exists(invoiceDir))
        {
            throw new DirectoryNotFoundException("发票目录不存在：" + invoiceDir);
        }

        var outputRoot = _workspace.Resolve(Path.Combine(config.OutputDir, $"PDF发票归集_{DateTime.Now:yyyyMMdd_HHmmss}"));
        var pdfRoot = Path.Combine(outputRoot, "PDF发票");
        Directory.CreateDirectory(pdfRoot);

        var rows = Directory.EnumerateFiles(invoiceDir, "*.pdf", SearchOption.AllDirectories)
            .Where(f => !IsInsideOutputOrArchive(f, invoiceDir))
            .Select(ReadPdf)
            .OrderBy(r => r.InvoiceDate)
            .ThenBy(r => r.Category)
            .ThenBy(r => r.InvoiceNumber)
            .ThenBy(r => r.SourcePath, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var unique = new List<PdfInvoiceRow>();
        var duplicates = new List<PdfInvoiceRow>();
        foreach (var group in rows.GroupBy(DedupeKey, StringComparer.OrdinalIgnoreCase))
        {
            var candidates = group.ToList();
            var primary = candidates
                .OrderByDescending(r => r.InvoiceNumber.Length)
                .ThenByDescending(r => r.Amount)
                .ThenByDescending(r => r.TextLength)
                .ThenBy(r => r.SourcePath, StringComparer.OrdinalIgnoreCase)
                .First();

            primary.CopiedPath = CopyUniquePdf(pdfRoot, primary);
            unique.Add(primary);

            foreach (var duplicate in candidates.Where(r => !ReferenceEquals(r, primary)))
            {
                duplicate.DuplicateOf = primary.CopiedPath;
                duplicate.ReviewFlag = string.IsNullOrWhiteSpace(duplicate.InvoiceNumber)
                    ? "相同PDF哈希重复"
                    : "同一发票号重复";
                duplicates.Add(duplicate);
            }
        }

        var previous = BuildPreviousIndex(config, invoiceDir);
        foreach (var row in unique)
        {
            row.ReviewFlag = BuildReviewFlag(row);
            row.MergeGroup = BuildMergeGroup(row);
            row.PreviousMatch = FindPreviousMatch(row, previous);
            if (!string.IsNullOrWhiteSpace(row.PreviousMatch.MatchType))
            {
                row.ReviewFlag = MergeFlag(row.ReviewFlag, "疑似上期已报销");
            }
        }

        var workbookPath = Path.Combine(outputRoot, "PDF发票归集核验.xlsx");
        WriteWorkbook(workbookPath, unique, duplicates, outputRoot);
        return workbookPath;
    }

    private static PdfInvoiceRow ReadPdf(string path)
    {
        var text = ExtractPdfText(path);
        var combined = $"{Path.GetFileName(path)} {path} {text}";
        var row = new PdfInvoiceRow
        {
            SourcePath = path,
            FileName = Path.GetFileName(path),
            Sha256 = Sha256(path),
            TextLength = text.Length,
            InvoiceNumber = ExtractInvoiceNo(combined),
            Amount = ExtractAmount(combined),
            InvoiceDate = ExtractDate(combined),
            Vendor = ExtractVendor(combined),
            Category = DetectCategory(combined)
        };
        return row;
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

    private static string DedupeKey(PdfInvoiceRow row)
    {
        return !string.IsNullOrWhiteSpace(row.InvoiceNumber)
            ? $"no:{row.InvoiceNumber}"
            : $"sha:{row.Sha256}";
    }

    private static string CopyUniquePdf(string pdfRoot, PdfInvoiceRow row)
    {
        var fileName = SafeFileName(string.Join("_", new[]
        {
            string.IsNullOrWhiteSpace(row.InvoiceDate) ? "未识别日期" : row.InvoiceDate.Replace("-", ""),
            row.Category,
            row.Amount > 0 ? row.Amount.ToString("0.##", CultureInfo.InvariantCulture) : "金额未识别",
            string.IsNullOrWhiteSpace(row.InvoiceNumber) ? row.Sha256[..12] : row.InvoiceNumber,
            Shorten(row.Vendor, 24)
        }.Where(v => !string.IsNullOrWhiteSpace(v)))) + ".pdf";

        var target = UniquePath(Path.Combine(pdfRoot, fileName));
        File.Copy(row.SourcePath, target, overwrite: false);
        return target;
    }

    private List<PreviousInvoiceRow> BuildPreviousIndex(AppConfig config, string currentInvoiceDir)
    {
        var roots = config.PreviousInvoiceDirs
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Select(_workspace.Resolve)
            .Where(Directory.Exists)
            .Where(dir => !Path.GetFullPath(dir).Equals(Path.GetFullPath(currentInvoiceDir), StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return roots
            .SelectMany(root => Directory.EnumerateFiles(root, "*.pdf", SearchOption.AllDirectories)
                .Select(file => new { Root = root, File = file }))
            .Select(x =>
            {
                var row = ReadPdf(x.File);
                return new PreviousInvoiceRow
                {
                    Root = x.Root,
                    SourcePath = row.SourcePath,
                    Sha256 = row.Sha256,
                    InvoiceNumber = row.InvoiceNumber,
                    Amount = row.Amount,
                    InvoiceDate = row.InvoiceDate,
                    Vendor = row.Vendor,
                    Category = row.Category
                };
            })
            .ToList();
    }

    private static PreviousMatchResult FindPreviousMatch(PdfInvoiceRow row, List<PreviousInvoiceRow> previous)
    {
        if (previous.Count == 0)
        {
            return new PreviousMatchResult();
        }

        if (!string.IsNullOrWhiteSpace(row.InvoiceNumber))
        {
            var exactNo = previous.FirstOrDefault(p => p.InvoiceNumber.Equals(row.InvoiceNumber, StringComparison.OrdinalIgnoreCase));
            if (exactNo != null)
            {
                return PreviousMatchResult.From("发票号一致", exactNo);
            }
        }

        var exactHash = previous.FirstOrDefault(p => !string.IsNullOrWhiteSpace(row.Sha256) && p.Sha256 == row.Sha256);
        if (exactHash != null)
        {
            return PreviousMatchResult.From("文件哈希一致", exactHash);
        }

        if (row.Amount > 0 && !string.IsNullOrWhiteSpace(row.InvoiceDate))
        {
            var fuzzy = previous.FirstOrDefault(p =>
                p.Amount == row.Amount
                && p.InvoiceDate == row.InvoiceDate
                && !string.IsNullOrWhiteSpace(row.Vendor)
                && !string.IsNullOrWhiteSpace(p.Vendor)
                && (row.Vendor.Contains(p.Vendor, StringComparison.OrdinalIgnoreCase)
                    || p.Vendor.Contains(row.Vendor, StringComparison.OrdinalIgnoreCase)));
            if (fuzzy != null)
            {
                return PreviousMatchResult.From("日期金额销售方疑似一致", fuzzy);
            }
        }

        return new PreviousMatchResult();
    }

    private static void WriteWorkbook(string path, List<PdfInvoiceRow> unique, List<PdfInvoiceRow> duplicates, string outputRoot)
    {
        using var wb = new XLWorkbook();
        WriteSummarySheet(wb, unique, duplicates, outputRoot);
        WriteInvoiceSheet(wb, "唯一PDF发票", unique);
        WriteInvoiceSheet(wb, "待补充核验", unique.Where(r => !string.IsNullOrWhiteSpace(r.ReviewFlag)).ToList());
        WritePreviousOverlapSheet(wb, unique.Where(r => !string.IsNullOrWhiteSpace(r.PreviousMatch.MatchType)).ToList());
        WriteDuplicateSheet(wb, duplicates);
        WriteMergeSheet(wb, BuildMergeGroups(unique));
        wb.SaveAs(path);
    }

    private static void WriteSummarySheet(XLWorkbook wb, List<PdfInvoiceRow> unique, List<PdfInvoiceRow> duplicates, string outputRoot)
    {
        var ws = wb.Worksheets.Add("总览");
        var metrics = new (string Name, object Value)[]
        {
            ("PDF归集目录", Path.Combine(outputRoot, "PDF发票")),
            ("唯一PDF发票数量", unique.Count),
            ("重复PDF数量", duplicates.Count),
            ("唯一PDF金额合计", unique.Sum(r => r.Amount)),
            ("待补充核验数量", unique.Count(r => !string.IsNullOrWhiteSpace(r.ReviewFlag))),
            ("疑似上期已报销数量", unique.Count(r => !string.IsNullOrWhiteSpace(r.PreviousMatch.MatchType))),
            ("疑似上期已报销金额", unique.Where(r => !string.IsNullOrWhiteSpace(r.PreviousMatch.MatchType)).Sum(r => r.Amount)),
            ("高速/通行费数量", unique.Count(r => r.Category == "高速/通行费")),
            ("高速/通行费金额", unique.Where(r => r.Category == "高速/通行费").Sum(r => r.Amount)),
            ("滴滴/网约车数量", unique.Count(r => r.Category == "滴滴/网约车")),
            ("滴滴/网约车金额", unique.Where(r => r.Category == "滴滴/网约车").Sum(r => r.Amount))
        };

        ws.Cell(1, 1).Value = "指标";
        ws.Cell(1, 2).Value = "值";
        ws.Range(1, 1, 1, 2).Style.Font.Bold = true;
        for (var i = 0; i < metrics.Length; i++)
        {
            ws.Cell(i + 2, 1).Value = metrics[i].Name;
            if (metrics[i].Value is decimal amount)
            {
                ws.Cell(i + 2, 2).Value = amount;
                ws.Cell(i + 2, 2).Style.NumberFormat.Format = "¥#,##0.00";
            }
            else
            {
                ws.Cell(i + 2, 2).Value = metrics[i].Value?.ToString() ?? "";
            }
        }
        ws.Columns().AdjustToContents(8, 80);
    }

    private static void WriteInvoiceSheet(XLWorkbook wb, string name, List<PdfInvoiceRow> rows)
    {
        var ws = wb.Worksheets.Add(name);
        var headers = new[]
        {
            "类别", "合并组", "开票日期", "金额", "发票号", "销售方/抬头识别", "核验提示",
            "上期匹配类型", "上期文件", "归集PDF", "原始PDF", "SHA256"
        };
        WriteHeaders(ws, headers);
        var r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.Category;
            ws.Cell(r, 2).Value = row.MergeGroup;
            ws.Cell(r, 3).Value = row.InvoiceDate;
            ws.Cell(r, 4).Value = row.Amount;
            ws.Cell(r, 4).Style.NumberFormat.Format = "¥#,##0.00";
            ws.Cell(r, 5).Value = row.InvoiceNumber;
            ws.Cell(r, 6).Value = row.Vendor;
            ws.Cell(r, 7).Value = row.ReviewFlag;
            ws.Cell(r, 8).Value = row.PreviousMatch.MatchType;
            ws.Cell(r, 9).Value = row.PreviousMatch.SourcePath;
            ws.Cell(r, 10).Value = row.CopiedPath;
            ws.Cell(r, 11).Value = row.SourcePath;
            ws.Cell(r, 12).Value = row.Sha256;
            if (!string.IsNullOrWhiteSpace(row.PreviousMatch.MatchType))
            {
                ws.Range(r, 1, r, headers.Length).Style.Fill.BackgroundColor = XLColor.FromHtml("#F4CCCC");
            }
            else if (!string.IsNullOrWhiteSpace(row.ReviewFlag))
            {
                ws.Range(r, 1, r, headers.Length).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF2CC");
            }
            r++;
        }
        FormatSheet(ws, headers.Length);
    }

    private static void WritePreviousOverlapSheet(XLWorkbook wb, List<PdfInvoiceRow> rows)
    {
        var ws = wb.Worksheets.Add("上期重叠核验");
        var headers = new[]
        {
            "匹配类型", "本期类别", "本期开票日期", "本期金额", "本期发票号", "本期销售方",
            "上期开票日期", "上期金额", "上期发票号", "上期销售方", "本期归集PDF", "上期文件"
        };
        WriteHeaders(ws, headers);
        var r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.PreviousMatch.MatchType;
            ws.Cell(r, 2).Value = row.Category;
            ws.Cell(r, 3).Value = row.InvoiceDate;
            ws.Cell(r, 4).Value = row.Amount;
            ws.Cell(r, 4).Style.NumberFormat.Format = "¥#,##0.00";
            ws.Cell(r, 5).Value = row.InvoiceNumber;
            ws.Cell(r, 6).Value = row.Vendor;
            ws.Cell(r, 7).Value = row.PreviousMatch.InvoiceDate;
            ws.Cell(r, 8).Value = row.PreviousMatch.Amount;
            ws.Cell(r, 8).Style.NumberFormat.Format = "¥#,##0.00";
            ws.Cell(r, 9).Value = row.PreviousMatch.InvoiceNumber;
            ws.Cell(r, 10).Value = row.PreviousMatch.Vendor;
            ws.Cell(r, 11).Value = row.CopiedPath;
            ws.Cell(r, 12).Value = row.PreviousMatch.SourcePath;
            ws.Range(r, 1, r, headers.Length).Style.Fill.BackgroundColor = XLColor.FromHtml("#F4CCCC");
            r++;
        }
        FormatSheet(ws, headers.Length);
    }

    private static void WriteDuplicateSheet(XLWorkbook wb, List<PdfInvoiceRow> rows)
    {
        var ws = wb.Worksheets.Add("重复PDF");
        var headers = new[] { "重复原因", "开票日期", "金额", "发票号", "重复文件", "保留文件", "SHA256" };
        WriteHeaders(ws, headers);
        var r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.ReviewFlag;
            ws.Cell(r, 2).Value = row.InvoiceDate;
            ws.Cell(r, 3).Value = row.Amount;
            ws.Cell(r, 3).Style.NumberFormat.Format = "¥#,##0.00";
            ws.Cell(r, 4).Value = row.InvoiceNumber;
            ws.Cell(r, 5).Value = row.SourcePath;
            ws.Cell(r, 6).Value = row.DuplicateOf;
            ws.Cell(r, 7).Value = row.Sha256;
            r++;
        }
        FormatSheet(ws, headers.Length);
    }

    private static void WriteMergeSheet(XLWorkbook wb, List<MergeGroupRow> rows)
    {
        var ws = wb.Worksheets.Add("合并报销建议");
        var headers = new[] { "建议合并组", "类别", "月份", "发票数量", "金额合计", "日期范围", "建议报销摘要", "归集PDF清单" };
        WriteHeaders(ws, headers);
        var r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.GroupName;
            ws.Cell(r, 2).Value = row.Category;
            ws.Cell(r, 3).Value = row.Month;
            ws.Cell(r, 4).Value = row.Count;
            ws.Cell(r, 5).Value = row.Amount;
            ws.Cell(r, 5).Style.NumberFormat.Format = "¥#,##0.00";
            ws.Cell(r, 6).Value = row.DateRange;
            ws.Cell(r, 7).Value = row.Summary;
            ws.Cell(r, 8).Value = row.Files;
            r++;
        }
        FormatSheet(ws, headers.Length);
    }

    private static List<MergeGroupRow> BuildMergeGroups(List<PdfInvoiceRow> rows)
    {
        return rows
            .Where(r => r.Category is "高速/通行费" or "滴滴/网约车")
            .GroupBy(r => new { r.Category, Month = InvoiceMonth(r.InvoiceDate) })
            .Select(g =>
            {
                var list = g.OrderBy(r => r.InvoiceDate).ThenBy(r => r.InvoiceNumber).ToList();
                var minDate = list.Select(r => r.InvoiceDate).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "";
                var maxDate = list.Select(r => r.InvoiceDate).LastOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "";
                return new MergeGroupRow
                {
                    Category = g.Key.Category,
                    Month = g.Key.Month,
                    GroupName = $"{g.Key.Month}_{g.Key.Category}",
                    Count = list.Count,
                    Amount = list.Sum(r => r.Amount),
                    DateRange = string.IsNullOrWhiteSpace(minDate) ? "" : $"{minDate} 至 {maxDate}",
                    Summary = g.Key.Category == "高速/通行费"
                        ? $"{g.Key.Month}高速通行费汇总"
                        : $"{g.Key.Month}滴滴/网约车出行费汇总",
                    Files = string.Join(Environment.NewLine, list.Select(r => r.CopiedPath))
                };
            })
            .OrderBy(r => r.Month)
            .ThenBy(r => r.Category)
            .ToList();
    }

    private static string BuildReviewFlag(PdfInvoiceRow row)
    {
        var flags = new List<string>();
        if (string.IsNullOrWhiteSpace(row.InvoiceNumber)) flags.Add("发票号未识别");
        if (row.Amount <= 0) flags.Add("金额未识别");
        if (string.IsNullOrWhiteSpace(row.InvoiceDate)) flags.Add("开票日期未识别");
        if (row.Category == "其他发票") flags.Add("类别待确认");
        return string.Join("；", flags);
    }

    private static string MergeFlag(string left, string right)
    {
        if (string.IsNullOrWhiteSpace(left)) return right;
        return left.Contains(right, StringComparison.OrdinalIgnoreCase) ? left : $"{left}；{right}";
    }

    private static string BuildMergeGroup(PdfInvoiceRow row)
    {
        if (row.Category is "高速/通行费" or "滴滴/网约车")
        {
            return $"{InvoiceMonth(row.InvoiceDate)}_{row.Category}";
        }
        return "";
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

    private static string ExtractInvoiceNo(string text)
    {
        var labeled = Regex.Match(text, @"(?:发票号码|发票号|电子发票号码|EIid|InvoiceNumber)[_：:\s-]*([0-9]{8,30})", RegexOptions.IgnoreCase);
        if (labeled.Success) return NormalizeInvoiceNo(labeled.Groups[1].Value);

        var dzfp = Regex.Match(text, @"dzfp[_-]([0-9]{12,30})", RegexOptions.IgnoreCase);
        if (dzfp.Success) return NormalizeInvoiceNo(dzfp.Groups[1].Value);

        return Regex.Matches(text, @"(?<!\d)([0-9]{18,24})(?!\d)")
            .Select(m => NormalizeInvoiceNo(m.Groups[1].Value))
            .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "";
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

    private static string InvoiceMonth(string date)
    {
        return DateTime.TryParse(date, out var parsed) ? parsed.ToString("yyyy-MM") : "日期未识别";
    }

    private static string NormalizeInvoiceNo(string value)
    {
        if (value.Length > 20 && !value.StartsWith("20", StringComparison.Ordinal))
        {
            value = value[..20];
        }
        return value;
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

    private static bool IsInsideOutputOrArchive(string file, string invoiceDir)
    {
        var relative = Path.GetRelativePath(invoiceDir, file).Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
        return relative.Contains($"{Path.DirectorySeparatorChar}非PDF重复格式_", StringComparison.OrdinalIgnoreCase)
            || relative.Contains($"{Path.DirectorySeparatorChar}非PDF图片清理_", StringComparison.OrdinalIgnoreCase)
            || relative.StartsWith("PDF发票归集_", StringComparison.OrdinalIgnoreCase);
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

    private static bool ContainsAny(string value, params string[] keywords)
    {
        return keywords.Any(k => value.Contains(k, StringComparison.OrdinalIgnoreCase));
    }

    private static string SafeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(name.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
        cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim(' ', '.');
        return string.IsNullOrWhiteSpace(cleaned) ? "file" : cleaned;
    }

    private static string Shorten(string value, int max)
    {
        value = value.Trim();
        return value.Length <= max ? value : value[..max];
    }

    private static string UniquePath(string path)
    {
        if (!File.Exists(path)) return path;
        var dir = Path.GetDirectoryName(path)!;
        var stem = Path.GetFileNameWithoutExtension(path);
        var ext = Path.GetExtension(path);
        for (var i = 1; i < 10000; i++)
        {
            var candidate = Path.Combine(dir, $"{stem}_{i}{ext}");
            if (!File.Exists(candidate)) return candidate;
        }
        throw new IOException("无法生成不重复文件名：" + path);
    }

    private sealed class PdfInvoiceRow
    {
        public string SourcePath { get; set; } = "";
        public string FileName { get; set; } = "";
        public string CopiedPath { get; set; } = "";
        public string DuplicateOf { get; set; } = "";
        public string Sha256 { get; set; } = "";
        public string InvoiceNumber { get; set; } = "";
        public decimal Amount { get; set; }
        public string InvoiceDate { get; set; } = "";
        public string Vendor { get; set; } = "";
        public string Category { get; set; } = "";
        public string MergeGroup { get; set; } = "";
        public string ReviewFlag { get; set; } = "";
        public PreviousMatchResult PreviousMatch { get; set; } = new();
        public int TextLength { get; set; }
    }

    private sealed class PreviousInvoiceRow
    {
        public string Root { get; set; } = "";
        public string SourcePath { get; set; } = "";
        public string Sha256 { get; set; } = "";
        public string InvoiceNumber { get; set; } = "";
        public decimal Amount { get; set; }
        public string InvoiceDate { get; set; } = "";
        public string Vendor { get; set; } = "";
        public string Category { get; set; } = "";
    }

    private sealed class PreviousMatchResult
    {
        public string MatchType { get; set; } = "";
        public string SourcePath { get; set; } = "";
        public string InvoiceNumber { get; set; } = "";
        public decimal Amount { get; set; }
        public string InvoiceDate { get; set; } = "";
        public string Vendor { get; set; } = "";

        public static PreviousMatchResult From(string matchType, PreviousInvoiceRow row) => new()
        {
            MatchType = matchType,
            SourcePath = row.SourcePath,
            InvoiceNumber = row.InvoiceNumber,
            Amount = row.Amount,
            InvoiceDate = row.InvoiceDate,
            Vendor = row.Vendor
        };
    }

    private sealed class MergeGroupRow
    {
        public string GroupName { get; set; } = "";
        public string Category { get; set; } = "";
        public string Month { get; set; } = "";
        public int Count { get; set; }
        public decimal Amount { get; set; }
        public string DateRange { get; set; } = "";
        public string Summary { get; set; } = "";
        public string Files { get; set; } = "";
    }
}
