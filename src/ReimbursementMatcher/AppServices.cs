using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ClosedXML.Excel;
using UglyToad.PdfPig;

namespace ReimbursementMatcher;

public sealed class WorkspaceService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    public WorkspaceService()
    {
        RootDir = FindRootDir();
        ConfigPath = Path.Combine(RootDir, "matcher_config.json");
    }

    public string RootDir { get; }
    public string ConfigPath { get; }

    public AppConfig LoadConfig()
    {
        if (!File.Exists(ConfigPath))
        {
            return new AppConfig();
        }

        var json = File.ReadAllText(ConfigPath, Encoding.UTF8);
        return JsonSerializer.Deserialize<AppConfig>(json, JsonOptions) ?? new AppConfig();
    }

    public void SaveConfig(AppConfig config)
    {
        File.WriteAllText(ConfigPath, JsonSerializer.Serialize(config, JsonOptions) + Environment.NewLine, Encoding.UTF8);
    }

    public string Resolve(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return "";
        }
        return Path.IsPathRooted(path) ? path : Path.GetFullPath(Path.Combine(RootDir, path));
    }

    public string ToRelative(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return "";
        }
        var full = Path.GetFullPath(path);
        var root = Path.GetFullPath(RootDir);
        if (!root.EndsWith(Path.DirectorySeparatorChar))
        {
            root += Path.DirectorySeparatorChar;
        }
        return full.StartsWith(root, StringComparison.OrdinalIgnoreCase)
            ? Path.GetRelativePath(RootDir, full)
            : full;
    }

    private static string FindRootDir()
    {
        var dir = new DirectoryInfo(Directory.GetCurrentDirectory());
        while (dir != null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "README.md")) && Directory.Exists(Path.Combine(dir.FullName, "报销准备的资料")))
            {
                return dir.FullName;
            }
            dir = dir.Parent;
        }
        return Directory.GetCurrentDirectory();
    }
}

public sealed class ConfirmationService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private readonly WorkspaceService _workspace;

    public ConfirmationService(WorkspaceService workspace)
    {
        _workspace = workspace;
    }

    public string ConfirmationPath(AppConfig config)
    {
        var dir = _workspace.Resolve(config.RuleDir);
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "确认记录.json");
    }

    public string WeightPath(AppConfig config)
    {
        var dir = _workspace.Resolve(config.RuleDir);
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "判断权重.json");
    }

    public ConfirmationStore Load(AppConfig config)
    {
        var path = ConfirmationPath(config);
        if (!File.Exists(path))
        {
            return new ConfirmationStore();
        }
        var json = File.ReadAllText(path, Encoding.UTF8);
        return JsonSerializer.Deserialize<ConfirmationStore>(json, JsonOptions) ?? new ConfirmationStore();
    }

    public void Save(AppConfig config, ConfirmationStore store)
    {
        store.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        File.WriteAllText(ConfirmationPath(config), JsonSerializer.Serialize(store, JsonOptions) + Environment.NewLine, Encoding.UTF8);
        SaveWeights(config, BuildWeights(store));
    }

    public void Apply(ConfirmationStore store, List<EvidenceItem> items, List<MatchCandidate> matches)
    {
        foreach (var item in items)
        {
            if (!store.Items.TryGetValue(item.Id, out var confirmation))
            {
                continue;
            }
            item.ReimburseDecision = confirmation.ReimburseDecision;
            item.FileDecision = confirmation.FileDecision;
            item.Project = confirmation.Project;
            item.Note = confirmation.Note;
        }

        foreach (var match in matches)
        {
            if (store.Matches.TryGetValue(match.Id, out var confirmation))
            {
                match.Decision = confirmation.Decision;
            }
        }

        var weights = BuildWeights(store).Weights
            .ToDictionary(w => $"{w.Platform}|{w.Key}", StringComparer.OrdinalIgnoreCase);
        foreach (var item in items)
        {
            var key = $"{item.Platform}|{NormalizeKey(item.Title)}";
            if (!weights.TryGetValue(key, out var weight))
            {
                continue;
            }
            item.Weight = weight.Weight;
            item.Suggestion = weight.Suggestion;
            if (item.ReimburseDecision == "待确认" && Math.Abs(weight.Weight) >= 2)
            {
                item.ReimburseDecision = weight.Weight > 0 ? "建议报销" : "建议不报销";
            }
        }
    }

    public void ConfirmItem(AppConfig config, ConfirmationStore store, EvidenceItem item, string? reimburseDecision, string? fileDecision, string note)
    {
        store.Items.TryGetValue(item.Id, out var old);
        var before = old == null ? "" : $"{old.ReimburseDecision}/{old.FileDecision}";

        var confirmation = old ?? new ItemConfirmation
        {
            ItemId = item.Id,
            Kind = item.Kind,
            Platform = item.Platform,
            Title = item.Title
        };
        if (!string.IsNullOrWhiteSpace(reimburseDecision))
        {
            confirmation.ReimburseDecision = reimburseDecision;
            item.ReimburseDecision = reimburseDecision;
        }
        if (!string.IsNullOrWhiteSpace(fileDecision))
        {
            confirmation.FileDecision = fileDecision;
            item.FileDecision = fileDecision;
        }
        if (!string.IsNullOrWhiteSpace(note))
        {
            confirmation.Note = note;
            item.Note = note;
        }
        confirmation.Project = item.Project;
        confirmation.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        store.Items[item.Id] = confirmation;

        store.Events.Add(new ConfirmationEvent
        {
            Time = confirmation.UpdatedAt,
            ObjectType = "素材",
            ObjectId = item.Id,
            Action = "确认素材",
            Before = before,
            After = $"{confirmation.ReimburseDecision}/{confirmation.FileDecision}",
            Note = note
        });
        Save(config, store);
    }

    public void ConfirmMatch(AppConfig config, ConfirmationStore store, MatchCandidate match, string decision, string note)
    {
        store.Matches.TryGetValue(match.Id, out var old);
        var before = old?.Decision ?? "";
        var confirmation = old ?? new MatchConfirmation
        {
            MatchId = match.Id,
            ExpenseId = match.ExpenseId,
            InvoiceId = match.InvoiceId
        };
        confirmation.Decision = decision;
        confirmation.Note = note;
        confirmation.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        store.Matches[match.Id] = confirmation;
        match.Decision = decision;

        store.Events.Add(new ConfirmationEvent
        {
            Time = confirmation.UpdatedAt,
            ObjectType = "匹配",
            ObjectId = match.Id,
            Action = "确认匹配",
            Before = before,
            After = decision,
            Note = note
        });
        Save(config, store);
    }

    private void SaveWeights(AppConfig config, DecisionWeightStore weights)
    {
        File.WriteAllText(WeightPath(config), JsonSerializer.Serialize(weights, JsonOptions) + Environment.NewLine, Encoding.UTF8);
    }

    private static DecisionWeightStore BuildWeights(ConfirmationStore store)
    {
        var result = new DecisionWeightStore { UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") };
        result.Weights = store.Items.Values
            .Where(i => !string.IsNullOrWhiteSpace(i.Title))
            .GroupBy(i => $"{i.Platform}|{NormalizeKey(i.Title)}", StringComparer.OrdinalIgnoreCase)
            .Select(g =>
            {
                var rows = g.ToList();
                var reimburse = rows.Count(r => r.ReimburseDecision is "报销" or "建议报销");
                var exclude = rows.Count(r => r.ReimburseDecision is "不报销" or "建议不报销");
                var pending = rows.Count(r => r.ReimburseDecision is "待确认");
                var first = rows[0];
                var weight = reimburse - exclude;
                return new DecisionWeight
                {
                    Platform = first.Platform,
                    Key = NormalizeKey(first.Title),
                    DisplayName = first.Title,
                    ReimburseCount = reimburse,
                    ExcludeCount = exclude,
                    PendingCount = pending,
                    Weight = weight,
                    Suggestion = weight > 0 ? "倾向报销" : weight < 0 ? "倾向不报销" : "需要复核"
                };
            })
            .OrderByDescending(w => Math.Abs(w.Weight))
            .ThenBy(w => w.Platform)
            .ThenBy(w => w.DisplayName)
            .ToList();
        return result;
    }

    public static string NormalizeKey(string value)
    {
        var builder = new StringBuilder();
        foreach (var ch in (value ?? "").ToLowerInvariant())
        {
            if (char.IsLetterOrDigit(ch) || (ch >= '\u4e00' && ch <= '\u9fff'))
            {
                builder.Append(ch);
            }
        }
        return builder.ToString();
    }
}

public sealed class MaterialScanner
{
    private readonly WorkspaceService _workspace;
    private static readonly HashSet<string> ExtractableInvoiceExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".pdf", ".ofd", ".xml"
    };

    public MaterialScanner(WorkspaceService workspace)
    {
        _workspace = workspace;
    }

    public List<EvidenceItem> Scan(AppConfig config)
    {
        var roots = new[]
            {
                _workspace.Resolve(config.SourceRoot),
                _workspace.Resolve(config.InvoiceDir),
                _workspace.Resolve(config.Email.OutputDir)
            }
            .Where(Directory.Exists)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var result = new List<EvidenceItem>();
        foreach (var root in roots)
        {
            ExtractZipInvoices(root);
            var files = Directory.EnumerateFiles(root, "*.*", SearchOption.AllDirectories)
                .Where(IsSupportedFile)
                .Where(f => !IsSkippedArchiveFile(f))
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToList();
            result.AddRange(files.Select(f => CreateItem(root, f)));
        }
        return result
            .GroupBy(i => i.FilePath, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .Pipe(DeduplicateInvoices)
            .ToList();
    }

    private EvidenceItem CreateItem(string sourceRoot, string file)
    {
        var relative = Path.GetRelativePath(sourceRoot, file);
        var text = $"{relative} {Path.GetFileNameWithoutExtension(file)}";
        var meta = ReadInvoiceMeta(file);
        var kind = DetectKind(file, text);
        var platform = DetectPlatform(text);
        var amount = meta.Amount > 0 ? meta.Amount : ExtractAmount(text);
        var date = !string.IsNullOrWhiteSpace(meta.Date) ? meta.Date : ExtractDate(text);
        var invoiceNo = !string.IsNullOrWhiteSpace(meta.InvoiceNumber) ? meta.InvoiceNumber : ExtractInvoiceNo(text);
        var title = BuildTitle(kind, platform, text, file);
        var vendor = !string.IsNullOrWhiteSpace(meta.Vendor) ? meta.Vendor : ExtractVendor(text);

        return new EvidenceItem
        {
            Id = StableId(relative),
            Kind = kind,
            Platform = platform,
            Date = date,
            Amount = amount,
            Title = title,
            Vendor = vendor,
            InvoiceNumber = invoiceNo,
            FilePath = file,
            RelativePath = relative,
            FileDecision = kind == EvidenceKinds.Template ? "系统文件" : "待确认",
            ReimburseDecision = kind == EvidenceKinds.Invoice || kind == EvidenceKinds.Template ? "不适用" : "待确认",
            Project = GuessProject(title),
            ExtractedText = text,
            FileHash = Sha256(file)
        };
    }

    private static bool IsSupportedFile(string file)
    {
        var ext = Path.GetExtension(file).ToLowerInvariant();
        return ext is ".pdf" or ".ofd" or ".xml" or ".zip" or ".jpg" or ".jpeg" or ".png" or ".webp" or ".xlsx" or ".csv" or ".json";
    }

    private static bool IsSkippedArchiveFile(string file)
    {
        var ext = Path.GetExtension(file).ToLowerInvariant();
        if (ext == ".zip")
        {
            return true;
        }
        var normalized = file.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
        return normalized.Contains($"{Path.DirectorySeparatorChar}非PDF重复格式_", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains($"{Path.DirectorySeparatorChar}日期不符发票_", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains($"{Path.DirectorySeparatorChar}_日期不符发票{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase);
    }

    private static void ExtractZipInvoices(string root)
    {
        var cacheRoot = Path.Combine(root, "_zip解压缓存");
        foreach (var zip in Directory.EnumerateFiles(root, "*.zip", SearchOption.AllDirectories).ToList())
        {
            var normalized = zip.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
            if (normalized.Contains($"{Path.DirectorySeparatorChar}_zip解压缓存{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase)
                || normalized.Contains($"{Path.DirectorySeparatorChar}非PDF重复格式_", StringComparison.OrdinalIgnoreCase)
                || normalized.Contains($"{Path.DirectorySeparatorChar}日期不符发票_", StringComparison.OrdinalIgnoreCase)
                || normalized.Contains($"{Path.DirectorySeparatorChar}_日期不符发票{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            try
            {
                var zipKey = $"zip_{Sha256(zip)[..12]}";
                var targetDir = Path.Combine(cacheRoot, SafePathName(zipKey));
                using var archive = ZipFile.OpenRead(zip);
                var entries = archive.Entries
                    .Where(e => !string.IsNullOrWhiteSpace(e.Name))
                    .Where(e => ExtractableInvoiceExtensions.Contains(Path.GetExtension(e.Name)))
                    .ToList();
                if (entries.Count == 0)
                {
                    continue;
                }

                if (entries.Any(e => Path.GetExtension(e.Name).Equals(".pdf", StringComparison.OrdinalIgnoreCase)))
                {
                    entries = entries
                        .Where(e => Path.GetExtension(e.Name).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                        .ToList();
                }

                Directory.CreateDirectory(targetDir);
                foreach (var entry in entries)
                {
                    var target = Path.Combine(targetDir, SafePathName(entry.Name));
                    if (File.Exists(target))
                    {
                        continue;
                    }
                    entry.ExtractToFile(target, overwrite: false);
                }
            }
            catch
            {
                // Keep scanning the rest of the folder even if one zip is broken or encrypted.
            }
        }
    }

    private static string SafePathName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(name.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
        cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim(' ', '.');
        return string.IsNullOrWhiteSpace(cleaned) ? "file" : cleaned;
    }

    private static string DetectKind(string file, string text)
    {
        var ext = Path.GetExtension(file).ToLowerInvariant();
        var lower = text.ToLowerInvariant();
        if (text.Contains("报销模版") || text.Contains("模板"))
        {
            return EvidenceKinds.Template;
        }
        if (ext is ".jpg" or ".jpeg" or ".png" or ".webp"
            && (lower.Contains("qrcode") || lower.Contains("qr_") || lower.Contains("logo") || lower.Contains("header") || lower.Contains("ads") || lower.Contains("wx_txf")))
        {
            return EvidenceKinds.Other;
        }
        if (text.Contains("发票") || text.Contains("invoice", StringComparison.OrdinalIgnoreCase) || text.Contains("dzfp") || ext is ".ofd" or ".xml")
        {
            return EvidenceKinds.Invoice;
        }
        if (text.Contains("付款截图") || text.Contains("截图") || ext is ".jpg" or ".jpeg" or ".png" or ".webp")
        {
            return EvidenceKinds.PaymentScreenshot;
        }
        if (text.Contains("交易明细") || text.Contains("微信支付"))
        {
            return EvidenceKinds.TransactionSource;
        }
        if (text.Contains("已买到的宝贝") || text.Contains("订单"))
        {
            return EvidenceKinds.OrderSource;
        }
        return EvidenceKinds.Other;
    }

    private static string DetectPlatform(string text)
    {
        if (text.Contains("微信")) return "微信";
        if (text.Contains("淘宝")) return "淘宝";
        if (text.Contains("美团")) return "美团";
        if (text.Contains("滴滴")) return "滴滴";
        if (text.Contains("发票")) return "发票";
        return "其他";
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
            @"发票金额\s*[：:】\]\s]*[¥￥]?\s*([0-9]+(?:[._][0-9]+)?)",
            @"价税合计.{0,30}?[小写）\)]\s*[¥￥]?\s*([0-9]+(?:[._][0-9]+)?)",
            @"小写\s*[）\)]?\s*[¥￥]?\s*([0-9]+(?:[._][0-9]+)?)",
            @"金额\s*[：:】\]\s]*[¥￥]?\s*([0-9]+(?:[._][0-9]+)?)",
            @"[¥￥]\s*([0-9]+(?:\.[0-9]+)?)"
        };
        foreach (var pattern in patterns)
        {
            var match = Regex.Match(text, pattern);
            var value = match.Success ? match.Groups[1].Value.Replace('_', '.') : "";
            if (match.Success && decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var amount))
            {
                return amount;
            }
        }
        return 0m;
    }

    private static string ExtractInvoiceNo(string text)
    {
        var labeled = Regex.Match(text, @"(?:发票号码|发票号|电子发票号码|EIid|InvoiceNumber)[_：:\s-]*([0-9]{8,24})", RegexOptions.IgnoreCase);
        if (labeled.Success) return labeled.Groups[1].Value;

        var dzfp = Regex.Match(text, @"dzfp[_-]([0-9]{12,24})", RegexOptions.IgnoreCase);
        if (dzfp.Success) return dzfp.Groups[1].Value;

        var candidates = Regex.Matches(text, @"(?<!\d)([0-9]{18,24})(?!\d)")
            .Select(m => m.Groups[1].Value)
            .ToList();
        if (candidates.Count > 0)
        {
            var candidate = candidates[0];
            if (candidate.Length > 20 && !candidate.StartsWith("20", StringComparison.Ordinal))
            {
                candidate = candidate[..20];
            }
            return candidate;
        }

        var fallback = Regex.Matches(text, @"(?<!\d)([0-9]{8,17})(?!\d)")
            .Select(m => m.Groups[1].Value)
            .OrderByDescending(v => v.Length)
            .FirstOrDefault();
        return fallback ?? "";
    }

    private static string ExtractVendorFromInvoiceText(string text)
    {
        var patterns = new[]
        {
            @"销售方信息.{0,120}?名称[:：]?\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,60})",
            @"销方名称[:：]\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,60})",
            @"销售方名称[:：]\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,60})",
            @"<[^>]*(?:SellerName|Seller|Xsfmc|XSFMC)[^>]*>\s*([^<]{4,80})\s*</"
        };

        foreach (var pattern in patterns)
        {
            var match = Regex.Match(text, pattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            if (match.Success)
            {
                var value = CleanVendor(match.Groups[1].Value);
                if (!string.IsNullOrWhiteSpace(value) && !LooksLikeInvoiceTableHeader(value))
                {
                    return value;
                }
            }
        }

        var betweenTaxIds = Regex.Match(text, @"91440300MA5ECG7E71\s+(.{4,80}?)\s+[0-9A-Z]{15,20}", RegexOptions.Singleline);
        if (betweenTaxIds.Success)
        {
            var value = CleanVendor(betweenTaxIds.Groups[1].Value);
            if (!string.IsNullOrWhiteSpace(value) && !LooksLikeInvoiceTableHeader(value))
            {
                return value;
            }
        }

        return "";
    }

    private static string CleanVendor(string value)
    {
        return Regex.Replace(value ?? "", @"\s+", "").Trim();
    }

    private static bool LooksLikeInvoiceTableHeader(string value)
    {
        var keywords = new[] { "项目名称", "规格型号", "统一社会信用代码", "纳税人识别号", "下载次数", "税率", "征收率", "单价", "金额", "名称名称" };
        return keywords.Any(k => value.Contains(k, StringComparison.OrdinalIgnoreCase));
    }

    private static string ExtractVendor(string text)
    {
        var fileName = Path.GetFileNameWithoutExtension(text);
        var parts = Regex.Split(fileName, @"[_\s]+").Where(p => p.Length >= 3).ToList();
        return parts.FirstOrDefault(p => p.Any(ch => ch >= '\u4e00' && ch <= '\u9fff')) ?? "";
    }

    private static string BuildTitle(string kind, string platform, string text, string file)
    {
        var name = Path.GetFileNameWithoutExtension(file);
        if (kind == EvidenceKinds.Invoice)
        {
            return Shorten(name, 80);
        }
        return Shorten($"{platform}-{name}", 80);
    }

    private static string GuessProject(string text)
    {
        var hardwareKeywords = new[] { "硬件", "采购", "设备", "仪器", "工具", "五金", "电子", "电源", "传感器", "模块", "芯片", "零件", "配件", "网络设备", "防静电", "工作台", "插排", "U盘" };
        return hardwareKeywords.Any(k => text.Contains(k, StringComparison.OrdinalIgnoreCase))
            ? "全自动压缩机固有频率检测设备原理样机研制"
            : "";
    }

    private static string Shorten(string value, int max)
    {
        value = value.Trim();
        return value.Length <= max ? value : value[..max];
    }

    private static string StableId(string value)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(value.ToLowerInvariant()));
        return Convert.ToHexString(bytes)[..16];
    }

    private static List<EvidenceItem> DeduplicateInvoices(IEnumerable<EvidenceItem> source)
    {
        var items = source.ToList();
        var nonInvoices = items.Where(i => i.Kind != EvidenceKinds.Invoice).ToList();
        var invoices = items.Where(i => i.Kind == EvidenceKinds.Invoice).ToList();

        var byInvoiceNo = invoices
            .Where(i => !string.IsNullOrWhiteSpace(i.InvoiceNumber))
            .GroupBy(i => i.InvoiceNumber, StringComparer.OrdinalIgnoreCase);

        var result = new List<EvidenceItem>(nonInvoices);
        var consumed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var group in byInvoiceNo)
        {
            var rows = group.ToList();
            foreach (var row in rows) consumed.Add(row.FilePath);
            var primary = rows
                .OrderByDescending(InvoiceFormatPriority)
                .ThenBy(i => i.FilePath.Contains("_1.", StringComparison.OrdinalIgnoreCase))
                .ThenBy(i => i.FilePath, StringComparer.OrdinalIgnoreCase)
                .First();
            if (rows.Count > 1)
            {
                primary.DuplicateInfo = $"同一发票号合并 {rows.Count} 个文件";
                primary.Note = MergeNote(primary.Note, primary.DuplicateInfo);
            }
            MarkInvoiceReviewSuggestion(primary);
            result.Add(primary);
        }

        var remainder = invoices.Where(i => !consumed.Contains(i.FilePath)).ToList();
        foreach (var group in remainder.GroupBy(i => string.IsNullOrWhiteSpace(i.FileHash) ? i.FilePath : i.FileHash, StringComparer.OrdinalIgnoreCase))
        {
            var rows = group.ToList();
            var primary = rows
                .OrderByDescending(InvoiceFormatPriority)
                .ThenBy(i => i.FilePath, StringComparer.OrdinalIgnoreCase)
                .First();
            if (rows.Count > 1)
            {
                primary.DuplicateInfo = $"相同文件合并 {rows.Count} 个副本";
                primary.Note = MergeNote(primary.Note, primary.DuplicateInfo);
            }
            MarkInvoiceReviewSuggestion(primary);
            result.Add(primary);
        }

        return result.OrderBy(i => i.Date).ThenBy(i => i.Platform).ThenBy(i => i.Title).ToList();
    }

    private static int InvoiceFormatPriority(EvidenceItem item)
    {
        return Path.GetExtension(item.FilePath).ToLowerInvariant() switch
        {
            ".pdf" => 5,
            ".ofd" => 4,
            ".xml" => 3,
            ".zip" => 2,
            _ => 1
        };
    }

    private static string MergeNote(string left, string right)
    {
        if (string.IsNullOrWhiteSpace(left)) return right;
        return left.Contains(right, StringComparison.OrdinalIgnoreCase) ? left : $"{left}；{right}";
    }

    private static void MarkInvoiceReviewSuggestion(EvidenceItem item)
    {
        var needs = new List<string>();
        if (item.Amount <= 0) needs.Add("金额未识别");
        if (string.IsNullOrWhiteSpace(item.InvoiceNumber)) needs.Add("发票号未识别");
        if (needs.Count == 0)
        {
            return;
        }
        item.Suggestion = string.Join("；", needs);
        item.Note = MergeNote(item.Note, item.Suggestion);
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

    private static InvoiceMeta ReadInvoiceMeta(string file)
    {
        var ext = Path.GetExtension(file).ToLowerInvariant();
        if (ext == ".xml")
        {
            return ReadInvoiceXml(file);
        }

        var siblingXml = Path.ChangeExtension(file, ".xml");
        if (File.Exists(siblingXml))
        {
            return ReadInvoiceXml(siblingXml);
        }

        if (ext == ".pdf")
        {
            return ReadInvoicePdf(file);
        }

        return new InvoiceMeta();
    }

    private static InvoiceMeta ReadInvoiceXml(string file)
    {
        try
        {
            var doc = XDocument.Load(file);
            string? Element(string name) => doc.Descendants().FirstOrDefault(e => e.Name.LocalName.Equals(name, StringComparison.OrdinalIgnoreCase))?.Value.Trim();
            var amountText = Element("TotalTax-includedAmount")
                ?? Element("TotaltaxIncludedAmount")
                ?? Element("TotalAmount")
                ?? Element("TotalTaxIncludedAmount");
            decimal.TryParse(amountText, NumberStyles.Number, CultureInfo.InvariantCulture, out var amount);

            var invoiceNo = Element("InvoiceNumber") ?? Element("EIid");
            var date = Element("IssueTime") ?? Element("RequestTime");
            if (!string.IsNullOrWhiteSpace(date) && DateTime.TryParse(date, out var parsed))
            {
                date = parsed.ToString("yyyy-MM-dd");
            }

            return new InvoiceMeta
            {
                InvoiceNumber = invoiceNo ?? "",
                Amount = amount,
                Date = date ?? "",
                Vendor = ExtractVendorFromInvoiceText(doc.ToString(System.Xml.Linq.SaveOptions.DisableFormatting))
            };
        }
        catch
        {
            return new InvoiceMeta();
        }
    }

    private static InvoiceMeta ReadInvoicePdf(string file)
    {
        try
        {
            using var document = PdfDocument.Open(file);
            var builder = new StringBuilder();
            foreach (var page in document.GetPages().Take(2))
            {
                builder.AppendLine(page.Text);
            }
            var text = Regex.Replace(builder.ToString(), @"\s+", " ");
            return new InvoiceMeta
            {
                InvoiceNumber = ExtractInvoiceNo(text),
                Amount = ExtractAmount(text),
                Date = ExtractDate(text),
                Vendor = ExtractVendorFromInvoiceText(text)
            };
        }
        catch
        {
            return new InvoiceMeta();
        }
    }

    private sealed class InvoiceMeta
    {
        public string InvoiceNumber { get; set; } = "";
        public decimal Amount { get; set; }
        public string Date { get; set; } = "";
        public string Vendor { get; set; } = "";
    }
}

public sealed class InvoiceFileCleaner
{
    private readonly WorkspaceService _workspace;

    public InvoiceFileCleaner(WorkspaceService workspace)
    {
        _workspace = workspace;
    }

    public int ArchiveNonPdfFormats(AppConfig config)
    {
        var invoiceDir = _workspace.Resolve(config.InvoiceDir);
        if (!Directory.Exists(invoiceDir))
        {
            return 0;
        }

        var pdfKeys = Directory.EnumerateFiles(invoiceDir, "*.pdf", SearchOption.AllDirectories)
            .Select(file => InvoiceFormatPolicy.InvoiceKey(Path.GetFileName(file), file))
            .Where(key => !string.IsNullOrWhiteSpace(key))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (pdfKeys.Count == 0)
        {
            return 0;
        }

        var archiveRoot = _workspace.Resolve(Path.Combine(config.ArchiveDir, $"非PDF重复格式_{DateTime.Now:yyyyMMdd_HHmmss}"));
        var moved = 0;
        foreach (var file in Directory.EnumerateFiles(invoiceDir, "*.*", SearchOption.AllDirectories).ToList())
        {
            var ext = Path.GetExtension(file).ToLowerInvariant();
            if (ext == ".ico")
            {
                var relativeIcon = Path.GetRelativePath(invoiceDir, file);
                var targetIcon = UniquePath(Path.Combine(archiveRoot, relativeIcon));
                Directory.CreateDirectory(Path.GetDirectoryName(targetIcon)!);
                File.Move(file, targetIcon);
                File.WriteAllText(targetIcon + ".reason.txt", "ICO 图标不是发票。" + Environment.NewLine, Encoding.UTF8);
                moved++;
                continue;
            }
            if (ext == ".pdf" || ext == ".csv" || ext == ".xlsx" || ext == ".json")
            {
                continue;
            }
            if (file.Contains($"{Path.DirectorySeparatorChar}非PDF重复格式_", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var key = InvoiceFormatPolicy.InvoiceKey(Path.GetFileName(file), file);
            if (string.IsNullOrWhiteSpace(key) || !pdfKeys.Contains(key))
            {
                continue;
            }

            var relative = Path.GetRelativePath(invoiceDir, file);
            var target = UniquePath(Path.Combine(archiveRoot, relative));
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            File.Move(file, target);
            moved++;
        }

        return moved;
    }

    public int ArchiveDateRejectedInvoices(AppConfig config)
    {
        var invoiceDir = _workspace.Resolve(config.InvoiceDir);
        if (!Directory.Exists(invoiceDir))
        {
            return 0;
        }

        var startText = !string.IsNullOrWhiteSpace(config.Email.Start) ? config.Email.Start : config.DateStart;
        if (!DateTime.TryParse(startText, out var start))
        {
            return 0;
        }

        var archiveRoot = _workspace.Resolve(Path.Combine(config.ArchiveDir, $"日期不符发票_{DateTime.Now:yyyyMMdd_HHmmss}"));
        var moved = 0;
        foreach (var file in Directory.EnumerateFiles(invoiceDir, "*.pdf", SearchOption.AllDirectories).ToList())
        {
            if (IsInsideCleanerArchive(file))
            {
                continue;
            }

            var invoiceDate = ReadPdfDate(file);
            if (string.IsNullOrWhiteSpace(invoiceDate)
                || !DateTime.TryParse(invoiceDate, out var parsed)
                || parsed.Date >= start.Date)
            {
                continue;
            }

            var relative = Path.GetRelativePath(invoiceDir, file);
            var target = UniquePath(Path.Combine(archiveRoot, relative));
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            File.Move(file, target);
            File.WriteAllText(target + ".reason.txt", $"发票日期 {invoiceDate} 早于开始时间 {start:yyyy-MM-dd}，不纳入本期报销。" + Environment.NewLine, Encoding.UTF8);
            moved++;
        }

        return moved;
    }

    private static bool IsInsideCleanerArchive(string file)
    {
        var normalized = file.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
        return normalized.Contains($"{Path.DirectorySeparatorChar}非PDF重复格式_", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains($"{Path.DirectorySeparatorChar}日期不符发票_", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains($"{Path.DirectorySeparatorChar}_日期不符发票{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase);
    }

    private static string ReadPdfDate(string file)
    {
        try
        {
            using var document = PdfDocument.Open(file);
            var builder = new StringBuilder();
            foreach (var page in document.GetPages().Take(3))
            {
                builder.AppendLine(page.Text);
            }
            return ExtractDate(Regex.Replace(builder.ToString(), @"\s+", " "));
        }
        catch
        {
            return "";
        }
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

    private static string UniquePath(string path)
    {
        if (!File.Exists(path))
        {
            return path;
        }

        var dir = Path.GetDirectoryName(path)!;
        var stem = Path.GetFileNameWithoutExtension(path);
        var ext = Path.GetExtension(path);
        for (var i = 1; i < 10000; i++)
        {
            var candidate = Path.Combine(dir, $"{stem}_{i}{ext}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }
        }
        throw new IOException("无法生成不重复归档路径：" + path);
    }
}

public static class InvoiceFormatPolicy
{
    public static bool ShouldSkipBecausePdfExists(string extension, string invoiceKey, IEnumerable<string> localPdfKeys, ISet<string> existingPdfKeys)
    {
        if (extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }
        if (string.IsNullOrWhiteSpace(invoiceKey))
        {
            return false;
        }
        return existingPdfKeys.Contains(invoiceKey) || localPdfKeys.Contains(invoiceKey, StringComparer.OrdinalIgnoreCase);
    }

    public static string InvoiceKey(params string[] values)
    {
        var text = string.Join(" ", values.Where(v => !string.IsNullOrWhiteSpace(v)));
        var labeled = Regex.Match(text, @"(?:发票号码|发票号|电子发票号码|EIid|InvoiceNumber)[_：:\s-]*([0-9]{8,24})", RegexOptions.IgnoreCase);
        if (labeled.Success) return NormalizeInvoiceNo(labeled.Groups[1].Value);

        var dzfp = Regex.Match(text, @"dzfp[_-]([0-9]{12,24})", RegexOptions.IgnoreCase);
        if (dzfp.Success) return NormalizeInvoiceNo(dzfp.Groups[1].Value);

        var candidates = Regex.Matches(text, @"(?<!\d)([0-9]{18,24})(?!\d)")
            .Select(m => NormalizeInvoiceNo(m.Groups[1].Value))
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .ToList();
        if (candidates.Count > 0) return candidates[0];

        return NormalizeStemKey(text);
    }

    private static string NormalizeInvoiceNo(string value)
    {
        if (value.Length > 20 && !value.StartsWith("20", StringComparison.Ordinal))
        {
            value = value[..20];
        }
        return value;
    }

    private static string NormalizeStemKey(string text)
    {
        var stem = Path.GetFileNameWithoutExtension(text);
        stem = Regex.Replace(stem, @"_\d+$", "", RegexOptions.IgnoreCase);
        stem = Regex.Replace(stem, @"(_查阅需OFD阅读器|_ofd_查阅需OFD阅读器)$", "", RegexOptions.IgnoreCase);
        stem = ConfirmationService.NormalizeKey(stem);
        return stem.Length >= 12 ? stem : "";
    }
}

internal static class EnumerablePipeExtensions
{
    public static TResult Pipe<T, TResult>(this T value, Func<T, TResult> func) => func(value);
}

public sealed class EmailAuditReportService
{
    private readonly WorkspaceService _workspace;

    public EmailAuditReportService(WorkspaceService workspace)
    {
        _workspace = workspace;
    }

    public string Generate(AppConfig config)
    {
        var invoiceDir = _workspace.Resolve(config.Email.OutputDir);
        var outputDir = _workspace.Resolve(config.OutputDir);
        Directory.CreateDirectory(outputDir);
        var latest = Directory.Exists(invoiceDir)
            ? Directory.EnumerateFiles(invoiceDir, "邮箱发票下载清单_*.csv", SearchOption.TopDirectoryOnly)
                .OrderByDescending(File.GetLastWriteTime)
                .FirstOrDefault()
            : null;
        if (latest == null)
        {
            throw new FileNotFoundException("没有找到邮箱发票下载清单，请先执行一次邮箱发票下载。");
        }

        var rows = ReadCsv(latest);
        var grouped = rows
            .GroupBy(r => FirstNonEmpty(Get(r, "message_key"), Get(r, "msg_id"), Get(r, "subject")), StringComparer.OrdinalIgnoreCase)
            .Select(g => BuildAuditRow(g.Key, g.ToList()))
            .OrderBy(r => r.Date)
            .ThenBy(r => r.Subject)
            .ToList();

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("逐封邮件统计");
        var headers = new[]
        {
            "邮件日期", "邮件主题", "邮件ID", "总体状态",
            "有发票关键字", "有美团关键字", "有淘宝关键字", "有京东关键字",
            "有附件", "附件数", "疑似发票附件", "疑似发票附件数",
            "疑似发票链接", "链接数", "已下载文件数", "重复/跳过数",
            "需要人工核验", "需要模型判断", "判断原因", "文件清单", "链接清单", "模型判断结果"
        };
        for (var c = 0; c < headers.Length; c++)
        {
            ws.Cell(1, c + 1).Value = headers[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
            ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#D9EAF7");
        }

        var r = 2;
        foreach (var row in grouped)
        {
            ws.Cell(r, 1).Value = row.Date;
            ws.Cell(r, 2).Value = row.Subject;
            ws.Cell(r, 3).Value = row.MessageId;
            ws.Cell(r, 4).Value = row.Status;
            ws.Cell(r, 5).Value = row.HasInvoiceKeyword ? "是" : "否";
            ws.Cell(r, 6).Value = row.HasMeituanKeyword ? "是" : "否";
            ws.Cell(r, 7).Value = row.HasTaobaoKeyword ? "是" : "否";
            ws.Cell(r, 8).Value = row.HasJdKeyword ? "是" : "否";
            ws.Cell(r, 9).Value = row.HasAttachment ? "是" : "否";
            ws.Cell(r, 10).Value = row.AttachmentCount;
            ws.Cell(r, 11).Value = row.HasLikelyInvoiceAttachment ? "是" : "否";
            ws.Cell(r, 12).Value = row.LikelyInvoiceAttachmentCount;
            ws.Cell(r, 13).Value = row.HasLikelyInvoiceLink ? "是" : "否";
            ws.Cell(r, 14).Value = row.LinkCount;
            ws.Cell(r, 15).Value = row.DownloadedFileCount;
            ws.Cell(r, 16).Value = row.SkippedOrDuplicateCount;
            ws.Cell(r, 17).Value = row.NeedsHumanReview ? "是" : "否";
            ws.Cell(r, 18).Value = row.NeedsModelReview ? "是" : "否";
            ws.Cell(r, 19).Value = row.Reason;
            ws.Cell(r, 20).Value = row.Files;
            ws.Cell(r, 21).Value = row.Urls;
            ws.Cell(r, 22).Value = "";

            var color = row.NeedsHumanReview
                ? XLColor.FromHtml("#FFF2CC")
                : row.DownloadedFileCount > 0 || row.SkippedOrDuplicateCount > 0
                    ? XLColor.FromHtml("#D9EAD3")
                    : XLColor.NoColor;
            ws.Range(r, 1, r, headers.Length).Style.Fill.BackgroundColor = color;
            r++;
        }

        ws.Columns().AdjustToContents(8, 60);
        ws.SheetView.FreezeRows(1);

        var summary = wb.Worksheets.Add("统计汇总");
        summary.Cell(1, 1).Value = "指标";
        summary.Cell(1, 2).Value = "数量";
        summary.Range(1, 1, 1, 2).Style.Font.Bold = true;
        var metrics = new (string Name, int Count)[]
        {
            ("邮件总数", grouped.Count),
            ("有发票关键字", grouped.Count(x => x.HasInvoiceKeyword)),
            ("有美团关键字", grouped.Count(x => x.HasMeituanKeyword)),
            ("有淘宝关键字", grouped.Count(x => x.HasTaobaoKeyword)),
            ("有京东关键字", grouped.Count(x => x.HasJdKeyword)),
            ("有附件", grouped.Count(x => x.HasAttachment)),
            ("有疑似发票附件", grouped.Count(x => x.HasLikelyInvoiceAttachment)),
            ("有疑似发票链接", grouped.Count(x => x.HasLikelyInvoiceLink)),
            ("需要人工核验", grouped.Count(x => x.NeedsHumanReview)),
            ("建议模型判断", grouped.Count(x => x.NeedsModelReview)),
            ("已下载到文件", grouped.Count(x => x.DownloadedFileCount > 0)),
            ("重复或跳过", grouped.Count(x => x.SkippedOrDuplicateCount > 0))
        };
        for (var i = 0; i < metrics.Length; i++)
        {
            summary.Cell(i + 2, 1).Value = metrics[i].Name;
            summary.Cell(i + 2, 2).Value = metrics[i].Count;
        }
        summary.Columns().AdjustToContents();

        var output = Path.Combine(outputDir, $"邮箱逐封统计_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        wb.SaveAs(output);
        return output;
    }

    private static EmailAuditRow BuildAuditRow(string key, List<Dictionary<string, string>> rows)
    {
        var subject = TextEncodingFixer.Fix(rows.Select(r => Get(r, "subject")).FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? "");
        var status = rows.Where(r => Get(r, "kind") == "message")
            .Select(r => Get(r, "status"))
            .FirstOrDefault(s => !string.IsNullOrWhiteSpace(s))
            ?? rows.Select(r => Get(r, "status")).FirstOrDefault(s => !string.IsNullOrWhiteSpace(s))
            ?? "";
        var date = rows.Select(r => Get(r, "date")).FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? "";
        var msgId = rows.Select(r => Get(r, "msg_id")).FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? key;

        var files = rows.Select(r => Get(r, "file")).Where(v => !string.IsNullOrWhiteSpace(v)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var urls = rows.Select(r => Get(r, "url")).Where(v => !string.IsNullOrWhiteSpace(v)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var kinds = rows.Select(r => Get(r, "kind")).ToList();
        var statuses = rows.Select(r => Get(r, "status")).ToList();
        var errors = rows.Select(r => Get(r, "error")).Where(e => !string.IsNullOrWhiteSpace(e)).ToList();

        var attachmentRows = rows.Where(r => Get(r, "kind") == "attachment").ToList();
        var attachmentCount = ParseInt(rows.Select(r => Get(r, "attachment_total")).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v))) ?? attachmentRows.Count;
        var linkCount = ParseInt(rows.Select(r => Get(r, "link_candidate_total")).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v))) ?? urls.Count;
        var likelyInvoiceAttachmentCount = attachmentRows.Count(r => LooksLikeInvoice(Get(r, "file"), subject) || LooksLikeInvoice(Get(r, "duplicate_of"), subject));
        var downloaded = rows.Count(r => !string.IsNullOrWhiteSpace(Get(r, "file")) && Get(r, "status") == "已下载");
        var skipped = rows.Count(r => Get(r, "status") is "重复跳过" or "重复已存在" or "PDF已存在跳过" or "已处理跳过");

        var hasInvoiceKeyword = LooksLikeInvoice(subject, string.Join(" ", files));
        var hasMeituan = ContainsAny(subject, "美团", "三快", "meituan");
        var hasTaobao = ContainsAny(subject, "淘宝", "天猫", "taobao", "tmall");
        var hasJd = ContainsAny(subject, "京东", "jd.com", "jingdong");
        var hasLikelyInvoiceLink = urls.Any(u => LooksLikeInvoice(u, subject));
        var hasRiskStatus = statuses.Any(s => s is "未下载到文件" or "失败" or "需人工确认" or "待核验");
        var needsHuman = hasRiskStatus || errors.Count > 0 || ((hasInvoiceKeyword || hasMeituan || hasTaobao || hasJd) && downloaded == 0 && skipped == 0 && likelyInvoiceAttachmentCount == 0 && !hasLikelyInvoiceLink);
        var needsModel = needsHuman || ((hasMeituan || hasTaobao || hasJd) && !hasInvoiceKeyword && downloaded == 0);
        var reasons = new List<string>();
        if (hasRiskStatus) reasons.Add("下载状态异常或未完成");
        if (errors.Count > 0) reasons.Add("存在错误说明");
        if (hasInvoiceKeyword) reasons.Add("命中发票关键字");
        if (hasMeituan || hasTaobao || hasJd) reasons.Add("命中平台关键字");
        if (attachmentCount > 0 && likelyInvoiceAttachmentCount == 0) reasons.Add("有附件但规则无法确认是否发票");
        if (needsModel) reasons.Add("建议模型二次判断");

        return new EmailAuditRow
        {
            Date = date,
            Subject = subject,
            MessageId = msgId,
            Status = status,
            HasInvoiceKeyword = hasInvoiceKeyword,
            HasMeituanKeyword = hasMeituan,
            HasTaobaoKeyword = hasTaobao,
            HasJdKeyword = hasJd,
            HasAttachment = attachmentCount > 0 || attachmentRows.Count > 0,
            AttachmentCount = attachmentCount,
            HasLikelyInvoiceAttachment = likelyInvoiceAttachmentCount > 0,
            LikelyInvoiceAttachmentCount = likelyInvoiceAttachmentCount,
            HasLikelyInvoiceLink = hasLikelyInvoiceLink,
            LinkCount = linkCount,
            DownloadedFileCount = downloaded,
            SkippedOrDuplicateCount = skipped,
            NeedsHumanReview = needsHuman,
            NeedsModelReview = needsModel,
            Reason = string.Join("；", reasons.Distinct()),
            Files = string.Join(Environment.NewLine, files),
            Urls = string.Join(Environment.NewLine, urls)
        };
    }

    private static List<Dictionary<string, string>> ReadCsv(string path)
    {
        var lines = File.ReadAllLines(path, Encoding.UTF8);
        if (lines.Length == 0) return [];
        var headers = ParseCsvLine(lines[0]);
        var result = new List<Dictionary<string, string>>();
        foreach (var line in lines.Skip(1))
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            var values = ParseCsvLine(line);
            var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < headers.Count; i++)
            {
                row[headers[i]] = i < values.Count ? values[i] : "";
            }
            result.Add(row);
        }
        return result;
    }

    private static List<string> ParseCsvLine(string line)
    {
        var result = new List<string>();
        var current = new StringBuilder();
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

    private static string Get(Dictionary<string, string> row, string key) => row.GetValueOrDefault(key, "");
    private static string FirstNonEmpty(params string[] values) => values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "";
    private static int? ParseInt(string? value) => int.TryParse(value, out var result) ? result : null;

    private static bool LooksLikeInvoice(string value, string context)
    {
        var text = $"{value} {context}";
        return ContainsAny(text, "发票", "电子发票", "invoice", "fapiao", "dzfp", ".ofd", ".pdf", ".xml", ".zip", "开票");
    }

    private static bool ContainsAny(string value, params string[] keywords)
    {
        return keywords.Any(k => value.Contains(k, StringComparison.OrdinalIgnoreCase));
    }

    private sealed class EmailAuditRow
    {
        public string Date { get; set; } = "";
        public string Subject { get; set; } = "";
        public string MessageId { get; set; } = "";
        public string Status { get; set; } = "";
        public bool HasInvoiceKeyword { get; set; }
        public bool HasMeituanKeyword { get; set; }
        public bool HasTaobaoKeyword { get; set; }
        public bool HasJdKeyword { get; set; }
        public bool HasAttachment { get; set; }
        public int AttachmentCount { get; set; }
        public bool HasLikelyInvoiceAttachment { get; set; }
        public int LikelyInvoiceAttachmentCount { get; set; }
        public bool HasLikelyInvoiceLink { get; set; }
        public int LinkCount { get; set; }
        public int DownloadedFileCount { get; set; }
        public int SkippedOrDuplicateCount { get; set; }
        public bool NeedsHumanReview { get; set; }
        public bool NeedsModelReview { get; set; }
        public string Reason { get; set; } = "";
        public string Files { get; set; } = "";
        public string Urls { get; set; } = "";
    }
}

public sealed class MatchingService
{
    public List<MatchCandidate> BuildCandidates(List<EvidenceItem> items)
    {
        var invoices = items.Where(i => i.Kind == EvidenceKinds.Invoice && i.FileDecision != "无效").ToList();
        var expenses = items.Where(i => i.Kind != EvidenceKinds.Invoice && i.Kind != EvidenceKinds.Template && i.ReimburseDecision != "不报销").ToList();
        var result = new List<MatchCandidate>();
        foreach (var expense in expenses)
        {
            foreach (var invoice in invoices)
            {
                var (score, reason) = Score(expense, invoice);
                if (score < 25)
                {
                    continue;
                }
                result.Add(new MatchCandidate
                {
                    Id = $"{expense.Id}-{invoice.Id}",
                    ExpenseId = expense.Id,
                    InvoiceId = invoice.Id,
                    ExpenseTitle = expense.Title,
                    InvoiceTitle = invoice.Title,
                    ExpenseAmount = expense.Amount,
                    InvoiceAmount = invoice.Amount,
                    ExpenseDate = expense.Date,
                    InvoiceDate = invoice.Date,
                    Score = score,
                    Reason = reason
                });
            }
        }
        return result.OrderByDescending(m => m.Score).ThenBy(m => m.ExpenseTitle).Take(500).ToList();
    }

    private static (int Score, string Reason) Score(EvidenceItem expense, EvidenceItem invoice)
    {
        var score = 0;
        var reasons = new List<string>();
        if (expense.Amount > 0 && invoice.Amount > 0)
        {
            var diff = Math.Abs(expense.Amount - invoice.Amount);
            if (diff == 0)
            {
                score += 45;
                reasons.Add("金额一致");
            }
            else if (diff <= 5 || diff <= expense.Amount * 0.05m)
            {
                score += 25;
                reasons.Add($"金额接近，差额{diff:0.##}");
            }
        }
        if (!string.IsNullOrWhiteSpace(expense.Date) && !string.IsNullOrWhiteSpace(invoice.Date)
            && DateTime.TryParse(expense.Date, out var d1) && DateTime.TryParse(invoice.Date, out var d2))
        {
            var days = Math.Abs((d1 - d2).TotalDays);
            if (days <= 3)
            {
                score += 25;
                reasons.Add("日期接近");
            }
            else if (days <= 30)
            {
                score += 10;
                reasons.Add("日期在30天内");
            }
        }
        var expenseKey = ConfirmationService.NormalizeKey(expense.Title);
        var invoiceKey = ConfirmationService.NormalizeKey(invoice.Title);
        if (!string.IsNullOrWhiteSpace(expenseKey) && !string.IsNullOrWhiteSpace(invoiceKey))
        {
            if (expenseKey.Contains(invoiceKey, StringComparison.OrdinalIgnoreCase) || invoiceKey.Contains(expenseKey, StringComparison.OrdinalIgnoreCase))
            {
                score += 20;
                reasons.Add("名称相似");
            }
            else if (SharedChinesePrefix(expenseKey, invoiceKey) >= 4)
            {
                score += 10;
                reasons.Add("关键词相似");
            }
        }
        if (!string.IsNullOrWhiteSpace(invoice.InvoiceNumber))
        {
            score += 5;
            reasons.Add("有发票号");
        }
        return (score, string.Join("；", reasons));
    }

    private static int SharedChinesePrefix(string left, string right)
    {
        var max = Math.Min(left.Length, right.Length);
        var count = 0;
        while (count < max && left[count] == right[count])
        {
            count++;
        }
        return count;
    }
}

public sealed class ReportService
{
    private readonly WorkspaceService _workspace;

    public ReportService(WorkspaceService workspace)
    {
        _workspace = workspace;
    }

    public string Generate(AppConfig config, List<EvidenceItem> items, List<MatchCandidate> matches)
    {
        var template = _workspace.Resolve(config.TemplatePath);
        var outputDir = _workspace.Resolve(config.OutputDir);
        Directory.CreateDirectory(outputDir);
        var output = Path.Combine(outputDir, $"报销整理_匹配工作台_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

        using var wb = File.Exists(template) ? new XLWorkbook(template) : new XLWorkbook();
        var ws = wb.Worksheets.FirstOrDefault(w => w.Name.Contains("报销明细")) ?? wb.Worksheets.Add("报销明细");
        PrepareExpenseSheet(ws);

        var row = 4;
        foreach (var item in items.Where(i => i.ReimburseDecision is "报销" or "建议报销").OrderBy(i => i.Date).ThenBy(i => i.Platform))
        {
            ws.Cell(row, 2).Value = item.Date;
            ws.Cell(row, 3).Value = item.Amount;
            ws.Cell(row, 4).Value = item.Kind == EvidenceKinds.Invoice ? "有" : "否";
            ws.Cell(row, 5).Value = BuildExpenseReason(item);
            ws.Cell(row, 6).Value = item.Project;
            ws.Cell(row, 8).Value = config.Operator;
            ws.Cell(row, 9).Value = item.InvoiceNumber;
            ws.Range(row, 2, row, 9).Style.Fill.BackgroundColor = item.ReimburseDecision == "报销" ? XLColor.FromHtml("#D9EAD3") : XLColor.FromHtml("#FFF2CC");
            row++;
        }
        ws.Columns().AdjustToContents(8, 45);

        WriteListSheet(wb, "软件确认记录", items);
        WriteMatchSheet(wb, "匹配候选", matches);
        wb.SaveAs(output);
        return output;
    }

    private static void PrepareExpenseSheet(IXLWorksheet ws)
    {
        if (ws.Cell(3, 2).IsEmpty())
        {
            ws.Cell(3, 2).Value = "开票日期";
            ws.Cell(3, 3).Value = "发票金额（RMB）";
            ws.Cell(3, 4).Value = "有无正式发票";
            ws.Cell(3, 5).Value = "支出事由";
            ws.Cell(3, 6).Value = "项目名称";
            ws.Cell(3, 7).Value = "归属研发支出";
            ws.Cell(3, 8).Value = "经办人";
            ws.Cell(3, 9).Value = "备注发票号";
        }
        var last = ws.LastRowUsed()?.RowNumber() ?? 4;
        if (last >= 4)
        {
            ws.Range(4, 2, Math.Max(last, 4), 9).Clear(XLClearOptions.Contents);
        }
    }

    private static string BuildExpenseReason(EvidenceItem item)
    {
        var title = item.Title.Trim();
        var vendor = CleanVendorForReason(item.Vendor);
        if (string.IsNullOrWhiteSpace(vendor) || !IsFoodOrHotelExpense(item))
        {
            return title;
        }

        return title.Contains(vendor, StringComparison.OrdinalIgnoreCase)
            ? title
            : $"{title}（{vendor}）";
    }

    private static bool IsFoodOrHotelExpense(EvidenceItem item)
    {
        var text = $"{item.Title} {item.Platform} {item.Vendor} {item.RelativePath}";
        return ContainsAny(text, "餐饮", "餐费", "饭店", "餐厅", "饮食", "外卖", "美团", "三快",
            "住宿", "酒店", "宾馆", "旅馆", "客房", "携程");
    }

    private static string CleanVendorForReason(string vendor)
    {
        vendor = Regex.Replace(vendor ?? "", @"\s+", "").Trim();
        return vendor is "" or "发票" or "电子发票" or "深圳博锐创科技有限公司"
            ? ""
            : ShortenText(vendor, 36);
    }

    private static bool ContainsAny(string value, params string[] keywords)
    {
        return keywords.Any(k => value.Contains(k, StringComparison.OrdinalIgnoreCase));
    }

    private static string ShortenText(string value, int max)
    {
        value = value.Trim();
        return value.Length <= max ? value : value[..max];
    }

    private static void WriteListSheet(XLWorkbook wb, string name, List<EvidenceItem> items)
    {
        if (wb.Worksheets.TryGetWorksheet(name, out var old)) old.Delete();
        var ws = wb.Worksheets.Add(name);
        var headers = new[] { "类型", "来源", "日期", "金额", "标题", "发票号", "报销确认", "文件确认", "匹配状态", "项目", "权重", "建议", "重复/合并", "路径", "备注" };
        for (var c = 0; c < headers.Length; c++)
        {
            ws.Cell(1, c + 1).Value = headers[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
        }
        var r = 2;
        foreach (var item in items)
        {
            ws.Cell(r, 1).Value = item.Kind;
            ws.Cell(r, 2).Value = item.Platform;
            ws.Cell(r, 3).Value = item.Date;
            ws.Cell(r, 4).Value = item.Amount;
            ws.Cell(r, 5).Value = item.Title;
            ws.Cell(r, 6).Value = item.InvoiceNumber;
            ws.Cell(r, 7).Value = item.ReimburseDecision;
            ws.Cell(r, 8).Value = item.FileDecision;
            ws.Cell(r, 9).Value = item.MatchStatus;
            ws.Cell(r, 10).Value = item.Project;
            ws.Cell(r, 11).Value = item.Weight;
            ws.Cell(r, 12).Value = item.Suggestion;
            ws.Cell(r, 13).Value = item.DuplicateInfo;
            ws.Cell(r, 14).Value = item.RelativePath;
            ws.Cell(r, 15).Value = item.Note;
            ws.Range(r, 1, r, headers.Length).Style.Fill.BackgroundColor = item.ReimburseDecision == "报销" ? XLColor.FromHtml("#D9EAD3") : item.ReimburseDecision.Contains("建议") || item.ReimburseDecision == "待确认" ? XLColor.FromHtml("#FFF2CC") : XLColor.NoColor;
            r++;
        }
        ws.Columns().AdjustToContents(8, 50);
    }

    private static void WriteMatchSheet(XLWorkbook wb, string name, List<MatchCandidate> matches)
    {
        if (wb.Worksheets.TryGetWorksheet(name, out var old)) old.Delete();
        var ws = wb.Worksheets.Add(name);
        var headers = new[] { "消费", "发票", "消费金额", "发票金额", "消费日期", "发票日期", "评分", "理由", "确认" };
        for (var c = 0; c < headers.Length; c++)
        {
            ws.Cell(1, c + 1).Value = headers[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
        }
        var r = 2;
        foreach (var match in matches)
        {
            ws.Cell(r, 1).Value = match.ExpenseTitle;
            ws.Cell(r, 2).Value = match.InvoiceTitle;
            ws.Cell(r, 3).Value = match.ExpenseAmount;
            ws.Cell(r, 4).Value = match.InvoiceAmount;
            ws.Cell(r, 5).Value = match.ExpenseDate;
            ws.Cell(r, 6).Value = match.InvoiceDate;
            ws.Cell(r, 7).Value = match.Score;
            ws.Cell(r, 8).Value = match.Reason;
            ws.Cell(r, 9).Value = match.Decision;
            r++;
        }
        ws.Columns().AdjustToContents(8, 50);
    }
}
