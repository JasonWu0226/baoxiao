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

public sealed class InvoiceRecognitionAuditService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private static readonly string[] SupportedExtensions =
    {
        ".pdf", ".ofd", ".xml", ".zip", ".jpg", ".jpeg", ".png", ".bmp", ".webp", ".ico"
    };

    private readonly WorkspaceService _workspace;

    public InvoiceRecognitionAuditService(WorkspaceService workspace)
    {
        _workspace = workspace;
    }

    public string Generate(AppConfig config, string? sourceDir = null)
    {
        var root = _workspace.Resolve(string.IsNullOrWhiteSpace(sourceDir) ? config.InvoiceDir : sourceDir);
        if (!Directory.Exists(root))
        {
            throw new DirectoryNotFoundException("发票扫描目录不存在：" + root);
        }

        var outputDir = _workspace.Resolve(config.OutputDir);
        Directory.CreateDirectory(outputDir);

        var rows = Scan(config, root);
        var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var xlsx = Path.Combine(outputDir, $"发票识别核验_{stamp}.xlsx");
        var json = Path.Combine(outputDir, $"发票识别核验_{stamp}.json");

        WriteWorkbook(xlsx, rows, root);
        File.WriteAllText(json, JsonSerializer.Serialize(rows, JsonOptions) + Environment.NewLine, Encoding.UTF8);
        return xlsx;
    }

    private List<InvoiceRecognitionAuditRow> Scan(AppConfig config, string root)
    {
        var files = Directory.EnumerateFiles(root, "*.*", SearchOption.AllDirectories)
            .Where(file => SupportedExtensions.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase))
            .Where(file => !IsMaintenanceArchiveFile(file))
            .OrderBy(file => file, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var rows = new List<InvoiceRecognitionAuditRow>();
        var byHash = new Dictionary<string, InvoiceRecognitionAuditRow>(StringComparer.OrdinalIgnoreCase);
        var byInvoiceNo = new Dictionary<string, InvoiceRecognitionAuditRow>(StringComparer.OrdinalIgnoreCase);

        foreach (var file in files)
        {
            var row = InspectFile(config, root, file);

            if (!string.IsNullOrWhiteSpace(row.Sha256))
            {
                if (byHash.TryGetValue(row.Sha256, out var duplicate))
                {
                    row.DuplicateStatus = "文件哈希重复";
                    row.DuplicateOf = duplicate.RelativePath;
                    row.Issues.Add("文件内容与其他发票完全相同");
                }
                else
                {
                    byHash[row.Sha256] = row;
                }
            }

            if (!string.IsNullOrWhiteSpace(row.InvoiceNumber))
            {
                if (byInvoiceNo.TryGetValue(row.InvoiceNumber, out var duplicate))
                {
                    row.DuplicateStatus = string.IsNullOrWhiteSpace(row.DuplicateStatus) ? "发票号重复" : row.DuplicateStatus + "；发票号重复";
                    row.DuplicateOf = string.IsNullOrWhiteSpace(row.DuplicateOf) ? duplicate.RelativePath : row.DuplicateOf;
                    row.Issues.Add("发票号码与其他文件重复");
                }
                else
                {
                    byInvoiceNo[row.InvoiceNumber] = row;
                }
            }

            FinalizeStatus(row);
            rows.Add(row);
        }

        return rows;
    }

    private InvoiceRecognitionAuditRow InspectFile(AppConfig config, string root, string file)
    {
        var row = new InvoiceRecognitionAuditRow
        {
            FileName = Path.GetFileName(file),
            FilePath = file,
            RelativePath = Path.GetRelativePath(root, file),
            Extension = Path.GetExtension(file).ToLowerInvariant(),
            SizeBytes = SafeLength(file),
            Sha256 = SafeSha256(file)
        };

        var textResult = ExtractText(file);
        row.TextStatus = textResult.Status;
        row.TextLength = textResult.Text.Length;
        row.TextPreview = Shorten(textResult.Text, 300);
        row.IsLikelyInvoice = LooksLikeInvoiceText(textResult.Text) || LooksLikeInvoiceName(row.FileName);

        if (row.Extension == ".ico")
        {
            row.Issues.Add("ICO图标文件，不是发票");
        }
        else if (IsImage(row.Extension))
        {
            row.Issues.Add("图片格式暂未做OCR，需人工确认是否为二维码或截图发票");
        }
        else if (row.Extension == ".zip")
        {
            row.Issues.Add(textResult.Extra);
        }
        else if (row.Extension == ".pdf" && row.TextLength == 0)
        {
            row.Issues.Add("PDF无可提取文本，可能是扫描件或图片型PDF");
        }
        else if (row.Extension == ".ofd" && row.TextLength == 0)
        {
            row.Issues.Add("OFD未提取到文本");
        }

        var combined = $"{row.FileName} {textResult.Text}";
        row.InvoiceNumber = ExtractInvoiceNo(combined);
        row.InvoiceDate = ExtractDate(combined);
        row.Amount = ExtractAmount(combined);
        row.BuyerName = ExtractBuyerName(combined);
        row.BuyerTaxId = ExtractBuyerTaxId(combined);
        row.Vendor = ExtractVendor(combined);
        row.Category = DetectCategory(combined);

        AddFieldIssues(config, row);
        row.Confidence = CalculateConfidence(row);
        return row;
    }

    private static void AddFieldIssues(AppConfig config, InvoiceRecognitionAuditRow row)
    {
        if (!row.IsLikelyInvoice && row.Extension is not ".ico")
        {
            row.Issues.Add("未命中发票关键字");
        }

        if (row.IsLikelyInvoice || row.Extension is ".pdf" or ".ofd" or ".xml")
        {
            if (string.IsNullOrWhiteSpace(row.InvoiceNumber))
            {
                row.Issues.Add("未识别发票号");
            }

            if (string.IsNullOrWhiteSpace(row.InvoiceDate))
            {
                row.Issues.Add("未识别开票日期");
            }

            if (row.Amount <= 0)
            {
                row.Issues.Add("未识别金额或金额为0");
            }

            if (string.IsNullOrWhiteSpace(row.Vendor))
            {
                row.Issues.Add("未识别销售方");
            }
        }

        if (!string.IsNullOrWhiteSpace(row.InvoiceDate)
            && DateTime.TryParse(row.InvoiceDate, out var invoiceDate))
        {
            if (DateTime.TryParse(config.DateStart, out var start) && invoiceDate < start)
            {
                row.Issues.Add("开票日期早于本期开始日期");
            }
            if (DateTime.TryParse(config.DateEnd, out var end) && invoiceDate > end)
            {
                row.Issues.Add("开票日期晚于本期结束日期");
            }
        }

        var companyName = (config.CompanyName ?? "").Trim();
        var companyTaxId = (config.CompanyTaxId ?? "").Trim();
        if (!string.IsNullOrWhiteSpace(row.BuyerTaxId)
            && !string.IsNullOrWhiteSpace(companyTaxId)
            && !row.BuyerTaxId.Equals(companyTaxId, StringComparison.OrdinalIgnoreCase))
        {
            row.Issues.Add("购方税号不匹配");
        }
        else if (!string.IsNullOrWhiteSpace(row.BuyerName)
            && !string.IsNullOrWhiteSpace(companyName)
            && !row.BuyerName.Contains(companyName, StringComparison.OrdinalIgnoreCase)
            && !companyName.Contains(row.BuyerName, StringComparison.OrdinalIgnoreCase))
        {
            row.Issues.Add("购方名称不匹配");
        }
    }

    private static void FinalizeStatus(InvoiceRecognitionAuditRow row)
    {
        var issueText = string.Join("；", row.Issues);
        row.IssueSummary = issueText;

        if (row.Extension == ".ico")
        {
            row.Status = "非发票/跳过";
            row.Risk = "低";
            return;
        }

        if (ContainsAny(issueText, "购方税号不匹配", "购方名称不匹配", "早于本期开始日期", "晚于本期结束日期"))
        {
            row.Status = "异常";
            row.Risk = "高";
            return;
        }

        if (ContainsAny(issueText, "未识别发票号", "未识别开票日期", "未识别金额", "PDF无可提取文本", "图片格式", "OFD未提取到文本"))
        {
            row.Status = "需复核";
            row.Risk = "中";
            return;
        }

        if (!string.IsNullOrWhiteSpace(row.DuplicateStatus))
        {
            row.Status = "需复核";
            row.Risk = "中";
            return;
        }

        row.Status = row.IsLikelyInvoice ? "通过" : "非发票/跳过";
        row.Risk = row.IsLikelyInvoice ? "低" : "低";
    }

    private static TextExtractionResult ExtractText(string file)
    {
        var ext = Path.GetExtension(file);
        if (ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            return ExtractPdfText(file);
        }

        if (ext.Equals(".xml", StringComparison.OrdinalIgnoreCase))
        {
            return new TextExtractionResult(ReadTextFile(file), "XML文本", "");
        }

        if (ext.Equals(".ofd", StringComparison.OrdinalIgnoreCase))
        {
            return ExtractOfdText(file);
        }

        if (ext.Equals(".zip", StringComparison.OrdinalIgnoreCase))
        {
            return ExtractZipText(file);
        }

        return new TextExtractionResult("", IsImage(ext) ? "图片未OCR" : "不支持文本提取", "");
    }

    private static TextExtractionResult ExtractPdfText(string file)
    {
        try
        {
            using var document = PdfDocument.Open(file);
            var builder = new StringBuilder();
            foreach (var page in document.GetPages())
            {
                builder.AppendLine(page.Text);
            }
            var text = NormalizeText(builder.ToString());
            return new TextExtractionResult(text, string.IsNullOrWhiteSpace(text) ? "PDF无文本" : "PDF文本", "");
        }
        catch (Exception ex)
        {
            return new TextExtractionResult("", "PDF读取失败", ex.Message);
        }
    }

    private static TextExtractionResult ExtractOfdText(string file)
    {
        try
        {
            using var stream = File.OpenRead(file);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            var builder = new StringBuilder();
            foreach (var entry in archive.Entries
                .Where(e => e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .Take(80))
            {
                using var entryStream = entry.Open();
                using var reader = new StreamReader(entryStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                builder.Append(' ').Append(reader.ReadToEnd());
            }
            var text = NormalizeText(Regex.Replace(builder.ToString(), "<[^>]+>", " "));
            return new TextExtractionResult(text, string.IsNullOrWhiteSpace(text) ? "OFD无文本" : "OFD文本", "");
        }
        catch (Exception ex)
        {
            return new TextExtractionResult(ReadTextFile(file), "OFD读取异常", ex.Message);
        }
    }

    private static TextExtractionResult ExtractZipText(string file)
    {
        try
        {
            using var stream = File.OpenRead(file);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            var entries = archive.Entries
                .Where(e => !string.IsNullOrWhiteSpace(e.Name))
                .ToList();
            var invoiceLike = entries
                .Where(e => SupportedExtensions.Contains(Path.GetExtension(e.Name), StringComparer.OrdinalIgnoreCase))
                .ToList();

            var builder = new StringBuilder();
            foreach (var entry in entries
                .Where(e => Path.GetExtension(e.Name).Equals(".xml", StringComparison.OrdinalIgnoreCase)
                    || Path.GetExtension(e.Name).Equals(".ofd", StringComparison.OrdinalIgnoreCase))
                .Take(50))
            {
                using var entryStream = entry.Open();
                using var reader = new StreamReader(entryStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                builder.Append(' ').Append(reader.ReadToEnd());
            }

            var text = NormalizeText(Regex.Replace(builder.ToString(), "<[^>]+>", " "));
            var extra = $"压缩包内文件{entries.Count}个，疑似发票文件{invoiceLike.Count}个";
            return new TextExtractionResult(text, "ZIP扫描", extra);
        }
        catch (Exception ex)
        {
            return new TextExtractionResult("", "ZIP读取失败", ex.Message);
        }
    }

    private static string ReadTextFile(string file)
    {
        try
        {
            var bytes = File.ReadAllBytes(file);
            var utf8 = Encoding.UTF8.GetString(bytes);
            return NormalizeText(TextEncodingFixer.Fix(utf8));
        }
        catch
        {
            return "";
        }
    }

    private static string ExtractInvoiceNo(string text)
    {
        return ExtractInvoiceNumbers(text).FirstOrDefault() ?? "";
    }

    private static List<string> ExtractInvoiceNumbers(string text)
    {
        var numbers = new List<string>();

        var oldStyle = Regex.Matches(text, @"发票代码\s*[:：]?\s*([0-9]{10,12}).{0,100}?发票号码\s*[:：]?\s*([0-9]{8})", RegexOptions.Singleline);
        numbers.AddRange(oldStyle.Select(m => NormalizeInvoiceNo(m.Groups[1].Value + m.Groups[2].Value)));

        var labeled = Regex.Matches(text, @"(?:发票号码|发票号|电子发票号码|数电票号码|发票No|InvoiceNumber|InvoiceNo|EIid)[_：:\s-]*([0-9]{8,30})", RegexOptions.IgnoreCase);
        numbers.AddRange(labeled.Select(m => NormalizeInvoiceNo(m.Groups[1].Value)));

        numbers.AddRange(Regex.Matches(text, @"(?<!\d)([0-9]{10,12})\s+([0-9]{8})(?!\d)")
            .Select(m => NormalizeInvoiceNo(m.Groups[1].Value + m.Groups[2].Value)));

        var dzfp = Regex.Matches(text, @"dzfp[_-]([0-9]{12,30})", RegexOptions.IgnoreCase);
        numbers.AddRange(dzfp.Select(m => NormalizeInvoiceNo(m.Groups[1].Value)));

        numbers.AddRange(Regex.Matches(text, @"(?<!\d)([0-9]{18,24})(?!\d)")
            .Select(m => NormalizeInvoiceNo(m.Groups[1].Value)));

        return numbers
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Where(v => v.Length >= 8)
            .Where(v => !LooksLikeNumericTaxId(v))
            .Where(v => !LooksLikeDateNoise(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(v => v.Length)
            .ToList();
    }

    private static string NormalizeInvoiceNo(string value)
    {
        value = Regex.Replace(value ?? "", @"\D", "");
        if (value.Length > 20 && !value.StartsWith("20", StringComparison.Ordinal))
        {
            value = value[..20];
        }
        return value;
    }

    private static string ExtractDate(string text)
    {
        var patterns = new[]
        {
            @"开票日期\s*[:：]?\s*(20\d{2})\s*[-_.年/]\s*(\d{1,2})\s*[-_.月/]\s*(\d{1,2})",
            @"开票时间\s*[:：]?\s*(20\d{2})\s*[-_.年/]\s*(\d{1,2})\s*[-_.月/]\s*(\d{1,2})",
            @"(20\d{2})\s*[-_.年/]\s*(\d{1,2})\s*[-_.月/]\s*(\d{1,2})"
        };

        foreach (var pattern in patterns)
        {
            var match = Regex.Match(text, pattern);
            if (match.Success)
            {
                var value = $"{int.Parse(match.Groups[1].Value):0000}-{int.Parse(match.Groups[2].Value):00}-{int.Parse(match.Groups[3].Value):00}";
                return DateTime.TryParse(value, out var parsed) ? parsed.ToString("yyyy-MM-dd") : "";
            }
        }

        var compact = Regex.Match(text, @"(?<!\d)(20\d{6})(?!\d)");
        if (compact.Success)
        {
            var value = compact.Groups[1].Value;
            var normalized = $"{value[..4]}-{value.Substring(4, 2)}-{value.Substring(6, 2)}";
            return DateTime.TryParse(normalized, out var parsed) ? parsed.ToString("yyyy-MM-dd") : "";
        }

        return "";
    }

    private static decimal ExtractAmount(string text)
    {
        var patterns = new[]
        {
            @"价税合计.{0,80}?[小写）\)]\s*[¥￥]?\s*([0-9]{1,8}(?:\.[0-9]{1,2})?)",
            @"小写\s*[）\)]?\s*[¥￥]?\s*([0-9]{1,8}(?:\.[0-9]{1,2})?)",
            @"发票金额\s*[：:】\]\s]*[¥￥]?\s*([0-9]{1,8}(?:\.[0-9]{1,2})?)",
            @"共\s*\d+\s*笔.{0,30}?合计\s*([0-9]{1,8}(?:\.[0-9]{1,2})?)\s*元",
            @"合计\s*([0-9]{1,8}(?:\.[0-9]{1,2})?)\s*元",
            @"金额\s*[：:】\]\s]*[¥￥]\s*([0-9]{1,8}(?:\.[0-9]{1,2})?)"
        };

        foreach (var pattern in patterns)
        {
            var match = Regex.Match(text, pattern, RegexOptions.Singleline);
            if (match.Success && TryParseReasonableAmount(match.Groups[1].Value, out var amount))
            {
                return amount;
            }
        }

        var currencyAmounts = Regex.Matches(text, @"[¥￥]\s*([0-9]{1,8}(?:\.[0-9]{1,2})?)")
            .Select(m => TryParseReasonableAmount(m.Groups[1].Value, out var amount) ? amount : 0m)
            .Where(amount => amount > 0)
            .ToList();
        return currencyAmounts.Count > 0 ? currencyAmounts.Max() : 0m;
    }

    private static string ExtractBuyerName(string text)
    {
        var patterns = new[]
        {
            @"购买方信息.{0,120}?名\s*称[:：]?\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,80})",
            @"购买方\s*名\s*称[:：]?\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,80})",
            @"购方\s*名\s*称[:：]?\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,80})",
            @"购\s*货\s*方.{0,40}?名\s*称[:：]?\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,80})",
            @"<[^>]*(?:BuyerName|Gmfmc|GMFMC)[^>]*>\s*([^<]{4,100})\s*</"
        };
        return ExtractNameByPatterns(text, patterns);
    }

    private static string ExtractBuyerTaxId(string text)
    {
        var patterns = new[]
        {
            @"购买方信息.{0,180}?(?:统一社会信用代码|纳税人识别号)[:：]?\s*([0-9A-Z]{15,20})",
            @"购买方.{0,180}?(?:统一社会信用代码|纳税人识别号)[:：]?\s*([0-9A-Z]{15,20})",
            @"购方.{0,180}?(?:统一社会信用代码|纳税人识别号)[:：]?\s*([0-9A-Z]{15,20})",
            @"<[^>]*(?:BuyerTaxNo|Gmfnsrsbh|GMFNSRSBH)[^>]*>\s*([^<]{15,20})\s*</"
        };

        foreach (var pattern in patterns)
        {
            var match = Regex.Match(text, pattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return Regex.Replace(match.Groups[1].Value.ToUpperInvariant(), @"[^0-9A-Z]", "");
            }
        }
        return "";
    }

    private static string ExtractVendor(string text)
    {
        var patterns = new[]
        {
            @"销售方信息.{0,140}?名\s*称[:：]?\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,80})",
            @"销售方\s*名\s*称[:：]?\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,80})",
            @"销方\s*名\s*称[:：]?\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,80})",
            @"销\s*货\s*方.{0,60}?名\s*称[:：]?\s*([\u4e00-\u9fffA-Za-z0-9（）()·\-]{4,80})",
            @"<[^>]*(?:SellerName|Seller|Xsfmc|XSFMC)[^>]*>\s*([^<]{4,100})\s*</"
        };

        var value = ExtractNameByPatterns(text, patterns);
        if (!string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        var fileNameVendor = Regex.Match(text, @"[0-9]+(?:\.[0-9]{1,2})?元-([^-\\/:]{4,60}?)-20\d{2}[._-]?\d{1,2}", RegexOptions.Singleline);
        if (fileNameVendor.Success)
        {
            value = CleanName(fileNameVendor.Groups[1].Value);
            if (LooksLikeValidVendor(value))
            {
                return value;
            }
        }

        return "";
    }

    private static string ExtractNameByPatterns(string text, IEnumerable<string> patterns)
    {
        foreach (var pattern in patterns)
        {
            var match = Regex.Match(text, pattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                continue;
            }

            var value = CleanName(match.Groups[1].Value);
            if (!string.IsNullOrWhiteSpace(value) && !LooksLikeNoiseName(value))
            {
                return value;
            }
        }
        return "";
    }

    private static string DetectCategory(string text)
    {
        if (ContainsAny(text, "通行费", "高速", "ETC", "收费公路", "过路费")) return "高速/通行费";
        if (ContainsAny(text, "滴滴", "DIDI", "网约车", "出行科技", "小桔")) return "滴滴/网约车";
        if (ContainsAny(text, "铁路", "火车票", "12306", "客票", "行程单")) return "火车/铁路";
        if (ContainsAny(text, "航空", "机票", "航旅", "机场")) return "飞机/机票";
        if (ContainsAny(text, "美团", "三快", "餐饮", "饭店", "餐厅", "饮食", "外卖")) return "餐饮";
        if (ContainsAny(text, "住宿", "酒店", "宾馆", "旅馆")) return "住宿";
        if (ContainsAny(text, "京东", "淘宝", "天猫", "拼多多", "采购", "五金", "电子", "科技")) return "采购";
        if (ContainsAny(text, "联通", "移动", "电信", "话费", "通信")) return "通信";
        return "其他发票";
    }

    private static bool LooksLikeInvoiceText(string text)
    {
        return ContainsAny(text, "发票", "电子发票", "数电票", "增值税", "价税合计", "开票日期", "发票号码", "发票代码");
    }

    private static bool LooksLikeInvoiceName(string fileName)
    {
        return ContainsAny(fileName, "发票", "invoice", "dzfp", "ofd", "数电票", "电子票");
    }

    private static bool LooksLikeValidVendor(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        value = CleanName(value);
        return value.Length is >= 4 and <= 80
            && !LooksLikeNoiseName(value)
            && ContainsAny(value, "公司", "分公司", "有限公司", "商贸", "科技", "餐", "饭", "酒店", "宾馆", "烧烤",
                "餐饮", "高速", "路桥", "服务区", "商行", "商店", "店", "中心");
    }

    private static bool LooksLikeNoiseName(string value)
    {
        if (ContainsAny(value, ":", "：", ";", "；", "项目名称", "规格型号", "统一社会信用代码", "纳税人识别号",
            "下载次数", "税率", "征收率", "单价", "金额", "发票代码", "发票号码", "开票日期", "银行账号", "账号",
            "复核人", "开票人", "收款人", "机器编号", "校验码", "电子支付标识", "密码区", "电话", "地址", "车牌号"))
        {
            return true;
        }

        var chineseCount = value.Count(ch => ch >= '\u4e00' && ch <= '\u9fff');
        return chineseCount < 2 || Regex.IsMatch(value, @"^[0-9A-Za-z%+\-*<>.]+$");
    }

    private static string CleanName(string value)
    {
        value = Regex.Replace(value ?? "", @"\s+", "").Trim();
        return value.Trim('：', ':', ';', '；', ',', '，');
    }

    private static decimal CalculateConfidence(InvoiceRecognitionAuditRow row)
    {
        decimal score = 0;
        if (row.IsLikelyInvoice) score += 0.2m;
        if (!string.IsNullOrWhiteSpace(row.InvoiceNumber)) score += 0.25m;
        if (!string.IsNullOrWhiteSpace(row.InvoiceDate)) score += 0.15m;
        if (row.Amount > 0) score += 0.2m;
        if (!string.IsNullOrWhiteSpace(row.Vendor)) score += 0.1m;
        if (!string.IsNullOrWhiteSpace(row.BuyerName) || !string.IsNullOrWhiteSpace(row.BuyerTaxId)) score += 0.1m;
        return Math.Min(score, 1m);
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

    private static bool IsImage(string ext)
    {
        return ext.Equals(".jpg", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".jpeg", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".png", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".bmp", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".webp", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".ico", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsMaintenanceArchiveFile(string file)
    {
        return file.Contains($"{Path.DirectorySeparatorChar}非PDF重复格式归档{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase)
            || file.Contains($"{Path.DirectorySeparatorChar}日期不符归档{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase)
            || file.Contains($"{Path.DirectorySeparatorChar}非发票归档{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase);
    }

    private static long SafeLength(string file)
    {
        try
        {
            return new FileInfo(file).Length;
        }
        catch
        {
            return 0;
        }
    }

    private static string SafeSha256(string file)
    {
        try
        {
            using var stream = File.OpenRead(file);
            return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
        }
        catch
        {
            return "";
        }
    }

    private static string NormalizeText(string text)
    {
        return TextEncodingFixer.Fix(Regex.Replace(text ?? "", @"\s+", " ")).Trim();
    }

    private static string Shorten(string value, int max)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "";
        }
        value = value.Trim();
        return value.Length <= max ? value : value[..max];
    }

    private static bool ContainsAny(string value, params string[] keywords)
    {
        return keywords.Any(k => value.Contains(k, StringComparison.OrdinalIgnoreCase));
    }

    private static void WriteWorkbook(string path, List<InvoiceRecognitionAuditRow> rows, string root)
    {
        using var workbook = new XLWorkbook();
        var summary = workbook.AddWorksheet("总览");
        summary.Cell(1, 1).Value = "指标";
        summary.Cell(1, 2).Value = "值";
        var summaryRows = new (string Key, object Value)[]
        {
            ("扫描目录", root),
            ("生成时间", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")),
            ("文件总数", rows.Count),
            ("通过", rows.Count(r => r.Status == "通过")),
            ("需复核", rows.Count(r => r.Status == "需复核")),
            ("异常", rows.Count(r => r.Status == "异常")),
            ("非发票/跳过", rows.Count(r => r.Status == "非发票/跳过")),
            ("缺发票号", rows.Count(r => r.IssueSummary.Contains("未识别发票号", StringComparison.OrdinalIgnoreCase))),
            ("金额为0/未识别", rows.Count(r => r.IssueSummary.Contains("未识别金额", StringComparison.OrdinalIgnoreCase))),
            ("购方不匹配", rows.Count(r => r.IssueSummary.Contains("购方", StringComparison.OrdinalIgnoreCase) && r.IssueSummary.Contains("不匹配", StringComparison.OrdinalIgnoreCase))),
            ("重复", rows.Count(r => !string.IsNullOrWhiteSpace(r.DuplicateStatus)))
        };
        for (var i = 0; i < summaryRows.Length; i++)
        {
            summary.Cell(i + 2, 1).Value = summaryRows[i].Key;
            summary.Cell(i + 2, 2).Value = summaryRows[i].Value.ToString();
        }
        summary.RangeUsed()?.SetAutoFilter();
        summary.Columns().AdjustToContents(10, 80);

        var sheet = workbook.AddWorksheet("识别明细");
        var headers = new[]
        {
            "状态", "风险", "置信度", "文件类型", "文本状态", "文本长度", "是否疑似发票",
            "发票号", "开票日期", "金额", "购方名称", "购方税号", "销售方", "类别",
            "重复状态", "重复来源", "问题说明", "文件名", "相对路径", "SHA256", "文本片段"
        };

        for (var c = 0; c < headers.Length; c++)
        {
            sheet.Cell(1, c + 1).Value = headers[c];
        }

        for (var r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            var values = new object[]
            {
                row.Status, row.Risk, row.Confidence, row.Extension, row.TextStatus, row.TextLength, row.IsLikelyInvoice ? "是" : "否",
                row.InvoiceNumber, row.InvoiceDate, row.Amount, row.BuyerName, row.BuyerTaxId, row.Vendor, row.Category,
                row.DuplicateStatus, row.DuplicateOf, row.IssueSummary, row.FileName, row.RelativePath, row.Sha256, row.TextPreview
            };

            for (var c = 0; c < values.Length; c++)
            {
                sheet.Cell(r + 2, c + 1).Value = values[c]?.ToString() ?? "";
            }
            sheet.Cell(r + 2, 10).Value = row.Amount;
        }

        var used = sheet.RangeUsed();
        if (used != null)
        {
            used.SetAutoFilter();
            used.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            used.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            used.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }
        sheet.SheetView.FreezeRows(1);
        sheet.Column(17).Width = 45;
        sheet.Column(19).Width = 60;
        sheet.Column(21).Width = 80;
        sheet.Columns(1, 16).AdjustToContents(8, 30);

        foreach (var row in rows.Select((value, index) => (value, index)))
        {
            var range = sheet.Range(row.index + 2, 1, row.index + 2, headers.Length);
            if (row.value.Status == "异常")
            {
                range.Style.Fill.BackgroundColor = XLColor.LightPink;
            }
            else if (row.value.Status == "需复核")
            {
                range.Style.Fill.BackgroundColor = XLColor.LightYellow;
            }
            else if (row.value.Status == "通过")
            {
                range.Style.Fill.BackgroundColor = XLColor.Honeydew;
            }
            else
            {
                range.Style.Fill.BackgroundColor = XLColor.LightGray;
            }
        }

        workbook.SaveAs(path);
    }

    private sealed record TextExtractionResult(string Text, string Status, string Extra);

    private sealed class InvoiceRecognitionAuditRow
    {
        public string Status { get; set; } = "";
        public string Risk { get; set; } = "";
        public decimal Confidence { get; set; }
        public string FileName { get; set; } = "";
        public string FilePath { get; set; } = "";
        public string RelativePath { get; set; } = "";
        public string Extension { get; set; } = "";
        public long SizeBytes { get; set; }
        public string Sha256 { get; set; } = "";
        public string TextStatus { get; set; } = "";
        public int TextLength { get; set; }
        public string TextPreview { get; set; } = "";
        public bool IsLikelyInvoice { get; set; }
        public string InvoiceNumber { get; set; } = "";
        public string InvoiceDate { get; set; } = "";
        public decimal Amount { get; set; }
        public string BuyerName { get; set; } = "";
        public string BuyerTaxId { get; set; } = "";
        public string Vendor { get; set; } = "";
        public string Category { get; set; } = "";
        public string DuplicateStatus { get; set; } = "";
        public string DuplicateOf { get; set; } = "";
        public string IssueSummary { get; set; } = "";
        public List<string> Issues { get; } = new();
    }
}
