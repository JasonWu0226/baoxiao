using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Drawing;
using HtmlAgilityPack;
using ZXing;
using ZXing.Common;

namespace ReimbursementMatcher;

public sealed class AiReviewService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private readonly WorkspaceService _workspace;
    private readonly Action<string> _log;

    public AiReviewService(WorkspaceService workspace, Action<string> log)
    {
        _workspace = workspace;
        _log = log;
    }

    public bool IsConfigured(AppConfig config)
    {
        return config.Ai.Enabled
            && !string.IsNullOrWhiteSpace(config.Ai.BaseUrl)
            && !string.IsNullOrWhiteSpace(config.Ai.ApiKey)
            && !string.IsNullOrWhiteSpace(config.Ai.Model);
    }

    public async Task<string> TestAsync(AppConfig config, CancellationToken ct = default)
    {
        EnsureConfigured(config);
        var content = await ChatAsync(config, config.Ai.Model, "你是报销发票判断助手。请只回复：OK", null, ct);
        return string.IsNullOrWhiteSpace(content) ? "模型无返回内容" : content.Trim();
    }

    public async Task<AiReviewResult> ReviewEmailAsync(AppConfig config, EmailAuditItem item, CancellationToken ct = default)
    {
        EnsureConfigured(config);
        var prompt = $@"你是中国报销发票邮件判断助手。请判断下面这封邮件是否包含发票或可下载发票。
只返回 JSON，不要解释，不要 Markdown。

JSON 格式：
{{""decision"":""has_invoice|no_invoice|needs_review"",""confidence"":0.0,""reason"":""一句中文原因"",""action"":""download|skip|manual_review"",""url"":""如果需要下载且有明确链接则填写，否则为空""}}

判断规则：
- 电子发票、数电票、增值税发票、报销凭证、行程单、PDF/OFD/XML/ZIP 发票附件，判断 has_invoice。
- 明确广告、银行还款提醒、普通通知、无附件无链接无发票语义，判断 no_invoice。
- 有平台/发票关键词但链接无法打开、附件信息冲突、只看到二维码或不确定，判断 needs_review。

邮件信息：
日期：{item.Date}
主题：{item.Subject}
下载状态：{item.Status}
规则判断：{item.RuleDecision}
有发票关键字：{item.HasInvoiceKeyword}
有平台关键字：{item.HasPlatformKeyword}
附件数：{item.AttachmentCount}
疑似发票附件：{item.HasLikelyInvoiceAttachment}
疑似发票链接：{item.HasLikelyInvoiceLink}
下载数：{item.DownloadedFileCount}
跳过数：{item.SkippedOrDuplicateCount}
判断原因：{item.Reason}
文件：
{item.Files}
链接：
{item.Urls}";

        var content = await ChatAsync(config, config.Ai.Model, prompt, null, ct);
        return ParseReviewResult(content);
    }

    public async Task<AiReviewResult> ReviewImageAsync(AppConfig config, string imagePath, string context, CancellationToken ct = default)
    {
        EnsureConfigured(config);
        var model = string.IsNullOrWhiteSpace(config.Ai.VisionModel) ? config.Ai.Model : config.Ai.VisionModel;
        var prompt = $@"你是中国报销发票图片判断助手。请判断这张图片是否是发票、电子发票截图、报销凭证、行程单，还是无关图片/二维码/图标。
只返回 JSON，不要解释，不要 Markdown。

JSON 格式：
{{""decision"":""has_invoice|no_invoice|needs_review"",""confidence"":0.0,""reason"":""一句中文原因"",""action"":""keep|archive|scan_qr|manual_review"",""url"":""""}}

判断规则：
- 如果图片本身是发票、电子发票、行程单、报销凭证，decision=has_invoice, action=keep。
- 如果只是 logo、favicon、广告、纯二维码且不能确认发票，decision=no_invoice 或 needs_review，action=archive 或 scan_qr。
- 看不清或无法确定，decision=needs_review, action=manual_review。

文件名/上下文：
{context}";

        var content = await ChatAsync(config, model, prompt, imagePath, ct);
        return ParseReviewResult(content);
    }

    private async Task<string> ChatAsync(AppConfig config, string model, string prompt, string? imagePath, CancellationToken ct)
    {
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
        using var request = new HttpRequestMessage(HttpMethod.Post, ChatEndpoint(config.Ai.BaseUrl));
        request.Headers.TryAddWithoutValidation("api-key", config.Ai.ApiKey);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", config.Ai.ApiKey);

        object userContent = prompt;
        if (!string.IsNullOrWhiteSpace(imagePath) && File.Exists(imagePath))
        {
            userContent = new object[]
            {
                new { type = "text", text = prompt },
                new { type = "image_url", image_url = new { url = ToDataUrl(imagePath) } }
            };
        }

        var body = new
        {
            model,
            messages = new object[]
            {
                new { role = "user", content = userContent }
            },
            temperature = 0.1,
            max_completion_tokens = 512
        };

        request.Content = new StringContent(JsonSerializer.Serialize(body, JsonOptions), Encoding.UTF8, "application/json");
        using var response = await http.SendAsync(request, ct);
        var responseText = await response.Content.ReadAsStringAsync(ct);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"大模型调用失败：HTTP {(int)response.StatusCode} {response.ReasonPhrase}；{Shorten(responseText, 500)}");
        }

        using var doc = JsonDocument.Parse(responseText);
        return doc.RootElement
            .GetProperty("choices")[0]
            .GetProperty("message")
            .GetProperty("content")
            .GetString() ?? "";
    }

    private static Uri ChatEndpoint(string baseUrl)
    {
        var trimmed = baseUrl.Trim().TrimEnd('/');
        if (trimmed.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
        {
            return new Uri(trimmed);
        }
        return new Uri($"{trimmed}/chat/completions");
    }

    private static string ToDataUrl(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        var mime = ext switch
        {
            ".jpg" or ".jpeg" => "image/jpeg",
            ".webp" => "image/webp",
            _ => "image/png"
        };
        return $"data:{mime};base64,{Convert.ToBase64String(File.ReadAllBytes(path))}";
    }

    private static AiReviewResult ParseReviewResult(string content)
    {
        var json = ExtractJson(content);
        try
        {
            var result = JsonSerializer.Deserialize<AiReviewResult>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new AiReviewResult();
            result.Decision = NormalizeDecision(result.Decision);
            result.Action ??= "";
            result.Reason ??= "";
            result.Url ??= "";
            return result;
        }
        catch
        {
            return new AiReviewResult
            {
                Decision = "needs_review",
                Confidence = 0,
                Reason = Shorten(content.Trim(), 300),
                Action = "manual_review"
            };
        }
    }

    private static string ExtractJson(string content)
    {
        content = content.Trim();
        content = Regex.Replace(content, "^```(?:json)?", "", RegexOptions.IgnoreCase).Trim();
        content = Regex.Replace(content, "```$", "").Trim();
        var match = Regex.Match(content, @"\{[\s\S]*\}");
        return match.Success ? match.Value : content;
    }

    private static string NormalizeDecision(string value)
    {
        value = (value ?? "").Trim().ToLowerInvariant();
        return value switch
        {
            "has_invoice" or "invoice" or "有发票" or "明确有发票" => "has_invoice",
            "no_invoice" or "none" or "无发票" or "明确无发票" => "no_invoice",
            _ => "needs_review"
        };
    }

    private void EnsureConfigured(AppConfig config)
    {
        if (!IsConfigured(config))
        {
            throw new InvalidOperationException("请先启用大模型，并填写 BaseUrl、API Key 和模型名称。");
        }
    }

    private static string Shorten(string value, int max)
    {
        return value.Length <= max ? value : value[..max];
    }
}

public sealed class NonPdfInvoiceMaintenanceService
{
    private readonly WorkspaceService _workspace;
    private readonly Action<string> _log;

    public NonPdfInvoiceMaintenanceService(WorkspaceService workspace, Action<string> log)
    {
        _workspace = workspace;
        _log = log;
    }

    public async Task<string> ProcessAsync(AppConfig config, CancellationToken ct = default)
    {
        var invoiceDir = _workspace.Resolve(config.InvoiceDir);
        if (!Directory.Exists(invoiceDir))
        {
            throw new DirectoryNotFoundException("发票目录不存在：" + invoiceDir);
        }

        var archiveRoot = _workspace.Resolve(Path.Combine(config.ArchiveDir, $"非PDF图片清理_{DateTime.Now:yyyyMMdd_HHmmss}"));
        var downloadedDir = Path.Combine(invoiceDir, "二维码取票下载");
        var ai = new AiReviewService(_workspace, _log);
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };

        var archived = 0;
        var qrDownloaded = 0;
        var aiKept = 0;
        var aiArchived = 0;
        var needsReview = 0;

        var files = Directory.EnumerateFiles(invoiceDir, "*.*", SearchOption.AllDirectories)
            .Where(f => !IsInsideGeneratedDir(f, invoiceDir))
            .Where(IsNonPdfCandidate)
            .ToList();

        foreach (var file in files)
        {
            ct.ThrowIfCancellationRequested();
            var ext = Path.GetExtension(file).ToLowerInvariant();
            if (ext == ".ico")
            {
                MoveToArchive(invoiceDir, archiveRoot, file, "ICO不是发票");
                archived++;
                continue;
            }

            if (IsImage(ext))
            {
                var qr = TryDecodeQr(file);
                if (!string.IsNullOrWhiteSpace(qr))
                {
                    var saved = await TryDownloadInvoiceFromUrlAsync(http, qr, downloadedDir, ct);
                    if (saved.Count > 0)
                    {
                        MoveToArchive(invoiceDir, archiveRoot, file, "二维码原图已取票");
                        qrDownloaded += saved.Count;
                        archived++;
                        continue;
                    }
                    needsReview++;
                    _log($"二维码未能直接取票：{Path.GetFileName(file)} -> {qr}");
                    continue;
                }

                if (ai.IsConfigured(config))
                {
                    var result = await ai.ReviewImageAsync(config, file, Path.GetRelativePath(invoiceDir, file), ct);
                    if (result.Decision == "no_invoice" && result.Confidence >= config.Ai.ConfidenceThreshold)
                    {
                        MoveToArchive(invoiceDir, archiveRoot, file, $"AI判断无发票：{result.Reason}");
                        aiArchived++;
                    }
                    else if (result.Decision == "has_invoice" && result.Confidence >= config.Ai.ConfidenceThreshold)
                    {
                        aiKept++;
                        _log($"AI保留图片发票：{Path.GetFileName(file)}；{result.Reason}");
                    }
                    else
                    {
                        needsReview++;
                        _log($"图片需要人工确认：{Path.GetFileName(file)}；{result.Reason}");
                    }
                }
                else
                {
                    needsReview++;
                }
            }
        }

        return $"非PDF/二维码处理完成：ICO/无关归档 {archived}；二维码下载发票 {qrDownloaded}；AI保留图片发票 {aiKept}；AI归档图片 {aiArchived}；待人工确认 {needsReview}。";
    }

    private static bool IsNonPdfCandidate(string file)
    {
        var ext = Path.GetExtension(file).ToLowerInvariant();
        return ext is ".ico" or ".png" or ".jpg" or ".jpeg" or ".webp";
    }

    private static bool IsImage(string ext)
    {
        return ext is ".png" or ".jpg" or ".jpeg" or ".webp";
    }

    private static bool IsInsideGeneratedDir(string file, string invoiceDir)
    {
        var relative = Path.GetRelativePath(invoiceDir, file).Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
        return relative.StartsWith("二维码取票下载" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
            || relative.Contains($"{Path.DirectorySeparatorChar}非PDF图片清理_", StringComparison.OrdinalIgnoreCase)
            || relative.Contains($"{Path.DirectorySeparatorChar}非PDF重复格式_", StringComparison.OrdinalIgnoreCase);
    }

    private static string TryDecodeQr(string file)
    {
        try
        {
            using var original = (Bitmap)Image.FromFile(file);
            using var bitmap = new Bitmap(original.Width, original.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.DrawImage(original, 0, 0, original.Width, original.Height);
            }

            var rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
            var data = bitmap.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, bitmap.PixelFormat);
            try
            {
                var bytes = new byte[Math.Abs(data.Stride) * data.Height];
                Marshal.Copy(data.Scan0, bytes, 0, bytes.Length);
                var source = new RGBLuminanceSource(bytes, data.Width, data.Height, RGBLuminanceSource.BitmapFormat.BGR24);
                var reader = new BarcodeReaderGeneric
                {
                    AutoRotate = true,
                    Options = new DecodingOptions { TryHarder = true }
                };
                return reader.Decode(source)?.Text ?? "";
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
        }
        catch
        {
            return "";
        }
    }

    private static async Task<List<string>> TryDownloadInvoiceFromUrlAsync(HttpClient http, string url, string outputDir, CancellationToken ct)
    {
        var saved = new List<string>();
        if (!LooksLikeInvoiceUrl(url))
        {
            return saved;
        }

        Directory.CreateDirectory(outputDir);
        try
        {
            using var response = await http.GetAsync(url, ct);
            var contentType = response.Content.Headers.ContentType?.MediaType ?? "";
            var bytes = await response.Content.ReadAsByteArrayAsync(ct);
            if (!response.IsSuccessStatusCode)
            {
                return saved;
            }

            if (IsDownloadableInvoice(contentType, url, bytes))
            {
                var path = UniquePath(Path.Combine(outputDir, SafeFileName($"二维码取票_{DateTime.Now:yyyyMMdd_HHmmss}{GuessExtension(contentType, url, bytes)}")));
                await File.WriteAllBytesAsync(path, bytes, ct);
                saved.Add(path);
                return saved;
            }

            if (contentType.Contains("html", StringComparison.OrdinalIgnoreCase))
            {
                var html = Encoding.UTF8.GetString(bytes);
                foreach (var link in ExtractLinksFromHtml(html, url).Where(LooksLikeInvoiceUrl).Take(8))
                {
                    saved.AddRange(await TryDownloadInvoiceFromUrlAsync(http, link, outputDir, ct));
                    if (saved.Count > 0) break;
                }
            }
        }
        catch
        {
            return saved;
        }

        return saved;
    }

    private static List<string> ExtractLinksFromHtml(string html, string baseUrl)
    {
        var doc = new HtmlAgilityPack.HtmlDocument();
        doc.LoadHtml(html);
        var result = new List<string>();
        var nodes = doc.DocumentNode.SelectNodes("//*[@href or @src]")?.AsEnumerable() ?? Enumerable.Empty<HtmlNode>();
        foreach (var attr in nodes)
        {
            foreach (var name in new[] { "href", "src" })
            {
                var value = attr.GetAttributeValue(name, "");
                if (string.IsNullOrWhiteSpace(value)) continue;
                if (Uri.TryCreate(new Uri(baseUrl), value, out var uri)) result.Add(uri.ToString());
            }
        }
        return result.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static bool LooksLikeInvoiceUrl(string url)
    {
        return url.Contains("invoice", StringComparison.OrdinalIgnoreCase)
            || url.Contains("fapiao", StringComparison.OrdinalIgnoreCase)
            || url.Contains("pdf", StringComparison.OrdinalIgnoreCase)
            || url.Contains("ofd", StringComparison.OrdinalIgnoreCase)
            || url.Contains("发票", StringComparison.OrdinalIgnoreCase)
            || url.Contains("download", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsDownloadableInvoice(string contentType, string url, byte[] bytes)
    {
        if (contentType.Contains("pdf", StringComparison.OrdinalIgnoreCase)
            || contentType.Contains("zip", StringComparison.OrdinalIgnoreCase)
            || contentType.Contains("xml", StringComparison.OrdinalIgnoreCase)
            || contentType.Contains("ofd", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var ext = Path.GetExtension(Uri.TryCreate(url, UriKind.Absolute, out var uri) ? uri.AbsolutePath : url).ToLowerInvariant();
        if (ext is ".pdf" or ".zip" or ".xml" or ".ofd")
        {
            return true;
        }

        return bytes.Length > 4 && bytes[0] == 0x25 && bytes[1] == 0x50 && bytes[2] == 0x44 && bytes[3] == 0x46;
    }

    private static string GuessExtension(string contentType, string url, byte[] bytes)
    {
        var ext = Path.GetExtension(Uri.TryCreate(url, UriKind.Absolute, out var uri) ? uri.AbsolutePath : url).ToLowerInvariant();
        if (ext is ".pdf" or ".zip" or ".xml" or ".ofd") return ext;
        if (contentType.Contains("zip", StringComparison.OrdinalIgnoreCase)) return ".zip";
        if (contentType.Contains("xml", StringComparison.OrdinalIgnoreCase)) return ".xml";
        if (contentType.Contains("ofd", StringComparison.OrdinalIgnoreCase)) return ".ofd";
        return ".pdf";
    }

    private static void MoveToArchive(string invoiceDir, string archiveRoot, string file, string reason)
    {
        var relative = Path.GetRelativePath(invoiceDir, file);
        var target = UniquePath(Path.Combine(archiveRoot, relative));
        Directory.CreateDirectory(Path.GetDirectoryName(target)!);
        File.Move(file, target);
        File.WriteAllText(target + ".reason.txt", reason + Environment.NewLine, Encoding.UTF8);
    }

    private static string SafeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(name.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
        cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim(' ', '.');
        return string.IsNullOrWhiteSpace(cleaned) ? "file" : cleaned;
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
}
