using System.Net.Http.Headers;
using System.IO.Compression;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using UglyToad.PdfPig;

namespace ReimbursementMatcher;

public sealed class EmailDownloader
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private static readonly HashSet<string> AllowedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".pdf", ".ofd", ".zip", ".png", ".jpg", ".jpeg", ".xlsx", ".xls", ".xml"
    };

    private static readonly string[] InvoiceUrlHints =
    [
        "invoice", "fapiao", "fp", "pdf", "download", "ofd", "发票", "下载"
    ];

    private static readonly string[] InvoiceMailKeywords =
    [
        "发票", "电子发票", "增值税发票", "专票", "普票", "数电票", "报销", "开票",
        "发票金额", "报销凭证", "电子报销凭证", "行程单", "客票行程单",
        "invoice", "fapiao", "dzfp", ".pdf", ".ofd", ".xml", ".zip"
    ];

    private readonly WorkspaceService _workspace;
    private readonly Action<string> _log;

    public EmailDownloader(WorkspaceService workspace, Action<string> log)
    {
        _workspace = workspace;
        _log = log;
    }

    public async Task<string> DownloadAsync(AppConfig config, string password, bool onlyAbnormal = false, CancellationToken ct = default)
    {
        if (!config.Email.Enabled)
        {
            throw new InvalidOperationException("邮箱下载未启用。");
        }
        if (string.IsNullOrWhiteSpace(config.Email.User))
        {
            throw new InvalidOperationException("邮箱账号为空。");
        }
        if (string.IsNullOrWhiteSpace(password))
        {
            throw new InvalidOperationException("请输入邮箱授权码。");
        }

        var outputDir = _workspace.Resolve(config.Email.OutputDir);
        Directory.CreateDirectory(outputDir);
        var rows = new List<Dictionary<string, string>>();
        var existingHashes = BuildExistingHashIndex(outputDir);
        var existingPdfKeys = BuildExistingPdfKeyIndex(outputDir);
        var processingStore = LoadProcessingStore(config);
        processingStore.Invoices = BuildExistingInvoiceIndex(outputDir);
        _log($"基于实际发票文件建立查重索引：{processingStore.Invoices.Count} 个");
        var decisionStore = LoadDecisionStore(config);

        using var client = new ImapClient();
        _log("连接邮箱 IMAP...");
        await client.ConnectAsync(config.Email.Host, config.Email.Port, SecureSocketOptions.SslOnConnect, ct);
        await client.AuthenticateAsync(config.Email.User, password, ct);

        var folder = string.Equals(config.Email.Folder, "inbox", StringComparison.OrdinalIgnoreCase)
            ? client.Inbox
            : await client.GetFolderAsync(config.Email.Folder, ct);
        await folder.OpenAsync(FolderAccess.ReadOnly, ct);

        var start = DateTime.Parse(config.Email.Start).AddDays(-1);
        var end = DateTime.Parse(config.Email.End).AddDays(2);
        var ids = await folder.SearchAsync(SearchQuery.DeliveredAfter(start).And(SearchQuery.DeliveredBefore(end)), ct);
        _log($"找到候选邮件 {ids.Count} 封");

        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(25) };
        for (var i = 0; i < ids.Count; i++)
        {
            var msg = await folder.GetMessageAsync(ids[i], ct);
            var date = msg.Date.LocalDateTime.ToString("yyyy-MM-dd");
            var subject = TextEncodingFixer.Fix(msg.Subject ?? "");
            var sender = msg.From.Mailboxes.FirstOrDefault()?.Name
                ?? msg.From.Mailboxes.FirstOrDefault()?.Address
                ?? msg.From.ToString();
            sender = TextEncodingFixer.Fix(sender);
            var messageId = msg.MessageId ?? "";
            var messageKey = $"{config.Email.Folder}:{ids[i]}";
            var prefix = SafeFileName($"发票_{date.Replace("-", "")}_{CleanSender(sender)}_{ids[i]}");

            var attachmentParts = msg.Attachments.OfType<MimePart>()
                .Where(p => AllowedExtensions.Contains(Path.GetExtension(p.FileName ?? "")))
                .ToList();
            var attachmentTotal = attachmentParts.Count;
            var attachmentNames = attachmentParts.Select(p => TextEncodingFixer.Fix(p.FileName ?? "attachment")).ToList();
            var links = ExtractLinks(msg).Where(LooksLikeInvoiceLink).GroupBy(l => l.Url).Select(g => g.First()).ToList();
            var keywordHits = FindKeywordHits($"{subject}\n{msg.TextBody}\n{StripHtml(msg.HtmlBody)}\n{string.Join("\n", attachmentNames)}\n{string.Join("\n", links.Select(l => l.Url))}");
            var isInvoiceMail = keywordHits.Count > 0;
            if (onlyAbnormal && !IsPreviouslyAbnormal(processingStore, messageKey, isInvoiceMail, attachmentTotal, links.Count))
            {
                continue;
            }
            if (IsManuallyMarkedNoInvoice(decisionStore, messageKey))
            {
                rows.Add(new Dictionary<string, string>
                {
                    ["kind"] = "message",
                    ["status"] = "人工确认无发票跳过",
                    ["date"] = date,
                    ["subject"] = subject,
                    ["sender"] = sender,
                    ["msg_id"] = ids[i].ToString(),
                    ["message_id"] = messageId,
                    ["message_key"] = messageKey,
                    ["attachment_total"] = attachmentTotal.ToString(),
                    ["link_candidate_total"] = links.Count.ToString(),
                    ["saved_file_count"] = "0",
                    ["error"] = "该邮件已被人工确认为无发票，本次不再下载"
                });
                continue;
            }
            if (IsPreviouslyCompleted(processingStore, messageKey, isInvoiceMail, attachmentTotal, links.Count))
            {
                var existingFiles = ExistingFilesForMessage(processingStore, messageKey);
                rows.Add(new Dictionary<string, string>
                {
                    ["kind"] = "message",
                    ["status"] = "已处理跳过",
                    ["download_status"] = existingFiles.Count > 0 ? "已存在" : "",
                    ["date"] = date,
                    ["subject"] = subject,
                    ["sender"] = sender,
                    ["msg_id"] = ids[i].ToString(),
                    ["message_id"] = messageId,
                    ["message_key"] = messageKey,
                    ["attachment_total"] = attachmentTotal.ToString(),
                    ["link_candidate_total"] = links.Count.ToString(),
                    ["saved_file_count"] = existingFiles.Count.ToString(),
                    ["file"] = string.Join("；", existingFiles),
                    ["error"] = existingFiles.Count > 0
                        ? $"该邮件此前已成功处理，本地已存在 {existingFiles.Count} 个文件，本次不重复下载"
                        : "该邮件此前已成功处理且无异常，本次不重复下载"
                });
                continue;
            }

            var existingRecord = processingStore.Messages.TryGetValue(messageKey, out var oldRecord)
                ? oldRecord
                : new EmailProcessingRecord { MessageKey = messageKey };
            existingRecord.Attempts += 1;

            if (attachmentTotal == 0 && links.Count == 0)
            {
                var status = isInvoiceMail ? "异常" : "非发票邮件";
                var error = isInvoiceMail ? "无附件" : "无关键词";
                rows.Add(MessageRow(status, date, subject, sender, ids[i].ToString(), messageId, messageKey, attachmentTotal, links.Count, 0, 0, error, "", keywordHits, attachmentNames));
                processingStore.Messages[messageKey] = BuildProcessingRecord(
                    existingRecord,
                    messageKey,
                    messageId,
                    date,
                    subject,
                    sender,
                    status,
                    isInvoiceMail ? "异常" : "",
                    error,
                    attachmentTotal,
                    links.Count,
                    0,
                    0,
                    keywordHits,
                    attachmentNames,
                    []);
                continue;
            }

            var before = rows.Count(r => r.TryGetValue("file", out var f) && !string.IsNullOrWhiteSpace(f));

            var attachmentRows = await SaveAttachmentsAsync(msg, outputDir, prefix, subject, sender, date, ids[i].ToString(), messageKey, config.Email.Start, existingHashes, existingPdfKeys, processingStore, ct);
            AttachMessageKey(attachmentRows, messageKey);
            rows.AddRange(attachmentRows);
            var linkRows = HasResolvedInvoiceFromAttachments(attachmentRows)
                ? BuildSkippedLinkRows(links, date, subject, sender, ids[i].ToString(), messageKey)
                : await DownloadLinksAsync(links, http, outputDir, prefix, subject, sender, date, ids[i].ToString(), messageKey, config.Email.Start, existingHashes, existingPdfKeys, processingStore, ct);
            AttachMessageKey(linkRows, messageKey);
            rows.AddRange(linkRows);

            var saved = rows.Count(r => r.TryGetValue("file", out var f) && !string.IsNullOrWhiteSpace(f)) - before;
            var duplicate = attachmentRows.Concat(linkRows).Count(r => r.GetValueOrDefault("status") is "重复跳过" or "重复发票");
            var hasRisk = attachmentRows.Concat(linkRows).Any(IsRiskRow);
            var processedRows = attachmentRows.Concat(linkRows).ToList();
            var onlyPdfSkipped = processedRows.Count > 0 && processedRows.All(r => r.GetValueOrDefault("status") == "PDF已存在跳过");
            var hasPdf = processedRows.Any(r => Path.GetExtension(r.GetValueOrDefault("file", "")).Equals(".pdf", StringComparison.OrdinalIgnoreCase));
            var hasNoInvoiceNumber = processedRows.Any(r => r.GetValueOrDefault("error", "").Contains("无发票号"));
            var downloadStatus = "";
            var messageStatus = saved > 0 && hasPdf && !hasNoInvoiceNumber
                ? "正常已下载"
                : duplicate > 0
                    ? "重复发票"
                    : onlyPdfSkipped
                        ? "PDF已存在跳过"
                        : saved > 0 && !hasPdf
                            ? "异常"
                            : hasRisk
                                ? "异常"
                                : "未下载到文件";
            if (messageStatus == "正常已下载") downloadStatus = "正常已下载";
            else if (messageStatus == "重复发票") downloadStatus = "重复已忽略";
            else if (messageStatus == "异常" || messageStatus == "未下载到文件") downloadStatus = "异常";

            var messageError = BuildMessageError(processedRows, saved, hasPdf, hasNoInvoiceNumber);
            rows.Add(MessageRow(messageStatus, date, subject, sender, ids[i].ToString(), messageId, messageKey, attachmentTotal, links.Count, saved, duplicate, messageError, downloadStatus, keywordHits, attachmentNames));

            processingStore.Messages[messageKey] = BuildProcessingRecord(
                existingRecord,
                messageKey,
                messageId,
                date,
                subject,
                sender,
                messageStatus,
                downloadStatus,
                messageError,
                attachmentTotal,
                links.Count,
                saved,
                duplicate,
                keywordHits,
                attachmentNames,
                processedRows);

            if ((i + 1) % 50 == 0)
            {
                _log($"已处理 {i + 1}/{ids.Count} 封");
            }

            var delay = Math.Max(0, config.Email.RequestIntervalSec);
            if (delay > 0)
            {
                await Task.Delay(TimeSpan.FromSeconds(delay), ct);
            }
        }

        await client.DisconnectAsync(true, ct);
        SaveProcessingStore(config, processingStore);
        var manifest = Path.Combine(outputDir, $"邮箱发票下载清单_{DateTime.Now:yyyyMMdd_HHmmss}.csv");
        WriteManifest(manifest, rows);
        _log($"下载完成：{rows.Count(r => r.ContainsKey("file"))} 个文件");
        _log($"清单：{manifest}");
        return manifest;
    }

    private EmailProcessingStore LoadProcessingStore(AppConfig config)
    {
        var path = ProcessingStorePath(config);
        if (!File.Exists(path))
        {
            return new EmailProcessingStore();
        }
        try
        {
            return JsonSerializer.Deserialize<EmailProcessingStore>(File.ReadAllText(path, Encoding.UTF8), JsonOptions) ?? new EmailProcessingStore();
        }
        catch
        {
            return new EmailProcessingStore();
        }
    }

    private EmailDecisionStore LoadDecisionStore(AppConfig config)
    {
        var path = Path.Combine(_workspace.Resolve(config.RuleDir), "邮件判断记录.json");
        if (!File.Exists(path))
        {
            return new EmailDecisionStore();
        }

        try
        {
            return JsonSerializer.Deserialize<EmailDecisionStore>(File.ReadAllText(path, Encoding.UTF8), JsonOptions) ?? new EmailDecisionStore();
        }
        catch
        {
            return new EmailDecisionStore();
        }
    }

    private void SaveProcessingStore(AppConfig config, EmailProcessingStore store)
    {
        store.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var path = ProcessingStorePath(config);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, JsonSerializer.Serialize(store, JsonOptions) + Environment.NewLine, Encoding.UTF8);
    }

    private string ProcessingStorePath(AppConfig config)
    {
        return Path.Combine(_workspace.Resolve(config.RuleDir), "邮箱处理记录.json");
    }

    private static bool IsPreviouslyCompleted(EmailProcessingStore store, string messageKey, bool currentIsInvoiceMail, int currentAttachmentTotal, int currentLinkCount)
    {
        if (!store.Messages.TryGetValue(messageKey, out var record))
        {
            return false;
        }
        if (record.HasRisk || record.Status is "失败" or "需人工确认" or "未下载到文件" or "异常" or "链接取票待处理")
        {
            return false;
        }

        var currentHasInvoiceSignal = currentIsInvoiceMail || currentAttachmentTotal > 0 || currentLinkCount > 0;
        if (record.Status is "无发票内容已跳过" or "非发票邮件" or "人工确认无发票跳过")
        {
            return !currentHasInvoiceSignal;
        }

        return !IsInvoiceRecordWithoutExistingFile(record, currentHasInvoiceSignal);
    }

    private static bool IsPreviouslyAbnormal(EmailProcessingStore store, string messageKey, bool currentIsInvoiceMail, int currentAttachmentTotal, int currentLinkCount)
    {
        var currentHasInvoiceSignal = currentIsInvoiceMail || currentAttachmentTotal > 0 || currentLinkCount > 0;
        return store.Messages.TryGetValue(messageKey, out var record)
            && (record.HasRisk
                || record.Status is "失败" or "需人工确认" or "未下载到文件" or "异常" or "链接取票待处理"
                || IsInvoiceRecordWithoutExistingFile(record, currentHasInvoiceSignal));
    }

    private static List<string> ExistingFilesForMessage(EmailProcessingStore store, string messageKey)
    {
        if (!store.Messages.TryGetValue(messageKey, out var record))
        {
            return [];
        }

        return record.Files
            .Concat(record.SavedFiles.Select(f => f.Path))
            .Where(p => !string.IsNullOrWhiteSpace(p) && File.Exists(p))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool IsInvoiceRecordWithoutExistingFile(EmailProcessingRecord record, bool currentHasInvoiceSignal)
    {
        var hasInvoiceSignal = currentHasInvoiceSignal
            || record.IsInvoiceMail
            || record.AttachmentTotal > 0
            || record.LinkCandidateTotal > 0
            || record.Status is "正常已下载" or "已下载" or "PDF已存在跳过" or "重复发票" or "重复已存在" or "已处理跳过";

        if (!hasInvoiceSignal)
        {
            return false;
        }

        var files = record.Files
            .Concat(record.SavedFiles.Select(f => f.Path))
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return files.Count == 0 || files.Any(p => !File.Exists(p));
    }

    private static bool IsManuallyMarkedNoInvoice(EmailDecisionStore store, string messageKey)
    {
        return store.Messages.TryGetValue(messageKey, out var record)
            && record.Decision == EmailAuditService.NoInvoice;
    }

    private static void AttachMessageKey(IEnumerable<Dictionary<string, string>> rows, string messageKey)
    {
        foreach (var row in rows)
        {
            row["message_key"] = messageKey;
        }
    }

    private static bool IsRiskRow(Dictionary<string, string> row)
    {
        var status = row.GetValueOrDefault("status", "");
        var error = row.GetValueOrDefault("error", "");
        return status is "失败" or "需人工确认" or "未下载到文件" or "待核验" or "异常"
            || (!string.IsNullOrWhiteSpace(error) && status is not "PDF已存在跳过" and not "重复发票");
    }

    private static bool HasResolvedInvoiceFromAttachments(IEnumerable<Dictionary<string, string>> rows)
    {
        foreach (var row in rows.Where(r => r.GetValueOrDefault("kind") is "attachment" or "zip_entry"))
        {
            var status = row.GetValueOrDefault("status");
            var file = row.GetValueOrDefault("file", "");
            var duplicateOf = row.GetValueOrDefault("duplicate_of", "");
            var hasPdfFile = Path.GetExtension(file).Equals(".pdf", StringComparison.OrdinalIgnoreCase)
                || Path.GetExtension(duplicateOf).Equals(".pdf", StringComparison.OrdinalIgnoreCase);
            if (hasPdfFile && status is "已下载" or "正常已下载" or "重复发票" or "重复跳过" or "PDF已存在跳过")
            {
                return true;
            }
        }
        return false;
    }

    private static List<Dictionary<string, string>> BuildSkippedLinkRows(List<LinkCandidate> links, string date, string subject, string sender, string msgId, string messageKey)
    {
        return links.Select(link => new Dictionary<string, string>
        {
            ["kind"] = "link",
            ["status"] = "附件已取得发票，链接跳过",
            ["download_status"] = "已跳过",
            ["date"] = date,
            ["subject"] = subject,
            ["sender"] = sender,
            ["msg_id"] = msgId,
            ["message_key"] = messageKey,
            ["url"] = link.Url,
            ["error"] = ""
        }).ToList();
    }

    private static Dictionary<string, string> SavedFileRow(
        string kind,
        SaveResult save,
        string date,
        string subject,
        string sender,
        string msgId,
        string messageKey,
        string url,
        string startDate,
        EmailProcessingStore processingStore)
    {
        var metadata = File.Exists(save.Path) ? ExtractInvoiceMetadata(save.Path) : new InvoiceMetadata();
        var row = new Dictionary<string, string>
        {
            ["kind"] = kind,
            ["status"] = save.Status,
            ["download_status"] = save.Status == "已下载" ? "正常已下载" : "",
            ["date"] = date,
            ["subject"] = subject,
            ["sender"] = sender,
            ["msg_id"] = msgId,
            ["message_key"] = messageKey,
            ["url"] = url,
            ["file"] = save.Path,
            ["md5"] = save.Md5,
            ["sha256"] = save.Sha256,
            ["duplicate_of"] = save.DuplicateOf,
            ["invoice_number"] = metadata.InvoiceNumber,
            ["invoice_amount"] = metadata.Amount,
            ["invoice_date"] = metadata.Date,
            ["text_extract_status"] = metadata.TextExtractStatus
        };

        if (save.Status != "已下载")
        {
            row["status"] = "重复发票";
            row["download_status"] = "重复已忽略";
            row["error"] = "重复发票";
            var duplicate = FindDuplicateByHash(processingStore, save.Md5, save.Sha256);
            row["duplicate_source_mail_id"] = duplicate.MessageKey;
            row["duplicate_of"] = string.IsNullOrWhiteSpace(duplicate.InvoiceKey) ? save.DuplicateOf : duplicate.InvoiceKey;
            return row;
        }

        if (Path.GetExtension(save.Path).Equals(".pdf", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(metadata.InvoiceNumber))
        {
            row["error"] = "无发票号";
        }
        if (IsInvoiceDateBeforeStart(metadata.Date, startDate))
        {
            row["status"] = "异常";
            row["download_status"] = "异常";
            row["error"] = AppendReason(row.GetValueOrDefault("error", ""), "发票日期早于开始时间");
        }

        var duplicateInvoice = FindDuplicateInvoice(processingStore, save.Md5, save.Sha256, metadata);
        if (!string.IsNullOrWhiteSpace(duplicateInvoice.InvoiceKey))
        {
            row["status"] = "重复发票";
            row["download_status"] = "重复已忽略";
            row["error"] = "重复发票";
            row["duplicate_of"] = duplicateInvoice.InvoiceKey;
            row["duplicate_source_mail_id"] = duplicateInvoice.MessageKey;
            TryDelete(save.Path);
            row["file"] = "";
            return row;
        }

        var invoiceKey = BuildInvoiceKey(save.Md5, metadata);
        processingStore.Invoices[invoiceKey] = new EmailInvoiceRecord
        {
            InvoiceKey = invoiceKey,
            MessageKey = messageKey,
            MessageId = msgId,
            Subject = subject,
            Sender = sender,
            Date = date,
            Path = save.Path,
            Md5 = save.Md5,
            Sha256 = save.Sha256,
            InvoiceNumber = metadata.InvoiceNumber,
            InvoiceAmount = metadata.Amount,
            InvoiceDate = metadata.Date,
            UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        };
        return row;
    }

    private static EmailInvoiceRecord FindDuplicateByHash(EmailProcessingStore store, string md5, string sha256)
    {
        return store.Invoices.Values.FirstOrDefault(i =>
            !string.IsNullOrWhiteSpace(md5) && i.Md5 == md5
            || !string.IsNullOrWhiteSpace(sha256) && i.Sha256 == sha256) ?? new EmailInvoiceRecord();
    }

    private static EmailInvoiceRecord FindDuplicateInvoice(EmailProcessingStore store, string md5, string sha256, InvoiceMetadata metadata)
    {
        var byHash = FindDuplicateByHash(store, md5, sha256);
        if (!string.IsNullOrWhiteSpace(byHash.InvoiceKey)) return byHash;

        if (string.IsNullOrWhiteSpace(metadata.InvoiceNumber))
        {
            return new EmailInvoiceRecord();
        }

        return store.Invoices.Values.FirstOrDefault(i =>
            i.InvoiceNumber == metadata.InvoiceNumber
            && (string.IsNullOrWhiteSpace(metadata.Amount) || string.IsNullOrWhiteSpace(i.InvoiceAmount) || i.InvoiceAmount == metadata.Amount)
            && (string.IsNullOrWhiteSpace(metadata.Date) || string.IsNullOrWhiteSpace(i.InvoiceDate) || i.InvoiceDate == metadata.Date)) ?? new EmailInvoiceRecord();
    }

    private static string BuildInvoiceKey(string md5, InvoiceMetadata metadata)
    {
        return !string.IsNullOrWhiteSpace(metadata.InvoiceNumber)
            ? $"number:{metadata.InvoiceNumber}|amount:{metadata.Amount}|date:{metadata.Date}"
            : $"md5:{md5}";
    }

    private static bool IsInvoiceDateBeforeStart(string invoiceDate, string startDate)
    {
        return DateTime.TryParse(invoiceDate, out var invoice)
            && DateTime.TryParse(startDate, out var start)
            && invoice.Date < start.Date;
    }

    private static string AppendReason(string current, string reason)
    {
        return string.IsNullOrWhiteSpace(current)
            ? reason
            : current.Contains(reason, StringComparison.OrdinalIgnoreCase)
                ? current
                : current + "；" + reason;
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(path) && File.Exists(path)) File.Delete(path);
        }
        catch
        {
            // If the duplicate file is locked, keep the status record and leave cleanup for later.
        }
    }

    private static Dictionary<string, string> MessageRow(
        string status,
        string date,
        string subject,
        string sender,
        string msgId,
        string messageId,
        string messageKey,
        int attachmentTotal,
        int linkCount,
        int saved,
        int duplicate,
        string error,
        string downloadStatus,
        List<string> keywordHits,
        List<string> attachmentNames) => new()
    {
        ["kind"] = "message",
        ["status"] = status,
        ["download_status"] = downloadStatus,
        ["date"] = date,
        ["subject"] = subject,
        ["sender"] = sender,
        ["msg_id"] = msgId,
        ["message_id"] = messageId,
        ["message_key"] = messageKey,
        ["attachment_total"] = attachmentTotal.ToString(),
        ["link_candidate_total"] = linkCount.ToString(),
        ["saved_file_count"] = saved.ToString(),
        ["duplicate_count"] = duplicate.ToString(),
        ["error"] = error,
        ["keyword_hits"] = string.Join("；", keywordHits),
        ["attachment_names"] = string.Join("；", attachmentNames)
    };

    private static EmailProcessingRecord BuildProcessingRecord(
        EmailProcessingRecord existing,
        string messageKey,
        string messageId,
        string date,
        string subject,
        string sender,
        string status,
        string downloadStatus,
        string errorReason,
        int attachmentTotal,
        int linkCount,
        int saved,
        int duplicate,
        List<string> keywordHits,
        List<string> attachmentNames,
        List<Dictionary<string, string>> processedRows)
    {
        var savedFiles = processedRows
            .Where(r => !string.IsNullOrWhiteSpace(r.GetValueOrDefault("file")))
            .Select(r => new EmailSavedFileRecord
            {
                Path = r.GetValueOrDefault("file", ""),
                FileName = Path.GetFileName(r.GetValueOrDefault("file", "")),
                Md5 = r.GetValueOrDefault("md5", ""),
                Sha256 = r.GetValueOrDefault("sha256", ""),
                Extension = Path.GetExtension(r.GetValueOrDefault("file", "")),
                InvoiceNumber = r.GetValueOrDefault("invoice_number", ""),
                InvoiceAmount = r.GetValueOrDefault("invoice_amount", ""),
                InvoiceDate = r.GetValueOrDefault("invoice_date", ""),
                TextExtractStatus = r.GetValueOrDefault("text_extract_status", ""),
                Error = r.GetValueOrDefault("error", "")
            })
            .ToList();

        return new EmailProcessingRecord
        {
            MessageKey = messageKey,
            MessageId = messageId,
            Date = date,
            Subject = subject,
            Sender = sender,
            Status = status,
            DownloadStatus = downloadStatus,
            ErrorReason = errorReason,
            HasAttachment = attachmentTotal > 0,
            IsInvoiceMail = keywordHits.Count > 0,
            HasRisk = status is "异常" or "失败" or "需人工确认" or "未下载到文件",
            AttachmentTotal = attachmentTotal,
            LinkCandidateTotal = linkCount,
            SavedFileCount = saved,
            DuplicateCount = duplicate,
            DuplicateOf = processedRows.Select(r => r.GetValueOrDefault("duplicate_of")).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "",
            DuplicateSourceMailId = processedRows.Select(r => r.GetValueOrDefault("duplicate_source_mail_id")).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "",
            Attempts = existing.Attempts,
            KeywordHits = keywordHits,
            AttachmentNames = attachmentNames,
            Files = savedFiles.Select(f => f.Path).Where(File.Exists).Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
            SavedFiles = savedFiles,
            UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        };
    }

    private static string BuildMessageError(List<Dictionary<string, string>> rows, int saved, bool hasPdf, bool hasNoInvoiceNumber)
    {
        var errors = rows.Select(r => r.GetValueOrDefault("error", ""))
            .Where(e => !string.IsNullOrWhiteSpace(e))
            .Distinct()
            .ToList();
        if (saved > 0 && !hasPdf) errors.Add("无 PDF");
        if (hasNoInvoiceNumber) errors.Add("无发票号");
        if (saved == 0 && errors.Count == 0) errors.Add("未下载到文件");
        return string.Join("；", errors.Distinct());
    }

    private static async Task<List<Dictionary<string, string>>> SaveAttachmentsAsync(
        MimeMessage msg,
        string outputDir,
        string prefix,
        string subject,
        string sender,
        string date,
        string msgId,
        string messageKey,
        string startDate,
        Dictionary<string, string> existingHashes,
        HashSet<string> existingPdfKeys,
        EmailProcessingStore processingStore,
        CancellationToken ct)
    {
        var rows = new List<Dictionary<string, string>>();
        var parts = msg.Attachments.OfType<MimePart>()
            .Where(p => AllowedExtensions.Contains(Path.GetExtension(p.FileName ?? "")))
            .ToList();
        var hasPdf = parts.Any(p => Path.GetExtension(p.FileName ?? "").Equals(".pdf", StringComparison.OrdinalIgnoreCase));
        if (hasPdf)
        {
            parts = parts.Where(p => Path.GetExtension(p.FileName ?? "").Equals(".pdf", StringComparison.OrdinalIgnoreCase)).ToList();
        }
        var messagePdfKeys = parts
            .Where(p => Path.GetExtension(p.FileName ?? "").Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            .Select(p => InvoiceFormatPolicy.InvoiceKey(TextEncodingFixer.Fix(p.FileName ?? ""), subject))
            .Where(k => !string.IsNullOrWhiteSpace(k))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var part in parts)
        {
            var fileName = TextEncodingFixer.Fix(part.FileName ?? "attachment.bin");
            var ext = Path.GetExtension(fileName);
            if (part.Content == null)
            {
                continue;
            }
            var invoiceKey = InvoiceFormatPolicy.InvoiceKey(fileName, subject);
            if (InvoiceFormatPolicy.ShouldSkipBecausePdfExists(ext, invoiceKey, messagePdfKeys, existingPdfKeys))
            {
                rows.Add(SkippedByPdfRow("attachment", date, subject, msgId, "", fileName, invoiceKey));
                continue;
            }

            await using var ms = new MemoryStream();
            await part.Content.DecodeToAsync(ms, ct);
            var bytes = ms.ToArray();
            var save = await SaveBytesAsync(bytes, outputDir, SafeFileName($"{prefix}_{fileName}"), existingHashes, ct);
            if (save.Status == "已下载" && ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(invoiceKey))
            {
                existingPdfKeys.Add(invoiceKey);
            }

            var row = SavedFileRow("attachment", save, date, subject, sender, msgId, messageKey, "", startDate, processingStore);
            rows.Add(row);

            rows.AddRange(await ExtractZipIfNeededAsync(save.Path, outputDir, prefix, subject, sender, date, msgId, messageKey, startDate, existingHashes, existingPdfKeys, processingStore, ct));
        }
        return rows;
    }

    private static async Task<List<Dictionary<string, string>>> DownloadLinksAsync(
        List<LinkCandidate> links,
        HttpClient http,
        string outputDir,
        string prefix,
        string subject,
        string sender,
        string date,
        string msgId,
        string messageKey,
        string startDate,
        Dictionary<string, string> existingHashes,
        HashSet<string> existingPdfKeys,
        EmailProcessingStore processingStore,
        CancellationToken ct)
    {
        var rows = new List<Dictionary<string, string>>();
        for (var i = 0; i < links.Count; i++)
        {
            var url = links[i].Url;
            try
            {
                using var response = await http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct);
                if (!response.IsSuccessStatusCode)
                {
                    rows.Add(ErrorRow("link_error", date, subject, msgId, url, $"HTTP {(int)response.StatusCode}"));
                    continue;
                }

                var bytes = await response.Content.ReadAsByteArrayAsync(ct);
                if (bytes.Length == 0)
                {
                    continue;
                }

                var contentType = response.Content.Headers.ContentType?.MediaType ?? "";
                var resolvedUrl = response.RequestMessage?.RequestUri?.ToString() ?? url;
                var suffix = Path.GetExtension(new Uri(resolvedUrl).AbsolutePath);
                if (!AllowedExtensions.Contains(suffix) && !IsDownloadContent(contentType))
                {
                    var parsedRows = await TryDownloadFromHtmlPageAsync(
                        bytes,
                        http,
                        outputDir,
                        prefix,
                        subject,
                        sender,
                        date,
                        msgId,
                        messageKey,
                        startDate,
                        url,
                        resolvedUrl,
                        existingHashes,
                        existingPdfKeys,
                        processingStore,
                        ct);
                    rows.AddRange(parsedRows);
                    if (parsedRows.Count == 0)
                    {
                        rows.Add(ErrorRow("link", date, subject, msgId, url, $"链接取票待处理：页面里没有找到可直接下载的 PDF/OFD/XML/ZIP，content-type={contentType}", "链接取票待处理"));
                    }
                    continue;
                }

                var fileName = FileNameFromResponse(response.Content.Headers, resolvedUrl, $"{prefix}_link{i + 1}");
                var invoiceKey = InvoiceFormatPolicy.InvoiceKey(fileName, subject, resolvedUrl);
                var ext = Path.GetExtension(fileName);
                if (InvoiceFormatPolicy.ShouldSkipBecausePdfExists(ext, invoiceKey, [], existingPdfKeys))
                {
                    rows.Add(SkippedByPdfRow("link", date, subject, msgId, url, fileName, invoiceKey));
                    continue;
                }
                var save = await SaveBytesAsync(bytes, outputDir, SafeFileName($"{prefix}_{fileName}"), existingHashes, ct);
                if (save.Status == "已下载" && ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(invoiceKey))
                {
                    existingPdfKeys.Add(invoiceKey);
                }

                var row = SavedFileRow("link", save, date, subject, sender, msgId, messageKey, url, startDate, processingStore);
                row["content_type"] = contentType;
                rows.Add(row);

                rows.AddRange(await ExtractZipIfNeededAsync(save.Path, outputDir, prefix, subject, sender, date, msgId, messageKey, startDate, existingHashes, existingPdfKeys, processingStore, ct));
            }
            catch (Exception ex)
            {
                rows.Add(ErrorRow("link_error", date, subject, msgId, url, ex.Message));
            }
        }
        return rows;
    }

    private static Dictionary<string, string> ErrorRow(string kind, string date, string subject, string msgId, string url, string error, string withStatus = "失败") => new()
    {
        ["kind"] = kind,
        ["status"] = withStatus,
        ["date"] = date,
        ["subject"] = subject,
        ["msg_id"] = msgId,
        ["url"] = url,
        ["error"] = error
    };

    private static Dictionary<string, string> SkippedByPdfRow(string kind, string date, string subject, string msgId, string url, string fileName, string invoiceKey) => new()
    {
        ["kind"] = kind,
        ["status"] = "PDF已存在跳过",
        ["date"] = date,
        ["subject"] = subject,
        ["msg_id"] = msgId,
        ["url"] = url,
        ["file"] = fileName,
        ["invoice_key"] = invoiceKey,
        ["error"] = "同一张发票已有PDF，非PDF格式不下载/不保留"
    };

    private static async Task<List<Dictionary<string, string>>> TryDownloadFromHtmlPageAsync(
        byte[] bytes,
        HttpClient http,
        string outputDir,
        string prefix,
        string subject,
        string sender,
        string date,
        string msgId,
        string messageKey,
        string startDate,
        string originalUrl,
        string resolvedUrl,
        Dictionary<string, string> existingHashes,
        HashSet<string> existingPdfKeys,
        EmailProcessingStore processingStore,
        CancellationToken ct)
    {
        var rows = new List<Dictionary<string, string>>();
        var html = DecodeHtml(bytes);
        if (string.IsNullOrWhiteSpace(html))
        {
            return rows;
        }

        var candidates = ExtractDownloadLinksFromHtml(html, resolvedUrl).Take(12).ToList();
        for (var i = 0; i < candidates.Count; i++)
        {
            var candidate = candidates[i];
            try
            {
                using var response = await http.GetAsync(candidate.Url, HttpCompletionOption.ResponseHeadersRead, ct);
                if (!response.IsSuccessStatusCode)
                {
                    rows.Add(ErrorRow("link_page_candidate", date, subject, msgId, candidate.Url, $"页面候选链接下载失败：HTTP {(int)response.StatusCode}", "链接候选失败"));
                    continue;
                }

                var contentType = response.Content.Headers.ContentType?.MediaType ?? "";
                var finalUrl = response.RequestMessage?.RequestUri?.ToString() ?? candidate.Url;
                var suffix = Path.GetExtension(new Uri(finalUrl).AbsolutePath);
                if (!AllowedExtensions.Contains(suffix) && !IsDownloadContent(contentType))
                {
                    continue;
                }

                var fileBytes = await response.Content.ReadAsByteArrayAsync(ct);
                if (fileBytes.Length == 0)
                {
                    continue;
                }

                var fileName = FileNameFromResponse(response.Content.Headers, finalUrl, $"{prefix}_page_link{i + 1}{GuessExtension(contentType, suffix)}");
                var invoiceKey = InvoiceFormatPolicy.InvoiceKey(fileName, subject, finalUrl);
                var ext = Path.GetExtension(fileName);
                if (InvoiceFormatPolicy.ShouldSkipBecausePdfExists(ext, invoiceKey, [], existingPdfKeys))
                {
                    rows.Add(SkippedByPdfRow("link_page_candidate", date, subject, msgId, candidate.Url, fileName, invoiceKey));
                    continue;
                }

                var save = await SaveBytesAsync(fileBytes, outputDir, SafeFileName($"{prefix}_{fileName}"), existingHashes, ct);
                if (save.Status == "已下载" && ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(invoiceKey))
                {
                    existingPdfKeys.Add(invoiceKey);
                }

                var row = SavedFileRow("link_page_candidate", save, date, subject, sender, msgId, messageKey, candidate.Url, startDate, processingStore);
                row["status"] = row.GetValueOrDefault("status") == "已下载" ? "页面解析已下载" : row.GetValueOrDefault("status", "");
                row["source_url"] = originalUrl;
                row["content_type"] = contentType;
                row["link_text"] = candidate.Text;
                rows.Add(row);

                rows.AddRange(await ExtractZipIfNeededAsync(save.Path, outputDir, prefix, subject, sender, date, msgId, messageKey, startDate, existingHashes, existingPdfKeys, processingStore, ct));
            }
            catch (Exception ex)
            {
                rows.Add(ErrorRow("link_page_candidate", date, subject, msgId, candidate.Url, ex.Message, "链接候选失败"));
            }
        }

        return rows.Where(r => r.GetValueOrDefault("status") != "链接候选失败").ToList();
    }

    private static string DecodeHtml(byte[] bytes)
    {
        try
        {
            return Encoding.UTF8.GetString(bytes);
        }
        catch
        {
            return Encoding.Default.GetString(bytes);
        }
    }

    private static List<LinkCandidate> ExtractDownloadLinksFromHtml(string html, string baseUrl)
    {
        var result = new List<LinkCandidate>();
        var doc = new HtmlAgilityPack.HtmlDocument();
        doc.LoadHtml(html);
        foreach (var node in doc.DocumentNode.SelectNodes("//*[@href or @src or @data-url or @data-href]") ?? Enumerable.Empty<HtmlNode>())
        {
            var raw = FirstNonEmpty(
                node.GetAttributeValue("href", ""),
                node.GetAttributeValue("src", ""),
                node.GetAttributeValue("data-url", ""),
                node.GetAttributeValue("data-href", ""));
            var text = HtmlEntity.DeEntitize(node.InnerText ?? "");
            AddCandidate(result, baseUrl, raw, text);
        }

        foreach (Match match in Regex.Matches(html, @"https?://[^\s""'<>]+", RegexOptions.IgnoreCase))
        {
            AddCandidate(result, baseUrl, match.Value, "");
        }

        return result
            .Where(c => LooksLikeDownloadCandidate(c))
            .GroupBy(c => c.Url, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .ToList();
    }

    private static void AddCandidate(List<LinkCandidate> result, string baseUrl, string raw, string text)
    {
        if (string.IsNullOrWhiteSpace(raw) || raw.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }
        try
        {
            var uri = new Uri(new Uri(baseUrl), HtmlEntity.DeEntitize(raw.Trim()));
            result.Add(new LinkCandidate(uri.ToString(), text));
        }
        catch
        {
            // Ignore malformed page links.
        }
    }

    private static bool LooksLikeDownloadCandidate(LinkCandidate candidate)
    {
        var text = $"{candidate.Url} {candidate.Text}";
        return ContainsAny(text, ".pdf", ".ofd", ".xml", ".zip", "download", "invoice", "fapiao", "发票", "下载", "pdf", "ofd");
    }

    private static List<LinkCandidate> ExtractLinks(MimeMessage msg)
    {
        var links = new List<LinkCandidate>();
        if (!string.IsNullOrWhiteSpace(msg.HtmlBody))
        {
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(msg.HtmlBody);
            links.AddRange(doc.DocumentNode.SelectNodes("//a[@href]")?.Select(n => new LinkCandidate(NormalizeUrl(n.GetAttributeValue("href", "")), HtmlEntity.DeEntitize(n.InnerText ?? ""))) ?? []);
            links.AddRange(Regex.Matches(msg.HtmlBody, @"https?://[^\s""'<>]+").Select(m => new LinkCandidate(NormalizeUrl(m.Value), "")));
        }
        if (!string.IsNullOrWhiteSpace(msg.TextBody))
        {
            links.AddRange(Regex.Matches(msg.TextBody, @"https?://[^\s""'<>]+").Select(m => new LinkCandidate(NormalizeUrl(m.Value), "")));
        }
        return links.SelectMany(ExpandRedirectLinks)
            .Where(l => l.Url.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            .GroupBy(l => l.Url, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .ToList();
    }

    private static IEnumerable<LinkCandidate> ExpandRedirectLinks(LinkCandidate link)
    {
        var normalized = link with { Url = NormalizeUrl(link.Url) };
        yield return normalized;

        var real = ExtractQueryValue(normalized.Url, "jump_to");
        if (string.IsNullOrWhiteSpace(real))
        {
            real = ExtractQueryValue(normalized.Url, "url");
        }
        if (string.IsNullOrWhiteSpace(real))
        {
            real = ExtractQueryValue(normalized.Url, "target");
        }
        if (!string.IsNullOrWhiteSpace(real) && real.StartsWith("http", StringComparison.OrdinalIgnoreCase))
        {
            yield return new LinkCandidate(NormalizeUrl(real), string.IsNullOrWhiteSpace(link.Text) ? "跳转后的真实下载地址" : link.Text);
        }
    }

    private static string NormalizeUrl(string url)
    {
        var decoded = HtmlEntity.DeEntitize(url).Trim().TrimEnd(')', '.', ',', ';', '，', '。');
        try
        {
            return Uri.UnescapeDataString(decoded);
        }
        catch
        {
            return decoded;
        }
    }

    private static string ExtractQueryValue(string url, string key)
    {
        try
        {
            var uri = new Uri(url);
            var query = uri.Query.TrimStart('?').Split('&', StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in query)
            {
                var pieces = part.Split('=', 2);
                if (pieces.Length == 2 && pieces[0].Equals(key, StringComparison.OrdinalIgnoreCase))
                {
                    return Uri.UnescapeDataString(pieces[1]);
                }
            }
        }
        catch
        {
            // Ignore malformed redirect links.
        }
        return "";
    }

    private static bool LooksLikeInvoiceLink(LinkCandidate link)
    {
        return InvoiceUrlHints.Any(h => link.Url.Contains(h, StringComparison.OrdinalIgnoreCase) || link.Text.Contains(h, StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsDownloadContent(string contentType)
    {
        return contentType.Contains("pdf", StringComparison.OrdinalIgnoreCase)
            || contentType.Contains("octet-stream", StringComparison.OrdinalIgnoreCase)
            || contentType.Contains("zip", StringComparison.OrdinalIgnoreCase)
            || contentType.Contains("image", StringComparison.OrdinalIgnoreCase)
            || contentType.Contains("xml", StringComparison.OrdinalIgnoreCase);
    }

    private static string FileNameFromResponse(HttpContentHeaders headers, string url, string fallback)
    {
        var fileName = headers.ContentDisposition?.FileNameStar
            ?? headers.ContentDisposition?.FileName?.Trim('"');
        if (!string.IsNullOrWhiteSpace(fileName))
        {
            return SafeFileName(Uri.UnescapeDataString(fileName));
        }

        var fromUrl = Path.GetFileName(new Uri(url).AbsolutePath);
        return string.IsNullOrWhiteSpace(fromUrl)
            ? fallback + ".pdf"
            : SafeFileName(Uri.UnescapeDataString(fromUrl));
    }

    private static string GuessExtension(string contentType, string suffix)
    {
        if (AllowedExtensions.Contains(suffix)) return suffix;
        if (contentType.Contains("pdf", StringComparison.OrdinalIgnoreCase)) return ".pdf";
        if (contentType.Contains("zip", StringComparison.OrdinalIgnoreCase)) return ".zip";
        if (contentType.Contains("xml", StringComparison.OrdinalIgnoreCase)) return ".xml";
        if (contentType.Contains("png", StringComparison.OrdinalIgnoreCase)) return ".png";
        if (contentType.Contains("jpeg", StringComparison.OrdinalIgnoreCase) || contentType.Contains("jpg", StringComparison.OrdinalIgnoreCase)) return ".jpg";
        return ".pdf";
    }

    private static InvoiceMetadata ExtractInvoiceMetadata(string path)
    {
        if (!Path.GetExtension(path).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            return new InvoiceMetadata();
        }

        var text = ExtractPdfText(path);
        return new InvoiceMetadata
        {
            InvoiceNumber = ExtractInvoiceNo(text),
            Amount = ExtractAmount(text),
            Date = ExtractDate(text),
            TextExtractStatus = string.IsNullOrWhiteSpace(text) ? "未提取到文本" : "成功"
        };
    }

    private static string ExtractPdfText(string path)
    {
        try
        {
            using var pdf = PdfDocument.Open(path);
            return string.Join(Environment.NewLine, pdf.GetPages().Select(p => p.Text));
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
        return "";
    }

    private static string ExtractAmount(string text)
    {
        var patterns = new[]
        {
            @"价税合计.{0,60}?[小写）\)]\s*[¥￥]?\s*([0-9]+(?:\.[0-9]{1,2})?)",
            @"小写\s*[）\)]?\s*[¥￥]?\s*([0-9]+(?:\.[0-9]{1,2})?)",
            @"合计\s*[¥￥]\s*([0-9]+(?:\.[0-9]{1,2})?)",
            @"[¥￥]\s*([0-9]+(?:\.[0-9]{1,2})?)"
        };
        foreach (var pattern in patterns)
        {
            var match = Regex.Match(text, pattern, RegexOptions.Singleline);
            if (match.Success && decimal.TryParse(match.Groups[1].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var amount))
            {
                return amount.ToString("0.##", CultureInfo.InvariantCulture);
            }
        }
        return "";
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

    private static string NormalizeInvoiceNo(string value)
    {
        if (value.Length > 20 && !value.StartsWith("20", StringComparison.Ordinal))
        {
            value = value[..20];
        }
        return value;
    }

    private static List<string> FindKeywordHits(string text)
    {
        return InvoiceMailKeywords
            .Where(k => text.Contains(k, StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string StripHtml(string? html)
    {
        return string.IsNullOrWhiteSpace(html)
            ? ""
            : Regex.Replace(html, "<[^>]+>", " ");
    }

    private static string CleanSender(string sender)
    {
        var cleaned = Regex.Replace(sender, "<[^>]+>", "").Trim(' ', '"');
        return string.IsNullOrWhiteSpace(cleaned) ? "未知发件人" : TrimForName(cleaned, 40);
    }

    private static string FirstNonEmpty(params string[] values)
    {
        return values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "";
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

    private static string TrimForName(string text, int max)
    {
        return text.Length <= max ? text : text[..max];
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
        throw new IOException("无法生成不重复文件名：" + path);
    }

    private static Dictionary<string, string> BuildExistingHashIndex(string outputDir)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var file in Directory.EnumerateFiles(outputDir, "*.*", SearchOption.AllDirectories))
        {
            try
            {
                var hash = Sha256(file);
                if (!result.ContainsKey(hash))
                {
                    result[hash] = file;
                }
            }
            catch
            {
                // Ignore files that are being written by another process.
            }
        }
        return result;
    }

    private static HashSet<string> BuildExistingPdfKeyIndex(string outputDir)
    {
        return Directory.EnumerateFiles(outputDir, "*.pdf", SearchOption.AllDirectories)
            .Select(file => InvoiceFormatPolicy.InvoiceKey(Path.GetFileName(file), file))
            .Where(key => !string.IsNullOrWhiteSpace(key))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    private static Dictionary<string, EmailInvoiceRecord> BuildExistingInvoiceIndex(string outputDir)
    {
        var result = new Dictionary<string, EmailInvoiceRecord>(StringComparer.OrdinalIgnoreCase);
        if (!Directory.Exists(outputDir))
        {
            return result;
        }

        foreach (var file in Directory.EnumerateFiles(outputDir, "*.*", SearchOption.AllDirectories))
        {
            var extension = Path.GetExtension(file);
            if (extension.Equals(".csv", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            try
            {
                var bytes = File.ReadAllBytes(file);
                var md5 = Md5(bytes);
                var sha256 = Sha256(bytes);
                var metadata = ExtractInvoiceMetadata(file);
                var invoiceKey = BuildInvoiceKey(md5, metadata);
                if (string.IsNullOrWhiteSpace(invoiceKey))
                {
                    invoiceKey = $"sha256:{sha256}";
                }

                result.TryAdd(invoiceKey, new EmailInvoiceRecord
                {
                    InvoiceKey = invoiceKey,
                    MessageKey = "existing-file",
                    MessageId = "",
                    Subject = Path.GetFileName(file),
                    Sender = "",
                    Date = File.GetLastWriteTime(file).ToString("yyyy-MM-dd"),
                    Path = file,
                    Md5 = md5,
                    Sha256 = sha256,
                    InvoiceNumber = metadata.InvoiceNumber,
                    InvoiceAmount = metadata.Amount,
                    InvoiceDate = metadata.Date,
                    UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                });
            }
            catch
            {
                // Ignore unreadable files; they will not participate in duplicate detection.
            }
        }

        return result;
    }

    private static async Task<SaveResult> SaveBytesAsync(byte[] bytes, string outputDir, string fileName, Dictionary<string, string> existingHashes, CancellationToken ct)
    {
        var sha256 = Sha256(bytes);
        var md5 = Md5(bytes);
        if (existingHashes.TryGetValue(sha256, out var duplicate))
        {
            return new SaveResult("重复跳过", "", md5, sha256, duplicate);
        }

        var path = UniquePath(Path.Combine(outputDir, SafeFileName(fileName)));
        await File.WriteAllBytesAsync(path, bytes, ct);
        existingHashes[sha256] = path;
        return new SaveResult("已下载", path, md5, sha256, "");
    }

    private static async Task<List<Dictionary<string, string>>> ExtractZipIfNeededAsync(
        string path,
        string outputDir,
        string prefix,
        string subject,
        string sender,
        string date,
        string msgId,
        string messageKey,
        string startDate,
        Dictionary<string, string> existingHashes,
        HashSet<string> existingPdfKeys,
        EmailProcessingStore processingStore,
        CancellationToken ct)
    {
        var rows = new List<Dictionary<string, string>>();
        if (string.IsNullOrWhiteSpace(path) || !Path.GetExtension(path).Equals(".zip", StringComparison.OrdinalIgnoreCase))
        {
            return rows;
        }

        var extractDir = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(path) + "_解压");
        Directory.CreateDirectory(extractDir);
        try
        {
            using var archive = ZipFile.OpenRead(path);
            var entries = archive.Entries.Where(e => AllowedExtensions.Contains(Path.GetExtension(e.Name))).ToList();
            var pdfKeysInZip = entries
                .Where(e => Path.GetExtension(e.Name).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                .Select(e => InvoiceFormatPolicy.InvoiceKey(e.Name, subject, path))
                .Where(k => !string.IsNullOrWhiteSpace(k))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in entries)
            {
                var ext = Path.GetExtension(entry.Name);
                var invoiceKey = InvoiceFormatPolicy.InvoiceKey(entry.Name, subject, path);
                if (InvoiceFormatPolicy.ShouldSkipBecausePdfExists(ext, invoiceKey, pdfKeysInZip, existingPdfKeys))
                {
                    rows.Add(SkippedByPdfRow("zip_entry", date, subject, msgId, path, entry.Name, invoiceKey));
                    continue;
                }
                await using var entryStream = entry.Open();
                await using var ms = new MemoryStream();
                await entryStream.CopyToAsync(ms, ct);
                var save = await SaveBytesAsync(ms.ToArray(), extractDir, SafeFileName($"{prefix}_{entry.Name}"), existingHashes, ct);
                if (save.Status == "已下载" && ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(invoiceKey))
                {
                    existingPdfKeys.Add(invoiceKey);
                }
                rows.Add(SavedFileRow("zip_entry", save, date, subject, sender, msgId, messageKey, path, startDate, processingStore));
            }
        }
        catch (Exception ex)
        {
            rows.Add(ErrorRow("zip_error", date, subject, msgId, path, ex.Message));
        }
        return rows;
    }

    private static string Sha256(string path)
    {
        using var sha = SHA256.Create();
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(sha.ComputeHash(stream)).ToLowerInvariant();
    }

    private static string Sha256(byte[] bytes)
    {
        return Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    }

    private static string Md5(byte[] bytes)
    {
        return Convert.ToHexString(MD5.HashData(bytes)).ToLowerInvariant();
    }

    private static void WriteManifest(string path, List<Dictionary<string, string>> rows)
    {
        var headers = rows.SelectMany(r => r.Keys).Distinct().OrderBy(x => x).ToList();
        if (headers.Count == 0)
        {
            headers.Add("message");
            rows.Add(new Dictionary<string, string> { ["message"] = "无记录" });
        }

        using var writer = new StreamWriter(path, false, new UTF8Encoding(true));
        writer.WriteLine(string.Join(",", headers.Select(EscapeCsv)));
        foreach (var row in rows)
        {
            writer.WriteLine(string.Join(",", headers.Select(h => EscapeCsv(row.GetValueOrDefault(h, "")))));
        }
    }

    private static string EscapeCsv(string value)
    {
        return value.Contains(',') || value.Contains('"') || value.Contains('\n') || value.Contains('\r')
            ? "\"" + value.Replace("\"", "\"\"") + "\""
            : value;
    }

    private readonly record struct LinkCandidate(string Url, string Text);
    private readonly record struct SaveResult(string Status, string Path, string Md5, string Sha256, string DuplicateOf);
    private sealed record InvoiceMetadata
    {
        public string InvoiceNumber { get; init; } = "";
        public string Amount { get; init; } = "";
        public string Date { get; init; } = "";
        public string TextExtractStatus { get; init; } = "";
    }
}
