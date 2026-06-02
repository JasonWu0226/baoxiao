using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace ReimbursementMatcher;

public sealed class EmailAuditService
{
    public const string HasInvoice = "明确有发票";
    public const string NoInvoice = "明确无发票";
    public const string NeedsReview = "需要人工确认";
    public const string PreviousReimbursed = "疑似上期已报销";
    public const string DateOutOfRange = "发票日期不符";

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private readonly WorkspaceService _workspace;

    public EmailAuditService(WorkspaceService workspace)
    {
        _workspace = workspace;
    }

    public EmailDecisionStore LoadDecisions(AppConfig config)
    {
        var path = DecisionPath(config);
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

    public void SaveDecisions(AppConfig config, EmailDecisionStore store)
    {
        store.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var path = DecisionPath(config);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, JsonSerializer.Serialize(store, JsonOptions) + Environment.NewLine, Encoding.UTF8);
    }

    public List<EmailAuditItem> LoadLatestAudit(AppConfig config, EmailDecisionStore decisions)
    {
        var latest = FindLatestManifest(config);
        if (latest == null)
        {
            return [];
        }

        var rows = ReadCsv(latest);
        var processingStore = LoadProcessingStore(config);
        var messageKeyById = rows
            .Where(r => Get(r, "kind") == "message")
            .Select(r => new { Id = Get(r, "msg_id"), Key = Get(r, "message_key") })
            .Where(x => !string.IsNullOrWhiteSpace(x.Id) && !string.IsNullOrWhiteSpace(x.Key))
            .GroupBy(x => x.Id, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First().Key, StringComparer.OrdinalIgnoreCase);

        return rows
            .GroupBy(r => GroupKey(r, messageKeyById), StringComparer.OrdinalIgnoreCase)
            .Select(g => BuildItem(g.Key, g.ToList(), decisions, processingStore.Messages.GetValueOrDefault(g.Key)))
            .OrderByDescending(i => i.NeedsHumanReview)
            .ThenBy(i => i.Date)
            .ThenBy(i => i.Subject)
            .ToList();
    }

    public void Confirm(AppConfig config, EmailDecisionStore store, EmailAuditItem item, string decision, string note)
    {
        store.Messages.TryGetValue(item.MessageKey, out var old);
        var before = old?.Decision ?? "";
        var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var record = new EmailDecisionRecord
        {
            MessageKey = item.MessageKey,
            Subject = item.Subject,
            Decision = decision,
            Note = note,
            UpdatedAt = now
        };
        store.Messages[item.MessageKey] = record;
        store.Events.Add(new ConfirmationEvent
        {
            Time = now,
            ObjectType = "邮件",
            ObjectId = item.MessageKey,
            Action = "确认邮件是否包含发票",
            Before = before,
            After = decision,
            Note = note
        });
        SaveDecisions(config, store);
        ApplyDecision(item, record);
    }

    public string GenerateChecklistExcel(AppConfig config, List<EmailAuditItem> items)
    {
        var outputDir = _workspace.Resolve(config.OutputDir);
        Directory.CreateDirectory(outputDir);
        var output = Path.Combine(outputDir, $"邮件发票判断清单_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

        using var wb = new XLWorkbook();
        WriteSheet(wb, "需要人工确认", items.Where(i => i.NeedsHumanReview).ToList());
        WriteSheet(wb, "全部邮件", items);
        WriteSummary(wb, items);
        wb.SaveAs(output);
        return output;
    }

    public string GenerateInvoicePresenceExcel(AppConfig config, List<EmailAuditItem> items)
    {
        var outputDir = _workspace.Resolve(config.OutputDir);
        Directory.CreateDirectory(outputDir);
        var output = Path.Combine(outputDir, $"发票存在性核验_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        var invoiceDir = _workspace.Resolve(config.InvoiceDir);

        var rows = items.Select(BuildPresenceRow).ToList();
        var needCheck = rows
            .Where(r => r.NextAction != "无需处理" && r.FinalDecision != PreviousReimbursed && r.FinalDecision != DateOutOfRange)
            .OrderBy(r => r.Date)
            .ThenBy(r => r.Subject)
            .ToList();
        var confirmed = rows
            .Where(r => r.FinalDecision == HasInvoice && r.ExistingFileCount > 0)
            .OrderBy(r => r.Date)
            .ThenBy(r => r.Subject)
            .ToList();
        var noInvoice = rows
            .Where(r => r.FinalDecision == NoInvoice)
            .OrderBy(r => r.Date)
            .ThenBy(r => r.Subject)
            .ToList();
        var previous = rows
            .Where(r => r.FinalDecision == PreviousReimbursed)
            .OrderBy(r => r.Date)
            .ThenBy(r => r.Subject)
            .ToList();
        var imageFiles = Directory.Exists(invoiceDir)
            ? Directory.EnumerateFiles(invoiceDir, "*.*", SearchOption.AllDirectories)
                .Where(IsImageOrIcon)
                .Where(f => !f.Contains($"{Path.DirectorySeparatorChar}非PDF图片清理_", StringComparison.OrdinalIgnoreCase))
                .Select(f => new NonPdfReviewRow
                {
                    Type = Path.GetExtension(f).Equals(".ico", StringComparison.OrdinalIgnoreCase) ? "ICO图标" : "图片/可能二维码",
                    NextAction = Path.GetExtension(f).Equals(".ico", StringComparison.OrdinalIgnoreCase)
                        ? "可归档，肯定不是发票"
                        : "打开查看；如为二维码则执行“处理非PDF/二维码/AI判断”",
                    File = f,
                    SizeKb = new FileInfo(f).Length / 1024
                })
                .OrderBy(r => r.Type)
                .ThenBy(r => r.File)
                .ToList()
            : [];

        using var wb = new XLWorkbook();
        WritePresenceSheet(wb, "需要逐封核验", needCheck);
        WritePresenceSheet(wb, "疑似上期已报销", previous);
        WritePresenceSheet(wb, "有发票且已存在", confirmed);
        WritePresenceSheet(wb, "明确无发票", noInvoice);
        WriteNonPdfSheet(wb, "二维码与非PDF图片", imageFiles);
        WritePresenceSummary(wb, rows, imageFiles);
        wb.SaveAs(output);
        return output;
    }

    public string SaveDownloadRules(AppConfig config, List<EmailAuditItem> items)
    {
        var path = Path.Combine(_workspace.Resolve(config.RuleDir), "邮件下载规则.json");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var rules = items
            .Where(i => !string.IsNullOrWhiteSpace(i.ManualDecision))
            .Select(i => new
            {
                i.MessageKey,
                i.Subject,
                i.ManualDecision,
                i.Note,
                Keywords = ExtractKeywords(i.Subject)
            })
            .ToList();
        File.WriteAllText(path, JsonSerializer.Serialize(new
        {
            UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
            Rules = rules
        }, JsonOptions) + Environment.NewLine, Encoding.UTF8);
        return path;
    }

    public string DecisionPath(AppConfig config)
    {
        return Path.Combine(_workspace.Resolve(config.RuleDir), "邮件判断记录.json");
    }

    private EmailProcessingStore LoadProcessingStore(AppConfig config)
    {
        var path = Path.Combine(_workspace.Resolve(config.RuleDir), "邮箱处理记录.json");
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

    private string? FindLatestManifest(AppConfig config)
    {
        var candidates = new[]
            {
                config.Email.OutputDir,
                config.InvoiceDir,
                "历史发票"
            }
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Select(_workspace.Resolve)
            .Where(Directory.Exists)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .SelectMany(dir => Directory.EnumerateFiles(dir, "邮箱发票下载清单_*.csv", SearchOption.TopDirectoryOnly))
            .OrderByDescending(File.GetLastWriteTime)
            .ToList();

        return candidates.FirstOrDefault();
    }

    private static string GroupKey(Dictionary<string, string> row, Dictionary<string, string> messageKeyById)
    {
        var messageKey = Get(row, "message_key");
        if (!string.IsNullOrWhiteSpace(messageKey)) return messageKey;

        var msgId = Get(row, "msg_id");
        if (!string.IsNullOrWhiteSpace(msgId) && messageKeyById.TryGetValue(msgId, out var mapped)) return mapped;
        if (!string.IsNullOrWhiteSpace(msgId)) return msgId;

        return FirstNonEmpty(Get(row, "subject"), Guid.NewGuid().ToString("N"));
    }

    private static EmailAuditItem BuildItem(string key, List<Dictionary<string, string>> rows, EmailDecisionStore decisions, EmailProcessingRecord? historical)
    {
        var subject = TextEncodingFixer.Fix(rows.Select(r => Get(r, "subject")).FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? historical?.Subject ?? "");
        var status = rows.Where(r => Get(r, "kind") == "message").Select(r => Get(r, "status")).FirstOrDefault(s => !string.IsNullOrWhiteSpace(s))
            ?? rows.Select(r => Get(r, "status")).FirstOrDefault(s => !string.IsNullOrWhiteSpace(s))
            ?? historical?.Status
            ?? "";
        var date = rows.Select(r => Get(r, "date")).FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? historical?.Date ?? "";
        var msgId = rows.Select(r => Get(r, "msg_id")).FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? historical?.MessageId ?? key;
        var files = rows.Select(r => Get(r, "file"))
            .Concat(rows.Select(r => Get(r, "duplicate_of")))
            .Concat(historical?.Files ?? [])
            .Concat(historical?.SavedFiles.Select(f => f.Path) ?? [])
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var urls = rows.Select(r => Get(r, "url")).Where(v => !string.IsNullOrWhiteSpace(v)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var attachmentRows = rows.Where(r => Get(r, "kind") is "attachment" or "zip_entry").ToList();
        var attachmentCount = Math.Max(ParseInt(rows.Select(r => Get(r, "attachment_total")).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v))) ?? attachmentRows.Count, historical?.AttachmentTotal ?? 0);
        var likelyAttachment = attachmentRows.Count(r => LooksLikeInvoice(Get(r, "file"), subject) || LooksLikeInvoice(Get(r, "duplicate_of"), subject))
            + (historical?.SavedFiles.Count(f => LooksLikeInvoice(f.Path, subject) || LooksLikeInvoice(f.FileName, subject)) ?? 0);
        var downloaded = Math.Max(rows.Count(r => !string.IsNullOrWhiteSpace(Get(r, "file")) && Get(r, "status") is "已下载" or "正常已下载" or "页面解析已下载"), historical?.SavedFileCount ?? 0);
        var skipped = rows.Count(r => Get(r, "status") is "重复跳过" or "重复已存在" or "重复发票" or "PDF已存在跳过" or "已处理跳过" or "人工确认无发票跳过" or "附件已取得发票，链接跳过" or "疑似上期已报销跳过" or "发票日期早于开始时间跳过");
        var hasInvoiceKeyword = LooksLikeInvoice(subject, string.Join(" ", files));
        var hasPlatform = ContainsAny(subject, "美团", "三快", "淘宝", "天猫", "京东", "携程", "ctrip", "jd.com", "taobao", "tmall");
        var hasLikelyLink = urls.Any(u => LooksLikeInvoice(u, subject));
        var hasAnyLink = urls.Count > 0 || (historical?.LinkCandidateTotal ?? 0) > 0;
        var hasRiskStatus = rows.Select(r => Get(r, "status")).Any(s => s is "未下载到文件" or "失败" or "需人工确认" or "待核验" or "异常" or "链接取票待处理");
        var errors = rows.Select(r => Get(r, "error")).Where(IsRealError).ToList();
        var statusSet = rows.Select(r => Get(r, "status")).Where(s => !string.IsNullOrWhiteSpace(s)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var historicalMatch = FirstNonEmpty(
            rows.Select(r => Get(r, "historical_match_summary")).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "",
            historical?.HistoricalMatchSummary ?? "");
        var historicalFile = FirstNonEmpty(
            rows.Select(r => Get(r, "historical_match_file")).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "",
            historical?.HistoricalMatchFile ?? "");
        var ruleDecision = !string.IsNullOrWhiteSpace(historicalMatch)
            ? PreviousReimbursed
            : DecideByRule(hasInvoiceKeyword, hasPlatform, attachmentCount, likelyAttachment, hasLikelyLink, hasAnyLink, downloaded, skipped, hasRiskStatus, errors.Count, statusSet);

        var item = new EmailAuditItem
        {
            MessageKey = key,
            Date = date,
            Subject = TextEncodingFixer.Fix(subject),
            MessageId = msgId,
            RuleDecision = ruleDecision,
            FinalDecision = ruleDecision,
            Status = status,
            HasInvoiceKeyword = hasInvoiceKeyword,
            HasPlatformKeyword = hasPlatform,
            HasAttachment = attachmentCount > 0 || attachmentRows.Count > 0,
            AttachmentCount = attachmentCount,
            HasLikelyInvoiceAttachment = likelyAttachment > 0,
            HasLikelyInvoiceLink = hasLikelyLink,
            DownloadedFileCount = downloaded,
            SkippedOrDuplicateCount = skipped,
            NeedsHumanReview = ruleDecision == NeedsReview,
            HistoricalMatch = historicalMatch,
            HistoricalMatchFile = historicalFile,
            Reason = MergeReason(BuildReason(hasInvoiceKeyword, hasPlatform, attachmentCount, likelyAttachment, hasLikelyLink, downloaded, skipped, hasRiskStatus, errors.Count), historicalMatch),
            Files = string.Join(Environment.NewLine, files),
            Urls = string.Join(Environment.NewLine, urls)
        };

        if (decisions.Messages.TryGetValue(key, out var manual))
        {
            ApplyDecision(item, manual);
        }
        return item;
    }

    private static void ApplyDecision(EmailAuditItem item, EmailDecisionRecord manual)
    {
        item.ManualDecision = manual.Decision;
        item.FinalDecision = string.IsNullOrWhiteSpace(manual.Decision) ? item.RuleDecision : manual.Decision;
        item.Note = manual.Note;
        item.NeedsHumanReview = item.FinalDecision == NeedsReview;
    }

    private static string DecideByRule(bool invoiceKeyword, bool platformKeyword, int attachmentCount, int likelyAttachment, bool likelyLink, bool anyLink, int downloaded, int skipped, bool riskStatus, int errors, HashSet<string> statuses)
    {
        if (statuses.Contains("疑似上期已报销跳过")) return PreviousReimbursed;
        if (statuses.Contains("发票日期早于开始时间跳过")) return DateOutOfRange;
        if (!invoiceKeyword && attachmentCount == 0 && !anyLink && downloaded == 0 && skipped == 0) return NoInvoice;
        if (invoiceKeyword && (downloaded > 0 || likelyAttachment > 0 || skipped > 0 || attachmentCount > 0 && HasBenignHandledStatus(statuses))) return HasInvoice;
        if (invoiceKeyword && platformKeyword && attachmentCount > 0 && skipped > 0) return HasInvoice;
        if (riskStatus || errors > 0) return NeedsReview;
        if (downloaded > 0 || likelyAttachment > 0 || likelyLink || skipped > 0 && invoiceKeyword) return HasInvoice;
        if (!invoiceKeyword && !platformKeyword && attachmentCount == 0 && downloaded == 0 && skipped == 0) return NoInvoice;
        if (!invoiceKeyword && !platformKeyword && attachmentCount > 0 && likelyAttachment == 0) return NoInvoice;
        return NeedsReview;
    }

    private static bool HasBenignHandledStatus(HashSet<string> statuses)
    {
        return statuses.Any(s => s is "已处理跳过" or "重复跳过" or "重复已存在" or "重复发票" or "PDF已存在跳过" or "附件已取得发票，链接跳过" or "疑似上期已报销跳过");
    }

    private static bool IsRealError(string error)
    {
        if (string.IsNullOrWhiteSpace(error)) return false;
        return !ContainsAny(error,
            "该邮件此前已成功处理",
            "本次不重复下载",
            "同一张发票已有PDF",
            "非PDF格式不下载",
            "附件已取得发票",
            "重复发票",
            "上期索引",
            "发票号命中上期",
            "金额和销售方命中上期",
            "早于开始时间",
            "人工确认为无发票");
    }

    private static string BuildReason(bool invoiceKeyword, bool platformKeyword, int attachmentCount, int likelyAttachment, bool likelyLink, int downloaded, int skipped, bool riskStatus, int errors)
    {
        var reasons = new List<string>();
        if (riskStatus) reasons.Add("下载状态异常");
        if (errors > 0) reasons.Add("存在错误说明");
        if (invoiceKeyword) reasons.Add("命中发票关键字");
        if (platformKeyword) reasons.Add("命中平台关键字");
        if (attachmentCount > 0) reasons.Add($"附件{attachmentCount}个");
        if (likelyAttachment > 0) reasons.Add($"疑似发票附件{likelyAttachment}个");
        if (likelyLink) reasons.Add("存在疑似发票链接");
        if (downloaded > 0) reasons.Add($"已下载{downloaded}个文件");
        if (skipped > 0) reasons.Add($"重复/跳过{skipped}项");
        return string.Join("；", reasons);
    }

    private static string MergeReason(string reason, string historicalMatch)
    {
        if (string.IsNullOrWhiteSpace(historicalMatch)) return reason;
        return string.IsNullOrWhiteSpace(reason) ? historicalMatch : $"{reason}；{historicalMatch}";
    }

    private static void WriteSheet(XLWorkbook wb, string name, List<EmailAuditItem> items)
    {
        var ws = wb.Worksheets.Add(name);
        var headers = new[] { "最终判断", "规则判断", "人工判断", "邮件日期", "邮件主题", "邮件ID", "状态", "有发票关键字", "有平台关键字", "有附件", "附件数", "疑似发票附件", "疑似发票链接", "下载数", "重复/跳过数", "上期匹配", "上期文件", "判断原因", "人工备注", "文件清单", "链接清单" };
        for (var c = 0; c < headers.Length; c++)
        {
            ws.Cell(1, c + 1).Value = headers[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
        }

        var r = 2;
        foreach (var item in items)
        {
            ws.Cell(r, 1).Value = item.FinalDecision;
            ws.Cell(r, 2).Value = item.RuleDecision;
            ws.Cell(r, 3).Value = item.ManualDecision;
            ws.Cell(r, 4).Value = item.Date;
            ws.Cell(r, 5).Value = item.Subject;
            ws.Cell(r, 6).Value = item.MessageId;
            ws.Cell(r, 7).Value = item.Status;
            ws.Cell(r, 8).Value = item.HasInvoiceKeyword ? "是" : "否";
            ws.Cell(r, 9).Value = item.HasPlatformKeyword ? "是" : "否";
            ws.Cell(r, 10).Value = item.HasAttachment ? "是" : "否";
            ws.Cell(r, 11).Value = item.AttachmentCount;
            ws.Cell(r, 12).Value = item.HasLikelyInvoiceAttachment ? "是" : "否";
            ws.Cell(r, 13).Value = item.HasLikelyInvoiceLink ? "是" : "否";
            ws.Cell(r, 14).Value = item.DownloadedFileCount;
            ws.Cell(r, 15).Value = item.SkippedOrDuplicateCount;
            ws.Cell(r, 16).Value = item.HistoricalMatch;
            ws.Cell(r, 17).Value = item.HistoricalMatchFile;
            ws.Cell(r, 18).Value = item.Reason;
            ws.Cell(r, 19).Value = item.Note;
            ws.Cell(r, 20).Value = item.Files;
            ws.Cell(r, 21).Value = item.Urls;
            ws.Range(r, 1, r, headers.Length).Style.Fill.BackgroundColor = item.FinalDecision switch
            {
                HasInvoice => XLColor.FromHtml("#D9EAD3"),
                NoInvoice => XLColor.FromHtml("#E5E7EB"),
                PreviousReimbursed => XLColor.FromHtml("#EADCF8"),
                DateOutOfRange => XLColor.FromHtml("#E5E7EB"),
                _ => XLColor.FromHtml("#FFF2CC")
            };
            r++;
        }

        ws.Columns().AdjustToContents(8, 60);
        ws.SheetView.FreezeRows(1);
    }

    private static void WriteSummary(XLWorkbook wb, List<EmailAuditItem> items)
    {
        var ws = wb.Worksheets.Add("统计汇总");
        var metrics = new (string Name, int Count)[]
        {
            ("邮件总数", items.Count),
            (HasInvoice, items.Count(i => i.FinalDecision == HasInvoice)),
            (NoInvoice, items.Count(i => i.FinalDecision == NoInvoice)),
            (PreviousReimbursed, items.Count(i => i.FinalDecision == PreviousReimbursed)),
            (DateOutOfRange, items.Count(i => i.FinalDecision == DateOutOfRange)),
            (NeedsReview, items.Count(i => i.FinalDecision == NeedsReview)),
            ("已有人工判断", items.Count(i => !string.IsNullOrWhiteSpace(i.ManualDecision))),
            ("命中发票关键字", items.Count(i => i.HasInvoiceKeyword)),
            ("命中平台关键字", items.Count(i => i.HasPlatformKeyword)),
            ("有附件", items.Count(i => i.HasAttachment)),
            ("有疑似发票附件", items.Count(i => i.HasLikelyInvoiceAttachment)),
            ("有疑似发票链接", items.Count(i => i.HasLikelyInvoiceLink))
        };
        ws.Cell(1, 1).Value = "指标";
        ws.Cell(1, 2).Value = "数量";
        ws.Range(1, 1, 1, 2).Style.Font.Bold = true;
        for (var i = 0; i < metrics.Length; i++)
        {
            ws.Cell(i + 2, 1).Value = metrics[i].Name;
            ws.Cell(i + 2, 2).Value = metrics[i].Count;
        }
        ws.Columns().AdjustToContents();
    }

    private static PresenceReviewRow BuildPresenceRow(EmailAuditItem item)
    {
        var existingFiles = SplitFileList(item.Files)
            .Where(File.Exists)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var hasLink = !string.IsNullOrWhiteSpace(item.Urls);
        var isLinkPending = item.Status.Contains("链接取票待处理", StringComparison.OrdinalIgnoreCase)
            || item.Reason.Contains("链接取票待处理", StringComparison.OrdinalIgnoreCase);
        var isPrevious = item.FinalDecision == PreviousReimbursed;
        var isDateOutOfRange = item.FinalDecision == DateOutOfRange;
        var hasInvoice = item.FinalDecision == HasInvoice;
        var needsReview = item.FinalDecision == NeedsReview || item.NeedsHumanReview;

        var status = isPrevious
            ? "疑似上期已报销"
            : isDateOutOfRange
                ? "发票日期不符"
            : hasInvoice && existingFiles.Count > 0
            ? "已确认存在"
            : hasInvoice && item.SkippedOrDuplicateCount > 0 && existingFiles.Count == 0
                ? "有发票但仅显示重复/跳过"
                : hasInvoice
                    ? "有发票但未找到本地文件"
                    : needsReview
                        ? "需要判断是否有发票"
                        : "明确无发票";

        var action = status switch
        {
            "已确认存在" => "无需处理",
            "明确无发票" => "无需处理",
            "发票日期不符" => "无需处理；不纳入本期报销",
            "疑似上期已报销" => "核对上期文件；确认后本期不重复下载和报销",
            "有发票但仅显示重复/跳过" => "打开下载清单核对重复来源；必要时重新下载异常邮件",
            "有发票但未找到本地文件" when isLinkPending || hasLink => "打开链接或二维码取票；下载后重新生成核验表",
            "有发票但未找到本地文件" => "重新下载该邮件或人工打开邮件核对附件",
            _ => "人工确认有无发票；确认后保存下载规则"
        };

        return new PresenceReviewRow
        {
            CheckStatus = status,
            NextAction = action,
            FinalDecision = item.FinalDecision,
            RuleDecision = item.RuleDecision,
            ManualDecision = item.ManualDecision,
            Date = item.Date,
            Subject = item.Subject,
            Status = item.Status,
            AttachmentCount = item.AttachmentCount,
            DownloadedFileCount = item.DownloadedFileCount,
            SkippedOrDuplicateCount = item.SkippedOrDuplicateCount,
            ExistingFileCount = existingFiles.Count,
            Reason = item.Reason,
            Files = item.Files,
            ExistingFiles = string.Join(Environment.NewLine, existingFiles),
            Urls = item.Urls,
            Note = item.Note
        };
    }

    private static void WritePresenceSheet(XLWorkbook wb, string name, List<PresenceReviewRow> rows)
    {
        var ws = wb.Worksheets.Add(name);
        var headers = new[]
        {
            "核验状态", "下一步动作", "最终判断", "规则判断", "人工判断", "日期", "邮件主题", "下载状态",
            "附件数", "下载数", "重复/跳过数", "本地存在文件数", "判断原因", "本地存在文件", "文件清单", "链接清单", "备注"
        };
        for (var c = 0; c < headers.Length; c++)
        {
            ws.Cell(1, c + 1).Value = headers[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
        }

        var r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.CheckStatus;
            ws.Cell(r, 2).Value = row.NextAction;
            ws.Cell(r, 3).Value = row.FinalDecision;
            ws.Cell(r, 4).Value = row.RuleDecision;
            ws.Cell(r, 5).Value = row.ManualDecision;
            ws.Cell(r, 6).Value = row.Date;
            ws.Cell(r, 7).Value = row.Subject;
            ws.Cell(r, 8).Value = row.Status;
            ws.Cell(r, 9).Value = row.AttachmentCount;
            ws.Cell(r, 10).Value = row.DownloadedFileCount;
            ws.Cell(r, 11).Value = row.SkippedOrDuplicateCount;
            ws.Cell(r, 12).Value = row.ExistingFileCount;
            ws.Cell(r, 13).Value = row.Reason;
            ws.Cell(r, 14).Value = row.ExistingFiles;
            ws.Cell(r, 15).Value = row.Files;
            ws.Cell(r, 16).Value = row.Urls;
            ws.Cell(r, 17).Value = row.Note;
            ws.Range(r, 1, r, headers.Length).Style.Fill.BackgroundColor = row.NextAction == "无需处理"
                ? XLColor.FromHtml("#D9EAD3")
                : XLColor.FromHtml("#FFF2CC");
            r++;
        }

        ws.Columns().AdjustToContents(8, 70);
        ws.SheetView.FreezeRows(1);
    }

    private static void WriteNonPdfSheet(XLWorkbook wb, string name, List<NonPdfReviewRow> rows)
    {
        var ws = wb.Worksheets.Add(name);
        var headers = new[] { "类型", "下一步动作", "大小KB", "文件" };
        for (var c = 0; c < headers.Length; c++)
        {
            ws.Cell(1, c + 1).Value = headers[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
        }
        var r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.Type;
            ws.Cell(r, 2).Value = row.NextAction;
            ws.Cell(r, 3).Value = row.SizeKb;
            ws.Cell(r, 4).Value = row.File;
            ws.Range(r, 1, r, headers.Length).Style.Fill.BackgroundColor = row.Type == "ICO图标"
                ? XLColor.FromHtml("#E5E7EB")
                : XLColor.FromHtml("#FFF2CC");
            r++;
        }
        ws.Columns().AdjustToContents(8, 80);
        ws.SheetView.FreezeRows(1);
    }

    private static void WritePresenceSummary(XLWorkbook wb, List<PresenceReviewRow> rows, List<NonPdfReviewRow> nonPdfRows)
    {
        var ws = wb.Worksheets.Add("统计汇总");
        var metrics = new (string Name, int Count)[]
        {
            ("邮件总数", rows.Count),
            ("明确有发票邮件", rows.Count(r => r.FinalDecision == HasInvoice)),
            ("有发票且本地已存在", rows.Count(r => r.FinalDecision == HasInvoice && r.ExistingFileCount > 0)),
            ("有发票但需继续核验", rows.Count(r => r.FinalDecision == HasInvoice && r.NextAction != "无需处理")),
            ("疑似上期已报销邮件", rows.Count(r => r.FinalDecision == PreviousReimbursed)),
            ("发票日期不符邮件", rows.Count(r => r.FinalDecision == DateOutOfRange)),
            ("需要人工判断邮件", rows.Count(r => r.FinalDecision == NeedsReview)),
            ("明确无发票邮件", rows.Count(r => r.FinalDecision == NoInvoice)),
            ("图片/二维码候选", nonPdfRows.Count(r => r.Type != "ICO图标")),
            ("ICO图标", nonPdfRows.Count(r => r.Type == "ICO图标"))
        };
        ws.Cell(1, 1).Value = "指标";
        ws.Cell(1, 2).Value = "数量";
        ws.Range(1, 1, 1, 2).Style.Font.Bold = true;
        for (var i = 0; i < metrics.Length; i++)
        {
            ws.Cell(i + 2, 1).Value = metrics[i].Name;
            ws.Cell(i + 2, 2).Value = metrics[i].Count;
        }
        ws.Columns().AdjustToContents();
    }

    private static List<string> SplitFileList(string files)
    {
        return Regex.Split(files ?? "", @"[\r\n；;]+")
            .Select(v => v.Trim())
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .ToList();
    }

    private static bool IsImageOrIcon(string file)
    {
        return Path.GetExtension(file).ToLowerInvariant() is ".ico" or ".png" or ".jpg" or ".jpeg" or ".webp";
    }

    private static List<Dictionary<string, string>> ReadCsv(string path)
    {
        var lines = File.ReadAllLines(path, Encoding.UTF8);
        if (lines.Length == 0) return [];
        var headers = ParseCsvLine(lines[0]);
        return lines.Skip(1).Where(line => !string.IsNullOrWhiteSpace(line)).Select(line =>
        {
            var values = ParseCsvLine(line);
            var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < headers.Count; i++) row[headers[i]] = i < values.Count ? values[i] : "";
            return row;
        }).ToList();
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
                if (ch == '"' && i + 1 < line.Length && line[i + 1] == '"') { current.Append('"'); i++; }
                else if (ch == '"') quoted = false;
                else current.Append(ch);
            }
            else if (ch == ',') { result.Add(current.ToString()); current.Clear(); }
            else if (ch == '"') quoted = true;
            else current.Append(ch);
        }
        result.Add(current.ToString());
        return result;
    }

    private static string Get(Dictionary<string, string> row, string key) => row.GetValueOrDefault(key, "");
    private static string FirstNonEmpty(params string[] values) => values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "";
    private static int? ParseInt(string? value) => int.TryParse(value, out var result) ? result : null;
    private static bool LooksLikeInvoice(string value, string context) => ContainsAny($"{value} {context}",
        "发票", "电子发票", "发票金额", "报销凭证", "电子报销凭证", "行程单", "客票行程单", "invoice", "fapiao", "dzfp", ".ofd", ".pdf", ".xml", ".zip", "开票");
    private static bool ContainsAny(string value, params string[] keywords) => keywords.Any(k => value.Contains(k, StringComparison.OrdinalIgnoreCase));
    private static List<string> ExtractKeywords(string subject) => Regex.Matches(subject, @"[\u4e00-\u9fffA-Za-z0-9]{2,}").Select(m => m.Value).Where(v => v.Length >= 2 && !Regex.IsMatch(v, @"^\d+$")).Distinct(StringComparer.OrdinalIgnoreCase).Take(12).ToList();

    private sealed class PresenceReviewRow
    {
        public string CheckStatus { get; set; } = "";
        public string NextAction { get; set; } = "";
        public string FinalDecision { get; set; } = "";
        public string RuleDecision { get; set; } = "";
        public string ManualDecision { get; set; } = "";
        public string Date { get; set; } = "";
        public string Subject { get; set; } = "";
        public string Status { get; set; } = "";
        public int AttachmentCount { get; set; }
        public int DownloadedFileCount { get; set; }
        public int SkippedOrDuplicateCount { get; set; }
        public int ExistingFileCount { get; set; }
        public string Reason { get; set; } = "";
        public string Files { get; set; } = "";
        public string ExistingFiles { get; set; } = "";
        public string Urls { get; set; } = "";
        public string Note { get; set; } = "";
    }

    private sealed class NonPdfReviewRow
    {
        public string Type { get; set; } = "";
        public string NextAction { get; set; } = "";
        public string File { get; set; } = "";
        public long SizeKb { get; set; }
    }
}
