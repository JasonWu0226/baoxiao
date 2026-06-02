using System.Text.Json.Serialization;

namespace ReimbursementMatcher;

public sealed class AppConfig
{
    public string SourceRoot { get; set; } = "报销准备的资料";
    public string InvoiceDir { get; set; } = "所有发票";
    public string TemplatePath { get; set; } = "报销准备的资料/报销输出文件模板/报销模版.xlsx";
    public string OutputDir { get; set; } = "输出";
    public string RuleDir { get; set; } = "规则";
    public string ArchiveDir { get; set; } = "输出/归档";
    public List<string> PreviousInvoiceDirs { get; set; } = new();
    public string CompanyName { get; set; } = "深圳博锐创科技有限公司";
    public string CompanyTaxId { get; set; } = "91440300MA5ECG7E71";
    public string Operator { get; set; } = "Jason";
    public string DateStart { get; set; } = "2026-01-01";
    public string DateEnd { get; set; } = "2026-04-30";
    public EmailConfig Email { get; set; } = new();
    public AiConfig Ai { get; set; } = new();
}

public sealed class EmailConfig
{
    public bool Enabled { get; set; } = true;
    public string Host { get; set; } = "imap.qq.com";
    public int Port { get; set; } = 993;
    public string Folder { get; set; } = "inbox";
    public string User { get; set; } = "290432815@qq.com";
    public string Password { get; set; } = "";
    public string OutputDir { get; set; } = "所有发票";
    public string Start { get; set; } = "2026-01-01";
    public string End { get; set; } = "2026-04-30";
    public double RequestIntervalSec { get; set; } = 0.1;
}

public sealed class AiConfig
{
    public bool Enabled { get; set; }
    public string BaseUrl { get; set; } = "https://api.mimo-v2.com/v1";
    public string ApiKey { get; set; } = "";
    public string ProxyUrl { get; set; } = "";
    public string Model { get; set; } = "mimo-v2.5-pro";
    public string VisionModel { get; set; } = "mimo-v2-omni";
    public double ConfidenceThreshold { get; set; } = 0.75;
    public int MaxItemsPerRun { get; set; } = 50;
}

public sealed class EvidenceItem
{
    public string Id { get; set; } = "";
    public string Kind { get; set; } = "";
    public string Platform { get; set; } = "";
    public string Date { get; set; } = "";
    public decimal Amount { get; set; }
    public string Title { get; set; } = "";
    public string Vendor { get; set; } = "";
    public string InvoiceNumber { get; set; } = "";
    public string FilePath { get; set; } = "";
    public string RelativePath { get; set; } = "";
    public string ReimburseDecision { get; set; } = "待确认";
    public string FileDecision { get; set; } = "待确认";
    public string MatchStatus { get; set; } = "未匹配";
    public string Project { get; set; } = "";
    public string Note { get; set; } = "";
    public int Weight { get; set; }
    public string Suggestion { get; set; } = "";
    public string ExtractedText { get; set; } = "";
    public string FileHash { get; set; } = "";
    public string DuplicateInfo { get; set; } = "";
}

public sealed class DownloadRecord
{
    public string Status { get; set; } = "";
    public string DownloadStatus { get; set; } = "";
    public string Kind { get; set; } = "";
    public string Date { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Sender { get; set; } = "";
    public string MessageId { get; set; } = "";
    public string MessageKey { get; set; } = "";
    public string File { get; set; } = "";
    public string Url { get; set; } = "";
    public string Error { get; set; } = "";
    public string InvoiceNumber { get; set; } = "";
    public string InvoiceAmount { get; set; } = "";
    public string InvoiceDate { get; set; } = "";
    public string Md5 { get; set; } = "";
    public string Sha256 { get; set; } = "";
    public string SavedFileCount { get; set; } = "";
    public string AttachmentTotal { get; set; } = "";
    public string LinkCandidateTotal { get; set; } = "";
    public string DuplicateOf { get; set; } = "";
}

public sealed class EmailProcessingStore
{
    public int Version { get; set; } = 1;
    public string UpdatedAt { get; set; } = "";
    public Dictionary<string, EmailProcessingRecord> Messages { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, EmailInvoiceRecord> Invoices { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class EmailProcessingRecord
{
    public string MessageKey { get; set; } = "";
    public string MessageId { get; set; } = "";
    public string Date { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Sender { get; set; } = "";
    public string Status { get; set; } = "";
    public string DownloadStatus { get; set; } = "";
    public string ErrorReason { get; set; } = "";
    public bool HasAttachment { get; set; }
    public bool IsInvoiceMail { get; set; }
    public bool HasRisk { get; set; }
    public int AttachmentTotal { get; set; }
    public int LinkCandidateTotal { get; set; }
    public int SavedFileCount { get; set; }
    public int DuplicateCount { get; set; }
    public string DuplicateOf { get; set; } = "";
    public string DuplicateSourceMailId { get; set; } = "";
    public string HistoricalMatchType { get; set; } = "";
    public string HistoricalMatchSummary { get; set; } = "";
    public string HistoricalMatchFile { get; set; } = "";
    public int Attempts { get; set; }
    public List<string> KeywordHits { get; set; } = new();
    public List<string> AttachmentNames { get; set; } = new();
    public List<string> Files { get; set; } = new();
    public List<EmailSavedFileRecord> SavedFiles { get; set; } = new();
    public string UpdatedAt { get; set; } = "";
}

public sealed class EmailSavedFileRecord
{
    public string Path { get; set; } = "";
    public string FileName { get; set; } = "";
    public string Md5 { get; set; } = "";
    public string Sha256 { get; set; } = "";
    public string Extension { get; set; } = "";
    public string InvoiceNumber { get; set; } = "";
    public string InvoiceAmount { get; set; } = "";
    public string InvoiceDate { get; set; } = "";
    public string TextExtractStatus { get; set; } = "";
    public string Error { get; set; } = "";
}

public sealed class EmailInvoiceRecord
{
    public string InvoiceKey { get; set; } = "";
    public string MessageKey { get; set; } = "";
    public string MessageId { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Sender { get; set; } = "";
    public string Date { get; set; } = "";
    public string Path { get; set; } = "";
    public string Md5 { get; set; } = "";
    public string Sha256 { get; set; } = "";
    public string InvoiceNumber { get; set; } = "";
    public string InvoiceAmount { get; set; } = "";
    public string InvoiceDate { get; set; } = "";
    public string UpdatedAt { get; set; } = "";
}

public sealed class EmailAuditItem
{
    public string MessageKey { get; set; } = "";
    public string Date { get; set; } = "";
    public string Subject { get; set; } = "";
    public string MessageId { get; set; } = "";
    public string RuleDecision { get; set; } = "";
    public string ManualDecision { get; set; } = "";
    public string FinalDecision { get; set; } = "";
    public string Status { get; set; } = "";
    public bool HasInvoiceKeyword { get; set; }
    public bool HasPlatformKeyword { get; set; }
    public bool HasAttachment { get; set; }
    public int AttachmentCount { get; set; }
    public bool HasLikelyInvoiceAttachment { get; set; }
    public bool HasLikelyInvoiceLink { get; set; }
    public int DownloadedFileCount { get; set; }
    public int SkippedOrDuplicateCount { get; set; }
    public bool NeedsHumanReview { get; set; }
    public string HistoricalMatch { get; set; } = "";
    public string HistoricalMatchFile { get; set; } = "";
    public string Reason { get; set; } = "";
    public string Files { get; set; } = "";
    public string Urls { get; set; } = "";
    public string Note { get; set; } = "";
}

public sealed class PreviousReimbursementIndex
{
    public int Version { get; set; } = 1;
    public string UpdatedAt { get; set; } = "";
    public List<string> SourceRoots { get; set; } = new();
    public List<PreviousReimbursementInvoice> Invoices { get; set; } = new();
    public List<PreviousReimbursementDocument> Documents { get; set; } = new();
}

public sealed class PreviousReimbursementInvoice
{
    public string SourceRoot { get; set; } = "";
    public string SourcePath { get; set; } = "";
    public string FileName { get; set; } = "";
    public string Sha256 { get; set; } = "";
    public string InvoiceNumber { get; set; } = "";
    public decimal Amount { get; set; }
    public string InvoiceDate { get; set; } = "";
    public string Vendor { get; set; } = "";
    public string Category { get; set; } = "";
    public string RelatedDocument { get; set; } = "";
}

public sealed class PreviousReimbursementDocument
{
    public string SourceRoot { get; set; } = "";
    public string SourcePath { get; set; } = "";
    public string FileName { get; set; } = "";
    public string Kind { get; set; } = "";
    public long SizeBytes { get; set; }
    public string UpdatedAt { get; set; } = "";
}

public sealed class PreviousReimbursementMatch
{
    public bool IsMatch { get; set; }
    public int Score { get; set; }
    public string MatchType { get; set; } = "";
    public string Summary { get; set; } = "";
    public string SourcePath { get; set; } = "";
    public string InvoiceNumber { get; set; } = "";
    public decimal Amount { get; set; }
    public string InvoiceDate { get; set; } = "";
    public string Vendor { get; set; } = "";
}

public sealed class AiReviewResult
{
    public string Decision { get; set; } = "needs_review";
    public double Confidence { get; set; }
    public string Reason { get; set; } = "";
    public string Action { get; set; } = "";
    public string Url { get; set; } = "";
}

public sealed class EmailDecisionStore
{
    public int Version { get; set; } = 1;
    public string UpdatedAt { get; set; } = "";
    public Dictionary<string, EmailDecisionRecord> Messages { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public List<ConfirmationEvent> Events { get; set; } = new();
}

public sealed class EmailDecisionRecord
{
    public string MessageKey { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Decision { get; set; } = "";
    public string Note { get; set; } = "";
    public string UpdatedAt { get; set; } = "";
}

public sealed class MatchCandidate
{
    public string Id { get; set; } = "";
    public string ExpenseId { get; set; } = "";
    public string InvoiceId { get; set; } = "";
    public string ExpenseTitle { get; set; } = "";
    public string InvoiceTitle { get; set; } = "";
    public decimal ExpenseAmount { get; set; }
    public decimal InvoiceAmount { get; set; }
    public string ExpenseDate { get; set; } = "";
    public string InvoiceDate { get; set; } = "";
    public int Score { get; set; }
    public string Reason { get; set; } = "";
    public string Decision { get; set; } = "待确认";
}

public sealed class ConfirmationStore
{
    public int Version { get; set; } = 1;
    public string UpdatedAt { get; set; } = "";
    public Dictionary<string, ItemConfirmation> Items { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, MatchConfirmation> Matches { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public List<ConfirmationEvent> Events { get; set; } = new();
}

public sealed class ItemConfirmation
{
    public string ItemId { get; set; } = "";
    public string Kind { get; set; } = "";
    public string Platform { get; set; } = "";
    public string Title { get; set; } = "";
    public string ReimburseDecision { get; set; } = "待确认";
    public string FileDecision { get; set; } = "待确认";
    public string Project { get; set; } = "";
    public string Note { get; set; } = "";
    public string UpdatedAt { get; set; } = "";
}

public sealed class MatchConfirmation
{
    public string MatchId { get; set; } = "";
    public string ExpenseId { get; set; } = "";
    public string InvoiceId { get; set; } = "";
    public string Decision { get; set; } = "待确认";
    public string Note { get; set; } = "";
    public string UpdatedAt { get; set; } = "";
}

public sealed class ConfirmationEvent
{
    public string Time { get; set; } = "";
    public string ObjectType { get; set; } = "";
    public string ObjectId { get; set; } = "";
    public string Action { get; set; } = "";
    public string Before { get; set; } = "";
    public string After { get; set; } = "";
    public string Note { get; set; } = "";
}

public sealed class DecisionWeightStore
{
    public int Version { get; set; } = 1;
    public string UpdatedAt { get; set; } = "";
    public List<DecisionWeight> Weights { get; set; } = new();
}

public sealed class DecisionWeight
{
    public string Platform { get; set; } = "";
    public string Key { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public int ReimburseCount { get; set; }
    public int ExcludeCount { get; set; }
    public int PendingCount { get; set; }
    public int Weight { get; set; }
    public string Suggestion { get; set; } = "";
}

public static class EvidenceKinds
{
    public const string Invoice = "发票";
    public const string PaymentScreenshot = "付款截图";
    public const string OrderSource = "订单素材";
    public const string TransactionSource = "交易流水";
    public const string Template = "模板";
    public const string Other = "其他";
}
