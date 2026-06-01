#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动化发票邮件下载与查重程序

功能：
1. 通过 IMAP 扫描指定时间之后的邮件。
2. 按邮件标题/正文关键字识别发票邮件。
3. 优先下载 PDF；同一封邮件同时存在 PDF 和图片时只保留 PDF。
4. 按文件 MD5、发票号码、金额、开票日期进行查重。
5. 保存 JSON/CSV 状态记录，支持增量处理和异常邮件重新处理。

运行：
    python invoice_mail_downloader.py --config invoice_mail_config.json
    python invoice_mail_downloader.py --config invoice_mail_config.json --reprocess-abnormal
    python invoice_mail_downloader.py --config invoice_mail_config.json --reprocess-message INBOX:12345
"""

from __future__ import annotations

import argparse
import csv
import getpass
import hashlib
import imaplib
import json
import os
import re
import shutil
import sys
from dataclasses import dataclass, field
from datetime import datetime
from email import policy
from email.header import decode_header, make_header
from email.message import EmailMessage, Message
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any


STATUS_UNPROCESSED = "未处理"
STATUS_DONE = "正常已下载"
STATUS_NO_INVOICE = "非发票邮件"
STATUS_DUPLICATE = "重复发票"
STATUS_ERROR = "异常"

REASON_NO_ATTACHMENT = "无附件"
REASON_NO_KEYWORD = "无关键词"
REASON_NO_PDF = "无 PDF"
REASON_BAD_FORMAT = "格式错误"
REASON_DOWNLOAD_FAILED = "下载失败"
REASON_UNSUPPORTED = "格式不支持"
REASON_NO_INVOICE_NUMBER = "无发票号"
REASON_OUT_OF_RANGE = "发票日期早于开始时间"


@dataclass
class Config:
    imap_server: str
    imap_port: int
    username: str
    password: str
    password_env: str
    folder: str
    start_time: str
    end_time: str
    use_ssl: bool
    save_dir: Path
    record_json: Path
    record_csv: Path
    invoice_keywords: list[str]
    legal_extensions: list[str]
    preferred_extensions: list[str]
    skip_completed: bool
    prefer_pdf: bool
    keep_duplicate_file: bool


@dataclass
class AttachmentCandidate:
    filename: str
    extension: str
    part: Message


@dataclass
class SavedFile:
    path: str = ""
    filename: str = ""
    md5: str = ""
    sha256: str = ""
    extension: str = ""
    invoice_number: str = ""
    invoice_amount: str = ""
    invoice_date: str = ""
    text_extract_status: str = ""
    error: str = ""


@dataclass
class MailRecord:
    mail_id: str
    message_id: str = ""
    subject: str = ""
    sender: str = ""
    received_at: str = ""
    has_attachment: bool = False
    is_invoice_mail: bool = False
    status: str = STATUS_UNPROCESSED
    error_reason: str = ""
    keyword_hits: list[str] = field(default_factory=list)
    attachment_names: list[str] = field(default_factory=list)
    saved_files: list[dict[str, Any]] = field(default_factory=list)
    duplicate_of: str = ""
    duplicate_source_mail_id: str = ""
    download_status: str = ""
    processed_at: str = ""
    attempts: int = 0
    notes: list[str] = field(default_factory=list)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "MailRecord":
        known = {f.name for f in cls.__dataclass_fields__.values()}  # type: ignore[attr-defined]
        return cls(**{k: v for k, v in data.items() if k in known})


def load_config(path: Path, require_password: bool = True) -> Config:
    if not path.exists():
        raise FileNotFoundError(f"配置文件不存在：{path}")

    raw = json.loads(path.read_text(encoding="utf-8"))
    email_cfg = raw.get("email", {})
    paths = raw.get("paths", {})
    rules = raw.get("rules", {})

    password_env = email_cfg.get("password_env", "INVOICE_MAIL_PASSWORD")
    password = email_cfg.get("password") or os.environ.get(password_env, "")
    if require_password and not password:
        password = getpass.getpass("请输入邮箱授权码/密码：")

    return Config(
        imap_server=email_cfg.get("imap_server", "imap.qq.com"),
        imap_port=int(email_cfg.get("imap_port", 993)),
        username=email_cfg.get("username", ""),
        password=password,
        password_env=password_env,
        folder=email_cfg.get("folder", "INBOX"),
        start_time=email_cfg.get("start_time", "2026-01-01"),
        end_time=email_cfg.get("end_time", ""),
        use_ssl=bool(email_cfg.get("use_ssl", True)),
        save_dir=Path(paths.get("save_dir", "所有发票")),
        record_json=Path(paths.get("record_json", "规则/发票邮件处理记录.json")),
        record_csv=Path(paths.get("record_csv", "输出/发票邮件处理记录.csv")),
        invoice_keywords=list(rules.get("invoice_keywords", ["发票", "电子发票", "增值税发票", "专票", "普票", "数电票", "报销", "invoice"])),
        legal_extensions=[x.lower() for x in rules.get("legal_extensions", [".pdf", ".ofd", ".xml", ".jpg", ".jpeg", ".png", ".zip"])],
        preferred_extensions=[x.lower() for x in rules.get("preferred_extensions", [".pdf", ".ofd", ".xml", ".jpg", ".jpeg", ".png", ".zip"])],
        skip_completed=bool(rules.get("skip_completed", True)),
        prefer_pdf=bool(rules.get("prefer_pdf", True)),
        keep_duplicate_file=bool(rules.get("keep_duplicate_file", False)),
    )


def load_store(record_json: Path) -> dict[str, Any]:
    if not record_json.exists():
        return {"version": 1, "updated_at": "", "mails": {}, "invoices": {}}
    try:
        data = json.loads(record_json.read_text(encoding="utf-8"))
        data.setdefault("version", 1)
        data.setdefault("mails", {})
        data.setdefault("invoices", {})
        return data
    except Exception:
        backup = record_json.with_suffix(".broken.json")
        shutil.copy2(record_json, backup)
        return {"version": 1, "updated_at": "", "mails": {}, "invoices": {}}


def save_store(store: dict[str, Any], config: Config) -> None:
    store["updated_at"] = now_text()
    config.record_json.parent.mkdir(parents=True, exist_ok=True)
    config.record_json.write_text(json.dumps(store, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    export_csv(store, config.record_csv)


def export_csv(store: dict[str, Any], csv_path: Path) -> None:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    rows = []
    for raw in store.get("mails", {}).values():
        rec = MailRecord.from_dict(raw)
        rows.append({
            "邮件ID": rec.mail_id,
            "邮件Message-ID": rec.message_id,
            "邮件主题": rec.subject,
            "发件人": rec.sender,
            "接收时间": rec.received_at,
            "是否含附件": "是" if rec.has_attachment else "否",
            "是否为发票邮件": "是" if rec.is_invoice_mail else "否",
            "处理状态": rec.status,
            "异常原因": rec.error_reason,
            "命中关键词": "；".join(rec.keyword_hits),
            "附件": "；".join(rec.attachment_names),
            "保存文件": "；".join(str(f.get("path", "")) for f in rec.saved_files),
            "重复对应发票": rec.duplicate_of,
            "重复来源邮件ID": rec.duplicate_source_mail_id,
            "下载状态": rec.download_status,
            "处理时间": rec.processed_at,
            "处理次数": rec.attempts,
        })

    headers = [
        "邮件ID", "邮件Message-ID", "邮件主题", "发件人", "接收时间", "是否含附件", "是否为发票邮件",
        "处理状态", "异常原因", "命中关键词", "附件", "保存文件", "重复对应发票", "重复来源邮件ID", "下载状态",
        "处理时间", "处理次数",
    ]
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        writer.writerows(rows)


def connect_imap(config: Config) -> imaplib.IMAP4:
    if not config.username:
        raise ValueError("配置缺少邮箱账号 username")
    client: imaplib.IMAP4
    if config.use_ssl:
        client = imaplib.IMAP4_SSL(config.imap_server, config.imap_port)
    else:
        client = imaplib.IMAP4(config.imap_server, config.imap_port)
    client.login(config.username, config.password)
    status, _ = client.select(config.folder, readonly=True)
    if status != "OK":
        raise RuntimeError(f"无法打开邮箱文件夹：{config.folder}")
    return client


def search_mail_uids(client: imaplib.IMAP4, config: Config) -> list[str]:
    criteria = ["SINCE", imap_date(config.start_time)]
    if config.end_time:
        criteria.extend(["BEFORE", imap_date(config.end_time, add_one_day=True)])
    status, data = client.uid("search", None, *criteria)
    if status != "OK":
        raise RuntimeError("IMAP 搜索失败")
    raw = data[0] or b""
    return [x.decode("ascii") for x in raw.split() if x]


def fetch_message(client: imaplib.IMAP4, uid: str) -> EmailMessage:
    status, data = client.uid("fetch", uid, "(RFC822)")
    if status != "OK" or not data:
        raise RuntimeError(f"邮件下载失败 UID={uid}")
    for item in data:
        if isinstance(item, tuple):
            return BytesParser(policy=policy.default).parsebytes(item[1])
    raise RuntimeError(f"邮件内容为空 UID={uid}")


def process_all(config: Config, reprocess_abnormal: bool = False, reprocess_message: str = "") -> None:
    config.save_dir.mkdir(parents=True, exist_ok=True)
    store = load_store(config.record_json)
    client = connect_imap(config)
    try:
        uids = search_mail_uids(client, config)
        print(f"找到候选邮件 {len(uids)} 封")
        for index, uid in enumerate(uids, start=1):
            mail_id = f"{config.folder}:{uid}"
            old = store["mails"].get(mail_id)
            if should_skip(old, config, reprocess_abnormal, reprocess_message, mail_id):
                continue
            try:
                msg = fetch_message(client, uid)
                rec = process_message(config, store, mail_id, msg)
                store["mails"][mail_id] = record_to_dict(rec)
            except Exception as exc:
                rec = MailRecord.from_dict(old) if old else MailRecord(mail_id=mail_id)
                rec.status = STATUS_ERROR
                rec.error_reason = f"{REASON_DOWNLOAD_FAILED}: {exc}"
                rec.processed_at = now_text()
                rec.attempts += 1
                store["mails"][mail_id] = record_to_dict(rec)
            if index % 20 == 0:
                save_store(store, config)
                print(f"已处理 {index}/{len(uids)} 封")
        save_store(store, config)
    finally:
        try:
            client.close()
        except Exception:
            pass
        client.logout()

    print(f"处理完成：JSON={config.record_json} CSV={config.record_csv}")


def should_skip(old: dict[str, Any] | None, config: Config, reprocess_abnormal: bool, reprocess_message: str, mail_id: str) -> bool:
    if reprocess_message:
        return mail_id != reprocess_message
    if not old or not config.skip_completed:
        return False
    status = old.get("status", "")
    if reprocess_abnormal:
        return status != STATUS_ERROR
    return status in {STATUS_DONE, STATUS_NO_INVOICE, STATUS_DUPLICATE}


def process_message(config: Config, store: dict[str, Any], mail_id: str, msg: EmailMessage) -> MailRecord:
    subject = decode_mime_header(msg.get("Subject", ""))
    sender = decode_mime_header(msg.get("From", ""))
    message_id = str(msg.get("Message-ID", ""))
    received_at = parse_received_time(msg)
    body_text = extract_body_text(msg)
    attachments = collect_attachments(msg, config)
    keyword_hits = find_keywords(f"{subject}\n{body_text}", config.invoice_keywords)

    old_raw = store.get("mails", {}).get(mail_id)
    rec = MailRecord.from_dict(old_raw) if old_raw else MailRecord(mail_id=mail_id)
    rec.mail_id = mail_id
    rec.message_id = message_id
    rec.subject = subject
    rec.sender = sender
    rec.received_at = received_at
    rec.has_attachment = bool(attachments)
    rec.keyword_hits = keyword_hits
    rec.attachment_names = [a.filename for a in attachments]
    rec.saved_files = []
    rec.duplicate_of = ""
    rec.duplicate_source_mail_id = ""
    rec.download_status = ""
    rec.processed_at = now_text()
    rec.attempts += 1

    if not attachments:
        rec.is_invoice_mail = bool(keyword_hits)
        rec.status = STATUS_ERROR if rec.is_invoice_mail else STATUS_NO_INVOICE
        rec.download_status = "异常" if rec.is_invoice_mail else ""
        rec.error_reason = REASON_NO_ATTACHMENT if rec.is_invoice_mail else REASON_NO_KEYWORD
        return rec

    if not keyword_hits:
        rec.is_invoice_mail = False
        rec.status = STATUS_NO_INVOICE
        rec.error_reason = REASON_NO_KEYWORD
        return rec

    rec.is_invoice_mail = True
    selected = choose_attachments(attachments, config)
    if not selected:
        rec.status = STATUS_ERROR
        rec.download_status = "异常"
        rec.error_reason = REASON_UNSUPPORTED
        return rec

    saved_files: list[SavedFile] = []
    errors: list[str] = []
    for candidate in selected:
        try:
            saved = save_attachment(config, mail_id, rec, candidate)
            saved_files.append(saved)
        except Exception as exc:
            errors.append(f"{candidate.filename}: {exc}")

    if not saved_files:
        rec.status = STATUS_ERROR
        rec.download_status = "异常"
        rec.error_reason = f"{REASON_DOWNLOAD_FAILED}: {'；'.join(errors)}" if errors else REASON_DOWNLOAD_FAILED
        return rec

    kept_files: list[SavedFile] = []
    for saved in saved_files:
        duplicate_invoice_key, duplicate_mail_id = detect_duplicate(store, saved)
        if duplicate_invoice_key:
            rec.status = STATUS_DUPLICATE
            rec.download_status = "重复已忽略"
            rec.error_reason = "重复发票"
            rec.duplicate_of = duplicate_invoice_key
            rec.duplicate_source_mail_id = duplicate_mail_id
            if not config.keep_duplicate_file and saved.path and Path(saved.path).exists():
                Path(saved.path).unlink()
                saved.path = ""
            rec.saved_files.append(saved_file_to_dict(saved))
            continue

        invoice_key = build_invoice_key(saved)
        if not saved.invoice_number and saved.extension == ".pdf":
            saved.error = REASON_NO_INVOICE_NUMBER
        if is_invoice_date_before_start(saved.invoice_date, config.start_time):
            rec.status = STATUS_ERROR
            rec.download_status = "异常"
            rec.error_reason = REASON_OUT_OF_RANGE
            rec.saved_files.append(saved_file_to_dict(saved))
            continue

        store.setdefault("invoices", {})[invoice_key] = {
            "invoice_key": invoice_key,
            "mail_id": rec.mail_id,
            "message_id": rec.message_id,
            "subject": rec.subject,
            "sender": rec.sender,
            "received_at": rec.received_at,
            "path": saved.path,
            "md5": saved.md5,
            "sha256": saved.sha256,
            "invoice_number": saved.invoice_number,
            "invoice_amount": saved.invoice_amount,
            "invoice_date": saved.invoice_date,
            "updated_at": now_text(),
        }
        kept_files.append(saved)
        rec.saved_files.append(saved_file_to_dict(saved))

    if rec.status == STATUS_DUPLICATE:
        return rec
    if kept_files:
        missing_number = [f.filename for f in kept_files if f.extension == ".pdf" and not f.invoice_number]
        has_pdf = any(f.extension == ".pdf" for f in kept_files)
        if missing_number:
            rec.status = STATUS_ERROR
            rec.download_status = "异常"
            rec.error_reason = f"{REASON_NO_INVOICE_NUMBER}: {'；'.join(missing_number)}"
        elif not has_pdf:
            rec.status = STATUS_ERROR
            rec.download_status = "异常"
            rec.error_reason = REASON_NO_PDF
        else:
            rec.status = STATUS_DONE
            rec.download_status = "正常已下载"
            rec.error_reason = ""
        return rec

    rec.status = STATUS_ERROR
    rec.download_status = "异常"
    rec.error_reason = rec.error_reason or "未保存有效发票"
    return rec


def choose_attachments(attachments: list[AttachmentCandidate], config: Config) -> list[AttachmentCandidate]:
    legal = [a for a in attachments if a.extension in config.legal_extensions]
    if not legal:
        return []
    if config.prefer_pdf:
        pdfs = [a for a in legal if a.extension == ".pdf"]
        if pdfs:
            return pdfs
    result: list[AttachmentCandidate] = []
    for ext in config.preferred_extensions:
        same_ext = [a for a in legal if a.extension == ext]
        if same_ext:
            result.extend(same_ext)
            break
    return result


def save_attachment(config: Config, mail_id: str, rec: MailRecord, candidate: AttachmentCandidate) -> SavedFile:
    payload = candidate.part.get_payload(decode=True)
    if not payload:
        raise ValueError(REASON_DOWNLOAD_FAILED)

    md5 = hashlib.md5(payload).hexdigest()
    sha256 = hashlib.sha256(payload).hexdigest()
    sender_name = safe_filename(clean_sender(rec.sender))
    date_text = date_for_filename(rec.received_at)
    unique = safe_filename(mail_id.replace(":", "_"))
    ext = candidate.extension
    filename = f"发票_{date_text}_{sender_name}_{unique}{ext}"
    target = unique_path(config.save_dir / filename)
    target.write_bytes(payload)

    saved = SavedFile(
        path=str(target),
        filename=target.name,
        md5=md5,
        sha256=sha256,
        extension=ext,
    )
    if ext == ".pdf":
        text = extract_pdf_text(target)
        saved.text_extract_status = "成功" if text else "未提取到文本"
        saved.invoice_number = extract_invoice_number(text)
        saved.invoice_amount = extract_invoice_amount(text)
        saved.invoice_date = extract_invoice_date(text)
    return saved


def detect_duplicate(store: dict[str, Any], saved: SavedFile) -> tuple[str, str]:
    invoices = store.setdefault("invoices", {})
    for key, invoice in invoices.items():
        if saved.md5 and invoice.get("md5") == saved.md5:
            return key, invoice.get("mail_id", "")

    new_key = build_invoice_key(saved)
    if new_key in invoices:
        return new_key, invoices[new_key].get("mail_id", "")

    for key, invoice in invoices.items():
        if saved.invoice_number and saved.invoice_number == invoice.get("invoice_number"):
            same_amount = not saved.invoice_amount or not invoice.get("invoice_amount") or saved.invoice_amount == invoice.get("invoice_amount")
            same_date = not saved.invoice_date or not invoice.get("invoice_date") or saved.invoice_date == invoice.get("invoice_date")
            if same_amount and same_date:
                return key, invoice.get("mail_id", "")
    return "", ""


def build_invoice_key(saved: SavedFile) -> str:
    if saved.invoice_number:
        return f"number:{saved.invoice_number}|amount:{saved.invoice_amount}|date:{saved.invoice_date}"
    return f"md5:{saved.md5}"


def collect_attachments(msg: EmailMessage, config: Config) -> list[AttachmentCandidate]:
    result: list[AttachmentCandidate] = []
    for part in msg.walk():
        if part.is_multipart():
            continue
        filename = part.get_filename()
        disposition = (part.get_content_disposition() or "").lower()
        if not filename and disposition != "attachment":
            continue
        filename = decode_mime_header(filename or "attachment")
        ext = Path(filename).suffix.lower()
        if ext in config.legal_extensions:
            result.append(AttachmentCandidate(filename=filename, extension=ext, part=part))
    return result


def extract_body_text(msg: EmailMessage) -> str:
    texts: list[str] = []
    for part in msg.walk():
        if part.is_multipart():
            continue
        content_type = part.get_content_type()
        if content_type not in {"text/plain", "text/html"}:
            continue
        try:
            text = part.get_content()
        except Exception:
            payload = part.get_payload(decode=True) or b""
            charset = part.get_content_charset() or "utf-8"
            text = payload.decode(charset, errors="ignore")
        if content_type == "text/html":
            text = re.sub(r"<[^>]+>", " ", text)
        texts.append(text)
    return "\n".join(texts)


def extract_pdf_text(path: Path) -> str:
    try:
        from pypdf import PdfReader  # type: ignore
    except Exception:
        try:
            from PyPDF2 import PdfReader  # type: ignore
        except Exception:
            return ""

    try:
        reader = PdfReader(str(path))
        texts = []
        for page in reader.pages:
            try:
                texts.append(page.extract_text() or "")
            except Exception:
                continue
        return "\n".join(texts)
    except Exception:
        return ""


def extract_invoice_number(text: str) -> str:
    if not text:
        return ""
    patterns = [
        r"发票号码[:：\s]*([0-9]{8,30})",
        r"发票号[:：\s]*([0-9]{8,30})",
        r"票据号码[:：\s]*([0-9]{8,30})",
        r"Invoice\s*No\.?[:：\s]*([A-Za-z0-9-]{6,30})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return ""


def extract_invoice_amount(text: str) -> str:
    if not text:
        return ""
    patterns = [
        r"价税合计[（(]大写[）)]?.*?[¥￥]\s*([0-9]+(?:\.[0-9]{1,2})?)",
        r"小写[）)]?\s*[¥￥]?\s*([0-9]+(?:\.[0-9]{1,2})?)",
        r"合计\s*[¥￥]\s*([0-9]+(?:\.[0-9]{1,2})?)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.DOTALL)
        if match:
            return match.group(1)
    return ""


def extract_invoice_date(text: str) -> str:
    if not text:
        return ""
    patterns = [
        r"开票日期[:：\s]*([0-9]{4})年([0-9]{1,2})月([0-9]{1,2})日",
        r"开票日期[:：\s]*([0-9]{4})[-/.]([0-9]{1,2})[-/.]([0-9]{1,2})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            y, m, d = match.groups()
            return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
    return ""


def is_invoice_date_before_start(invoice_date: str, start_time: str) -> bool:
    if not invoice_date:
        return False
    try:
        return datetime.fromisoformat(invoice_date).date() < datetime.fromisoformat(start_time[:10]).date()
    except Exception:
        return False


def find_keywords(text: str, keywords: list[str]) -> list[str]:
    return [kw for kw in keywords if kw and kw.lower() in text.lower()]


def decode_mime_header(value: str) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def parse_received_time(msg: EmailMessage) -> str:
    value = msg.get("Date", "")
    if not value:
        return ""
    try:
        dt = parsedate_to_datetime(value)
        if dt.tzinfo:
            dt = dt.astimezone()
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return value


def clean_sender(sender: str) -> str:
    sender = re.sub(r"<[^>]+>", "", sender).strip().strip('"')
    return sender or "未知发件人"


def safe_filename(value: str) -> str:
    value = re.sub(r'[\\/:*?"<>|]+', "_", value)
    value = re.sub(r"\s+", " ", value).strip(" .")
    return value[:60] or "unknown"


def date_for_filename(received_at: str) -> str:
    match = re.match(r"(\d{4})-(\d{2})-(\d{2})", received_at or "")
    if match:
        return "".join(match.groups())
    return datetime.now().strftime("%Y%m%d")


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    for i in range(1, 10000):
        candidate = path.with_name(f"{stem}_{i}{suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"无法生成唯一文件名：{path}")


def imap_date(value: str, add_one_day: bool = False) -> str:
    dt = datetime.fromisoformat(value[:10])
    if add_one_day:
        from datetime import timedelta
        dt = dt + timedelta(days=1)
    return dt.strftime("%d-%b-%Y")


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def saved_file_to_dict(saved: SavedFile) -> dict[str, Any]:
    return {
        "path": saved.path,
        "filename": saved.filename,
        "md5": saved.md5,
        "sha256": saved.sha256,
        "extension": saved.extension,
        "invoice_number": saved.invoice_number,
        "invoice_amount": saved.invoice_amount,
        "invoice_date": saved.invoice_date,
        "text_extract_status": saved.text_extract_status,
        "error": saved.error,
    }


def record_to_dict(rec: MailRecord) -> dict[str, Any]:
    return {
        "mail_id": rec.mail_id,
        "message_id": rec.message_id,
        "subject": rec.subject,
        "sender": rec.sender,
        "received_at": rec.received_at,
        "has_attachment": rec.has_attachment,
        "is_invoice_mail": rec.is_invoice_mail,
        "status": rec.status,
        "error_reason": rec.error_reason,
        "keyword_hits": rec.keyword_hits,
        "attachment_names": rec.attachment_names,
        "saved_files": rec.saved_files,
        "duplicate_of": rec.duplicate_of,
        "duplicate_source_mail_id": rec.duplicate_source_mail_id,
        "download_status": rec.download_status,
        "processed_at": rec.processed_at,
        "attempts": rec.attempts,
        "notes": rec.notes,
    }


def ensure_config(path: Path) -> None:
    if path.exists():
        return
    example = Path("invoice_mail_config.example.json")
    if example.exists():
        shutil.copy2(example, path)
        print(f"已创建配置文件：{path}")
    else:
        raise FileNotFoundError("缺少 invoice_mail_config.example.json")


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="自动化发票邮件下载与查重程序")
    parser.add_argument("--config", default="invoice_mail_config.json", help="配置文件路径")
    parser.add_argument("--init-config", action="store_true", help="从示例创建配置文件")
    parser.add_argument("--reprocess-abnormal", action="store_true", help="只重新处理异常邮件")
    parser.add_argument("--reprocess-message", default="", help="重新处理指定邮件 ID，例如 INBOX:12345")
    parser.add_argument("--export-csv", action="store_true", help="仅根据 JSON 记录重新导出 CSV")
    args = parser.parse_args(argv)

    config_path = Path(args.config)
    if args.init_config:
        ensure_config(config_path)
        return 0

    ensure_config(config_path)
    if args.export_csv:
        config = load_config(config_path, require_password=False)
        store = load_store(config.record_json)
        export_csv(store, config.record_csv)
        print(f"已导出 CSV：{config.record_csv}")
        return 0

    config = load_config(config_path, require_password=True)
    process_all(config, reprocess_abnormal=args.reprocess_abnormal, reprocess_message=args.reprocess_message)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
