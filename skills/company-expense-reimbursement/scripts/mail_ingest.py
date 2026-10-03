#!/usr/bin/env python3
"""Read a configured mailbox and download likely invoice attachments.

The command is deliberately read-only.  It never marks mail as read, moves,
deletes, or sends messages.  Gmail uses the read-only Gmail API when its
optional dependencies and OAuth client file are configured.  Any IMAP server
(including NetEase enterprise mail and Gmail with an app password) is also
supported without a provider-specific Python package.

Credentials, OAuth tokens, and state files should stay outside this public
repository.  The output directory may be a case's ignored ``03-发票凭证``
directory.
"""

from __future__ import annotations

import argparse
import base64
import email
import getpass
import hashlib
import html
import imaplib
import json
import os
import re
import sys
import tempfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from email import policy
from email.header import decode_header
from email.message import Message
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any, Callable, Iterable

import yaml


SUPPORTED_EXTENSIONS = {".pdf", ".png", ".jpg", ".jpeg", ".ofd"}
STRONG_KEYWORDS = (
    "发票",
    "电子发票",
    "增值税",
    "行程单",
    "invoice",
    "tax invoice",
    "receipt",
    "dzfp",
    "itinerary",
)
TRAVEL_KEYWORDS = (
    "滴滴",
    "打车",
    "网约车",
    "住宿",
    "酒店",
    "机票",
    "飞机",
    "火车",
    "高铁",
    "飞猪",
    "出行",
    "taxi",
    "hotel",
    "flight",
    "train",
)


def default_private_dir() -> Path:
    user_profile = os.environ.get("USERPROFILE")
    if user_profile:
        return Path(user_profile) / ".codex" / "private"
    return Path.home() / ".codex" / "private"


def default_config_path() -> Path:
    return default_private_dir() / "mail-config.yaml"


def expand_path(value: str | Path) -> Path:
    expanded = os.path.expandvars(os.path.expanduser(str(value)))
    return Path(expanded).resolve()


def decode_header_value(value: str | None) -> str:
    if not value:
        return ""
    parts: list[str] = []
    for fragment, charset in decode_header(value):
        if isinstance(fragment, bytes):
            try:
                parts.append(fragment.decode(charset or "utf-8", errors="replace"))
            except LookupError:
                parts.append(fragment.decode("utf-8", errors="replace"))
        else:
            parts.append(fragment)
    return "".join(parts).strip()


def parse_received_at(value: str | None) -> str:
    if not value:
        return datetime.now(timezone.utc).isoformat()
    try:
        parsed = parsedate_to_datetime(value)
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc).isoformat()
    except (TypeError, ValueError, OverflowError):
        return value


def parse_date(value: str | None) -> date | None:
    if not value:
        return None
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise ValueError(f"日期必须是 YYYY-MM-DD：{value}") from exc


def imap_date(value: date | None) -> str | None:
    return value.strftime("%d-%b-%Y") if value else None


def gmail_query_date(value: date | None) -> str | None:
    return value.isoformat() if value else None


def urlsafe_b64decode(value: str) -> bytes:
    padding = "=" * (-len(value) % 4)
    return base64.urlsafe_b64decode(value + padding)


def strip_html(value: str) -> str:
    value = re.sub(r"<script[\s\S]*?</script>", " ", value, flags=re.IGNORECASE)
    value = re.sub(r"<style[\s\S]*?</style>", " ", value, flags=re.IGNORECASE)
    value = re.sub(r"<[^>]+>", " ", value)
    return html.unescape(re.sub(r"\s+", " ", value)).strip()


def text_from_mime_part(part: Message) -> str:
    payload = part.get_payload(decode=True)
    if payload is None:
        raw = part.get_payload()
        return raw if isinstance(raw, str) else ""
    charset = part.get_content_charset() or "utf-8"
    try:
        text = payload.decode(charset, errors="replace")
    except LookupError:
        text = payload.decode("utf-8", errors="replace")
    return strip_html(text) if part.get_content_type() == "text/html" else text


def safe_component(value: str, fallback: str = "item") -> str:
    value = value.replace("\\", "_").replace("/", "_")
    value = re.sub(r"[^0-9A-Za-z一-龥._-]+", "_", value).strip("._")
    return value[:120] or fallback


def normalized_text(*values: str) -> str:
    return " ".join(v for v in values if v).lower()


def classify_attachment(
    filename: str,
    mime_type: str,
    subject: str,
    sender: str,
    body: str,
    include_all: bool = False,
) -> tuple[bool, int, list[str]]:
    """Return (is_candidate, score, reasons) without inspecting file contents."""

    suffix = Path(filename).suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        return False, 0, ["unsupported-extension"]
    if include_all:
        return True, 1, ["include-all-attachments"]

    name_and_mime = normalized_text(filename, mime_type)
    message_text = normalized_text(subject, sender, body)
    score = 1  # supported document type
    reasons = ["supported-document"]
    strong_hits = [word for word in STRONG_KEYWORDS if word in name_and_mime or word in message_text]
    travel_hits = [word for word in TRAVEL_KEYWORDS if word in name_and_mime or word in message_text]
    if strong_hits:
        score += 3
        reasons.append("strong:" + ",".join(sorted(set(strong_hits))))
    if travel_hits:
        score += 1
        reasons.append("travel:" + ",".join(sorted(set(travel_hits))))
    # A filename that explicitly says invoice/receipt is sufficient even when
    # the sender's message is empty (a common case with automated mail).
    filename_strong = [word for word in STRONG_KEYWORDS if word in name_and_mime]
    candidate = score >= 4 or bool(filename_strong)
    return candidate, score, reasons


@dataclass
class AttachmentRef:
    attachment_id: str
    filename: str
    mime_type: str
    size: int | None
    loader: Callable[[], bytes]


@dataclass
class MailMessage:
    message_id: str
    provider: str
    account: str
    folder: str
    sender: str
    subject: str
    received_at: str
    body: str
    attachments: list[AttachmentRef]


class ImapProvider:
    name = "imap"

    def __init__(self, config: dict[str, Any], folder_override: str | None = None):
        section = config.get("imap") or {}
        self.host = str(section.get("host") or "").strip()
        if not self.host:
            raise ValueError("mail-config.yaml 缺少 imap.host")
        self.port = int(section.get("port") or 993)
        self.use_ssl = bool(section.get("ssl", True))
        self.username = str(section.get("username") or config.get("account") or "").strip()
        if not self.username:
            raise ValueError("mail-config.yaml 缺少 imap.username")
        self.password = read_secret(section, "password", "IMAP_PASSWORD")
        self.folder = folder_override or str(section.get("folder") or "INBOX")
        self.connection: imaplib.IMAP4 | imaplib.IMAP4_SSL | None = None

    def connect(self) -> None:
        if self.use_ssl:
            self.connection = imaplib.IMAP4_SSL(self.host, self.port)
        else:
            self.connection = imaplib.IMAP4(self.host, self.port)
        try:
            self.connection.login(self.username, self.password)
        except imaplib.IMAP4.error as exc:
            message = str(exc)
            if "DOMAINNOTEXIST" in message.upper():
                raise RuntimeError(
                    "网易邮箱服务器拒绝了登录：邮箱地址中的域名不存在。"
                    "请把 mail-config.yaml 中的 account 和 imap.username 改成真实的完整企业邮箱地址，"
                    "并确认 imap.host 是企业邮箱后台提供的服务器。当前错误不是密码错误。"
                ) from exc
            if "AUTH" in message.upper() or "LOGIN" in message.upper():
                raise RuntimeError(
                    "网易邮箱登录失败。请确认邮箱地址、密码/客户端专用密码，以及企业邮箱是否允许 IMAP 登录。"
                ) from exc
            raise
        status, _ = self.connection.select(self.folder, readonly=True)
        if status != "OK":
            raise RuntimeError(f"无法以只读方式打开邮箱文件夹：{self.folder}")

    def close(self) -> None:
        if self.connection is None:
            return
        try:
            self.connection.close()
        except imaplib.IMAP4.error:
            pass
        try:
            self.connection.logout()
        except imaplib.IMAP4.error:
            pass
        self.connection = None

    def list_messages(
        self, since: date | None, until: date | None, limit: int | None
    ) -> Iterable[MailMessage]:
        if self.connection is None:
            self.connect()
        assert self.connection is not None
        criteria: list[str] = []
        if since:
            criteria.extend(["SINCE", imap_date(since) or ""])
        if until:
            criteria.extend(["BEFORE", imap_date(until + timedelta(days=1)) or ""])
        if not criteria:
            criteria = ["ALL"]
        status, data = self.connection.uid("search", None, *criteria)
        if status != "OK":
            raise RuntimeError("IMAP 搜索邮件失败")
        uids = (data[0] or b"").split()
        if limit:
            uids = uids[-limit:]
        for uid in uids:
            yield self._fetch_message(uid)

    def _fetch_message(self, uid: bytes) -> MailMessage:
        assert self.connection is not None
        status, data = self.connection.uid("fetch", uid, "(RFC822)")
        if status != "OK":
            raise RuntimeError(f"IMAP 获取邮件失败：{uid!r}")
        raw = next((item[1] for item in data if isinstance(item, tuple) and isinstance(item[1], bytes)), None)
        if raw is None:
            raise RuntimeError(f"IMAP 邮件没有 RFC822 内容：{uid!r}")
        message = email.message_from_bytes(raw, policy=policy.default)
        sender = decode_header_value(message.get("From"))
        subject = decode_header_value(message.get("Subject"))
        received_at = parse_received_at(message.get("Date"))
        body_parts: list[str] = []
        attachments: list[AttachmentRef] = []
        part_index = 0
        for part in message.walk():
            if part.is_multipart():
                continue
            filename = decode_header_value(part.get_filename())
            disposition = part.get_content_disposition()
            if filename or disposition == "attachment":
                part_index += 1
                filename = filename or f"attachment-{part_index}"
                payload = part.get_payload(decode=True) or b""
                attachments.append(
                    AttachmentRef(
                        attachment_id=f"part-{part_index}",
                        filename=filename,
                        mime_type=part.get_content_type(),
                        size=len(payload),
                        loader=lambda payload=payload: payload,
                    )
                )
            elif part.get_content_type() in {"text/plain", "text/html"}:
                body_parts.append(text_from_mime_part(part))
        return MailMessage(
            message_id=f"imap:{uid.decode(errors='replace')}",
            provider=self.name,
            account=self.username,
            folder=self.folder,
            sender=sender,
            subject=subject,
            received_at=received_at,
            body="\n".join(body_parts),
            attachments=attachments,
        )


class GmailProvider:
    name = "gmail"
    SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

    def __init__(self, config: dict[str, Any], folder_override: str | None = None, query_override: str | None = None):
        section = config.get("gmail") or {}
        credentials_value = section.get("credentials_json")
        if not credentials_value:
            raise ValueError("mail-config.yaml 缺少 gmail.credentials_json")
        self.credentials_path = expand_path(credentials_value)
        self.token_path = expand_path(section.get("token_json") or default_private_dir() / "gmail-token.json")
        self.user_id = str(section.get("user_id") or "me")
        self.account = str(section.get("account") or config.get("account") or self.user_id)
        self.folder = folder_override or str(section.get("folder") or "")
        self.query_override = query_override or section.get("query")
        self.service = self._build_service()

    def _build_service(self):
        try:
            from google.auth.transport.requests import Request
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import InstalledAppFlow
            from googleapiclient.discovery import build
        except ImportError as exc:
            raise RuntimeError(
                "Gmail API 依赖未安装。运行：python -m pip install "
                "google-api-python-client google-auth-httplib2 google-auth-oauthlib"
            ) from exc
        if not self.credentials_path.exists():
            raise FileNotFoundError(f"找不到 Gmail OAuth 客户端文件：{self.credentials_path}")
        credentials = None
        if self.token_path.exists():
            credentials = Credentials.from_authorized_user_file(str(self.token_path), self.SCOPES)
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        if not credentials or not credentials.valid:
            self.token_path.parent.mkdir(parents=True, exist_ok=True)
            flow = InstalledAppFlow.from_client_secrets_file(str(self.credentials_path), self.SCOPES)
            credentials = flow.run_local_server(port=0)
            self.token_path.write_text(credentials.to_json(), encoding="utf-8")
        return build("gmail", "v1", credentials=credentials, cache_discovery=False)

    def list_messages(
        self, since: date | None, until: date | None, limit: int | None
    ) -> Iterable[MailMessage]:
        query = str(self.query_override or "has:attachment")
        if since:
            query += f" after:{gmail_query_date(since)}"
        if until:
            query += f" before:{gmail_query_date(until + timedelta(days=1))}"
        if self.folder:
            folder_query = self.folder if self.folder.startswith("in:") else f"in:{self.folder}"
            query += f" {folder_query}"
        page_token: str | None = None
        count = 0
        while True:
            request = self.service.users().messages().list(
                userId=self.user_id, q=query, pageToken=page_token, maxResults=100
            )
            response = request.execute()
            for item in response.get("messages", []):
                message_id = item["id"]
                full = self.service.users().messages().get(
                    userId=self.user_id, id=message_id, format="full"
                ).execute()
                yield self._convert_message(full)
                count += 1
                if limit and count >= limit:
                    return
            page_token = response.get("nextPageToken")
            if not page_token:
                return

    def _convert_message(self, raw: dict[str, Any]) -> MailMessage:
        headers = {h.get("name", "").lower(): h.get("value", "") for h in raw.get("payload", {}).get("headers", [])}
        sender = headers.get("from", "")
        subject = decode_header_value(headers.get("subject"))
        received_at = parse_received_at(headers.get("date"))
        if raw.get("internalDate"):
            try:
                received_at = datetime.fromtimestamp(
                    int(raw["internalDate"]) / 1000, tz=timezone.utc
                ).isoformat()
            except (TypeError, ValueError, OSError):
                pass
        body_parts: list[str] = []
        attachments: list[AttachmentRef] = []
        self._walk_payload(raw.get("payload", {}), raw["id"], body_parts, attachments)
        return MailMessage(
            message_id=f"gmail:{raw['id']}",
            provider=self.name,
            account=self.account,
            folder=self.folder or "mailbox",
            sender=sender,
            subject=subject,
            received_at=received_at,
            body="\n".join(body_parts),
            attachments=attachments,
        )

    def _walk_payload(
        self,
        payload: dict[str, Any],
        message_id: str,
        body_parts: list[str],
        attachments: list[AttachmentRef],
    ) -> None:
        mime_type = payload.get("mimeType") or "application/octet-stream"
        body = payload.get("body") or {}
        filename = decode_header_value(payload.get("filename"))
        if filename or body.get("attachmentId"):
            attachment_id = str(body.get("attachmentId") or f"inline-{len(attachments) + 1}")
            if not filename:
                filename = f"attachment-{len(attachments) + 1}"
            inline_data = body.get("data")
            size = body.get("size")

            def load(
                inline_data: str | None = inline_data,
                message_id: str = message_id,
                attachment_id: str = attachment_id,
            ) -> bytes:
                if inline_data:
                    return urlsafe_b64decode(inline_data)
                response = self.service.users().messages().attachments().get(
                    userId=self.user_id, messageId=message_id, id=attachment_id
                ).execute()
                return urlsafe_b64decode(response.get("data", ""))

            attachments.append(
                AttachmentRef(
                    attachment_id=attachment_id,
                    filename=filename,
                    mime_type=mime_type,
                    size=int(size) if size is not None else None,
                    loader=load,
                )
            )
        elif mime_type in {"text/plain", "text/html"} and body.get("data"):
            try:
                text = urlsafe_b64decode(body["data"]).decode("utf-8", errors="replace")
                body_parts.append(strip_html(text) if mime_type == "text/html" else text)
            except (ValueError, UnicodeError):
                pass
        for child in payload.get("parts", []) or []:
            self._walk_payload(child, message_id, body_parts, attachments)


def read_secret(section: dict[str, Any], key: str, default_env: str) -> str:
    env_name = str(section.get(f"{key}_env") or default_env)
    value = os.environ.get(env_name)
    if value:
        return value
    password_file = section.get(f"{key}_file")
    if password_file:
        path = expand_path(password_file)
        if path.exists():
            return path.read_text(encoding="utf-8").strip()
    plain = section.get(key)
    if plain:
        return str(plain)
    if sys.stdin.isatty():
        value = getpass.getpass(f"请输入邮箱密码/应用专用密码（{env_name}，不会显示）：")
        if value:
            return value
    raise ValueError(
        f"缺少邮箱密码。优先在交互式终端输入，或设置环境变量 {env_name}，"
        f"或在本地配置中使用 {key}_file；不要把密码提交到 GitHub。"
    )


def load_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(
            f"找不到邮箱配置：{path}\n请参考 profiles/mail-config.example.yaml 创建本地配置。"
        )
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    if not isinstance(data, dict):
        raise ValueError("邮箱配置必须是 YAML 对象")
    return data


def is_placeholder(value: Any) -> bool:
    text = str(value or "").strip().lower()
    return not text or "example" in text or "your-" in text or text in {"name@company.example", "name@company.com"}


def configure_imap_settings(config: dict[str, Any], args: argparse.Namespace, config_path: Path) -> None:
    """Collect non-secret IMAP settings on first use; never save a password."""
    provider_name = str(args.provider or config.get("provider") or "").lower()
    if provider_name not in {"imap", "netease", "netease-imap", "gmail-imap"}:
        return
    section = config.setdefault("imap", {})
    account = str(config.get("account") or section.get("username") or "").strip()
    host = str(section.get("host") or "").strip()
    needs_setup = bool(args.setup or args.email or args.imap_host or is_placeholder(account) or is_placeholder(host))
    if not needs_setup:
        return
    if not sys.stdin.isatty() and (is_placeholder(account) or is_placeholder(host)) and not (args.email and args.imap_host):
        raise ValueError(
            "邮箱配置仍是示例值。请在交互式终端运行，或传入 --email 和 --imap-host。"
        )
    if args.email:
        account = args.email.strip()
    elif is_placeholder(account):
        account = input("请输入完整企业邮箱地址（例如 name@your-company.com）：").strip()
    if "@" not in account or account.endswith("@"):
        raise ValueError("邮箱地址格式不完整，请输入完整的企业邮箱地址")
    if args.imap_host:
        host = args.imap_host.strip()
    elif args.setup or is_placeholder(host):
        default_host = host if not is_placeholder(host) else "imap.qiye.163.com"
        entered = input(f"请输入 IMAP 服务器地址（直接回车使用 {default_host}）：").strip()
        host = entered or default_host
    section["username"] = account
    section["host"] = host
    config["account"] = account
    if args.save_mail_settings:
        config_path.write_text(
            yaml.safe_dump(config, allow_unicode=True, sort_keys=False), encoding="utf-8"
        )
        print(f"已保存邮箱地址和 IMAP 服务器到：{config_path}", file=sys.stderr)


def load_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"sources": {}, "hashes": {}}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"无法读取采集状态文件：{path}") from exc
    if not isinstance(data, dict):
        raise ValueError(f"采集状态文件格式错误：{path}")
    data.setdefault("sources", {})
    data.setdefault("hashes", {})
    return data


def atomic_write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        mode="w", encoding="utf-8", dir=path.parent, prefix=f".{path.name}.", delete=False
    ) as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
        temporary = Path(handle.name)
    temporary.replace(path)


def make_output_name(message: MailMessage, attachment: AttachmentRef) -> str:
    stamp = safe_component(message.received_at[:10].replace("-", ""), "date")
    message_part = safe_component(message.message_id.replace(":", "_"), "message")
    attachment_part = safe_component(attachment.attachment_id, "attachment")
    filename = safe_component(attachment.filename, "invoice")
    return f"{stamp}_{message_part}_{attachment_part}_{filename}"


def source_key(message: MailMessage, attachment: AttachmentRef) -> str:
    return f"{message.provider}:{message.message_id}:{attachment.attachment_id}"


def source_record(
    message: MailMessage,
    attachment: AttachmentRef,
    status: str,
    score: int,
    reasons: list[str],
    **extra: Any,
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "source_key": source_key(message, attachment),
        "provider": message.provider,
        "account": message.account,
        "folder": message.folder,
        "message_id": message.message_id,
        "attachment_id": attachment.attachment_id,
        "received_at": message.received_at,
        "sender": message.sender,
        "subject": message.subject,
        "attachment_name": attachment.filename,
        "mime_type": attachment.mime_type,
        "size": attachment.size,
        "invoice_score": score,
        "reasons": reasons,
        "status": status,
    }
    result.update(extra)
    return result


def build_provider(
    config: dict[str, Any], provider_override: str | None, folder_override: str | None, query_override: str | None
):
    provider_name = str(provider_override or config.get("provider") or "").lower()
    if provider_name == "gmail":
        return GmailProvider(config, folder_override, query_override)
    if provider_name in {"imap", "netease", "netease-imap", "gmail-imap"}:
        return ImapProvider(config, folder_override)
    raise ValueError("provider 必须是 gmail 或 imap；网易企业邮箱请使用 imap 并填写服务器地址")


def run(args: argparse.Namespace) -> int:
    config_path = expand_path(args.config)
    config = load_config(config_path)
    configure_imap_settings(config, args, config_path)
    provider = build_provider(config, args.provider, args.folder, args.query)
    since = parse_date(args.since)
    until = parse_date(args.until)
    if since and until and until < since:
        raise ValueError("--until 不能早于 --since")

    output_dir = expand_path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)
    state_path = expand_path(args.state) if args.state else output_dir / ".mail-ingest-state.json"
    state = {"sources": {}, "hashes": {}} if args.reset_state else load_state(state_path)
    sources: dict[str, Any] = state.setdefault("sources", {})
    hashes: dict[str, Any] = state.setdefault("hashes", {})
    downloaded: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []
    warnings: list[str] = []

    try:
        for message in provider.list_messages(since, until, args.limit):
            mail_text = normalized_text(message.subject, message.sender, message.body)
            for attachment in message.attachments:
                candidate, score, reasons = classify_attachment(
                    attachment.filename,
                    attachment.mime_type,
                    message.subject,
                    message.sender,
                    mail_text,
                    include_all=args.include_all_attachments,
                )
                if not candidate:
                    if args.verbose:
                        print(f"跳过非发票附件：{message.subject} / {attachment.filename}", file=sys.stderr)
                    continue
                key = source_key(message, attachment)
                if key in sources:
                    skipped.append(source_record(message, attachment, "duplicate-source", score, reasons))
                    continue
                if args.dry_run:
                    record = source_record(message, attachment, "dry-run", score, reasons)
                    downloaded.append(record)
                    continue
                payload = attachment.loader()
                digest = hashlib.sha256(payload).hexdigest()
                duplicate_key = hashes.get(digest)
                if duplicate_key:
                    record = source_record(
                        message,
                        attachment,
                        "duplicate-content",
                        score,
                        reasons,
                        sha256=digest,
                        duplicate_of=duplicate_key,
                    )
                    sources[key] = record
                    skipped.append(record)
                    continue
                filename = make_output_name(message, attachment)
                destination = output_dir / filename
                destination.write_bytes(payload)
                record = source_record(
                    message,
                    attachment,
                    "downloaded",
                    score,
                    reasons,
                    sha256=digest,
                    saved_path=str(destination),
                )
                sources[key] = record
                hashes[digest] = key
                downloaded.append(record)
    finally:
        close = getattr(provider, "close", None)
        if close:
            close()

    if not args.dry_run:
        state["updated_at"] = datetime.now(timezone.utc).isoformat()
        atomic_write_json(state_path, state)
        atomic_write_json(output_dir / "mail-index.json", list(sources.values()))

    result = {
        "provider": getattr(provider, "name", "unknown"),
        "config": str(config_path),
        "output": str(output_dir),
        "downloaded": downloaded,
        "skipped": skipped,
        "warnings": warnings,
        "dry_run": args.dry_run,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="只读扫描邮箱并下载疑似发票附件")
    parser.add_argument("--config", type=Path, default=default_config_path(), help="本地邮箱 YAML 配置")
    parser.add_argument("--provider", choices=["gmail", "imap", "netease", "netease-imap", "gmail-imap"])
    parser.add_argument("--email", help="覆盖 IMAP 登录用的完整企业邮箱地址")
    parser.add_argument("--imap-host", help="覆盖 IMAP 服务器地址")
    parser.add_argument("--setup", action="store_true", help="交互式询问邮箱地址和 IMAP 服务器")
    parser.add_argument("--save-mail-settings", action="store_true", help="保存邮箱地址和 IMAP 服务器；不保存密码")
    parser.add_argument("--folder", help="覆盖配置中的文件夹或 Gmail 标签")
    parser.add_argument("--query", help="覆盖 Gmail 搜索语句")
    parser.add_argument("--since", help="起始日期 YYYY-MM-DD")
    parser.add_argument("--until", help="结束日期 YYYY-MM-DD（含当天）")
    parser.add_argument("--limit", type=int, help="最多读取多少封邮件")
    parser.add_argument("--output", type=Path, default=Path(".mail-inbox"), help="附件输出目录")
    parser.add_argument("--state", type=Path, help="覆盖去重状态 JSON 路径")
    parser.add_argument("--reset-state", action="store_true", help="重置输出目录中的去重状态")
    parser.add_argument("--include-all-attachments", action="store_true", help="下载所有支持格式的附件，不做发票关键词过滤")
    parser.add_argument("--dry-run", action="store_true", help="只列出候选，不下载和写入状态")
    parser.add_argument("--verbose", action="store_true", help="显示被过滤的附件")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        return run(args)
    except KeyboardInterrupt:
        print("已中止。", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
