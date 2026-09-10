from __future__ import annotations

import csv
import re
from html import escape
from pathlib import Path

import pythoncom
import win32com.client

from .models import BulkEmailRequest, EmailRequest, TableBlock


OL_MAIL_ITEM = 0
EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")


def _table_to_html(table: TableBlock) -> str:
    title = f"<h3>{escape(table.title)}</h3>" if table.title else ""
    header = "".join(f"<th>{escape(str(col))}</th>" for col in table.columns)
    rows = []
    for row in table.rows:
        cells = "".join(f"<td>{escape(str(cell))}</td>" for cell in row)
        rows.append(f"<tr>{cells}</tr>")

    return (
        f"{title}"
        "<table style='border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:10.5pt'>"
        f"<thead><tr>{header}</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody>"
        "</table>"
    )


def _build_html_body(body: str, tables: list[TableBlock]) -> str:
    safe_body = escape(body).replace("\n", "<br>")
    rendered_tables = "<br>".join(_table_to_html(table) for table in tables)
    return (
        "<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:10.5pt'>"
        f"<div>{safe_body}</div>"
        f"{'<br>' + rendered_tables if rendered_tables else ''}"
        "</body></html>"
    )


def _validate_attachments(paths: list[str]) -> list[Path]:
    resolved: list[Path] = []
    for raw_path in paths:
        path = Path(raw_path).expanduser().resolve()
        if not path.is_file():
            raise FileNotFoundError(f"Attachment not found: {path}")
        resolved.append(path)
    return resolved


def _normalize_addresses(addresses: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()

    for value in addresses:
        for address in re.split(r"[;,\s]+", value.strip()):
            if not address:
                continue
            normalized = address.lower()
            if not EMAIL_RE.match(address):
                raise ValueError(f"Invalid email address: {address}")
            if normalized not in seen:
                seen.add(normalized)
                result.append(address)

    return result


def _load_recipients_from_file(file_path: str, column: str = "email") -> list[str]:
    path = Path(file_path).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Recipient file not found: {path}")

    suffix = path.suffix.lower()

    if suffix == ".txt":
        content = path.read_text(encoding="utf-8-sig")
        values = [line.strip() for line in content.splitlines() if line.strip()]
        return _normalize_addresses(values)

    if suffix == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            sample = handle.read(4096)
            handle.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            except csv.Error:
                dialect = csv.excel

            reader = csv.DictReader(handle, dialect=dialect)
            if not reader.fieldnames:
                raise ValueError("CSV file has no header row")

            field_map = {name.strip().lower(): name for name in reader.fieldnames if name}
            requested = column.strip().lower()
            actual_column = field_map.get(requested)
            if actual_column is None:
                raise ValueError(
                    f"Column '{column}' not found in CSV. Available columns: {', '.join(reader.fieldnames)}"
                )

            values = [str(row.get(actual_column, "")).strip() for row in reader]
            return _normalize_addresses([value for value in values if value])

    raise ValueError("Recipient file must be .txt or .csv")


def _merge_recipients(addresses: list[str], file_path: str | None, column: str) -> list[str]:
    combined = list(addresses)
    if file_path:
        combined.extend(_load_recipients_from_file(file_path, column))
    return _normalize_addresses(combined)


def _new_mail(subject: str, body: str, tables: list[TableBlock], attachments: list[str]):
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(OL_MAIL_ITEM)
    mail.Subject = subject
    mail.HTMLBody = _build_html_body(body, tables)

    for path in _validate_attachments(attachments):
        mail.Attachments.Add(str(path))

    return mail


def _create_mail(request: EmailRequest):
    recipients = _merge_recipients(
        request.to,
        request.recipient_file,
        request.recipient_file_column,
    )
    if not recipients:
        raise ValueError("At least one recipient is required")

    mail = _new_mail(request.subject, request.body, request.tables, request.attachments)
    mail.To = "; ".join(recipients)
    mail.CC = "; ".join(_normalize_addresses(request.cc))
    mail.BCC = "; ".join(_normalize_addresses(request.bcc))
    return mail


def create_draft(request: EmailRequest) -> dict:
    mail = _create_mail(request)
    mail.Save()
    entry_id = getattr(mail, "EntryID", None)
    return {"status": "draft_created", "entry_id": entry_id}


def send_email(request: EmailRequest) -> dict:
    mail = _create_mail(request)
    mail.Send()
    return {"status": "sent"}


def send_bulk_email(request: BulkEmailRequest) -> dict:
    recipients = _merge_recipients(
        request.recipients,
        request.recipient_file,
        request.recipient_file_column,
    )
    if not recipients:
        raise ValueError("At least one recipient is required")

    sent = 0
    failed: list[dict[str, str]] = []

    for recipient in recipients:
        try:
            mail = _new_mail(request.subject, request.body, request.tables, request.attachments)
            mail.To = recipient
            mail.Send()
            sent += 1
        except Exception as exc:  # COM errors need to be returned per recipient
            failed.append({"recipient": recipient, "error": str(exc)})

    return {
        "status": "completed" if not failed else "completed_with_errors",
        "total": len(recipients),
        "sent": sent,
        "failed": failed,
    }


def get_outlook_status() -> dict:
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    accounts = [account.SmtpAddress for account in namespace.Accounts]
    return {"status": "ok", "accounts": accounts}
