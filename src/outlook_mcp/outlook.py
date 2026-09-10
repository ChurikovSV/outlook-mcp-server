from __future__ import annotations

from html import escape
from pathlib import Path

import pythoncom
import win32com.client

from .models import EmailRequest, TableBlock


OL_MAIL_ITEM = 0


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


def _build_html_body(request: EmailRequest) -> str:
    body = escape(request.body).replace("\n", "<br>")
    tables = "<br>".join(_table_to_html(table) for table in request.tables)
    return (
        "<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:10.5pt'>"
        f"<div>{body}</div>"
        f"{'<br>' + tables if tables else ''}"
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


def _create_mail(request: EmailRequest):
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(OL_MAIL_ITEM)

    mail.To = "; ".join(request.to)
    mail.CC = "; ".join(request.cc)
    mail.BCC = "; ".join(request.bcc)
    mail.Subject = request.subject
    mail.HTMLBody = _build_html_body(request)

    for path in _validate_attachments(request.attachments):
        mail.Attachments.Add(str(path))

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


def get_outlook_status() -> dict:
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    accounts = [account.SmtpAddress for account in namespace.Accounts]
    return {"status": "ok", "accounts": accounts}
