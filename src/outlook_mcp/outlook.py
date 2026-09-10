from __future__ import annotations

import csv
import re
from html import escape
from pathlib import Path

import pythoncom
import win32com.client as win32
from openpyxl import load_workbook

from .models import BulkEmailRequest, EmailRequest, TableBlock


OL_MAIL_ITEM = 0
OL_DISCARD = 1
EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")

# Match the COM setup used by the existing working Outlook automation script.
win32.gencache.is_readonly = True


def _get_outlook():
    """Return Outlook.Application using the same simple COM path as the known working script."""
    pythoncom.CoInitialize()
    return win32.Dispatch("Outlook.Application")


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


def _load_recipients_from_xlsx(path: Path, column: str, sheet: str | None) -> list[str]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet:
            if sheet not in workbook.sheetnames:
                raise ValueError(
                    f"Sheet '{sheet}' not found in XLSX. Available sheets: {', '.join(workbook.sheetnames)}"
                )
            worksheet = workbook[sheet]
        else:
            worksheet = workbook[workbook.sheetnames[0]]

        rows = worksheet.iter_rows(values_only=True)
        try:
            header_row = next(rows)
        except StopIteration as exc:
            raise ValueError("XLSX file is empty") from exc

        headers = [str(value).strip() if value is not None else "" for value in header_row]
        header_map = {name.lower(): index for index, name in enumerate(headers) if name}
        requested = column.strip().lower()

        if requested not in header_map:
            available = ", ".join(name for name in headers if name)
            raise ValueError(
                f"Column '{column}' not found in XLSX. Available columns: {available}"
            )

        column_index = header_map[requested]
        values: list[str] = []

        for row in rows:
            if column_index >= len(row):
                continue
            value = row[column_index]
            if value is None:
                continue
            text = str(value).strip()
            if text:
                values.append(text)

        return _normalize_addresses(values)
    finally:
        workbook.close()


def _load_recipients_from_file(
    file_path: str,
    column: str = "email",
    sheet: str | None = None,
) -> list[str]:
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

    if suffix == ".xlsx":
        return _load_recipients_from_xlsx(path, column, sheet)

    raise ValueError("Recipient file must be .txt, .csv or .xlsx")


def _merge_recipients(
    addresses: list[str],
    file_path: str | None,
    column: str,
    sheet: str | None = None,
) -> list[str]:
    combined = list(addresses)
    if file_path:
        combined.extend(_load_recipients_from_file(file_path, column, sheet))
    return _normalize_addresses(combined)


def _new_mail(subject: str, body: str, tables: list[TableBlock], attachments: list[str]):
    outlook = _get_outlook()
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
        request.recipient_file_sheet,
    )
    if not recipients:
        raise ValueError("At least one recipient is required")

    mail = _new_mail(request.subject, request.body, request.tables, request.attachments)
    mail.To = "; ".join(recipients)
    mail.CC = "; ".join(_normalize_addresses(request.cc))
    mail.BCC = "; ".join(_normalize_addresses(request.bcc))
    return mail


def create_draft(request: EmailRequest) -> dict:
    """Create one Outlook draft. No programmatic send is performed."""
    mail = _create_mail(request)
    mail.Save()
    entry_id = getattr(mail, "EntryID", None)
    return {"status": "draft_created", "entry_id": entry_id}


def create_bulk_drafts(request: BulkEmailRequest) -> dict:
    """Create one separate Outlook draft per recipient. No programmatic send is performed."""
    recipients = _merge_recipients(
        request.recipients,
        request.recipient_file,
        request.recipient_file_column,
        request.recipient_file_sheet,
    )
    if not recipients:
        raise ValueError("At least one recipient is required")

    created = 0
    failed: list[dict[str, str]] = []
    drafts: list[dict[str, str | None]] = []

    for recipient in recipients:
        try:
            mail = _new_mail(request.subject, request.body, request.tables, request.attachments)
            mail.To = recipient
            mail.Save()
            entry_id = getattr(mail, "EntryID", None)
            created += 1
            drafts.append({"recipient": recipient, "entry_id": entry_id})
        except Exception as exc:
            failed.append({"recipient": recipient, "error": str(exc)})

    return {
        "status": "completed" if not failed else "completed_with_errors",
        "total": len(recipients),
        "created": created,
        "failed": failed,
        "drafts": drafts,
    }


def get_outlook_status() -> dict:
    """Check the exact COM operation the draft workflow needs: Outlook.Application + CreateItem(0)."""
    outlook = _get_outlook()
    mail = outlook.CreateItem(OL_MAIL_ITEM)

    version = None
    try:
        version = str(outlook.Version)
    except Exception:
        pass

    try:
        mail.Close(OL_DISCARD)
    except Exception:
        pass

    return {
        "status": "ok",
        "create_item": True,
        "outlook_version": version,
        "mode": "draft_only",
    }
