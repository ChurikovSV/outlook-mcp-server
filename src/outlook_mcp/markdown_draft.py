from __future__ import annotations

from pathlib import Path

from .markdown_email import markdown_to_html
from .models import EmailRequest
from .outlook import (
    _apply_sender,
    _get_outlook,
    _materialize_uploaded_attachments,
    _merge_recipients,
    _normalize_addresses,
    _validate_attachments,
)


OL_MAIL_ITEM = 0


def _read_markdown_file(file_path: str) -> str:
    path = Path(file_path).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Markdown file not found: {path}")
    if path.suffix.lower() not in {".md", ".markdown"}:
        raise ValueError("Markdown file must have .md or .markdown extension")
    return path.read_text(encoding="utf-8-sig")


def create_draft_from_markdown(
    request: EmailRequest,
    markdown: str | None = None,
    markdown_file: str | None = None,
) -> dict:
    """Create an Outlook draft from Markdown without sending it."""
    if bool(markdown) == bool(markdown_file):
        raise ValueError("Provide exactly one of markdown or markdown_file")

    markdown_text = markdown if markdown is not None else _read_markdown_file(markdown_file or "")
    html_body = markdown_to_html(markdown_text)

    recipients = _merge_recipients(
        request.to,
        request.recipient_file,
        request.recipient_file_column,
        request.recipient_file_sheet,
    )
    if not recipients:
        raise ValueError("At least one recipient is required")

    with _materialize_uploaded_attachments(request.uploaded_attachments) as uploaded_paths:
        outlook = _get_outlook()
        mail = outlook.CreateItem(OL_MAIL_ITEM)
        sender = _apply_sender(mail, request.from_email)

        mail.Subject = request.subject
        mail.To = "; ".join(recipients)
        mail.CC = "; ".join(_normalize_addresses(request.cc))
        mail.BCC = "; ".join(_normalize_addresses(request.bcc))
        mail.HTMLBody = html_body

        for path in _validate_attachments(request.attachments):
            mail.Attachments.Add(str(path))
        for path in uploaded_paths:
            mail.Attachments.Add(str(path))

        mail.Save()
        entry_id = getattr(mail, "EntryID", None)

    return {
        "status": "draft_created",
        "entry_id": entry_id,
        "format": "markdown",
        "from_email": sender,
        "sender_mode": "sent_on_behalf_of" if sender else "default_account",
        "markdown_source": "file" if markdown_file else "inline",
        "uploaded_attachments": len(request.uploaded_attachments),
    }
