from __future__ import annotations

import re
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
_TEMPLATE_VAR_RE = re.compile(r"{{\s*([^{}]+?)\s*}}")


def _read_markdown_file(file_path: str) -> str:
    path = Path(file_path).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Markdown file not found: {path}")
    if path.suffix.lower() not in {".md", ".markdown"}:
        raise ValueError("Markdown file must have .md or .markdown extension")
    return path.read_text(encoding="utf-8-sig")


def _render_template(text: str, variables: dict[str, str] | None = None) -> tuple[str, list[str]]:
    values = {str(key).strip(): str(value) for key, value in (variables or {}).items()}
    used: list[str] = []
    missing: set[str] = set()

    def replace(match: re.Match[str]) -> str:
        key = match.group(1).strip()
        if key not in values:
            missing.add(key)
            return match.group(0)
        if key not in used:
            used.append(key)
        return values[key]

    rendered = _TEMPLATE_VAR_RE.sub(replace, text)
    if missing:
        names = ", ".join(sorted(missing))
        raise ValueError(f"Missing template variables: {names}")
    return rendered, used


def create_draft_from_markdown(
    request: EmailRequest,
    markdown: str | None = None,
    markdown_file: str | None = None,
    variables: dict[str, str] | None = None,
) -> dict:
    """Create an Outlook draft from Markdown without sending it."""
    if bool(markdown) == bool(markdown_file):
        raise ValueError("Provide exactly one of markdown or markdown_file")

    recipients = _merge_recipients(
        request.to,
        request.recipient_file,
        request.recipient_file_column,
        request.recipient_file_sheet,
    )
    if not recipients:
        raise ValueError("At least one recipient is required")

    template_values = dict(variables or {})
    if len(recipients) == 1:
        template_values.setdefault("email", recipients[0])

    markdown_text = markdown if markdown is not None else _read_markdown_file(markdown_file or "")
    rendered_markdown, used_body_variables = _render_template(markdown_text, template_values)
    rendered_subject, used_subject_variables = _render_template(request.subject, template_values)
    html_body = markdown_to_html(rendered_markdown)

    variables_used = list(dict.fromkeys(used_subject_variables + used_body_variables))

    with _materialize_uploaded_attachments(request.uploaded_attachments) as uploaded_paths:
        outlook = _get_outlook()
        mail = outlook.CreateItem(OL_MAIL_ITEM)
        sender = _apply_sender(mail, request.from_email)

        mail.Subject = rendered_subject
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
        "template_variables_used": variables_used,
        "uploaded_attachments": len(request.uploaded_attachments),
    }
