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
_TEMPLATE_NAME_RE = re.compile(r"^[A-Za-z0-9А-Яа-яЁё_-]+$")
TEMPLATES_DIR = Path(__file__).resolve().parents[2] / "templates"


def _read_markdown_file(file_path: str) -> str:
    path = Path(file_path).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Markdown file not found: {path}")
    if path.suffix.lower() not in {".md", ".markdown"}:
        raise ValueError("Markdown file must have .md or .markdown extension")
    return path.read_text(encoding="utf-8-sig")


def _read_named_template(template_name: str) -> tuple[str, Path]:
    name = template_name.strip()
    if not name:
        raise ValueError("template_name cannot be empty")
    if name.lower().endswith(".md"):
        name = name[:-3]
    elif name.lower().endswith(".markdown"):
        name = name[:-9]
    if not _TEMPLATE_NAME_RE.fullmatch(name):
        raise ValueError("template_name may contain only letters, digits, hyphens and underscores")

    path = (TEMPLATES_DIR / f"{name}.md").resolve()
    templates_root = TEMPLATES_DIR.resolve()
    if path.parent != templates_root:
        raise ValueError("Invalid template_name")
    if not path.is_file():
        available = sorted(p.stem for p in TEMPLATES_DIR.glob("*.md")) if TEMPLATES_DIR.is_dir() else []
        suffix = f" Available templates: {', '.join(available)}" if available else " No templates are installed."
        raise FileNotFoundError(f"Markdown template not found: {name}.{suffix}")
    return path.read_text(encoding="utf-8-sig"), path


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
    template_name: str | None = None,
    variables: dict[str, str] | None = None,
) -> dict:
    """Create an Outlook draft from inline Markdown, a local Markdown file, or a named local template. Never sends it."""
    sources_selected = sum(bool(value) for value in (markdown, markdown_file, template_name))
    if sources_selected != 1:
        raise ValueError("Provide exactly one of markdown, markdown_file or template_name")

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

    source_label: str
    template_path: str | None = None
    if markdown is not None:
        markdown_text = markdown
        source_label = "inline"
    elif markdown_file is not None:
        markdown_text = _read_markdown_file(markdown_file)
        source_label = "file"
        template_path = str(Path(markdown_file).expanduser().resolve())
    else:
        markdown_text, resolved_template = _read_named_template(template_name or "")
        source_label = "template"
        template_path = str(resolved_template)

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
        "markdown_source": source_label,
        "template_name": template_name if source_label == "template" else None,
        "template_path": template_path,
        "template_variables_used": variables_used,
        "uploaded_attachments": len(request.uploaded_attachments),
    }
