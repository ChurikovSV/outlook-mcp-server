from __future__ import annotations

import re
from html import escape


BODY_STYLE = "font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#222222;line-height:1.4"
H1_STYLE = "font-family:Calibri,Arial,sans-serif;font-size:18pt;font-weight:bold;color:#1f4e79;margin:0 0 12px 0"
H2_STYLE = "font-family:Calibri,Arial,sans-serif;font-size:14pt;font-weight:bold;color:#2f5597;margin:18px 0 8px 0"
H3_STYLE = "font-family:Calibri,Arial,sans-serif;font-size:12pt;font-weight:bold;color:#333333;margin:14px 0 6px 0"
P_STYLE = "font-family:Calibri,Arial,sans-serif;font-size:11pt;margin:8px 0;line-height:1.4"
LIST_STYLE = "font-family:Calibri,Arial,sans-serif;font-size:11pt;margin:8px 0 8px 24px;padding:0"
TABLE_STYLE = "border-collapse:collapse;width:100%;font-family:Calibri,Arial,sans-serif;font-size:10.5pt;margin:12px 0"
TH_STYLE = "background:#1f4e79;color:#ffffff;font-weight:bold;text-align:left;border:1px solid #c9c9c9;padding:7px"
TD_STYLE = "border:1px solid #d9d9d9;padding:7px;vertical-align:top"


_LINK_RE = re.compile(r"\[([^\]]+)\]\((https?://[^\s)]+)\)")
_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
_ITALIC_RE = re.compile(r"(?<!\*)\*([^*\n]+?)\*(?!\*)")
_ORDERED_RE = re.compile(r"^\s*\d+[.)]\s+(.+)$")
_TABLE_SEPARATOR_RE = re.compile(r"^\s*\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)+\|?\s*$")


def _inline(text: str) -> str:
    """Render a conservative subset of inline Markdown suitable for Outlook HTML."""
    safe = escape(text, quote=True)
    safe = _LINK_RE.sub(
        lambda match: (
            f'<a href="{match.group(2)}" style="color:#0563C1;text-decoration:underline">'
            f'{match.group(1)}</a>'
        ),
        safe,
    )
    safe = _BOLD_RE.sub(r"<strong>\1</strong>", safe)
    safe = _ITALIC_RE.sub(r"<em>\1</em>", safe)
    return safe


def _split_table_row(line: str) -> list[str]:
    value = line.strip()
    if value.startswith("|"):
        value = value[1:]
    if value.endswith("|"):
        value = value[:-1]
    return [cell.strip() for cell in value.split("|")]


def _is_table_start(lines: list[str], index: int) -> bool:
    if index + 1 >= len(lines):
        return False
    return "|" in lines[index] and bool(_TABLE_SEPARATOR_RE.match(lines[index + 1]))


def _render_table(lines: list[str], index: int) -> tuple[str, int]:
    headers = _split_table_row(lines[index])
    index += 2  # skip the separator row
    rows: list[list[str]] = []

    while index < len(lines):
        line = lines[index]
        if not line.strip() or "|" not in line:
            break
        row = _split_table_row(line)
        if len(row) < len(headers):
            row.extend([""] * (len(headers) - len(row)))
        rows.append(row[: len(headers)])
        index += 1

    header_html = "".join(f'<th style="{TH_STYLE}">{_inline(cell)}</th>' for cell in headers)
    body_rows = []
    for row in rows:
        cells = "".join(f'<td style="{TD_STYLE}">{_inline(cell)}</td>' for cell in row)
        body_rows.append(f"<tr>{cells}</tr>")

    table_html = (
        f'<table role="presentation" cellpadding="0" cellspacing="0" style="{TABLE_STYLE}">'
        f"<thead><tr>{header_html}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table>"
    )
    return table_html, index


def markdown_to_html(markdown: str) -> str:
    """Convert practical business-email Markdown to Outlook-friendly HTML.

    Supported: H1-H3, paragraphs, bold, italic, links, unordered and ordered lists,
    horizontal rules, and pipe tables. HTML from the input is escaped deliberately.
    """
    lines = markdown.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    html: list[str] = [f'<html><body style="{BODY_STYLE}">']
    list_type: str | None = None

    def close_list() -> None:
        nonlocal list_type
        if list_type:
            html.append(f"</{list_type}>")
            list_type = None

    def open_list(kind: str) -> None:
        nonlocal list_type
        if list_type == kind:
            return
        close_list()
        html.append(f'<{kind} style="{LIST_STYLE}">')
        list_type = kind

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if _is_table_start(lines, i):
            close_list()
            table_html, i = _render_table(lines, i)
            html.append(table_html)
            continue

        if stripped.startswith("### "):
            close_list()
            html.append(f'<h3 style="{H3_STYLE}">{_inline(stripped[4:])}</h3>')
        elif stripped.startswith("## "):
            close_list()
            html.append(f'<h2 style="{H2_STYLE}">{_inline(stripped[3:])}</h2>')
        elif stripped.startswith("# "):
            close_list()
            html.append(f'<h1 style="{H1_STYLE}">{_inline(stripped[2:])}</h1>')
        elif stripped in {"---", "***", "___"}:
            close_list()
            html.append('<hr style="border:0;border-top:1px solid #d9d9d9;margin:14px 0">')
        elif stripped.startswith(("- ", "* ", "+ ")):
            open_list("ul")
            html.append(f"<li>{_inline(stripped[2:])}</li>")
        else:
            ordered_match = _ORDERED_RE.match(line)
            if ordered_match:
                open_list("ol")
                html.append(f"<li>{_inline(ordered_match.group(1))}</li>")
            elif not stripped:
                close_list()
            else:
                close_list()
                html.append(f'<p style="{P_STYLE}">{_inline(stripped)}</p>')

        i += 1

    close_list()
    html.append("</body></html>")
    return "".join(html)
