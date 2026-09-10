# outlook-mcp-server

Local MCP server for Microsoft Outlook on Windows using `pywin32` / `win32com`.

The server does not connect directly to Exchange. It uses the locally configured Outlook profile and Outlook COM automation.

## Requirements

- Windows
- Microsoft Outlook Desktop configured with the corporate Exchange account
- Python 3.11+
- Permission to use Outlook COM automation

## Install

```powershell
git clone https://github.com/ChurikovSV/outlook-mcp-server.git
cd outlook-mcp-server
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e .
```

## Run over HTTP (default)

```powershell
python -m outlook_mcp.server
```

Default MCP endpoint:

```text
http://127.0.0.1:8000/mcp
```

A different port can be selected with:

```powershell
python -m outlook_mcp.server --port 8765
```

## Draft-only mode

The server intentionally does not call Outlook `Send()`. Corporate Outlook policies may block programmatic sending, while draft creation remains allowed.

Available workflow tools:

- `get_outlook_status`
- `diagnose_outlook`
- `create_draft`
- `create_bulk_drafts`

`create_bulk_drafts` creates one separate Outlook draft per recipient.

## Create one draft

```json
{
  "to": ["user@company.ru"],
  "subject": "Meeting protocol",
  "body": "Colleagues, sending the meeting protocol."
}
```

## Recipients from TXT / CSV / XLSX

Recipients can be supplied directly and/or loaded from a local file.

Excel example:

```json
{
  "to": [],
  "recipient_file": "C:\\Work\\mail\\users.xlsx",
  "recipient_file_sheet": "Получатели",
  "recipient_file_column": "Почта",
  "subject": "Meeting protocol",
  "body": "Colleagues, sending the meeting protocol."
}
```

Supported recipient files:

- `.txt`
- `.csv`
- `.xlsx`

## Create separate drafts for a distribution list

```json
{
  "recipients": [
    "user1@company.ru",
    "user2@company.ru"
  ],
  "subject": "Notification",
  "body": "Message text"
}
```

Or use `recipient_file` with `create_bulk_drafts`.

## Local file attachments

Files already available on the Windows machine can be attached by path:

```json
{
  "to": ["user@company.ru"],
  "subject": "Report",
  "body": "Report is attached.",
  "attachments": [
    "C:\\Work\\reports\\report.xlsx"
  ]
}
```

## Files uploaded through the MCP client

A client that can read an uploaded file and pass its contents to the MCP tool can use `uploaded_attachments`.

```json
{
  "to": ["user@company.ru"],
  "subject": "Report",
  "body": "Report is attached.",
  "uploaded_attachments": [
    {
      "filename": "report.xlsx",
      "content_base64": "UEsDBBQAAAAI..."
    }
  ]
}
```

The server:

1. validates the Base64 data;
2. writes it to a temporary directory;
3. adds it to the Outlook draft;
4. saves the draft;
5. deletes the temporary copy.

Uploaded files are limited to 20 MB per attachment by the MCP server. Base64 data URLs are also accepted.

Both `attachments` and `uploaded_attachments` may be used in the same request.

For `create_bulk_drafts`, an uploaded attachment is materialized once and attached to every generated draft.

## Tables in the message body

Structured tables may be supplied using the `tables` parameter. The server renders them as Outlook-compatible HTML.

```json
{
  "to": ["user@company.ru"],
  "subject": "Meeting protocol",
  "body": "Agreed actions:",
  "tables": [
    {
      "title": "Actions",
      "columns": ["Task", "Owner", "Due date"],
      "rows": [
        ["Prepare report", "Ivanov", "12.09.2026"]
      ]
    }
  ]
}
```
