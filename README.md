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

## Draft-only mail mode

The server intentionally does not call Outlook `Send()`. Corporate Outlook policies may block programmatic sending, while draft creation remains allowed.

Available mail tools:

- `get_outlook_status`
- `diagnose_outlook`
- `create_draft`
- `create_bulk_drafts`
- `create_drafts_batch`

`create_bulk_drafts` creates one separate Outlook draft per recipient with common content.

`create_drafts_batch` accepts many fully prepared messages in one call, allowing each draft to have its own recipients, subject, body, tables and attachments.

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

The server validates the Base64 data, writes it to a temporary directory, attaches it to the Outlook draft, saves the draft and removes the temporary copy. Uploaded files are limited to 20 MB per attachment. Base64 data URLs are also accepted.

Both `attachments` and `uploaded_attachments` may be used in the same request.

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

## Outlook calendar tools

The server also exposes Outlook calendar operations:

- `list_calendar_events`
- `create_calendar_event`
- `update_calendar_event`
- `delete_calendar_event`

Calendar date/time arguments use local ISO format, for example:

```text
2026-09-10T15:30:00
```

### List events

```json
{
  "start": "2026-09-10T00:00:00",
  "end": "2026-09-11T00:00:00"
}
```

### Create an event

```json
{
  "subject": "Project sync",
  "start": "2026-09-10T15:30:00",
  "end": "2026-09-10T16:00:00",
  "location": "Teams",
  "body": "Discuss project status",
  "attendees": ["user1@company.ru", "user2@company.ru"],
  "reminder_minutes": 15
}
```

The event is saved in the local Outlook calendar. If attendees are provided, the item is prepared as a meeting, but the MCP server deliberately does **not** call `Send()`, so invitations are not sent automatically.

### Update an event

Use the `entry_id` returned by `list_calendar_events` or `create_calendar_event`:

```json
{
  "entry_id": "OUTLOOK_ENTRY_ID",
  "location": "Room 301",
  "start": "2026-09-10T16:00:00",
  "end": "2026-09-10T16:30:00"
}
```

The item is saved without automatically sending meeting updates.

### Delete an event

```json
{
  "entry_id": "OUTLOOK_ENTRY_ID"
}
```

The item is deleted from the local calendar. The server does not automatically send a meeting cancellation.
