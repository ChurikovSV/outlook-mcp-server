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

## Run

```powershell
outlook-mcp
```

or:

```powershell
python -m outlook_mcp.server
```

## MCP tools

### `get_outlook_status`

Checks that Outlook COM is available and returns configured Outlook accounts.

### `create_draft`

Creates one draft. Recipients can be passed directly and/or loaded from a local `.txt`, `.csv` or `.xlsx` file.

### `send_email`

Sends one email. All `to` recipients are placed into the same message and can see each other.

### `send_bulk_email`

Sends a separate email to every recipient. Use this for distribution lists when recipients must not see the other addresses.

## Send to a recipient list

```json
{
  "to": [
    "user1@company.ru",
    "user2@company.ru",
    "user3@company.ru"
  ],
  "subject": "Meeting protocol",
  "body": "Colleagues, sending the meeting protocol.",
  "tables": [],
  "attachments": []
}
```

This request can be used with `create_draft` or `send_email`.

## Send to recipients from TXT

Example `C:\\Work\\mail\\recipients.txt`:

```text
user1@company.ru
user2@company.ru
user3@company.ru
```

Request:

```json
{
  "to": [],
  "recipient_file": "C:\\Work\\mail\\recipients.txt",
  "subject": "Meeting protocol",
  "body": "Colleagues, sending the meeting protocol.",
  "tables": [],
  "attachments": []
}
```

A TXT file is read as UTF-8. One address per line is recommended. Commas, semicolons and spaces inside a line are also accepted as separators.

## Send to recipients from CSV

Example `C:\\Work\\mail\\recipients.csv`:

```csv
email,name
user1@company.ru,Ivan Ivanov
user2@company.ru,Petr Petrov
user3@company.ru,Anna Sidorova
```

Request:

```json
{
  "to": [],
  "recipient_file": "C:\\Work\\mail\\recipients.csv",
  "recipient_file_column": "email",
  "subject": "Meeting protocol",
  "body": "Colleagues, sending the meeting protocol.",
  "tables": [],
  "attachments": []
}
```

CSV delimiters `,`, `;` and tab are detected automatically. The header is required. The default recipient column is `email`.

## Send to recipients from Excel XLSX

The `.xlsx` file is read directly with `openpyxl`; Microsoft Excel does not need to be started.

Example workbook `C:\\Work\\mail\\users.xlsx`:

| ФИО | Почта | Подразделение |
| --- | --- | --- |
| Иванов Иван | user1@company.ru | Support |
| Петров Петр | user2@company.ru | DevOps |
| Сидорова Анна | user3@company.ru | Analytics |

Request:

```json
{
  "to": [],
  "recipient_file": "C:\\Work\\mail\\users.xlsx",
  "recipient_file_column": "Почта",
  "subject": "Meeting protocol",
  "body": "Colleagues, sending the meeting protocol."
}
```

By default the first worksheet is used. To select a specific worksheet:

```json
{
  "to": [],
  "recipient_file": "C:\\Work\\mail\\users.xlsx",
  "recipient_file_sheet": "Получатели",
  "recipient_file_column": "Почта",
  "subject": "Notification",
  "body": "Message text"
}
```

Column and worksheet names are passed exactly as they appear in the workbook. Column lookup is case-insensitive and ignores leading/trailing spaces.

Excel recipient files also work with `send_bulk_email`:

```json
{
  "recipients": [],
  "recipient_file": "C:\\Work\\mail\\users.xlsx",
  "recipient_file_sheet": "Получатели",
  "recipient_file_column": "Почта",
  "subject": "Notification",
  "body": "Message text"
}
```

## Combine a direct list and a file

The two sources can be combined:

```json
{
  "to": ["manager@company.ru"],
  "recipient_file": "C:\\Work\\mail\\team.xlsx",
  "recipient_file_column": "email",
  "subject": "Meeting protocol",
  "body": "Colleagues, sending the meeting protocol."
}
```

Duplicate addresses are removed automatically.

## Individual bulk sending

Use `send_bulk_email` when each recipient should receive a separate message:

```json
{
  "recipients": ["manager@company.ru"],
  "recipient_file": "C:\\Work\\mail\\team.csv",
  "recipient_file_column": "email",
  "subject": "Notification",
  "body": "Message text",
  "tables": [],
  "attachments": []
}
```

The result contains the total number of recipients, successfully sent messages and per-recipient failures:

```json
{
  "status": "completed",
  "total": 15,
  "sent": 15,
  "failed": []
}
```

## Email with a table

The agent does not need to generate Outlook-compatible HTML itself. It sends structured table data and the MCP server renders it into HTML.

```json
{
  "to": ["team@company.ru"],
  "subject": "Meeting protocol — 10.09.2026",
  "body": "Colleagues, here are the agreed actions.",
  "tables": [
    {
      "title": "Actions",
      "columns": ["Task", "Owner", "Due date"],
      "rows": [
        ["Prepare report", "Ivan Ivanov", "12.09.2026"],
        ["Check metrics", "Petr Petrov", "15.09.2026"]
      ]
    }
  ],
  "attachments": []
}
```

## Email with attachments

Attachments are local Windows file paths accessible to the Windows user running Outlook and the MCP server.

```json
{
  "to": ["team@company.ru"],
  "subject": "Meeting protocol",
  "body": "The full protocol is attached.",
  "attachments": [
    "C:\\Users\\sergey\\Documents\\protocol.docx",
    "C:\\Users\\sergey\\Documents\\metrics.xlsx"
  ]
}
```

The server validates that each attachment exists before creating or sending the email.

## Recommended agent policy

For meeting-protocol workflows, prefer `create_draft` by default. Use `send_email` or `send_bulk_email` only when the user explicitly asks to send the message.

## Current scope

Version `0.1.0` supports:

- Outlook availability check
- configured account listing
- draft creation
- direct sending
- individual bulk sending
- recipients passed as a list
- recipients loaded from TXT
- recipients loaded from CSV
- recipients loaded from Excel XLSX
- selecting an XLSX worksheet and recipient column
- recipient deduplication and basic email validation
- To / CC / BCC
- plain text body rendered as HTML
- structured HTML tables
- local file attachments

Planned extensions can include reply/forward, selecting a specific sending account, inline images, draft lookup and meeting-oriented templates.
