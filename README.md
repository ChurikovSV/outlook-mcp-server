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
outlook-mcp
```

By default the server starts with Streamable HTTP on localhost:

```text
http://127.0.0.1:8000/mcp
```

Equivalent explicit command:

```powershell
outlook-mcp --transport streamable-http --host 127.0.0.1 --port 8000
```

Use a different port if needed:

```powershell
outlook-mcp --port 8765
```

Then the MCP endpoint is:

```text
http://127.0.0.1:8765/mcp
```

For security, `127.0.0.1` is the default bind address. Do not expose the server on `0.0.0.0` unless access is intentionally protected and allowed by corporate policy.

## Other transports

stdio is still available:

```powershell
outlook-mcp --transport stdio
```

SSE is available only for compatibility with older MCP clients:

```powershell
outlook-mcp --transport sse --host 127.0.0.1 --port 8000
```

For new HTTP integrations, use Streamable HTTP rather than SSE.

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

## Send to recipients from CSV

```json
{
  "to": [],
  "recipient_file": "C:\\Work\\mail\\recipients.csv",
  "recipient_file_column": "email",
  "subject": "Meeting protocol",
  "body": "Colleagues, sending the meeting protocol."
}
```

CSV delimiters `,`, `;` and tab are detected automatically. The header is required.

## Send to recipients from Excel

Example workbook:

```text
Sheet: Получатели

ФИО             Почта                 Подразделение
Иванов Иван     ivanov@company.ru     Support
Петров Петр     petrov@company.ru     DevOps
```

Request:

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

If `recipient_file_sheet` is omitted, the first worksheet is used.

## Individual bulk sending

Use `send_bulk_email` when each recipient should receive a separate message:

```json
{
  "recipients": [],
  "recipient_file": "C:\\Work\\mail\\users.xlsx",
  "recipient_file_sheet": "Получатели",
  "recipient_file_column": "Почта",
  "subject": "Notification",
  "body": "Message text",
  "tables": [],
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

## Recommended agent policy

For meeting-protocol workflows, prefer `create_draft` by default. Use `send_email` or `send_bulk_email` only when the user explicitly asks to send the message.
