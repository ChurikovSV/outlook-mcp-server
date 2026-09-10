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

Creates a draft in Outlook. Recommended as the default operation for agents.

### `send_email`

Sends an email immediately through Outlook.

## Request model

```json
{
  "to": ["user1@company.ru"],
  "cc": ["user2@company.ru"],
  "bcc": [],
  "subject": "Meeting protocol",
  "body": "Hello!\nPlease find the meeting results below.",
  "tables": [],
  "attachments": []
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

Multiple tables can be passed in the same request by adding more objects to `tables`.

## Email with attachments

Attachments are local Windows file paths accessible to the Windows user running Outlook and the MCP server.

```json
{
  "to": ["team@company.ru"],
  "subject": "Meeting protocol",
  "body": "The full protocol is attached.",
  "tables": [],
  "attachments": [
    "C:\\Users\\sergey\\Documents\\protocol.docx",
    "C:\\Users\\sergey\\Documents\\metrics.xlsx"
  ]
}
```

The server validates that each attachment exists before creating or sending the email.

## Email with a table and attachments

```json
{
  "to": ["team@company.ru"],
  "cc": ["manager@company.ru"],
  "subject": "Meeting results",
  "body": "Colleagues, below is a short summary. The detailed report is attached.",
  "tables": [
    {
      "title": "Decisions",
      "columns": ["Decision", "Responsible", "Status"],
      "rows": [
        ["Update dashboard", "Ivanov", "In progress"],
        ["Validate source data", "Petrov", "Planned"]
      ]
    }
  ],
  "attachments": [
    "C:\\Work\\meeting\\report.xlsx"
  ]
}
```

## Recommended agent policy

For meeting-protocol workflows, prefer `create_draft` by default. Use `send_email` only when the user explicitly asks to send the message.

This allows the user to review recipients, subject, body, tables and attachments in Outlook before sending.

## Current scope

Version `0.1.0` supports:

- Outlook availability check
- configured account listing
- draft creation
- direct sending
- To / CC / BCC
- plain text body rendered as HTML
- structured HTML tables
- local file attachments

Planned extensions can include reply/forward, selecting a specific sending account, inline images, draft lookup and meeting-oriented templates.
