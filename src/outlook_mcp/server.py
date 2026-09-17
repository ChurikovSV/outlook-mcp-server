import argparse

from mcp.server.fastmcp import FastMCP

from .browser_owa_freebusy import (
    diagnose_browser_owa as outlook_diagnose_browser_owa,
    get_browser_owa_free_busy as outlook_get_browser_owa_free_busy,
)
from .calendar import (
    create_calendar_event as outlook_create_calendar_event,
    delete_calendar_event as outlook_delete_calendar_event,
    diagnose_meeting_attendee as outlook_diagnose_meeting_attendee,
    list_calendar_events as outlook_list_calendar_events,
    prepare_calendar_meeting as outlook_prepare_calendar_meeting,
    update_calendar_event as outlook_update_calendar_event,
)
from .diagnostics import (
    diagnose_calendar as run_calendar_diagnostics,
    diagnose_outlook as run_outlook_diagnostics,
)
from .freebusy import (
    diagnose_free_busy as outlook_diagnose_free_busy,
    get_employee_free_busy as outlook_get_employee_free_busy,
)
from .mailbox_calendar import (
    list_mailbox_calendar_events as outlook_list_mailbox_calendar_events,
)
from .mailbox_calendar_diagnostics import (
    diagnose_mailbox_calendar as outlook_diagnose_mailbox_calendar,
)
from .markdown_draft import (
    create_draft_from_markdown as outlook_create_draft_from_markdown,
)
from .owa_freebusy import (
    diagnose_owa_free_busy as outlook_diagnose_owa_free_busy,
    get_owa_free_busy as outlook_get_owa_free_busy,
)
from .send_as_diagnostics import (
    diagnose_send_as_account as outlook_diagnose_send_as_account,
)
from .models import BatchDraftRequest, BulkEmailRequest, EmailRequest
from .outlook import create_bulk_drafts as outlook_create_bulk_drafts
from .outlook import create_draft as outlook_create_draft
from .outlook import create_drafts_batch as outlook_create_drafts_batch
from .outlook import get_outlook_status as outlook_get_status


mcp: FastMCP | None = None


def _build_server(host: str, port: int) -> FastMCP:
    server = FastMCP(
        "outlook-mcp-server",
        host=host,
        port=port,
    )

    @server.tool()
    def diagnose_outlook() -> dict:
        """Run step-by-step Outlook COM diagnostics without raising COM errors."""
        return run_outlook_diagnostics()

    @server.tool()
    def diagnose_calendar() -> dict:
        """Run step-by-step Outlook calendar diagnostics and identify the failing COM operation."""
        return run_calendar_diagnostics()

    @server.tool()
    def diagnose_mailbox_calendar(
        store_name: str = "Управление программами",
        sample_limit: int = 5,
    ) -> dict:
        """Check read-only access to an additional/shared Outlook Store calendar. Never creates, changes, deletes, saves, or sends anything."""
        return outlook_diagnose_mailbox_calendar(
            store_name=store_name,
            sample_limit=sample_limit,
        )

    @server.tool()
    def diagnose_send_as_account(
        email: str = "up_onework@sberbank.ru",
    ) -> dict:
        """Check whether Outlook can prepare a draft on behalf of another mailbox without saving or sending it."""
        return outlook_diagnose_send_as_account(email=email)

    @server.tool()
    def diagnose_meeting_attendee(attendee: str) -> dict:
        """Test adding and resolving a meeting attendee without saving or sending anything."""
        return outlook_diagnose_meeting_attendee(attendee)

    @server.tool()
    def diagnose_free_busy(email: str, slot_minutes: int = 30) -> dict:
        """Test whether Outlook COM can resolve an employee and read Exchange free/busy data."""
        return outlook_diagnose_free_busy(email=email, slot_minutes=slot_minutes)

    @server.tool()
    def get_employee_free_busy(
        email: str,
        start: str,
        end: str,
        slot_minutes: int = 30,
    ) -> dict:
        """Get an employee's free/busy using Outlook COM. Corporate policy may block this path."""
        return outlook_get_employee_free_busy(
            email=email,
            start=start,
            end=end,
            slot_minutes=slot_minutes,
        )

    @server.tool()
    def diagnose_owa_free_busy(email: str) -> dict:
        """Test OWA GetUserAvailabilityInternal through direct HTTP. May be blocked by browser-bound authentication."""
        return outlook_diagnose_owa_free_busy(email=email)

    @server.tool()
    def get_owa_free_busy(
        emails: list[str],
        start: str,
        end: str,
        slot_minutes: int = 30,
    ) -> dict:
        """Get free/busy through direct OWA HTTP. May be blocked by browser-bound authentication."""
        return outlook_get_owa_free_busy(
            emails=emails,
            start=start,
            end=end,
            slot_minutes=slot_minutes,
        )

    @server.tool()
    def diagnose_browser_owa(cdp_url: str | None = None) -> dict:
        """Connect to an already-running Chromium/SberBrowser via CDP, locate the authenticated mail.sberbank.ru tab, and verify that the active OWA session can be observed. Does not return cookies or canary values."""
        return outlook_diagnose_browser_owa(cdp_url=cdp_url)

    @server.tool()
    def get_browser_owa_free_busy(
        emails: list[str],
        start: str,
        end: str,
        slot_minutes: int = 30,
        cdp_url: str | None = None,
    ) -> dict:
        """Get OWA free/busy by executing GetUserAvailabilityInternal from inside an already-authenticated OWA browser tab connected through Chrome DevTools Protocol. Prefer this when direct HTTP returns 401."""
        return outlook_get_browser_owa_free_busy(
            emails=emails,
            start=start,
            end=end,
            slot_minutes=slot_minutes,
            cdp_url=cdp_url,
        )

    @server.tool()
    def get_outlook_status() -> dict:
        """Check whether Outlook COM can create mail items for draft-only workflows."""
        return outlook_get_status()

    @server.tool()
    def create_draft(
        subject: str,
        to: list[str] | None = None,
        body: str = "",
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        from_email: str | None = None,
        recipient_file: str | None = None,
        recipient_file_column: str = "email",
        recipient_file_sheet: str | None = None,
        attachments: list[str] | None = None,
        uploaded_attachments: list[dict] | None = None,
        tables: list[dict] | None = None,
    ) -> dict:
        """Create one Outlook draft. Optionally set from_email for a shared/delegated mailbox. Never calls Send()."""
        request = EmailRequest(
            to=to or [],
            cc=cc or [],
            bcc=bcc or [],
            from_email=from_email,
            recipient_file=recipient_file,
            recipient_file_column=recipient_file_column,
            recipient_file_sheet=recipient_file_sheet,
            subject=subject,
            body=body,
            attachments=attachments or [],
            uploaded_attachments=uploaded_attachments or [],
            tables=tables or [],
        )
        return outlook_create_draft(request)

    @server.tool()
    def create_draft_from_markdown(
        subject: str,
        to: list[str] | None = None,
        markdown: str | None = None,
        markdown_file: str | None = None,
        variables: dict[str, str] | None = None,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        from_email: str | None = None,
        recipient_file: str | None = None,
        recipient_file_column: str = "email",
        recipient_file_sheet: str | None = None,
        attachments: list[str] | None = None,
        uploaded_attachments: list[dict] | None = None,
    ) -> dict:
        """Create a formatted Outlook draft from Markdown text or a local .md file. Use {{name}} placeholders with variables for personalization. For one recipient {{email}} is populated automatically. Supports from_email for a shared/delegated mailbox such as a GPYa. Never calls Send()."""
        request = EmailRequest(
            to=to or [],
            cc=cc or [],
            bcc=bcc or [],
            from_email=from_email,
            recipient_file=recipient_file,
            recipient_file_column=recipient_file_column,
            recipient_file_sheet=recipient_file_sheet,
            subject=subject,
            attachments=attachments or [],
            uploaded_attachments=uploaded_attachments or [],
        )
        return outlook_create_draft_from_markdown(
            request=request,
            markdown=markdown,
            markdown_file=markdown_file,
            variables=variables,
        )

    @server.tool()
    def create_bulk_drafts(
        subject: str,
        recipients: list[str] | None = None,
        body: str = "",
        from_email: str | None = None,
        recipient_file: str | None = None,
        recipient_file_column: str = "email",
        recipient_file_sheet: str | None = None,
        attachments: list[str] | None = None,
        uploaded_attachments: list[dict] | None = None,
        tables: list[dict] | None = None,
    ) -> dict:
        """Create one Outlook draft per recipient with common content. Optionally set from_email for a shared/delegated mailbox. Never calls Send()."""
        request = BulkEmailRequest(
            recipients=recipients or [],
            from_email=from_email,
            recipient_file=recipient_file,
            recipient_file_column=recipient_file_column,
            recipient_file_sheet=recipient_file_sheet,
            subject=subject,
            body=body,
            attachments=attachments or [],
            uploaded_attachments=uploaded_attachments or [],
            tables=tables or [],
        )
        return outlook_create_bulk_drafts(request)

    @server.tool()
    def create_drafts_batch(drafts: list[dict]) -> dict:
        """Create many fully prepared Outlook drafts in one call. Each item can have its own from_email, recipients, subject, body, tables and attachments. Never calls Send()."""
        request = BatchDraftRequest(drafts=drafts)
        return outlook_create_drafts_batch(request)

    @server.tool()
    def list_calendar_events(
        start: str,
        end: str,
        limit: int = 100,
        store_name: str | None = None,
    ) -> dict:
        """List Outlook calendar events. Omit store_name for the default calendar, or pass an additional/shared Outlook Store such as 'Управление программами'."""
        if store_name:
            return outlook_list_mailbox_calendar_events(
                store_name=store_name,
                start=start,
                end=end,
                limit=limit,
            )
        return outlook_list_calendar_events(start=start, end=end, limit=limit)

    @server.tool()
    def prepare_calendar_meeting(
        subject: str,
        start: str,
        end: str,
        attendees: list[str],
        location: str = "",
        body: str = "",
        all_day: bool = False,
        reminder_minutes: int | None = 15,
    ) -> dict:
        """Open a prepared Outlook meeting window for manual review/send. Pass the requested attendees exactly as provided by the user. Do not omit them and do not replace them with an empty list. If Outlook blocks Recipients.Add, this tool automatically uses the RequiredAttendees fallback. It does not save or send programmatically."""
        if not attendees:
            raise ValueError("prepare_calendar_meeting requires at least one attendee")
        return outlook_prepare_calendar_meeting(
            subject=subject,
            start=start,
            end=end,
            location=location,
            body=body,
            attendees=attendees,
            all_day=all_day,
            reminder_minutes=reminder_minutes,
        )

    @server.tool()
    def create_calendar_event(
        subject: str,
        start: str,
        end: str,
        location: str = "",
        body: str = "",
        attendees: list[str] | None = None,
        all_day: bool = False,
        reminder_minutes: int | None = 15,
    ) -> dict:
        """Create and save an Outlook calendar event. If saving a meeting with attendees is blocked, opens the prepared meeting window instead. Never calls Send()."""
        return outlook_create_calendar_event(
            subject=subject,
            start=start,
            end=end,
            location=location,
            body=body,
            attendees=attendees,
            all_day=all_day,
            reminder_minutes=reminder_minutes,
        )

    @server.tool()
    def update_calendar_event(
        entry_id: str,
        subject: str | None = None,
        start: str | None = None,
        end: str | None = None,
        location: str | None = None,
        body: str | None = None,
        all_day: bool | None = None,
        reminder_minutes: int | None = None,
        disable_reminder: bool = False,
    ) -> dict:
        """Update and save an Outlook calendar event without sending meeting updates."""
        return outlook_update_calendar_event(
            entry_id=entry_id,
            subject=subject,
            start=start,
            end=end,
            location=location,
            body=body,
            all_day=all_day,
            reminder_minutes=reminder_minutes,
            disable_reminder=disable_reminder,
        )

    @server.tool()
    def delete_calendar_event(entry_id: str) -> dict:
        """Delete an Outlook calendar item locally without sending a cancellation."""
        return outlook_delete_calendar_event(entry_id=entry_id)

    return server


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Outlook MCP server")
    parser.add_argument(
        "--transport",
        choices=("streamable-http", "stdio", "sse"),
        default="streamable-http",
        help="MCP transport (default: streamable-http)",
    )
    parser.add_argument(
        "--host",
        default="127.0.0.1",
        help="HTTP/SSE bind address (default: 127.0.0.1)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="HTTP/SSE port (default: 8000)",
    )
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    server = _build_server(args.host, args.port)

    if args.transport == "stdio":
        server.run(transport="stdio")
        return

    server.run(transport=args.transport)


if __name__ == "__main__":
    main()
