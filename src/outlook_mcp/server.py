import argparse

from mcp.server.fastmcp import FastMCP

from .diagnostics import diagnose_outlook as run_outlook_diagnostics
from .models import BulkEmailRequest, EmailRequest
from .outlook import create_bulk_drafts as outlook_create_bulk_drafts
from .outlook import create_draft as outlook_create_draft
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
    def get_outlook_status() -> dict:
        """Check whether Outlook COM can create mail items for draft-only workflows."""
        return outlook_get_status()

    @server.tool()
    def create_draft(
        to: list[str],
        subject: str,
        body: str = "",
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        recipient_file: str | None = None,
        recipient_file_column: str = "email",
        recipient_file_sheet: str | None = None,
        attachments: list[str] | None = None,
        tables: list[dict] | None = None,
    ) -> dict:
        """Create one Outlook draft. The server never calls Send()."""
        request = EmailRequest(
            to=to,
            cc=cc or [],
            bcc=bcc or [],
            recipient_file=recipient_file,
            recipient_file_column=recipient_file_column,
            recipient_file_sheet=recipient_file_sheet,
            subject=subject,
            body=body,
            attachments=attachments or [],
            tables=tables or [],
        )
        return outlook_create_draft(request)

    @server.tool()
    def create_bulk_drafts(
        recipients: list[str],
        subject: str,
        body: str = "",
        recipient_file: str | None = None,
        recipient_file_column: str = "email",
        recipient_file_sheet: str | None = None,
        attachments: list[str] | None = None,
        tables: list[dict] | None = None,
    ) -> dict:
        """Create a separate Outlook draft for every recipient. The server never calls Send()."""
        request = BulkEmailRequest(
            recipients=recipients,
            recipient_file=recipient_file,
            recipient_file_column=recipient_file_column,
            recipient_file_sheet=recipient_file_sheet,
            subject=subject,
            body=body,
            attachments=attachments or [],
            tables=tables or [],
        )
        return outlook_create_bulk_drafts(request)

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
