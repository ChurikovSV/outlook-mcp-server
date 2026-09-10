import argparse

from mcp.server.fastmcp import FastMCP

from .models import BulkEmailRequest, EmailRequest
from .outlook import create_draft as outlook_create_draft
from .outlook import get_outlook_status as outlook_get_status
from .outlook import send_bulk_email as outlook_send_bulk_email
from .outlook import send_email as outlook_send_email


mcp = FastMCP("outlook-mcp-server")


@mcp.tool()
def get_outlook_status() -> dict:
    """Check Outlook COM availability and return configured accounts."""
    return outlook_get_status()


@mcp.tool()
def create_draft(request: EmailRequest) -> dict:
    """Create an Outlook draft using recipients from a list and/or TXT/CSV/XLSX file."""
    return outlook_create_draft(request)


@mcp.tool()
def send_email(request: EmailRequest) -> dict:
    """Send one Outlook email using recipients from a list and/or TXT/CSV/XLSX file."""
    return outlook_send_email(request)


@mcp.tool()
def send_bulk_email(request: BulkEmailRequest) -> dict:
    """Send a separate Outlook email to each recipient from a list and/or TXT/CSV/XLSX file."""
    return outlook_send_bulk_email(request)


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

    if args.transport == "stdio":
        mcp.run(transport="stdio")
        return

    mcp.run(
        transport=args.transport,
        host=args.host,
        port=args.port,
    )


if __name__ == "__main__":
    main()
