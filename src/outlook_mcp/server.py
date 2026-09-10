from mcp.server.fastmcp import FastMCP

from .models import EmailRequest
from .outlook import create_draft as outlook_create_draft
from .outlook import get_outlook_status as outlook_get_status
from .outlook import send_email as outlook_send_email


mcp = FastMCP("outlook-mcp-server")


@mcp.tool()
def get_outlook_status() -> dict:
    """Check Outlook COM availability and return configured accounts."""
    return outlook_get_status()


@mcp.tool()
def create_draft(request: EmailRequest) -> dict:
    """Create an Outlook draft with optional HTML tables and local file attachments."""
    return outlook_create_draft(request)


@mcp.tool()
def send_email(request: EmailRequest) -> dict:
    """Send an Outlook email with optional HTML tables and local file attachments."""
    return outlook_send_email(request)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
