"""Outlook MCP Server - fork (multi-account allowlist).

Fork of Abhishek-Aditya-bs/Outlook-MCP-Server @ bb641f7.
Adds: namespace.Accounts enumeration gated by an SMTP allowlist
(`included_accounts` in config.properties). Fail-closed.
"""

import asyncio
import logging
import platform
import sys
from typing import Any, Sequence

# Check if running on Windows
if platform.system() != 'Windows':
    print("[ERROR] Outlook MCP Server requires Windows with Microsoft Outlook installed")
    print(f"   Current platform: {platform.system()}")
    print("\n[INFO] To use this server:")
    print("   1. Run on a Windows machine with Outlook installed")
    print("   2. Or use a Windows virtual machine")
    print("   3. Or access a remote Windows desktop")
    sys.exit(1)

from mcp import server, types
from mcp.server import Server
from mcp.server.stdio import stdio_server

try:
    from src.config.config_reader import config
    from src.utils.outlook_client import outlook_client
    from src.utils.email_formatter import format_mailbox_status, format_email_chain
except ImportError as e:
    print(f"[ERROR] Import Error: {e}")
    print("\n[INFO] Please install required dependencies:")
    print("   pip install -r requirements.txt")
    print("\nNote: pywin32 is required and only works on Windows")
    sys.exit(1)

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Create MCP server
app = Server("outlook-mcp-server")


@app.list_tools()
async def list_tools() -> list[types.Tool]:
    """List available MCP tools."""
    return [
        types.Tool(
            name="check_mailbox_access",
            description=(
                "Check connection status and per-account accessibility for every "
                "Outlook account on the SMTP allowlist (included_accounts)."
            ),
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),
        types.Tool(
            name="get_email_chain",
            description=(
                "Search emails containing the specified text in BOTH subject and body "
                "using exact phrase matching. Searches every Outlook account on the "
                "SMTP allowlist (included_accounts) unless `accounts` is provided to "
                "restrict to a subset. Returns full email content with originating-account "
                "labeling. Use specific search terms (error codes, unique phrases) for "
                "best results."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "search_text": {
                        "type": "string",
                        "description": (
                            "Exact text pattern to search for in email subject and body. "
                            "The search looks for this exact phrase."
                        )
                    },
                    "accounts": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": (
                            "Optional list of SMTP addresses to restrict the search to "
                            "(must be a subset of included_accounts). If omitted, all "
                            "allowlisted accounts are searched."
                        )
                    },
                    "include_personal": {
                        "type": "boolean",
                        "description": (
                            "DEPRECATED. Retained for backward compatibility. "
                            "Ignored when `accounts` is set."
                        ),
                        "default": True
                    },
                    "include_shared": {
                        "type": "boolean",
                        "description": (
                            "DEPRECATED. Retained for backward compatibility. "
                            "Ignored when `accounts` is set."
                        ),
                        "default": True
                    }
                },
                "required": ["search_text"]
            }
        )
    ]


@app.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> Sequence[types.TextContent]:
    """Handle tool calls."""
    logger.info(f"Executing tool: {name}")

    try:
        if name == "check_mailbox_access":
            return await handle_check_mailbox_access()

        elif name == "get_email_chain":
            search_text = arguments.get("search_text")
            if not search_text:
                raise ValueError("search_text parameter is required")

            accounts = arguments.get("accounts")
            include_personal = arguments.get("include_personal", True)
            include_shared = arguments.get("include_shared", True)

            return await handle_get_email_chain(
                search_text, accounts, include_personal, include_shared
            )

        else:
            raise ValueError(f"Unknown tool: {name}")

    except Exception as e:
        logger.error(f"Error in tool {name}: {e}")
        error_response = {
            "status": "error",
            "tool": name,
            "error": str(e),
            "message": f"Failed to execute {name}: {str(e)}"
        }
        return [types.TextContent(type="text", text=str(error_response))]


async def handle_check_mailbox_access():
    """Handle mailbox access check."""
    logger.info("Checking mailbox access...")
    try:
        access_result = await asyncio.to_thread(outlook_client.check_access)
        formatted_result = format_mailbox_status(access_result)
        logger.info("Mailbox access check completed")
        return [types.TextContent(type="text", text=str(formatted_result))]
    except Exception as e:
        logger.error(f"Error checking mailbox access: {e}")
        error_response = {
            "status": "error",
            "message": f"Could not check mailbox access: {str(e)}",
            "troubleshooting": [
                "Make sure Outlook is running",
                "Grant permission when security dialog appears",
                "Verify included_accounts is set in config.properties",
            ]
        }
        return [types.TextContent(type="text", text=str(error_response))]


async def handle_get_email_chain(search_text: str, accounts,
                                 include_personal: bool, include_shared: bool):
    """Handle email search and retrieval."""
    logger.info(f"Searching for emails containing: {search_text}")
    try:
        emails = await asyncio.to_thread(
            outlook_client.search_emails,
            search_text=search_text,
            accounts=accounts,
            include_personal=include_personal,
            include_shared=include_shared,
        )
        formatted_result = format_email_chain(emails, search_text)
        logger.info(f"Found {len(emails)} emails containing '{search_text}'")
        return [types.TextContent(type="text", text=str(formatted_result))]
    except Exception as e:
        logger.error(f"Error searching emails: {e}")
        error_response = {
            "status": "error",
            "search_text": search_text,
            "message": f"Could not search emails: {str(e)}",
            "troubleshooting": [
                "Verify Outlook connection",
                "Check that included_accounts is configured",
                "Use specific search terms for best results",
            ]
        }
        return [types.TextContent(type="text", text=str(error_response))]


@app.list_resources()
async def list_resources() -> list[types.Resource]:
    """List available resources."""
    return [
        types.Resource(
            uri="outlook-mcp://config",
            name="Current Configuration",
            description="Show current configuration settings",
            mimeType="text/plain"
        )
    ]


@app.read_resource()
async def read_resource(uri: str) -> str:
    """Read resource content."""
    if uri == "outlook-mcp://config":
        config.show_config()
        return "Configuration displayed in console"
    raise ValueError(f"Unknown resource: {uri}")


async def main():
    """Main entry point."""
    print("=" * 60)
    print("[STARTING] Outlook MCP Server (fork: multi-account allowlist)")
    print("=" * 60)

    config.show_config()

    allowlist = config.get_list('included_accounts', [])
    print("\n[INFO] Important Notes:")
    print("   * Make sure Microsoft Outlook is running")
    print("   * Grant permission when security dialog appears")

    if not allowlist:
        print("\n[WARNING] included_accounts is empty!")
        print("   The server will return empty results until SMTP addresses are added")
        print("   to config.properties (comma-separated).")
    else:
        print(f"\n[INFO] SMTP allowlist ({len(allowlist)} account(s)):")
        for smtp in allowlist:
            print(f"   * {smtp}")

    # Upstream config key deprecation warning
    legacy_shared = config.get('shared_mailbox_email')
    if legacy_shared and 'example.com' not in str(legacy_shared):
        print("\n[WARNING] 'shared_mailbox_email' is set but is DEPRECATED in this fork.")
        print("   The value is IGNORED. Migrate to included_accounts.")

    print("\n[TOOLS] Available Tools:")
    print("   1. check_mailbox_access  - Test connection and per-account access")
    print("   2. get_email_chain       - Search by phrase across allowlisted accounts")

    print("\n[READY] Server ready! Listening for MCP client connections...")
    print("=" * 60)

    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n[INFO] Server stopped by user")
    except Exception as e:
        print(f"\n[ERROR] Server error: {e}")
        logger.error(f"Server error: {e}")
