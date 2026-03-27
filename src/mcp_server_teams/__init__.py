"""MCP Server Teams — Microsoft Teams integration for AI agents via MCP."""

from .server import serve

__version__ = "0.1.0"


def main():
    """MCP Server Teams — read and send Microsoft Teams messages."""
    import argparse
    import asyncio

    parser = argparse.ArgumentParser(
        description="Give a model the ability to interact with Microsoft Teams",
    )
    parser.add_argument(
        "--client-id",
        type=str,
        help="Azure AD application (client) ID (overrides TEAMS_CLIENT_ID env var)",
    )
    parser.add_argument(
        "--tenant-id",
        type=str,
        help="Azure AD directory (tenant) ID (overrides TEAMS_TENANT_ID env var)",
    )

    args = parser.parse_args()
    asyncio.run(serve(args.client_id, args.tenant_id))


if __name__ == "__main__":
    main()
