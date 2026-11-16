"""Command line: `serve` (default), `auth`, `status`, `logout`."""

from __future__ import annotations

import argparse
import json
import logging
import sys

from outlook_graph_mcp import __version__
from outlook_graph_mcp.auth import AuthError, TokenProvider, eprint
from outlook_graph_mcp.config import ConfigError, load_settings


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="outlook-graph-mcp-plus",
        description="Microsoft 365 Outlook (mail, calendar, To Do) as a local MCP server.",
    )
    parser.add_argument("--version", action="version", version=__version__)
    sub = parser.add_subparsers(dest="command")
    sub.add_parser("serve", help="run the MCP server on stdio (default)")
    auth = sub.add_parser("auth", help="sign in to Microsoft 365 (once per user)")
    auth.add_argument(
        "--device-code",
        action="store_true",
        help="sign in with a code on another device instead of a local browser",
    )
    sub.add_parser("status", help="show whether you are signed in")
    sub.add_parser("logout", help="remove cached sign-in tokens")
    args = parser.parse_args(argv)
    command = args.command or "serve"

    if command == "serve":
        # stdout belongs to the MCP protocol; logs go to stderr.
        logging.basicConfig(
            level=logging.INFO,
            stream=sys.stderr,
            format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
        )
        from outlook_graph_mcp import tools  # noqa: F401 - registers the tools
        from outlook_graph_mcp.app import mcp

        mcp.run()
        return 0

    try:
        settings = load_settings()
        tokens = TokenProvider(settings)
        if command == "auth":
            user = tokens.login(device_code=args.device_code, echo=eprint)
            eprint(f"Signed in as {user}.")
        elif command == "status":
            print(
                json.dumps(
                    {
                        **tokens.status(),
                        "tenant": settings.tenant_id,
                        "attachment_dir": str(settings.attachment_dir),
                        "timezone": settings.timezone,
                    },
                    indent=2,
                )
            )
        elif command == "logout":
            eprint(f"Removed {tokens.logout()} cached account(s).")
    except (ConfigError, AuthError) as e:
        eprint(f"Error: {e}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
