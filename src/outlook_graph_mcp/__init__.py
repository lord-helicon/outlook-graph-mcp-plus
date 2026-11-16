"""Outlook (Microsoft 365) as a local MCP server via Microsoft Graph."""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("outlook-graph-mcp-plus")
except PackageNotFoundError:  # running from a source checkout without install
    __version__ = "0.0.0"
