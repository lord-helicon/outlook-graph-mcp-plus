"""Settings, read from environment variables.

Each organisation registers its own app in Microsoft Entra ID and gives every
user its client ID and tenant ID. Nothing in this package is tied to one
tenant, and no secret is involved: sign-in is delegated (public client).
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from pathlib import Path

ENV_CLIENT_ID = "OUTLOOK_GRAPH_CLIENT_ID"
ENV_TENANT_ID = "OUTLOOK_GRAPH_TENANT_ID"
ENV_ATTACHMENT_DIR = "OUTLOOK_GRAPH_ATTACHMENT_DIR"
ENV_TIMEZONE = "OUTLOOK_GRAPH_TIMEZONE"
ENV_CACHE_DIR = "OUTLOOK_GRAPH_CACHE_DIR"
ENV_ALLOW_PLAINTEXT = "OUTLOOK_GRAPH_ALLOW_PLAINTEXT_TOKEN_CACHE"

# Delegated Microsoft Graph permissions. offline_access is added by MSAL itself.
SCOPES = [
    "User.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Tasks.ReadWrite",
]

_GUID = re.compile(r"[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}")
_DOMAIN = re.compile(r"[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+")


class ConfigError(RuntimeError):
    """Missing or invalid configuration."""


@dataclass(frozen=True)
class Settings:
    client_id: str
    tenant_id: str
    attachment_dir: Path
    timezone: str
    cache_dir: Path
    allow_plaintext_token_cache: bool

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}"


def _local_timezone() -> str:
    try:
        from tzlocal import get_localzone_name

        return get_localzone_name() or "UTC"
    except Exception:  # noqa: BLE001 - any failure falls back to UTC
        return "UTC"


def load_settings(env: dict[str, str] | None = None) -> Settings:
    """Read settings from the environment; raise ConfigError with a fix-it hint."""
    env = dict(os.environ if env is None else env)
    client_id = env.get(ENV_CLIENT_ID, "").strip()
    tenant_id = env.get(ENV_TENANT_ID, "").strip()
    if not _GUID.fullmatch(client_id):
        raise ConfigError(
            f"{ENV_CLIENT_ID} must be the Application (client) ID of your organisation's "
            "app registration (a GUID). Ask your Microsoft 365 administrator for it."
        )
    if tenant_id.lower() in ("", "common", "consumers"):
        raise ConfigError(
            f"{ENV_TENANT_ID} must be your organisation's Directory (tenant) ID or domain "
            "(e.g. contoso.onmicrosoft.com). Personal Microsoft accounts are not supported."
        )
    if not (
        _GUID.fullmatch(tenant_id) or _DOMAIN.fullmatch(tenant_id) or tenant_id == "organizations"
    ):
        raise ConfigError(f"{ENV_TENANT_ID} '{tenant_id}' is not a tenant ID or domain.")
    home = Path.home()
    attachment_dir = Path(env.get(ENV_ATTACHMENT_DIR) or home / "Downloads").expanduser()
    cache_dir = Path(env.get(ENV_CACHE_DIR) or home / ".outlook-graph-mcp-plus").expanduser()
    allow_plain = env.get(ENV_ALLOW_PLAINTEXT, "").strip().lower() in ("1", "true", "yes")
    return Settings(
        client_id=client_id,
        tenant_id=tenant_id,
        attachment_dir=attachment_dir,
        timezone=(env.get(ENV_TIMEZONE) or "").strip() or _local_timezone(),
        cache_dir=cache_dir,
        allow_plaintext_token_cache=allow_plain,
    )
