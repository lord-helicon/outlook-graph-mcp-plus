"""Delegated sign-in with MSAL (public client, no secret).

`outlook-graph-mcp-plus auth` signs the user in once: in the browser
(authorization code + PKCE on a localhost redirect), or with a device code
when no browser is available. Tokens live in an MSAL cache persisted in the
OS secret store (macOS Keychain, Windows DPAPI, Linux libsecret). A plaintext
cache file is used only when explicitly allowed.

The MCP server itself never prompts: it only refreshes silently and asks the
user to run `auth` again when that fails.
"""

from __future__ import annotations

import json
import sys
from collections.abc import Callable
from pathlib import Path

import msal
from msal_extensions import FilePersistence, PersistedTokenCache, build_encrypted_persistence

from outlook_graph_mcp.config import SCOPES, Settings


class AuthError(RuntimeError):
    """Not signed in, or the sign-in can no longer be refreshed."""


NOT_SIGNED_IN = (
    "Not signed in to Microsoft 365. Run `outlook-graph-mcp-plus auth` in a terminal "
    "(with the same OUTLOOK_GRAPH_CLIENT_ID / OUTLOOK_GRAPH_TENANT_ID), then try again."
)


def _persistence(settings: Settings):
    settings.cache_dir.mkdir(parents=True, exist_ok=True, mode=0o700)
    location = str(settings.cache_dir / "token_cache.bin")
    try:
        return build_encrypted_persistence(location)
    except Exception as e:  # noqa: BLE001 - no OS secret store available
        if not settings.allow_plaintext_token_cache:
            raise AuthError(
                "No OS secret store is available for the token cache "
                f"({type(e).__name__}: {e}). Install/unlock one (e.g. libsecret on Linux), "
                "or set OUTLOOK_GRAPH_ALLOW_PLAINTEXT_TOKEN_CACHE=true to accept a "
                "plaintext cache file readable by your user account."
            ) from e
        persistence = FilePersistence(location)
        Path(location).touch(mode=0o600, exist_ok=True)
        return persistence


class TokenProvider:
    """Builds the MSAL app lazily and hands out access tokens."""

    def __init__(self, settings: Settings):
        self.settings = settings
        self._app: msal.PublicClientApplication | None = None
        self._account_file = settings.cache_dir / "account.json"

    @property
    def app(self) -> msal.PublicClientApplication:
        if self._app is None:
            cache = PersistedTokenCache(_persistence(self.settings))
            self._app = msal.PublicClientApplication(
                self.settings.client_id,
                authority=self.settings.authority,
                token_cache=cache,
            )
        return self._app

    # --- account selection -------------------------------------------------

    def _remember(self, result: dict) -> str:
        username = (result.get("id_token_claims") or {}).get("preferred_username", "")
        self.settings.cache_dir.mkdir(parents=True, exist_ok=True, mode=0o700)
        self._account_file.write_text(json.dumps({"username": username}))
        return username

    def _account(self):
        accounts = self.app.get_accounts()
        if not accounts:
            return None
        wanted = ""
        if self._account_file.exists():
            try:
                wanted = json.loads(self._account_file.read_text()).get("username", "")
            except (OSError, ValueError):
                wanted = ""
        for a in accounts:
            if wanted and a.get("username", "").lower() == wanted.lower():
                return a
        return accounts[0] if len(accounts) == 1 else None

    # --- used by the server -------------------------------------------------

    def access_token(self) -> str:
        """A valid access token, refreshed silently; AuthError if sign-in is needed."""
        account = self._account()
        if account is None:
            raise AuthError(NOT_SIGNED_IN)
        result = self.app.acquire_token_silent(SCOPES, account=account)
        if not result or "access_token" not in result:
            raise AuthError(NOT_SIGNED_IN)
        return result["access_token"]

    def status(self) -> dict:
        account = self._account()
        if account is None:
            return {"signed_in": False}
        try:
            self.access_token()
            ok = True
        except AuthError:
            ok = False
        return {"signed_in": ok, "username": account.get("username", "")}

    # --- used by the CLI ----------------------------------------------------

    def login(self, device_code: bool = False, echo: Callable[[str], None] = print) -> str:
        """Interactive sign-in. Returns the signed-in username."""
        result = None
        if not device_code:
            try:
                result = self.app.acquire_token_interactive(
                    SCOPES, prompt="select_account", timeout=300
                )
            except Exception as e:  # noqa: BLE001 - no browser / no display
                echo(f"Browser sign-in not possible ({type(e).__name__}); using a device code.")
                result = None
        if result is None or "access_token" not in result:
            if result is not None and not device_code:
                echo(f"Browser sign-in failed: {result.get('error_description', result)}")
            flow = self.app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                raise AuthError(f"Could not start device code sign-in: {flow}")
            echo(flow["message"])
            result = self.app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise AuthError(
                "Sign-in failed: "
                + str(result.get("error_description") or result.get("error") or result)
            )
        return self._remember(result)

    def logout(self) -> int:
        """Remove all cached accounts and tokens. Returns how many were removed."""
        accounts = self.app.get_accounts()
        for a in accounts:
            self.app.remove_account(a)
        try:
            self._account_file.unlink()
        except FileNotFoundError:
            pass
        return len(accounts)


def eprint(msg: str) -> None:
    print(msg, file=sys.stderr, flush=True)
