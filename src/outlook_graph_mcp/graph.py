"""Small async Microsoft Graph client (v1.0) on httpx.

- Every id that goes into a URL path is percent-encoded with `seg()`, so an id
  coming from mailbox content can never add path segments or a query.
- Errors become GraphError with a short, readable message.
- One retry on 429/503, honouring Retry-After (capped).
"""

from __future__ import annotations

import asyncio
from collections.abc import Callable
from pathlib import Path
from typing import Any
from urllib.parse import quote

import httpx

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
_MAX_RETRY_AFTER = 10.0


class GraphError(RuntimeError):
    def __init__(self, message: str, status: int = 0, code: str = ""):
        super().__init__(message)
        self.status, self.code = status, code


def seg(value: str) -> str:
    """Percent-encode one URL path segment (ids, folder names)."""
    return quote(str(value), safe="")


def _message(resp: httpx.Response) -> tuple[str, str]:
    try:
        err = resp.json().get("error", {})
        return err.get("code", ""), err.get("message", "")
    except ValueError:
        return "", resp.text[:200]


def _friendly(resp: httpx.Response) -> GraphError:
    code, msg = _message(resp)
    s = resp.status_code
    if s == 401:
        text = "Microsoft 365 rejected the sign-in. Run `outlook-graph-mcp-plus auth` again."
    elif s == 403:
        text = (
            "Access denied by Microsoft 365. The app registration may be missing a "
            f"permission, or an admin has to grant consent. ({code or s}: {msg})"
        )
    elif code in ("ErrorInvalidIdMalformed", "ErrorInvalidId"):
        text = "The id is not valid for this mailbox. Use an entry_id returned by this server."
    elif s == 404 or code in ("ErrorItemNotFound", "ResourceNotFound"):
        text = "Not found: the item does not exist (or no longer exists) in this mailbox."
    elif s == 429:
        text = "Microsoft 365 is throttling requests; try again in a minute."
    else:
        text = f"Microsoft Graph error {s}{' ' + code if code else ''}: {msg}".strip()
    return GraphError(text, s, code)


class GraphClient:
    def __init__(
        self,
        token: Callable[[], str],
        timezone: str = "UTC",
        transport: httpx.AsyncBaseTransport | None = None,
        timeout: float = 60.0,
    ):
        self._token = token
        self.timezone = timezone
        self._client = httpx.AsyncClient(base_url=GRAPH_ROOT, transport=transport, timeout=timeout)

    async def _headers(self, extra: dict | None = None) -> dict:
        # MSAL is synchronous (and may refresh over the network): keep it off the loop.
        token = await asyncio.to_thread(self._token)
        h = {
            "Authorization": f"Bearer {token}",
            "Prefer": f'outlook.timezone="{self.timezone}", outlook.body-content-type="text"',
        }
        if extra:
            h.update(extra)
        return h

    async def request(
        self,
        method: str,
        path: str,
        *,
        params: dict | None = None,
        json: Any = None,
        headers: dict | None = None,
    ) -> Any:
        for attempt in (1, 2):
            resp = await self._client.request(
                method, path, params=params, json=json, headers=await self._headers(headers)
            )
            if resp.status_code in (429, 503) and attempt == 1:
                try:
                    wait = float(resp.headers.get("Retry-After", "2"))
                except ValueError:
                    wait = 2.0
                await asyncio.sleep(min(max(wait, 0.0), _MAX_RETRY_AFTER))
                continue
            break
        if resp.status_code >= 400:
            raise _friendly(resp)
        if resp.status_code == 204 or not resp.content:
            return None
        return resp.json()

    async def get(self, path: str, **params: Any) -> Any:
        return await self.request(
            "GET", path, params={k: v for k, v in params.items() if v is not None}
        )

    async def download(self, path: str, target: Path, max_bytes: int) -> int:
        """Stream a binary resource to `target` (created exclusively). Returns bytes written."""
        written = 0
        created = False
        try:
            async with self._client.stream("GET", path, headers=await self._headers()) as resp:
                if resp.status_code >= 400:
                    await resp.aread()
                    raise _friendly(resp)
                with open(target, "xb") as f:  # x: never overwrite
                    created = True
                    async for chunk in resp.aiter_bytes():
                        written += len(chunk)
                        if written > max_bytes:
                            raise GraphError(
                                f"Attachment is larger than the {max_bytes} byte limit."
                            )
                        f.write(chunk)
        except BaseException:
            if created:
                target.unlink(missing_ok=True)  # no half-written files left behind
            raise
        return written

    async def aclose(self) -> None:
        await self._client.aclose()
