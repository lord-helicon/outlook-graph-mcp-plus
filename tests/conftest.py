"""A fake Microsoft Graph for tests: route requests, record what was sent."""

from __future__ import annotations

import asyncio
import json
import re
from pathlib import Path

import httpx
import pytest

from outlook_graph_mcp import app
from outlook_graph_mcp.config import Settings
from outlook_graph_mcp.graph import GraphClient


class FakeGraph:
    def __init__(self):
        self.routes: list[tuple[str, re.Pattern, object, int, dict]] = []
        self.calls: list[dict] = []

    def on(self, method: str, pattern: str, reply=None, status: int = 200, headers=None):
        """reply: dict/list (JSON), bytes, or callable(request_dict) -> (status, body)."""
        self.routes.insert(0, (method, re.compile(pattern), reply, status, headers or {}))
        return self

    def handler(self, request: httpx.Request) -> httpx.Response:
        path = request.url.path.removeprefix("/v1.0")
        body = json.loads(request.content) if request.content else None
        call = {
            "method": request.method,
            "path": path,
            "raw_path": request.url.raw_path.decode(),
            "params": dict(request.url.params),
            "json": body,
            "headers": dict(request.headers),
        }
        self.calls.append(call)
        for method, pattern, reply, status, headers in self.routes:
            if method == request.method and pattern.fullmatch(path):
                if callable(reply):
                    status, reply = reply(call)
                if isinstance(reply, bytes):
                    return httpx.Response(status, content=reply, headers=headers)
                if reply is None:
                    return httpx.Response(status if status != 200 else 204, headers=headers)
                return httpx.Response(status, json=reply, headers=headers)
        return httpx.Response(404, json={"error": {"code": "ErrorItemNotFound", "message": "nope"}})

    def last(self, method: str | None = None) -> dict:
        calls = [c for c in self.calls if method is None or c["method"] == method]
        return calls[-1]


@pytest.fixture
def fake(tmp_path: Path):
    fg = FakeGraph()
    cfg = Settings(
        client_id="00000000-0000-0000-0000-000000000001",
        tenant_id="contoso.onmicrosoft.com",
        attachment_dir=tmp_path / "att",
        timezone="Europe/Berlin",
        cache_dir=tmp_path / "cache",
        allow_plaintext_token_cache=False,
    )
    (tmp_path / "att").mkdir()
    client = GraphClient(
        lambda: "test-token", timezone=cfg.timezone, transport=httpx.MockTransport(fg.handler)
    )
    app._state.clear()
    app.set_graph(client, cfg)
    fg.settings = cfg
    yield fg
    app._state.clear()


def run(coro):
    return asyncio.run(coro)


def as_json(text: str):
    assert not text.startswith("Error"), text
    return json.loads(text)
