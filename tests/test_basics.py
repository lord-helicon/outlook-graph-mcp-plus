"""Config, input validation, attachment path rules and the Graph client."""

import os
from pathlib import Path

import httpx
import pytest
from conftest import FakeGraph, run

from outlook_graph_mcp import util
from outlook_graph_mcp.config import ConfigError, load_settings
from outlook_graph_mcp.graph import GraphClient, GraphError, seg

GUID = "00000000-0000-0000-0000-000000000001"


# --- config ----------------------------------------------------------------


def test_config_requires_client_and_org_tenant():
    with pytest.raises(ConfigError, match="CLIENT_ID"):
        load_settings({"OUTLOOK_GRAPH_TENANT_ID": "contoso.com"})
    for bad in ("", "common", "consumers"):
        with pytest.raises(ConfigError, match="tenant"):
            load_settings({"OUTLOOK_GRAPH_CLIENT_ID": GUID, "OUTLOOK_GRAPH_TENANT_ID": bad})
    with pytest.raises(ConfigError):
        load_settings({"OUTLOOK_GRAPH_CLIENT_ID": GUID, "OUTLOOK_GRAPH_TENANT_ID": "a/b"})
    s = load_settings(
        {
            "OUTLOOK_GRAPH_CLIENT_ID": GUID,
            "OUTLOOK_GRAPH_TENANT_ID": "contoso.com",
            "OUTLOOK_GRAPH_TIMEZONE": "Europe/Berlin",
        }
    )
    assert s.authority == "https://login.microsoftonline.com/contoso.com"
    assert s.timezone == "Europe/Berlin" and s.attachment_dir.name == "Downloads"


# --- validation ------------------------------------------------------------


def test_parse_when_and_counts():
    assert util.parse_when("2026-12-01 14:00", "start")[0].hour == 14
    assert util.parse_when("01.12.2026", "start")[0].month == 12
    dt, problem = util.parse_when("next tuesday", "start")
    assert dt is None and "YYYY-MM-DD" in problem
    assert util.bad_count("count", 0) and util.bad_count("count", True)
    assert util.bad_count("count", 101, 1, 100) and util.bad_count("count", 5) is None


@pytest.mark.parametrize("bad", ["", "a b", "../x", "x?y=1", "a\nb", "x" * 600])
def test_bad_ids_refused(bad):
    assert util.bad_id(bad)


def test_good_id_accepted():
    assert util.bad_id("AAMkAGI2THVSAAA=") is None
    assert util.bad_id("AQMkAD-_ab+/cd==") is None


def test_seg_encodes_path_separators():
    assert seg("a/b?c#d e") == "a%2Fb%3Fc%23d%20e"


def test_addresses():
    assert util.split_addresses("a@x.com; b@y.org,") == ["a@x.com", "b@y.org"]
    assert util.bad_addresses("a@x.com", "not-an-address")
    assert util.bad_addresses("a@x.com; b@y.org", "") is None


# --- attachment paths ------------------------------------------------------


def test_save_dir_confined(tmp_path: Path, monkeypatch):
    root = tmp_path / "root"
    root.mkdir()
    calls = []
    real = os.path.realpath
    monkeypatch.setattr(util.os.path, "realpath", lambda p: (calls.append(p), real(p))[1])
    for bad in (r"\\attacker.example\share", "//attacker.example/share", r"\\?\UNC\x\y"):
        calls.clear()
        d, problem = util.resolve_save_dir(bad, root)
        assert d is None and "network" in problem
        assert all("attacker" not in str(c) and "UNC" not in str(c) for c in calls)
    assert util.resolve_save_dir("", root)[0] == Path(real(root))
    assert util.resolve_save_dir("sub", root)[0] == Path(real(root)) / "sub"
    for bad in (str(tmp_path), "../escape", str(tmp_path / "other")):
        assert util.resolve_save_dir(bad, root)[0] is None
    (root / "link").symlink_to(tmp_path)
    assert util.resolve_save_dir("link", root)[0] is None


def test_safe_filename_and_unique(tmp_path: Path):
    assert util.safe_filename("../../evil.sh") == "evil.sh"
    assert util.safe_filename("a\\b\\c.txt") == "c.txt"
    assert util.safe_filename("..") == "attachment"
    assert util.safe_filename("CON.txt") == "_CON.txt"
    assert util.safe_filename("x:y?.txt") == "x_y_.txt"
    (tmp_path / "r.pdf").write_text("x")
    assert util.unique_path(tmp_path, "r.pdf").name == "r (1).pdf"


# --- Graph client ----------------------------------------------------------


def _client(fg: FakeGraph) -> GraphClient:
    return GraphClient(
        lambda: "tok", timezone="Europe/Berlin", transport=httpx.MockTransport(fg.handler)
    )


def test_headers_and_errors():
    fg = FakeGraph().on("GET", "/me", {"ok": 1})
    c = _client(fg)
    assert run(c.get("/me")) == {"ok": 1}
    h = fg.last()["headers"]
    assert h["authorization"] == "Bearer tok"
    assert 'outlook.timezone="Europe/Berlin"' in h["prefer"]
    fg.on(
        "GET",
        "/x",
        {"error": {"code": "ErrorInvalidIdMalformed", "message": "Id is malformed."}},
        status=400,
    )
    with pytest.raises(GraphError, match="not valid for this mailbox"):
        run(c.get("/x"))
    for status, text in ((404, "Not found"), (403, "Access denied"), (401, "rejected the sign-in")):
        fg.on("GET", "/x", {"error": {"code": "E", "message": "m"}}, status=status)
        with pytest.raises(GraphError, match=text):
            run(c.get("/x"))


def test_retry_once_on_429(monkeypatch):
    sleeps = []

    async def no_sleep(s):
        sleeps.append(s)

    monkeypatch.setattr("outlook_graph_mcp.graph.asyncio.sleep", no_sleep)
    attempts = []

    def flaky(call):
        attempts.append(1)
        return (429, {"error": {}}) if len(attempts) == 1 else (200, {"ok": True})

    fg = FakeGraph().on("GET", "/me", flaky, headers={"Retry-After": "120"})
    assert run(_client(fg).get("/me")) == {"ok": True}
    assert len(attempts) == 2 and sleeps == [10.0]


def test_download_never_overwrites_and_cleans_up(tmp_path: Path):
    fg = FakeGraph().on("GET", "/blob", b"x" * 100)
    c = _client(fg)
    target = tmp_path / "f.bin"
    assert run(c.download("/blob", target, 1000)) == 100
    with pytest.raises(FileExistsError):
        run(c.download("/blob", target, 1000))
    too_big = tmp_path / "big.bin"
    with pytest.raises(GraphError, match="larger"):
        run(c.download("/blob", too_big, 10))
    assert not too_big.exists()
