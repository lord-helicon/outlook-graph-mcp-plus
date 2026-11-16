"""Shared mailboxes: configuration, allowlist and request paths."""

from dataclasses import replace

import pytest
from conftest import as_json, run

from outlook_graph_mcp import app
from outlook_graph_mcp.config import SCOPES, SHARED_SCOPES, ConfigError, load_settings
from outlook_graph_mcp.tools import calendar, mail, tasks
from outlook_graph_mcp.tools.common import mailbox_base

GUID = "00000000-0000-0000-0000-000000000001"
BASE_ENV = {"OUTLOOK_GRAPH_CLIENT_ID": GUID, "OUTLOOK_GRAPH_TENANT_ID": "contoso.com"}
SHARED = "info@contoso.com"


def _with_shared(fake, *boxes):
    app._state["settings"] = replace(fake.settings, shared_mailboxes=tuple(boxes))


def test_config_parses_shared_mailboxes_and_scopes():
    s = load_settings(BASE_ENV)
    assert s.shared_mailboxes == () and s.scopes == SCOPES
    s = load_settings(
        {**BASE_ENV, "OUTLOOK_GRAPH_SHARED_MAILBOXES": "Info@Contoso.com; support@contoso.com,"}
    )
    assert s.shared_mailboxes == ("info@contoso.com", "support@contoso.com")
    assert s.scopes == SCOPES + SHARED_SCOPES
    with pytest.raises(ConfigError, match="not an email"):
        load_settings({**BASE_ENV, "OUTLOOK_GRAPH_SHARED_MAILBOXES": "info"})


def test_mailbox_allowlist(fake):
    assert mailbox_base("") == "/me" and mailbox_base("me") == "/me"
    with pytest.raises(Exception, match="none configured"):
        mailbox_base(SHARED)
    _with_shared(fake, SHARED)
    assert mailbox_base("INFO@contoso.com") == "/users/info%40contoso.com"
    with pytest.raises(Exception, match="not one of the shared mailboxes"):
        mailbox_base("ceo@contoso.com")


def test_unlisted_mailbox_refused_without_request(fake):
    _with_shared(fake, SHARED)
    out = run(mail.list_emails(mailbox="ceo@contoso.com"))
    assert out.startswith("Error:") and "list_mailboxes" in out
    out = run(calendar.list_events(mailbox="../me"))
    assert out.startswith("Error:")
    assert fake.calls == []


def test_mail_in_shared_mailbox(fake):
    _with_shared(fake, SHARED)
    msg = {"id": "M1=", "subject": "Hi", "from": {"emailAddress": {"address": "x@y.com"}}}
    fake.on("GET", f"/users/{SHARED}/mailFolders/inbox/messages", {"value": [msg]})
    res = as_json(run(mail.list_emails(mailbox=SHARED)))
    assert res[0]["entry_id"] == "M1="
    assert fake.last()["raw_path"].startswith("/v1.0/users/info%40contoso.com/")
    fake.on("PATCH", f"/users/{SHARED}/messages/M1=", {"subject": "Hi"})
    run(mail.mark_as_read("M1=", mailbox=SHARED))
    assert fake.last()["path"] == f"/users/{SHARED}/messages/M1="
    fake.on("GET", f"/users/{SHARED}/mailFolders", {"value": [{"id": "S1", "displayName": "Done"}]})
    fake.on("POST", f"/users/{SHARED}/messages/M1=/move", {"id": "M2=", "subject": "Hi"})
    run(mail.move_email("M1=", "Done", mailbox=SHARED))
    assert fake.last()["json"] == {"destinationId": "S1"}


def test_send_from_shared_mailbox(fake):
    _with_shared(fake, SHARED)
    fake.on("POST", f"/users/{SHARED}/sendMail", None, status=202)
    res = as_json(run(mail.send_email(to="a@x.com", subject="S", body="b", mailbox=SHARED)))
    assert res["from"] == SHARED and fake.last()["path"] == f"/users/{SHARED}/sendMail"


def test_calendar_in_shared_mailbox(fake):
    _with_shared(fake, SHARED)
    fake.on("GET", f"/users/{SHARED}/calendarView", {"value": []})
    run(calendar.list_events("2026-12-01", "2026-12-02", mailbox=SHARED))
    assert fake.last()["path"] == f"/users/{SHARED}/calendarView"
    ev = {
        "id": "E1=",
        "subject": "S",
        "start": {"dateTime": "2026-12-01T10:00:00"},
        "end": {"dateTime": "2026-12-01T11:00:00"},
    }
    fake.on("POST", f"/users/{SHARED}/events", ev)
    run(calendar.create_event("S", "2026-12-01 10:00", "2026-12-01 11:00", mailbox=SHARED))
    assert fake.last()["path"] == f"/users/{SHARED}/events"


def test_list_mailboxes(fake):
    fake.on("GET", "/me", {"mail": "me@contoso.com"})
    assert as_json(run(mail.list_mailboxes())) == [{"mailbox": "me@contoso.com", "kind": "own"}]
    _with_shared(fake, SHARED)
    assert as_json(run(mail.list_mailboxes()))[1] == {"mailbox": SHARED, "kind": "shared"}


def test_tasks_stay_on_own_mailbox(fake):
    assert "mailbox" not in tasks.list_tasks.__signature__.parameters


def test_exchange_internal_address_hidden():
    from outlook_graph_mcp.tools.common import address

    x500 = {
        "emailAddress": {
            "address": "/O=EXCHANGELABS/OU=X/CN=RECIPIENTS/CN=AB-DOMAINS",
            "name": "Domains",
        }
    }
    assert address(x500) == ("", "Domains")
    assert address({"emailAddress": {"address": "a@b.com", "name": "A"}}) == ("a@b.com", "A")
