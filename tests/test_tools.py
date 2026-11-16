"""The MCP tools against a fake Microsoft Graph."""

import json
import sys
from pathlib import Path

from conftest import as_json, run

from outlook_graph_mcp import app
from outlook_graph_mcp.auth import NOT_SIGNED_IN, AuthError
from outlook_graph_mcp.tools import calendar, mail, tasks

MSG = {
    "id": "AAMk1=",
    "subject": "Hello",
    "from": {"emailAddress": {"address": "a@x.com", "name": "Al"}},
    "receivedDateTime": "2026-11-16T10:00:00Z",
    "isRead": False,
    "hasAttachments": True,
}
FOLDERS = {
    "value": [
        {"id": "F1", "displayName": "Inbox", "childFolderCount": 1},
        {"id": "F2", "displayName": "Projects", "childFolderCount": 0},
        {"id": "F4", "displayName": "Archive", "childFolderCount": 0},
    ]
}
CHILDREN = {"value": [{"id": "F3", "displayName": "Projects", "childFolderCount": 0}]}


# --- mail --------------------------------------------------------------------


def test_list_emails_unread_filter_and_shape(fake):
    fake.on("GET", "/me/mailFolders/inbox/messages", {"value": [MSG]})
    res = as_json(run(mail.list_emails(count=5, unread_only=True)))
    assert res == [
        {
            "entry_id": "AAMk1=",
            "subject": "Hello",
            "sender": "a@x.com",
            "sender_name": "Al",
            "received_time": "2026-11-16T10:00:00",
            "unread": True,
            "has_attachments": True,
        }
    ]
    p = fake.last()["params"]
    assert p["$top"] == "5" and p["$filter"].startswith("receivedDateTime ge")
    assert "isRead eq false" in p["$filter"] and p["$orderby"] == "receivedDateTime desc"


def test_well_known_and_custom_folders(fake):
    fake.on("GET", "/me/mailFolders/sentitems/messages", {"value": []})
    assert as_json(run(mail.list_emails(folder="sent"))) == []
    fake.on("GET", "/me/mailFolders", FOLDERS)
    fake.on("GET", "/me/mailFolders/F1/childFolders", CHILDREN)
    fake.on("GET", "/me/mailFolders/F3/messages", {"value": []})
    run(mail.list_emails(folder="Inbox/Projects"))
    assert fake.last()["path"] == "/me/mailFolders/F3/messages"
    out = run(mail.list_emails(folder="Projects"))
    assert out.startswith("Error: Several folders") and "Inbox/Projects" in out
    assert "No mail folder" in run(mail.list_emails(folder="Nope"))
    fake.on("GET", "/me/mailFolders/F2/messages", {"value": []})
    run(mail.list_emails(folder="/Projects"))
    assert fake.last()["path"] == "/me/mailFolders/F2/messages"


def test_counts_validated(fake):
    for out in (
        run(mail.list_emails(count=0)),
        run(mail.search_emails(query="x", count=-1)),
        run(mail.list_drafts(offset=-1)),
        run(calendar.list_events(count=0)),
        run(tasks.list_tasks(count=0)),
        run(mail.save_attachment("AAMk1=", 0)),
    ):
        assert out.startswith("Error:") and "whole number" in out
    assert fake.calls == []


def test_ids_are_validated_and_encoded(fake):
    assert run(mail.read_email(entry_id="../me/sendMail")).startswith("Error:")
    assert fake.calls == []
    fake.on("GET", "/me/messages/a/b=", {**MSG, "body": {"content": "hi"}})
    run(mail.read_email(entry_id="a/b="))
    assert "/me/messages/a%2Fb%3D" in fake.last()["raw_path"]


def test_read_email_by_id_and_search(fake):
    fake.on(
        "GET",
        "/me/messages/AAMk1=",
        {
            **MSG,
            "toRecipients": [{"emailAddress": {"address": "me@x.com"}}],
            "body": {"content": "Body text"},
        },
    )
    res = as_json(run(mail.read_email(entry_id="AAMk1=")))
    assert res["body"] == "Body text" and res["to"] == "me@x.com"
    fake.on("GET", "/me/mailFolders/inbox/messages", {"value": []})
    out = run(mail.read_email(subject_search='He said "hi"'))
    assert out.startswith("Error: No email")
    assert fake.last()["params"]["$search"] == '"subject:He said  hi"'


def test_mark_read_unread_and_move(fake):
    fake.on("PATCH", "/me/messages/AAMk1=", {"subject": "Hello"})
    assert as_json(run(mail.mark_as_unread("AAMk1=")))["status"] == "marked_unread"
    assert fake.last()["json"] == {"isRead": False}
    assert as_json(run(mail.mark_as_read("AAMk1=")))["status"] == "marked_read"
    assert fake.last()["json"] == {"isRead": True}
    fake.on("POST", "/me/messages/AAMk1=/move", {"id": "NEW=", "subject": "Hello"})
    res = as_json(run(mail.move_email("AAMk1=", "junk")))
    assert res["entry_id"] == "NEW=" and fake.last()["json"] == {"destinationId": "junkemail"}


def test_send_email_and_drafts(fake):
    fake.on("POST", "/me/sendMail", None, status=202)
    res = as_json(run(mail.send_email(to="a@x.com; b@y.org", subject="S", body="<b>plain</b>")))
    assert res["status"] == "sent"
    msg = fake.last()["json"]["message"]
    assert msg["body"] == {"contentType": "Text", "content": "<b>plain</b>"}
    assert [r["emailAddress"]["address"] for r in msg["toRecipients"]] == ["a@x.com", "b@y.org"]
    run(mail.send_email(to="a@x.com", subject="S", body="", html_body="<p>x</p>"))
    assert fake.last()["json"]["message"]["body"]["contentType"] == "HTML"
    assert run(mail.send_email(to="", subject="S", body="b")).startswith("Error:")
    assert run(mail.send_email(to="nope", subject="S", body="b")).startswith("Error:")

    fake.on("POST", "/me/messages", {"id": "D1="})
    assert as_json(run(mail.save_draft(to="a@x.com", subject="S", body="b")))["entry_id"] == "D1="
    fake.on("GET", "/me/messages/D1=", {"id": "D1=", "isDraft": False})
    assert "not an unsent draft" in run(mail.send_draft("D1="))
    fake.on(
        "GET",
        "/me/messages/D1=",
        {
            "id": "D1=",
            "isDraft": True,
            "subject": "S",
            "toRecipients": [{"emailAddress": {"address": "a@x.com"}}],
        },
    )
    fake.on("POST", "/me/messages/D1=/send", None, status=202)
    assert as_json(run(mail.send_draft("D1=")))["status"] == "sent"


def test_replies(fake):
    fake.on("POST", "/me/messages/AAMk1=/replyAll", None, status=202)
    run(mail.reply_email("AAMk1=", "thanks", reply_all=True))
    assert fake.last()["json"] == {"comment": "thanks"}
    fake.on("POST", "/me/messages/AAMk1=/createReply", {"id": "R1=", "subject": "RE: Hello"})
    assert as_json(run(mail.create_draft_reply("AAMk1=", "draft text")))["entry_id"] == "R1="


ATTS = {
    "value": [
        {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "id": "A1",
            "name": "../../evil.sh",
            "contentType": "text/plain",
            "size": 5,
            "isInline": False,
        },
        {"@odata.type": "#microsoft.graph.itemAttachment", "id": "A2", "name": "fwd", "size": 9},
    ]
}


def test_attachments_saved_confined(fake):
    fake.on("GET", "/me/messages/AAMk1=/attachments", ATTS)
    fake.on("GET", "/me/messages/AAMk1=/attachments/A1/\\$value", b"hello")
    listed = as_json(run(mail.list_attachments("AAMk1=")))
    assert [a["kind"] for a in listed] == ["file", "item"] and listed[0]["index"] == 1
    res = as_json(run(mail.save_attachment("AAMk1=", 1)))
    root = fake.settings.attachment_dir.resolve()
    assert Path(res["path"]).parent == root and res["filename"] == "evil.sh" and res["size"] == 5
    res2 = as_json(run(mail.save_attachment("AAMk1=", 1)))
    assert res2["filename"] == "evil (1).sh"
    if sys.platform == "darwin":
        assert res["marked_as_downloaded"] is True
    assert "cannot be saved" in run(mail.save_attachment("AAMk1=", 2))
    assert "too high" in run(mail.save_attachment("AAMk1=", 3))
    assert "inside the attachment folder" in run(mail.save_attachment("AAMk1=", 1, "/tmp"))
    sub = as_json(run(mail.save_attachment("AAMk1=", 1, "inv/2026")))
    assert Path(sub["path"]).parent == root / "inv" / "2026"


def test_auth_errors_become_tool_errors(fake, monkeypatch):
    def boom():
        raise AuthError(NOT_SIGNED_IN)

    monkeypatch.setattr(app.graph(), "_token", boom)
    out = run(mail.list_emails())
    assert out.startswith("Error: Not signed in") and "outlook-graph-mcp-plus auth" in out


# --- calendar ----------------------------------------------------------------

EVENT = {
    "id": "E1=",
    "subject": "Standup",
    "start": {"dateTime": "2026-12-01T10:00:00.0000000"},
    "end": {"dateTime": "2026-12-01T10:15:00.0000000"},
    "location": {"displayName": "Room"},
    "organizer": {"emailAddress": {"address": "a@x.com"}},
    "isAllDay": False,
}


def test_list_events_range_in_utc(fake):
    fake.on("GET", "/me/calendarView", {"value": [EVENT]})
    res = as_json(run(calendar.list_events("2026-12-01", "2026-12-01")))
    assert res[0]["start"] == "2026-12-01T10:00:00" and res[0]["location"] == "Room"
    p = fake.last()["params"]
    # Berlin is UTC+1 in December; a bare end date includes the whole day.
    assert (
        p["startDateTime"] == "2026-11-30T23:00:00Z" and p["endDateTime"] == "2026-12-01T23:00:00Z"
    )
    assert run(calendar.list_events("2026-12-10", "2026-12-01")).startswith("Error: end_date")
    assert "YYYY-MM-DD" in run(calendar.list_events("someday"))


def test_search_events_filters_subject(fake):
    other = {**EVENT, "id": "E2=", "subject": "Lunch"}
    fake.on("GET", "/me/calendarView", {"value": [EVENT, other]})
    res = as_json(run(calendar.search_events("stand", "2026-11-30", "2026-12-05")))
    assert [e["entry_id"] for e in res] == ["E1="]


def test_create_events_and_meetings(fake):
    fake.on("POST", "/me/events", lambda call: (201, {**EVENT, **call["json"]}))
    as_json(
        run(calendar.create_event("S", "2026-12-01 10:00", "2026-12-01 11:00", body="<i>x</i>"))
    )
    sent = fake.last()["json"]
    assert sent["start"] == {"dateTime": "2026-12-01T10:00:00", "timeZone": "Europe/Berlin"}
    assert sent["body"] == {"contentType": "Text", "content": "<i>x</i>"}
    assert sent["isReminderOn"] and sent["reminderMinutesBeforeStart"] == 15
    run(calendar.create_event("S", "2026-12-02", "2026-12-02", all_day=True, reminder_minutes=0))
    sent = fake.last()["json"]
    assert sent["isAllDay"] and sent["end"]["dateTime"] == "2026-12-03T00:00:00"
    assert sent["isReminderOn"] is False
    assert run(calendar.create_event("S", "2026-12-01 12:00", "2026-12-01 11:00")).startswith(
        "Error: end is before start"
    )
    assert run(
        calendar.create_event("S", "2026-12-01 10:00", "2026-12-01 11:00", reminder_minutes=-5)
    ).startswith("Error:")
    run(
        calendar.create_meeting(
            "M",
            "2026-12-01 10:00",
            "2026-12-01 11:00",
            "a@x.com; b@y.org",
            optional_attendees="c@z.com",
        )
    )
    types = [(a["emailAddress"]["address"], a["type"]) for a in fake.last()["json"]["attendees"]]
    assert types == [("a@x.com", "required"), ("b@y.org", "required"), ("c@z.com", "optional")]
    assert run(calendar.create_meeting("M", "2026-12-01 10:00", "2026-12-01 11:00", "")).startswith(
        "Error:"
    )


def test_update_and_delete_event(fake):
    fake.on("PATCH", "/me/events/E1=", lambda call: (200, {**EVENT, **call["json"]}))
    as_json(run(calendar.update_event("E1=", location="Room 2")))
    assert fake.last()["json"] == {"location": {"displayName": "Room 2"}}
    assert run(calendar.update_event("E1=")).startswith("Error: Nothing")
    assert run(
        calendar.update_event("E1=", start="2026-12-01 12:00", end="2026-12-01 11:00")
    ).startswith("Error:")
    fake.on("GET", "/me/events/E1=", EVENT)
    fake.on("DELETE", "/me/events/E1=", None, status=204)
    assert as_json(run(calendar.delete_event("E1=")))["subject"] == "Standup"


# --- tasks -------------------------------------------------------------------


def test_tasks(fake):
    fake.on(
        "GET",
        "/me/todo/lists",
        {
            "value": [
                {"id": "L0", "wellknownListName": "none"},
                {"id": "L1", "wellknownListName": "defaultList"},
            ]
        },
    )
    task = {
        "id": "T1",
        "title": "Do it",
        "status": "notStarted",
        "importance": "high",
        "dueDateTime": {"dateTime": "2026-12-05T00:00:00.0000000"},
    }
    fake.on("GET", "/me/todo/lists/L1/tasks", {"value": [task]})
    res = as_json(run(tasks.list_tasks()))
    assert res == [
        {
            "entry_id": "T1",
            "subject": "Do it",
            "due_date": "2026-12-05",
            "complete": False,
            "priority": "high",
        }
    ]
    assert fake.last()["params"]["$filter"] == "status ne 'completed'"
    fake.on("POST", "/me/todo/lists/L1/tasks", lambda call: (201, {**task, **call["json"]}))
    run(tasks.create_task("New", due_date="05.12.2026", importance="HIGH"))
    sent = fake.last()["json"]
    assert sent["importance"] == "high" and sent["dueDateTime"]["dateTime"] == "2026-12-05T00:00:00"
    assert run(tasks.create_task("x", importance="urgent")).startswith("Error: importance")
    assert run(tasks.create_task("x", due_date="someday")).startswith("Error: due_date")
    fake.on("PATCH", "/me/todo/lists/L1/tasks/T1", {"title": "Do it", "status": "completed"})
    assert as_json(run(tasks.complete_task("T1")))["status"] == "completed"
    lookups = [c for c in fake.calls if c["path"] == "/me/todo/lists"]
    assert len(lookups) == 1  # default list looked up once


def test_server_reports_package_version():
    from outlook_graph_mcp import __version__

    assert app.mcp._mcp_server.version == __version__
    assert json.loads(json.dumps(__version__))
