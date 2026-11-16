"""Calendar tools: list, read, create (events and meetings), update, delete, search."""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from outlook_graph_mcp.app import graph, settings, tool
from outlook_graph_mcp.graph import seg
from outlook_graph_mcp.tools.common import address, addresses, get_all, mailbox_base
from outlook_graph_mcp.util import (
    bad_addresses,
    bad_count,
    bad_id,
    err,
    iso,
    out,
    parse_when,
    split_addresses,
    truncate,
)

_SELECT = "id,subject,start,end,location,organizer,isAllDay"


def _tz():
    try:
        return ZoneInfo(settings().timezone)
    except (ZoneInfoNotFoundError, ValueError):
        return UTC


def _utc(dt: datetime) -> str:
    """Naive local datetime (configured timezone) -> UTC ISO string for calendarView."""
    return dt.replace(tzinfo=_tz()).astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def _when(dt: datetime) -> dict:
    return {"dateTime": dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": settings().timezone}


def _summary(e: dict) -> dict:
    return {
        "entry_id": e.get("id", ""),
        "subject": e.get("subject") or "(no subject)",
        "start": iso((e.get("start") or {}).get("dateTime")),
        "end": iso((e.get("end") or {}).get("dateTime")),
        "location": (e.get("location") or {}).get("displayName", "") or "",
        "organizer": address(e.get("organizer"))[0],
        "all_day": bool(e.get("isAllDay")),
    }


def _range(start_date: str, end_date: str, default_back: int, default_ahead: int):
    """(start, end, None) as naive local datetimes, or (None, None, error)."""
    now = datetime.now(_tz()).replace(tzinfo=None, microsecond=0)
    if start_date:
        start, problem = parse_when(start_date, "start_date")
        if problem:
            return None, None, problem
    else:
        start = now - timedelta(days=default_back)
    if end_date:
        end, problem = parse_when(end_date, "end_date")
        if problem:
            return None, None, problem
        if ":" not in end_date:
            end = end + timedelta(days=1)  # a date without a time includes that whole day
    else:
        end = start + timedelta(days=default_ahead)
    if end <= start:
        return None, None, err("end_date is before start_date.")
    return start, end, None


async def _view(b: str, start: datetime, end: datetime, limit: int) -> list[dict]:
    return await get_all(
        f"{b}/calendarView",
        {
            "startDateTime": _utc(start),
            "endDateTime": _utc(end),
            "$top": min(limit, 250),
            "$orderby": "start/dateTime",
            "$select": _SELECT,
        },
        limit=limit,
    )


@tool
async def list_events(
    start_date: str = "", end_date: str = "", count: int = 20, mailbox: str = ""
) -> str:
    """List calendar events in a date range (recurring events expanded), earliest first.

    Args:
        start_date: Start, e.g. "2026-12-01" or "2026-12-01 08:00". Default: now.
        end_date: End (a bare date includes that whole day). Default: 7 days after start.
        count: Maximum events (1-500). Default 20.
        mailbox: Optional shared mailbox address (see list_mailboxes). Empty means
            your own mailbox. entry_ids belong to the mailbox they came from.

    Returns:
        JSON array of {entry_id, subject, start, end, location, organizer, all_day}.
    """
    b = mailbox_base(mailbox)
    if e := bad_count("count", count, 1, 500):
        return e
    start, end, problem = _range(start_date, end_date, 0, 7)
    if problem:
        return problem
    return out([_summary(e) for e in await _view(b, start, end, count)])


@tool
async def search_events(
    query: str, start_date: str = "", end_date: str = "", count: int = 10, mailbox: str = ""
) -> str:
    """Find calendar events whose subject contains `query` in a date range.

    Args:
        query: Text to look for in the subject (case-insensitive).
        start_date: Start of the range. Default: 30 days ago.
        end_date: End of the range. Default: 60 days after start.
        count: Maximum results (1-100). Default 10.
        mailbox: Optional shared mailbox address (see list_mailboxes). Empty means
            your own mailbox. entry_ids belong to the mailbox they came from.
    """
    b = mailbox_base(mailbox)
    if e := bad_count("count", count, 1, 100):
        return e
    start, end, problem = _range(start_date, end_date, 30, 60)
    if problem:
        return problem
    q = (query or "").strip().lower()
    events = await _view(b, start, end, 1000)
    return out([_summary(e) for e in events if q in (e.get("subject") or "").lower()][:count])


@tool
async def get_event(entry_id: str, mailbox: str = "") -> str:
    """Read one calendar event in full, including attendees and description.

    Args:
        entry_id: The event's entry_id.
        mailbox: Optional shared mailbox address (see list_mailboxes). Empty means
            your own mailbox. entry_ids belong to the mailbox they came from.
    """
    b = mailbox_base(mailbox)
    if e := bad_id(entry_id, "event"):
        return e
    ev = await graph().get(
        f"{b}/events/{seg(entry_id.strip())}",
        **{"$select": _SELECT + ",body,attendees,isOnlineMeeting,onlineMeeting"},
    )
    data = _summary(ev)
    data["attendees"] = addresses(ev.get("attendees"))
    data["body"] = truncate(((ev.get("body") or {}).get("content") or "").strip())
    data["online_meeting_url"] = ((ev.get("onlineMeeting") or {}).get("joinUrl")) or ""
    return out(data)


def _times(start: str, end: str, all_day: bool):
    s, problem = parse_when(start, "start")
    if problem:
        return None, None, problem
    e, problem = parse_when(end, "end")
    if problem:
        return None, None, problem
    if all_day:
        s = s.replace(hour=0, minute=0, second=0)
        e = e.replace(hour=0, minute=0, second=0)
        if e <= s:
            e = s + timedelta(days=1)  # Graph's all-day end is exclusive
    elif e < s:
        return None, None, err("end is before start.")
    return s, e, None


def _event_payload(subject, s, e, location, body, all_day) -> dict:
    payload = {"subject": subject or "", "start": _when(s), "end": _when(e), "isAllDay": all_day}
    if location:
        payload["location"] = {"displayName": location}
    if body:
        payload["body"] = {"contentType": "Text", "content": body}
    return payload


@tool
async def create_event(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    all_day: bool = False,
    reminder_minutes: int = 15,
    mailbox: str = "",
) -> str:
    """Create a personal calendar appointment (no invitations are sent).

    Args:
        subject: Title.
        start: Start, e.g. "2026-12-01 14:00" (all-day: "2026-12-01").
        end: End, e.g. "2026-12-01 15:00".
        location: Optional location.
        body: Optional plain-text description.
        all_day: All-day event. Default false.
        reminder_minutes: Minutes before the start for a reminder; 0 turns it off.
        mailbox: Optional shared mailbox address (see list_mailboxes). Empty means
            your own mailbox. entry_ids belong to the mailbox they came from.
    """
    b = mailbox_base(mailbox)
    if e := bad_count("reminder_minutes", reminder_minutes, 0, 40320):
        return e
    s, en, problem = _times(start, end, all_day)
    if problem:
        return problem
    payload = _event_payload(subject, s, en, location, body, all_day)
    payload["isReminderOn"] = reminder_minutes > 0
    if reminder_minutes > 0:
        payload["reminderMinutesBeforeStart"] = reminder_minutes
    ev = await graph().request("POST", f"{b}/events", json=payload)
    return out({"status": "created", **_summary(ev)})


@tool
async def create_meeting(
    subject: str,
    start: str,
    end: str,
    required_attendees: str,
    location: str = "",
    body: str = "",
    optional_attendees: str = "",
    mailbox: str = "",
) -> str:
    """Create a meeting and send invitations to the attendees.

    Args:
        subject: Title.
        start: Start, e.g. "2026-12-01 14:00".
        end: End, e.g. "2026-12-01 15:00".
        required_attendees: Addresses separated by ';' or ','.
        location: Optional location.
        body: Optional plain-text agenda.
        optional_attendees: Optional addresses separated by ';' or ','.
        mailbox: Optional shared mailbox address (see list_mailboxes). Empty means
            your own mailbox. entry_ids belong to the mailbox they came from.
    """
    b = mailbox_base(mailbox)
    if not split_addresses(required_attendees):
        return err("required_attendees needs at least one address.")
    if e := bad_addresses(required_attendees, optional_attendees):
        return e
    s, en, problem = _times(start, end, False)
    if problem:
        return problem
    payload = _event_payload(subject, s, en, location, body, False)
    payload["attendees"] = [
        {"emailAddress": {"address": a}, "type": "required"}
        for a in split_addresses(required_attendees)
    ] + [
        {"emailAddress": {"address": a}, "type": "optional"}
        for a in split_addresses(optional_attendees)
    ]
    ev = await graph().request("POST", f"{b}/events", json=payload)
    return out(
        {"status": "invitations_sent", **_summary(ev), "attendees": addresses(ev.get("attendees"))}
    )


@tool
async def update_event(
    entry_id: str,
    subject: str = "",
    start: str = "",
    end: str = "",
    location: str = "",
    body: str = "",
    mailbox: str = "",
) -> str:
    """Change an event. Only the fields you pass are changed.

    Args:
        entry_id: The event's entry_id.
        subject: New title.
        start: New start.
        end: New end.
        location: New location.
        body: New plain-text description.
        mailbox: Optional shared mailbox address (see list_mailboxes). Empty means
            your own mailbox. entry_ids belong to the mailbox they came from.
    """
    b = mailbox_base(mailbox)
    if e := bad_id(entry_id, "event"):
        return e
    patch: dict = {}
    s = en = None
    if start:
        s, problem = parse_when(start, "start")
        if problem:
            return problem
        patch["start"] = _when(s)
    if end:
        en, problem = parse_when(end, "end")
        if problem:
            return problem
        patch["end"] = _when(en)
    if s and en and en < s:
        return err("end is before start.")
    if subject:
        patch["subject"] = subject
    if location:
        patch["location"] = {"displayName": location}
    if body:
        patch["body"] = {"contentType": "Text", "content": body}
    if not patch:
        return err("Nothing to update: pass at least one field.")
    ev = await graph().request("PATCH", f"{b}/events/{seg(entry_id.strip())}", json=patch)
    return out({"status": "updated", **_summary(ev)})


@tool
async def delete_event(entry_id: str, mailbox: str = "") -> str:
    """Delete a calendar event.

    Args:
        entry_id: The event's entry_id.
        mailbox: Optional shared mailbox address (see list_mailboxes). Empty means
            your own mailbox. entry_ids belong to the mailbox they came from.
    """
    b = mailbox_base(mailbox)
    if e := bad_id(entry_id, "event"):
        return e
    path = f"{b}/events/{seg(entry_id.strip())}"
    ev = await graph().get(path, **{"$select": "subject"})
    await graph().request("DELETE", path)
    return out({"status": "deleted", "subject": ev.get("subject", "")})
