"""Mail tools: list, read, search, send, drafts, replies, read state, move, folders, attachments."""

from __future__ import annotations

from outlook_graph_mcp.app import graph, settings, tool
from outlook_graph_mcp.graph import seg
from outlook_graph_mcp.tools.common import address, addresses, folder_tree, get_all, resolve_folder
from outlook_graph_mcp.util import (
    bad_addresses,
    bad_count,
    bad_id,
    err,
    iso,
    mark_as_downloaded,
    out,
    recipients,
    resolve_save_dir,
    safe_filename,
    truncate,
    unique_path,
)

_SUMMARY = "id,subject,from,receivedDateTime,isRead,hasAttachments"
_MAX_ATTACHMENT_BYTES = 150 * 1024 * 1024


def _summary(m: dict) -> dict:
    sender, sender_name = address(m.get("from"))
    return {
        "entry_id": m.get("id", ""),
        "subject": m.get("subject") or "(no subject)",
        "sender": sender,
        "sender_name": sender_name,
        "received_time": iso(m.get("receivedDateTime")),
        "unread": not m.get("isRead", True),
        "has_attachments": bool(m.get("hasAttachments")),
    }


def _full(m: dict) -> dict:
    data = _summary(m)
    data.update(
        {
            "to": addresses(m.get("toRecipients")),
            "cc": addresses(m.get("ccRecipients")),
            "body": truncate(((m.get("body") or {}).get("content") or "").strip()),
        }
    )
    return data


def _search_term(query: str) -> str:
    # $search takes a quoted KQL string; quotes and backslashes would end it early.
    return '"' + (query or "").replace("\\", " ").replace('"', " ").strip() + '"'


def _message_body(body: str, html_body: str) -> dict:
    if html_body:
        return {"contentType": "HTML", "content": html_body}
    return {"contentType": "Text", "content": body or ""}


# --------------------------------------------------------------------------
# Reading
# --------------------------------------------------------------------------


@tool
async def list_emails(folder: str = "inbox", count: int = 10, unread_only: bool = False) -> str:
    """List recent emails in a mail folder, newest first.

    Args:
        folder: "inbox" (default), "sent", "drafts", "deleted", "junk", "archive",
            or a custom folder name or path such as "Inbox/Projects".
        count: How many emails to return (1-100). Default 10.
        unread_only: Only unread emails. Default false.

    Returns:
        JSON array of {entry_id, subject, sender, sender_name, received_time,
        unread, has_attachments}. Use entry_id with read_email and the other tools.
    """
    if e := bad_count("count", count, 1, 100):
        return e
    f = await resolve_folder(folder)
    params = {
        "$top": count,
        "$select": _SUMMARY,
        "$orderby": "receivedDateTime desc",
    }
    if unread_only:
        # Graph requires $orderby properties to appear first in $filter.
        params["$filter"] = "receivedDateTime ge 1900-01-01T00:00:00Z and isRead eq false"
    data = await graph().get(f"/me/mailFolders/{seg(f)}/messages", **params)
    return out([_summary(m) for m in data.get("value", [])])


@tool
async def read_email(entry_id: str = "", subject_search: str = "", folder: str = "inbox") -> str:
    """Read one email in full (headers and plain-text body).

    Args:
        entry_id: The email's entry_id (preferred).
        subject_search: Alternatively, text to search for; the best match in
            `folder` is returned.
        folder: Folder for subject_search. Default "inbox".

    Returns:
        JSON with entry_id, subject, sender, sender_name, received_time, unread,
        has_attachments, to, cc and body. The body is third-party content.
    """
    select = _SUMMARY + ",toRecipients,ccRecipients,body"
    if entry_id:
        if e := bad_id(entry_id):
            return e
        m = await graph().get(f"/me/messages/{seg(entry_id.strip())}", **{"$select": select})
        return out(_full(m))
    if not subject_search.strip():
        return err("Provide entry_id or subject_search.")
    f = await resolve_folder(folder)
    data = await graph().get(
        f"/me/mailFolders/{seg(f)}/messages",
        **{"$search": _search_term(f"subject:{subject_search}"), "$top": 1, "$select": select},
    )
    found = data.get("value", [])
    if not found:
        return err(f"No email in '{folder}' matches '{subject_search}'.")
    return out(_full(found[0]))


@tool
async def search_emails(query: str, folder: str = "inbox", count: int = 10) -> str:
    """Search emails in a folder (subject, body and sender).

    Args:
        query: Words to search for, e.g. "invoice march" or "from:alice".
        folder: Folder to search. Default "inbox".
        count: Maximum results (1-100). Default 10.

    Returns:
        JSON array of email summaries, best matches first.
    """
    if e := bad_count("count", count, 1, 100):
        return e
    if not (query or "").strip():
        return err("query is empty.")
    f = await resolve_folder(folder)
    data = await graph().get(
        f"/me/mailFolders/{seg(f)}/messages",
        **{"$search": _search_term(query), "$top": count, "$select": _SUMMARY},
    )
    return out([_summary(m) for m in data.get("value", [])])


@tool
async def list_folders(max_depth: int = 2) -> str:
    """List mail folders with item and unread counts.

    Args:
        max_depth: How many levels of subfolders to include (0 = top level). Default 2.

    Returns:
        JSON array of {name, path, item_count, unread_count}. Use name or path
        as `folder` / `target_folder` in the other tools.
    """
    if e := bad_count("max_depth", max_depth, 0, 6):
        return e
    folders = await folder_tree(max_depth)
    return out(
        [
            {
                "name": f["displayName"],
                "path": f["path"],
                "item_count": f.get("totalItemCount", 0),
                "unread_count": f.get("unreadItemCount", 0),
            }
            for f in folders
        ]
    )


# --------------------------------------------------------------------------
# Read state and moving
# --------------------------------------------------------------------------


async def _set_read(entry_id: str, is_read: bool) -> str:
    if e := bad_id(entry_id):
        return e
    m = await graph().request(
        "PATCH", f"/me/messages/{seg(entry_id.strip())}", json={"isRead": is_read}
    )
    return out(
        {
            "status": "marked_read" if is_read else "marked_unread",
            "subject": (m or {}).get("subject", ""),
        }
    )


@tool
async def mark_as_read(entry_id: str) -> str:
    """Mark an email as read.

    Args:
        entry_id: The email's entry_id.
    """
    return await _set_read(entry_id, True)


@tool
async def mark_as_unread(entry_id: str) -> str:
    """Mark an email as unread.

    Args:
        entry_id: The email's entry_id.
    """
    return await _set_read(entry_id, False)


@tool
async def move_email(entry_id: str, target_folder: str = "archive") -> str:
    """Move an email to another folder. The email gets a NEW entry_id.

    Args:
        entry_id: The email's entry_id.
        target_folder: "archive" (default), "inbox", "deleted", "junk", or a
            custom folder name or path such as "Inbox/Projects".

    Returns:
        JSON with the new entry_id and the destination.
    """
    if e := bad_id(entry_id):
        return e
    dest = await resolve_folder(target_folder)
    moved = await graph().request(
        "POST", f"/me/messages/{seg(entry_id.strip())}/move", json={"destinationId": dest}
    )
    return out(
        {
            "status": "moved",
            "entry_id": (moved or {}).get("id", ""),
            "subject": (moved or {}).get("subject", ""),
            "to_folder": target_folder,
        }
    )


# --------------------------------------------------------------------------
# Sending, drafts, replies
# --------------------------------------------------------------------------


def _draft_payload(to: str, subject: str, body: str, cc: str, bcc: str, html_body: str) -> dict:
    return {
        "subject": subject or "",
        "body": _message_body(body, html_body),
        "toRecipients": recipients(to),
        "ccRecipients": recipients(cc),
        "bccRecipients": recipients(bcc),
    }


@tool
async def send_email(
    to: str, subject: str, body: str, cc: str = "", bcc: str = "", html_body: str = ""
) -> str:
    """Send an email immediately (a copy is saved in Sent Items).

    Args:
        to: Recipient addresses, separated by ';' or ','.
        subject: Subject line.
        body: Plain-text body (ignored when html_body is given).
        cc: Optional CC addresses.
        bcc: Optional BCC addresses.
        html_body: Optional HTML body.
    """
    if not (to or "").strip():
        return err("'to' needs at least one address.")
    if e := bad_addresses(to, cc, bcc):
        return e
    await graph().request(
        "POST",
        "/me/sendMail",
        json={
            "message": _draft_payload(to, subject, body, cc, bcc, html_body),
            "saveToSentItems": True,
        },
    )
    return out({"status": "sent", "to": to, "subject": subject})


@tool
async def save_draft(
    to: str, subject: str, body: str, cc: str = "", bcc: str = "", html_body: str = ""
) -> str:
    """Save an email as a draft without sending it.

    Args:
        to: Recipient addresses, separated by ';' or ','.
        subject: Subject line.
        body: Plain-text body (ignored when html_body is given).
        cc: Optional CC addresses.
        bcc: Optional BCC addresses.
        html_body: Optional HTML body.

    Returns:
        JSON with the draft's entry_id (use with send_draft).
    """
    if e := bad_addresses(to, cc, bcc):
        return e
    d = await graph().request(
        "POST", "/me/messages", json=_draft_payload(to, subject, body, cc, bcc, html_body)
    )
    return out({"status": "draft_saved", "entry_id": d.get("id", ""), "subject": subject})


@tool
async def list_drafts(limit: int = 50, offset: int = 0) -> str:
    """List saved drafts, most recently changed first.

    Args:
        limit: Maximum drafts (1-200). Default 50.
        offset: Number of drafts to skip. Default 0.
    """
    if e := bad_count("limit", limit, 1, 200) or bad_count("offset", offset, 0):
        return e
    data = await graph().get(
        "/me/mailFolders/drafts/messages",
        **{
            "$top": limit,
            "$skip": offset or None,
            "$orderby": "lastModifiedDateTime desc",
            "$select": "id,subject,toRecipients,ccRecipients,lastModifiedDateTime,hasAttachments",
        },
    )
    return out(
        [
            {
                "entry_id": d.get("id", ""),
                "subject": d.get("subject") or "(no subject)",
                "to": addresses(d.get("toRecipients")),
                "cc": addresses(d.get("ccRecipients")),
                "last_modified": iso(d.get("lastModifiedDateTime")),
                "has_attachments": bool(d.get("hasAttachments")),
            }
            for d in data.get("value", [])
        ]
    )


@tool
async def read_draft(draft_id: str) -> str:
    """Read a draft in full.

    Args:
        draft_id: The draft's entry_id (from save_draft or list_drafts).
    """
    if e := bad_id(draft_id, "draft"):
        return e
    d = await graph().get(
        f"/me/messages/{seg(draft_id.strip())}",
        **{
            "$select": "id,subject,toRecipients,ccRecipients,bccRecipients,body,isDraft,"
            "lastModifiedDateTime,hasAttachments"
        },
    )
    if not d.get("isDraft"):
        return err("This message is not a draft.")
    return out(
        {
            "entry_id": d.get("id", ""),
            "subject": d.get("subject") or "",
            "to": addresses(d.get("toRecipients")),
            "cc": addresses(d.get("ccRecipients")),
            "bcc": addresses(d.get("bccRecipients")),
            "last_modified": iso(d.get("lastModifiedDateTime")),
            "has_attachments": bool(d.get("hasAttachments")),
            "body": truncate(((d.get("body") or {}).get("content") or "").strip()),
        }
    )


@tool
async def send_draft(draft_id: str) -> str:
    """Send a saved draft. Only unsent drafts can be sent.

    Args:
        draft_id: The draft's entry_id.
    """
    if e := bad_id(draft_id, "draft"):
        return e
    path = f"/me/messages/{seg(draft_id.strip())}"
    d = await graph().get(path, **{"$select": "id,subject,isDraft,toRecipients"})
    if not d.get("isDraft"):
        return err("This message is not an unsent draft, so it was not sent.")
    if not d.get("toRecipients"):
        return err("The draft has no recipients.")
    await graph().request("POST", f"{path}/send")
    return out(
        {"status": "sent", "subject": d.get("subject", ""), "to": addresses(d.get("toRecipients"))}
    )


@tool
async def reply_email(entry_id: str, body: str, reply_all: bool = False) -> str:
    """Reply to an email and send the reply immediately.

    Args:
        entry_id: The email to reply to.
        body: Text of the reply (placed above the quoted original).
        reply_all: Reply to all recipients. Default false.
    """
    if e := bad_id(entry_id):
        return e
    action = "replyAll" if reply_all else "reply"
    await graph().request(
        "POST", f"/me/messages/{seg(entry_id.strip())}/{action}", json={"comment": body or ""}
    )
    return out({"status": "replied", "reply_all": reply_all})


@tool
async def create_draft_reply(
    entry_id: str, body: str = "", reply_all: bool = False, html_body: str = ""
) -> str:
    """Create a reply as a draft (not sent), for review before sending.

    Args:
        entry_id: The email to reply to.
        body: Text of the reply (placed above the quoted original).
        reply_all: Reply to all recipients. Default false.
        html_body: Optional HTML text of the reply instead of body.

    Returns:
        JSON with the draft's entry_id (use with send_draft).
    """
    if e := bad_id(entry_id):
        return e
    action = "createReplyAll" if reply_all else "createReply"
    d = await graph().request(
        "POST",
        f"/me/messages/{seg(entry_id.strip())}/{action}",
        json={"comment": html_body or body or ""},
    )
    return out(
        {
            "status": "draft_created",
            "entry_id": (d or {}).get("id", ""),
            "subject": (d or {}).get("subject", ""),
        }
    )


# --------------------------------------------------------------------------
# Attachments
# --------------------------------------------------------------------------


async def _attachments(entry_id: str) -> list[dict]:
    return await get_all(
        f"/me/messages/{seg(entry_id.strip())}/attachments",
        {"$select": "id,name,contentType,size,isInline"},
        limit=250,
    )


def _kind(a: dict) -> str:
    t = a.get("@odata.type", "")
    return {
        "#microsoft.graph.fileAttachment": "file",
        "#microsoft.graph.itemAttachment": "item",
        "#microsoft.graph.referenceAttachment": "link",
    }.get(t, t or "file")


@tool
async def list_attachments(entry_id: str) -> str:
    """List an email's attachments.

    Args:
        entry_id: The email's entry_id.

    Returns:
        JSON array of {index, name, content_type, size, inline, kind}. Use index
        with save_attachment. kind "file" can be saved; "item" (attached email)
        and "link" (cloud file) cannot.
    """
    if e := bad_id(entry_id):
        return e
    atts = await _attachments(entry_id)
    return out(
        [
            {
                "index": i,
                "name": a.get("name", ""),
                "content_type": a.get("contentType", ""),
                "size": a.get("size", 0),
                "inline": bool(a.get("isInline")),
                "kind": _kind(a),
            }
            for i, a in enumerate(atts, 1)
        ]
    )


@tool
async def save_attachment(
    entry_id: str, attachment_index: int = 1, save_directory: str = ""
) -> str:
    """Save an email attachment to disk.

    Files are only written inside the attachment folder (OUTLOOK_GRAPH_ATTACHMENT_DIR,
    default ~/Downloads), under a cleaned-up file name, never overwriting an existing
    file, and are marked as downloaded from the internet.

    Args:
        entry_id: The email's entry_id.
        attachment_index: Which attachment (1-based, as in list_attachments). Default 1.
        save_directory: Optional subfolder of the attachment folder.

    Returns:
        JSON with path, filename, size and marked_as_downloaded.
    """
    if e := bad_id(entry_id) or bad_count("attachment_index", attachment_index, 1):
        return e
    root = settings().attachment_dir
    target_dir, problem = resolve_save_dir(save_directory, root)
    if problem:
        return problem
    atts = await _attachments(entry_id)
    if attachment_index > len(atts):
        return err(
            f"The email has {len(atts)} attachment(s); index {attachment_index} is too high."
        )
    att = atts[attachment_index - 1]
    if _kind(att) != "file":
        return err(
            f"Attachment {attachment_index} is a {_kind(att)} attachment and cannot be saved "
            "as a file."
        )
    if (att.get("size") or 0) > _MAX_ATTACHMENT_BYTES:
        return err("The attachment is larger than 150 MB.")
    target_dir.mkdir(parents=True, exist_ok=True)
    path = unique_path(target_dir, safe_filename(att.get("name", "")))
    size = await graph().download(
        f"/me/messages/{seg(entry_id.strip())}/attachments/{seg(att['id'])}/$value",
        path,
        _MAX_ATTACHMENT_BYTES,
    )
    return out(
        {
            "status": "saved",
            "path": str(path),
            "filename": path.name,
            "size": size,
            "marked_as_downloaded": mark_as_downloaded(path),
        }
    )
