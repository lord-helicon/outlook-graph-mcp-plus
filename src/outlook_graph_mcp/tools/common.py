"""Helpers shared by the tool modules."""

from __future__ import annotations

from typing import Any

from outlook_graph_mcp.app import graph, settings
from outlook_graph_mcp.graph import GraphError, seg

# Names accepted for Outlook's standard folders -> Graph well-known folder names.
WELL_KNOWN_FOLDERS = {
    "inbox": "inbox",
    "sent": "sentitems",
    "sent items": "sentitems",
    "sentitems": "sentitems",
    "sentmail": "sentitems",
    "drafts": "drafts",
    "deleted": "deleteditems",
    "deleted items": "deleteditems",
    "deleteditems": "deleteditems",
    "trash": "deleteditems",
    "junk": "junkemail",
    "junk email": "junkemail",
    "junk e-mail": "junkemail",
    "junk mail": "junkemail",
    "junkemail": "junkemail",
    "spam": "junkemail",
    "archive": "archive",
    "outbox": "outbox",
}

_MAX_FOLDER_DEPTH = 6


def mailbox_base(mailbox: str = "") -> str:
    """Graph path prefix for the signed-in user's mailbox or an allowed shared one.

    Only shared mailboxes listed in OUTLOOK_GRAPH_SHARED_MAILBOXES are accepted.
    Exchange permissions (Full Access, Send As) are checked by Microsoft anyway; the
    allowlist also stops text in an email from steering the model to another mailbox.
    """
    m = (mailbox or "").strip().lower()
    if m in ("", "me"):
        return "/me"
    allowed = settings().shared_mailboxes
    if m not in allowed:
        listed = ", ".join(allowed) if allowed else "none configured"
        raise GraphError(
            f"'{mailbox}' is not one of the shared mailboxes this server may use "
            f"({listed}). Call list_mailboxes; an administrator can add mailboxes "
            "with OUTLOOK_GRAPH_SHARED_MAILBOXES."
        )
    return f"/users/{seg(m)}"


async def get_all(path: str, params: dict, limit: int) -> list[dict]:
    """GET a collection, following @odata.nextLink until `limit` items."""
    items: list[dict] = []
    data = await graph().get(path, **params)
    while True:
        items.extend((data or {}).get("value", []))
        nxt = (data or {}).get("@odata.nextLink")
        if len(items) >= limit or not nxt:
            return items[:limit]
        data = await graph().request("GET", nxt)


async def folder_tree(max_depth: int, base: str = "/me") -> list[dict]:
    """All mail folders down to max_depth, each with its path ('Inbox/Projects')."""
    out: list[dict] = []

    async def walk(path: str, parent: str, depth: int) -> None:
        folders = await get_all(
            path,
            {
                "$top": 200,
                "$select": "id,displayName,totalItemCount,unreadItemCount,childFolderCount",
            },
            limit=2000,
        )
        for f in folders:
            full = f"{parent}/{f['displayName']}" if parent else f["displayName"]
            out.append({**f, "path": full, "depth": depth})
            if f.get("childFolderCount") and depth < max_depth:
                await walk(f"{base}/mailFolders/{seg(f['id'])}/childFolders", full, depth + 1)

    await walk(f"{base}/mailFolders", "", 0)
    return out


async def resolve_folder(name: str, base: str = "/me") -> str:
    """A folder name or path -> a value usable in /me/mailFolders/{value} and as destinationId.

    Standard names map to Graph's well-known names (language-independent).
    Custom folders match by display name or full path ('Inbox/Projects'),
    case-insensitively; ambiguous names are refused.
    """
    key = (name or "").strip()
    if not key:
        raise GraphError("The folder name is empty.")
    well_known = WELL_KNOWN_FOLDERS.get(key.lower())
    if well_known:
        return well_known
    folders = await folder_tree(_MAX_FOLDER_DEPTH, base)
    wanted = key.strip("/").lower()
    if "/" in key:
        # A path ('Inbox/Projects', or '/Projects' for a top-level folder) is taken literally.
        matches = [f for f in folders if f["path"].lower() == wanted]
    else:
        # A bare name must be unique, so mail never lands in a folder the user did not mean.
        matches = [f for f in folders if f["displayName"].lower() == wanted]
    if len(matches) == 1:
        return matches[0]["id"]
    if not matches:
        raise GraphError(f"No mail folder named '{name}'. Call list_folders to see the folders.")
    paths = ", ".join(f["path"] for f in matches[:5])
    raise GraphError(
        f"Several folders are named '{name}' ({paths}). Use the full path, "
        "e.g. 'Inbox/Projects', or '/Projects' for a top-level folder."
    )


def address(entry: Any) -> tuple[str, str]:
    """(address, name) from a Graph recipient/from object."""
    ea = (entry or {}).get("emailAddress") or {}
    return ea.get("address", "") or "", ea.get("name", "") or ""


def addresses(entries: list | None) -> str:
    return "; ".join(a for a, _ in (address(e) for e in entries or []) if a)
