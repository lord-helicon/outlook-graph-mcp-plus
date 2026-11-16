"""Task tools on Microsoft To Do's default list ("Tasks")."""

from __future__ import annotations

from outlook_graph_mcp.app import _state, graph, settings, tool
from outlook_graph_mcp.graph import GraphError, seg
from outlook_graph_mcp.util import bad_count, bad_id, err, iso, out, parse_when, truncate

_IMPORTANCE = ("low", "normal", "high")


async def _default_list() -> str:
    if "todo_list" not in _state:
        data = await graph().get("/me/todo/lists")
        lists = data.get("value", [])
        default = next((x for x in lists if x.get("wellknownListName") == "defaultList"), None)
        if default is None:
            raise GraphError("Microsoft To Do has no default task list for this account.")
        _state["todo_list"] = default["id"]
    return _state["todo_list"]


def _summary(t: dict) -> dict:
    return {
        "entry_id": t.get("id", ""),
        "subject": t.get("title") or "(no title)",
        "due_date": iso((t.get("dueDateTime") or {}).get("dateTime"))[:10] or None,
        "complete": t.get("status") == "completed",
        "priority": t.get("importance", "normal"),
    }


async def _task_path(entry_id: str) -> str:
    return f"/me/todo/lists/{seg(await _default_list())}/tasks/{seg(entry_id.strip())}"


@tool
async def list_tasks(include_completed: bool = False, count: int = 20) -> str:
    """List tasks in Microsoft To Do's default list ("Tasks").

    Args:
        include_completed: Also list completed tasks. Default false.
        count: Maximum tasks (1-200). Default 20.

    Returns:
        JSON array of {entry_id, subject, due_date, complete, priority}.
    """
    if e := bad_count("count", count, 1, 200):
        return e
    params = {"$top": count}
    if not include_completed:
        params["$filter"] = "status ne 'completed'"
    data = await graph().get(f"/me/todo/lists/{seg(await _default_list())}/tasks", **params)
    return out([_summary(t) for t in data.get("value", [])])


@tool
async def get_task(entry_id: str) -> str:
    """Read one task in full.

    Args:
        entry_id: The task's entry_id.
    """
    if e := bad_id(entry_id, "task"):
        return e
    t = await graph().get(await _task_path(entry_id))
    data = _summary(t)
    data["body"] = truncate(((t.get("body") or {}).get("content") or "").strip())
    data["created"] = iso(t.get("createdDateTime"))
    return out(data)


@tool
async def create_task(
    subject: str, body: str = "", due_date: str = "", importance: str = "normal"
) -> str:
    """Create a task in Microsoft To Do's default list.

    Args:
        subject: Task title.
        body: Optional notes.
        due_date: Optional due date, e.g. "2026-12-05".
        importance: "low", "normal" (default) or "high".
    """
    imp = (importance or "normal").strip().lower()
    if imp not in _IMPORTANCE:
        return err(f"importance must be 'low', 'normal' or 'high' (got {importance!r}).")
    payload: dict = {"title": subject or "", "importance": imp}
    if due_date:
        due, problem = parse_when(due_date, "due_date")
        if problem:
            return problem
        payload["dueDateTime"] = {
            "dateTime": due.strftime("%Y-%m-%dT00:00:00"),
            "timeZone": settings().timezone,
        }
    if body:
        payload["body"] = {"contentType": "text", "content": body}
    t = await graph().request(
        "POST", f"/me/todo/lists/{seg(await _default_list())}/tasks", json=payload
    )
    return out({"status": "created", **_summary(t)})


@tool
async def complete_task(entry_id: str) -> str:
    """Mark a task as completed.

    Args:
        entry_id: The task's entry_id.
    """
    if e := bad_id(entry_id, "task"):
        return e
    t = await graph().request("PATCH", await _task_path(entry_id), json={"status": "completed"})
    return out({"status": "completed", "subject": (t or {}).get("title", "")})


@tool
async def delete_task(entry_id: str) -> str:
    """Delete a task.

    Args:
        entry_id: The task's entry_id.
    """
    if e := bad_id(entry_id, "task"):
        return e
    path = await _task_path(entry_id)
    t = await graph().get(path)
    await graph().request("DELETE", path)
    return out({"status": "deleted", "subject": t.get("title", "")})
