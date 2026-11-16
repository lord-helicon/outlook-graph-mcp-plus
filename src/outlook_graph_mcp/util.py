"""Input validation, output formatting and attachment-saving rules."""

from __future__ import annotations

import json
import os
import re
import subprocess  # noqa: S404 - only /usr/bin/xattr with fixed arguments
import sys
import time
from datetime import datetime
from pathlib import Path

# --------------------------------------------------------------------------
# Output
# --------------------------------------------------------------------------


def out(data) -> str:
    """Compact JSON for tool results (they go straight into the model's context)."""
    return json.dumps(data, separators=(",", ":"), ensure_ascii=False, default=str)


def err(text: str) -> str:
    return f"Error: {text}"


def iso(value: str | None) -> str:
    """Graph dateTime '2026-12-01T10:00:00.0000000' (or with Z) -> '2026-12-01T10:00:00'."""
    if not value:
        return ""
    return value.rstrip("Z")[:19]


def truncate(text: str, limit: int = 20000) -> str:
    text = text or ""
    return text if len(text) <= limit else text[:limit] + f"\n[... truncated, {len(text)} chars]"


# --------------------------------------------------------------------------
# Input validation
# --------------------------------------------------------------------------

DATE_HINT = "Use YYYY-MM-DD or YYYY-MM-DD HH:MM (DD.MM.YYYY [HH:MM] is also accepted)."


def parse_when(value: str, field: str) -> tuple[datetime | None, str | None]:
    """(datetime, None) or (None, error text). Naive datetimes = configured timezone."""
    text = (value or "").strip()
    try:
        dt = datetime.fromisoformat(text)
        return dt.replace(tzinfo=None) if dt.tzinfo else dt, None
    except ValueError:
        pass
    for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S", "%d.%m.%Y"):
        try:
            return datetime.strptime(text, fmt), None
        except ValueError:
            continue
    return None, err(f"{field} '{value}' is not a valid date. {DATE_HINT}")


def bad_count(name: str, value, minimum: int = 1, maximum: int | None = None) -> str | None:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        return err(f"{name} must be a whole number of at least {minimum} (got {value!r}).")
    if maximum is not None and value > maximum:
        return err(f"{name} must be at most {maximum} (got {value}).")
    return None


def bad_id(value: str, kind: str = "message") -> str | None:
    """Graph ids are opaque base64url-like strings; refuse anything else."""
    v = (value or "").strip()
    if not v or len(v) > 512 or not re.fullmatch(r"[A-Za-z0-9+/=_\-]+", v):
        return err(f"'{value}' is not a valid {kind} id. Use an entry_id returned by this server.")
    return None


def split_addresses(value: str) -> list[str]:
    return [a.strip() for a in re.split(r"[;,]", value or "") if a.strip()]


def recipients(value: str) -> list[dict]:
    return [{"emailAddress": {"address": a}} for a in split_addresses(value)]


def bad_addresses(*values: str) -> str | None:
    for v in values:
        for a in split_addresses(v):
            if not re.fullmatch(r"[^@\s<>\"]+@[^@\s<>\"]+\.[^@\s<>\"]+", a):
                return err(f"'{a}' is not a valid email address.")
    return None


# --------------------------------------------------------------------------
# Attachments: where files may be written
# --------------------------------------------------------------------------


def _is_network_or_device_path(path: str) -> bool:
    """UNC (\\\\host\\share, //host/share), WebDAV and device paths (\\\\?\\, \\\\.\\).

    Checked on the text before any filesystem call: resolving such a path makes
    Windows contact the host (SMB/WebDAV, possibly sending NTLM credentials).
    """
    return path.replace("/", "\\").startswith("\\\\")


def resolve_save_dir(save_directory: str, root: Path) -> tuple[Path | None, str | None]:
    """The directory to save into, inside `root` (resolved), or an error.

    Empty = the root itself; relative = relative to the root. Symlinks are
    resolved before the containment check, so they cannot lead outside.
    """
    root = Path(os.path.realpath(root))
    if not save_directory:
        return root, None
    if _is_network_or_device_path(save_directory):
        return None, err(
            f"save_directory must be a local folder inside '{root}'; network and device "
            "paths are not allowed."
        )
    candidate = Path(save_directory).expanduser()
    if not candidate.is_absolute():
        candidate = root / candidate
    resolved = Path(os.path.realpath(candidate))
    if resolved != root and root not in resolved.parents:
        return None, err(
            f"save_directory must be inside the attachment folder '{root}' "
            "(set OUTLOOK_GRAPH_ATTACHMENT_DIR to change it)."
        )
    return resolved, None


def safe_filename(name: str) -> str:
    """Reduce a sender-chosen attachment name to one harmless file name."""
    name = (name or "").replace("\\", "/").rsplit("/", 1)[-1]
    name = re.sub(r"[^\w.\- ()]", "_", name).strip(" .")
    if name.upper().split(".")[0] in {
        "CON",
        "PRN",
        "AUX",
        "NUL",
        *(f"COM{i}" for i in range(1, 10)),
        *(f"LPT{i}" for i in range(1, 10)),
    }:
        name = "_" + name
    return name[:180] or "attachment"


def unique_path(directory: Path, filename: str) -> Path:
    """A path in `directory` that does not exist yet ('name (1).ext', ...)."""
    path = directory / filename
    if not os.path.lexists(path):
        return path
    stem, ext = os.path.splitext(filename)
    n = 1
    while os.path.lexists(directory / f"{stem} ({n}){ext}"):
        n += 1
    return directory / f"{stem} ({n}){ext}"


def mark_as_downloaded(path: Path) -> bool:
    """Mark a saved attachment as downloaded from the internet.

    Windows: Zone.Identifier stream (ZoneId=3), checked by SmartScreen and
    Office Protected View. macOS: com.apple.quarantine, checked by Gatekeeper.
    """
    try:
        if os.name == "nt":
            with open(f"{path}:Zone.Identifier", "w") as f:
                f.write("[ZoneTransfer]\r\nZoneId=3\r\n")
            return True
        if sys.platform == "darwin":
            value = f"0081;{int(time.time()):08x};outlook-graph-mcp-plus;"
            done = subprocess.run(  # noqa: S603 - fixed binary and arguments
                ["/usr/bin/xattr", "-w", "com.apple.quarantine", value, str(path)],
                capture_output=True,
                timeout=10,
            )
            return done.returncode == 0
    except (OSError, subprocess.SubprocessError):
        return False
    return False
