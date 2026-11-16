# Changelog

## Unreleased

### Fixed
- Exchange-internal X.500 sender addresses (`/O=EXCHANGELABS/...`, seen in a shared mailbox's
  own Sent Items) are no longer returned as `sender`; the display name is kept.

## 0.2.0 — 2026-11-16

### Added
- **Shared mailboxes.** Mail and calendar tools take an optional `mailbox` parameter. Only
  mailboxes listed in `OUTLOOK_GRAPH_SHARED_MAILBOXES` are accepted. Sending from a shared
  mailbox sends as that mailbox, which needs *Send As* in Exchange. Tasks stay on the user's
  own Microsoft To Do.
- `list_mailboxes` tool: the user's own mailbox plus the configured shared ones.
- The `.Shared` Graph permissions are requested only when shared mailboxes are configured.
- Clear error when an address isn't a mailbox (`ErrorInvalidUser`).

Tested live with a shared mailbox: reading (folders, inbox, read, search), drafts, read state,
moving, an event in the shared calendar, and sending from the shared mailbox (Send As: the mail
arrived from the shared address, with a copy in its Sent Items).

## 0.1.1 — 2026-11-16

First release on PyPI. All 28 tools were tested against a real Microsoft 365 tenant:
read-only on real mail, writes on test items, and sending to the tester's own address.

### Fixed
- An id that doesn't belong to the signed-in mailbox now gives a readable error instead of
  Graph's raw `ErrorInvalidIdMalformed`.
- "Access denied" on an item that was already sent, moved or deleted no longer suggests a
  missing app permission.

### Security
- Request URLs are no longer logged. httpx logged every Graph URL, including search terms,
  filters and ids, at INFO level, and MCP clients keep server output in their log files.

## 0.1.0 — 2026-11-16

Initial version (GitHub only): 28 tools for mail, drafts, folders, attachments, calendar and
Microsoft To Do via Microsoft Graph, with delegated sign-in per organisation.
