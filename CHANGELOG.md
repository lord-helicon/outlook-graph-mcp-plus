# Changelog

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
