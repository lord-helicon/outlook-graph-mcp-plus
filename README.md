# outlook-graph-mcp-plus

Microsoft 365 Outlook as a **local MCP server** for Claude (Desktop and Code): mail, folders, attachments, calendar and Microsoft To Do, through **Microsoft Graph**.

- Works with **Microsoft 365 work and school accounts**. Each organisation uses **its own app registration** in its own Entra ID tenant.
- **Delegated sign-in only**: each user signs in once and the server can reach **only that user's own** mailbox, calendar and tasks. There's no client secret and no access to other mailboxes.
- Runs **on the user's computer** (macOS, Windows, Linux). Outlook doesn't need to be installed. Tokens stay in the operating system's secret store.
- Offers the same 28 tools, with the same names, as its sibling project [outlook-desktop-mcp-plus](https://github.com/lord-helicon/outlook-desktop-mcp-plus), which controls a locally installed Outlook instead.

> **Status: alpha (0.1.0).** Unit-tested against a simulated Microsoft Graph, but not yet run against a real tenant.

---

## 1. For the Microsoft 365 administrator: app registration

You do this once per organisation, in the [Microsoft Entra admin center](https://entra.microsoft.com).

1. Go to **Entra ID → App registrations → New registration**.
   - **Name:** for example `Claude Outlook (MCP)`.
   - **Supported account types:** *Accounts in this organizational directory only (single tenant)*.
   - **Redirect URI:** platform **Public client/native (mobile & desktop)**, value `http://localhost`.
   - Click **Register**.
2. In **Authentication**, under *Advanced settings*, set **Allow public client flows** to **Yes**. This enables the device-code fallback for users without a local browser.
3. In **API permissions → Add a permission → Microsoft Graph → Delegated permissions**, add:
   - `User.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Calendars.ReadWrite`
   - `Tasks.ReadWrite`

   Then click **Grant admin consent for \<your organisation\>**, so users aren't asked to consent individually.
4. Note the **Application (client) ID** and the **Directory (tenant) ID** from the *Overview* page and give both to your users. They aren't secret.

No client secret or certificate is needed.

**Optional hardening**
- **Limit who can use it:** in *Enterprise applications → (the app) → Properties*, set **Assignment required** to *Yes*, then assign the users or groups who may use it.
- **Conditional Access** applies to the browser sign-in (MFA, compliant device, and so on). If your policies block the *device code* flow, users simply sign in with the browser.
- **Revoking access:** remove the user's assignment or delete the app registration. Existing refresh tokens stop working.

---

## 2. For each user: install, sign in, connect Claude

**Install** (Python 3.12+ via [uv](https://docs.astral.sh/uv/)):

```bash
uv tool install outlook-graph-mcp-plus
```

Until the first PyPI release, install from GitHub:

```bash
uv tool install git+https://github.com/lord-helicon/outlook-graph-mcp-plus
```

**Sign in once.** Use the two IDs from your administrator. A browser window opens for the normal Microsoft sign-in:

```bash
OUTLOOK_GRAPH_CLIENT_ID=<application-client-id> OUTLOOK_GRAPH_TENANT_ID=<tenant-id> outlook-graph-mcp-plus auth
```

On a machine without a browser, add `--device-code`. You then open a URL on any device and enter the code it shows. `outlook-graph-mcp-plus status` shows who is signed in, and `outlook-graph-mcp-plus logout` removes the cached tokens. On Windows PowerShell, set the variables first with `$env:OUTLOOK_GRAPH_CLIENT_ID="..."`.

**Connect Claude Code:**

```bash
claude mcp add outlook-graph -e OUTLOOK_GRAPH_CLIENT_ID=<application-client-id> -e OUTLOOK_GRAPH_TENANT_ID=<tenant-id> -- outlook-graph-mcp-plus
```

**Connect Claude Desktop** (*Settings → Developer → Edit Config*, `claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "outlook-graph": {
      "command": "outlook-graph-mcp-plus",
      "env": {
        "OUTLOOK_GRAPH_CLIENT_ID": "<application-client-id>",
        "OUTLOOK_GRAPH_TENANT_ID": "<tenant-id>"
      }
    }
  }
}
```

If Claude can't find the command, use the full path shown by `which outlook-graph-mcp-plus` (macOS and Linux) or `where outlook-graph-mcp-plus` (Windows). Restart Claude afterwards.

---

## Configuration

| Variable | Required | Meaning |
|---|---|---|
| `OUTLOOK_GRAPH_CLIENT_ID` | yes | Application (client) ID of the organisation's app registration |
| `OUTLOOK_GRAPH_TENANT_ID` | yes | Directory (tenant) ID or domain, for example `contoso.onmicrosoft.com`. Personal accounts (`common`, `consumers`) are refused |
| `OUTLOOK_GRAPH_ATTACHMENT_DIR` | no | The only folder attachments are saved into (default `~/Downloads`) |
| `OUTLOOK_GRAPH_TIMEZONE` | no | IANA timezone for dates (default: the computer's timezone) |
| `OUTLOOK_GRAPH_CACHE_DIR` | no | Where the sign-in state is kept (default `~/.outlook-graph-mcp-plus`) |
| `OUTLOOK_GRAPH_ALLOW_PLAINTEXT_TOKEN_CACHE` | no | `true` only if no OS secret store exists (for example a Linux box without libsecret) |

## Tools

| Area | Tools |
|---|---|
| Mail | `list_emails`, `read_email`, `search_emails`, `send_email`, `reply_email`, `mark_as_read`, `mark_as_unread`, `move_email`, `list_folders` |
| Drafts | `save_draft`, `list_drafts`, `read_draft`, `send_draft`, `create_draft_reply` |
| Attachments | `list_attachments`, `save_attachment` |
| Calendar | `list_events`, `search_events`, `get_event`, `create_event`, `create_meeting`, `update_event`, `delete_event` |
| Tasks (Microsoft To Do, default list) | `list_tasks`, `get_task`, `create_task`, `complete_task`, `delete_task` |

Notes on the tools:
- **Folders:** standard names work in any Outlook language: `inbox`, `sent`, `drafts`, `deleted`, `junk`, `archive`, `outbox`. Custom folders are given by name. If a name exists more than once, give the path, for example `Inbox/Projects`, or `/Projects` for a top-level folder.
- **Moving:** `move_email` gives the email a new `entry_id`.
- **Dates:** inputs take `YYYY-MM-DD`, `YYYY-MM-DD HH:MM` or `DD.MM.YYYY [HH:MM]`. Outputs are ISO 8601 in your timezone.

## Security

- **Delegated access:** no application permissions and no secrets. The server acts only as the signed-in user, within the permissions the administrator granted.
- **Tokens:** kept in the macOS Keychain, Windows DPAPI or Linux libsecret. A plaintext cache is only used when explicitly allowed.
- **Network:** the server only contacts `login.microsoftonline.com` and `graph.microsoft.com`. There's no telemetry.
- **Attachments:** only ever saved inside the attachment folder:
  - The path is resolved before checking, so symlinks and `..` can't escape it.
  - Network and device paths are refused before they're touched.
  - Sender-chosen file names are cleaned, and existing files are never overwritten.
  - Saved files are marked as downloaded from the internet (Windows Mark-of-the-Web, macOS quarantine).
- **Ids** from mailbox content are validated and URL-encoded, so they can't redirect a request to another Graph endpoint.
- **Third-party content:** email subjects, bodies and attachment names come from outside your organisation. A model can be steered by text in an email ("prompt injection"). The server tells the model to treat such content as data and to confirm before sending, replying, moving or deleting. Keep those confirmations switched on in your Claude client.

## Development

```bash
uv sync --group dev
uv run pytest -q
uv run ruff check src tests && uv run black --check src tests
```

## License

MIT, © 2026 Alexander Rolli.
