# google-drive-mcp

A Google Drive MCP server, mirroring the `xero-mcp` pattern (see `~/xero-mcp`):
this service is the **sole owner** of a Google OAuth refresh token, hosted on Railway
behind a bearer token, exposed at `/mcp` for Claude Code / opencode, with a thin
Paperclip connector (see "Paperclip connector" below) proxying to it — so the OAuth
credential lives in exactly one place regardless of how many agents/platforms use it.

**Primary use case this was built for:** taking an already-filled `.xlsx` (e.g. from
`xlsx-templater-mcp`) and turning it into a live, collaboratively-editable Google
Sheet for human review/sign-off — without rebuilding the spreadsheet from scratch and
without losing the original file (the Sheet is a converted copy; the source `.xlsx`
stays wherever it already lives, e.g. Supabase Storage). See `skills/xlsx-live-review/`.

## Tools

| Tool | Purpose |
|---|---|
| `list_files` | Search/list Drive files |
| `get_file` | Metadata for one file |
| `upload_and_convert_to_sheet` | **The key tool** — upload xlsx bytes, get back a live Google Sheet (`webViewLink`), optionally share it in the same call |
| `upload_file` | Upload bytes as-is, no conversion (e.g. a PDF) |
| `share_file` | Grant a role (reader/commenter/writer) to a user or "anyone with the link" |
| `download_file` | Raw bytes back, base64 |
| `export_sheet_to_xlsx` | Convert a (possibly human-edited) Sheet back to `.xlsx` bytes |
| `delete_file` | Permanent delete |

## What this does NOT do

- It does not generate a spreadsheet from structured data (that's a different job —
  Mila's `generate_*` tools do that, from scratch, with no Excel support today; see
  the SES engagement notes on why Mila couldn't be used for register review).
- `upload_and_convert_to_sheet` converts the file Drive receives — it is not a
  live two-way sync. If a human edits the Sheet, call `export_sheet_to_xlsx` to pull
  the edited version back out as the new `.xlsx` of record.
- Google's xlsx→Sheets conversion can flatten some Excel-only features (complex pivot
  cache internals, certain conditional formatting). For a document where **byte-exact**
  Excel fidelity is the whole point (e.g. an official government/utility program
  template), keep the original `.xlsx` as the system of record and treat the Sheet as
  a review copy, not a replacement.

## Setup — REQUIRES a manual step (Google Cloud Console)

Unlike Railway (which has a scriptable signup/login flow), Google Cloud OAuth app
creation needs a human in a browser. This cannot be fully self-served by an agent.
Someone with access to a Google Cloud account needs to:

1. Go to console.cloud.google.com → create (or pick) a project.
2. **APIs & Services → Enabled APIs** → enable the "Google Drive API".
3. **APIs & Services → OAuth consent screen**:
   - User type: Internal (if using Google Workspace) or External.
   - Scopes: add `https://www.googleapis.com/auth/drive.file` (recommended — the app
     can only see files it creates/opens, not the whole Drive; broaden to
     `https://www.googleapis.com/auth/drive` only if a workflow genuinely needs to
     browse/manage files it didn't create).
   - If "External" + "Testing" mode: refresh tokens expire after 7 days unless the
     app is moved to "In production" (may need Google verification for sensitive
     scopes — `drive.file` is non-sensitive and shouldn't require full verification).
4. **APIs & Services → Credentials → Create Credentials → OAuth client ID**:
   - Application type: Web application.
   - Authorized redirect URI: `https://google-drive.mcp.soapbox.build/callback`
     (or whatever this service's actual deployed domain is).
   - Save the **Client ID** and **Client Secret**.
5. Deploy this service to Railway (`soapbox-mcps` project, matching `xero-mcp`/
   `costing-mcp`) with env vars:
   - `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET` — from step 4
   - `GOOGLE_REDIRECT_URI` — must exactly match step 4's redirect URI
   - `MCP_AUTH_TOKEN` — a random bearer token this service requires on `/mcp`
   - `DRIVE_TOKEN_FILE` — defaults to `/data/google_token.json`; attach a Railway
     volume at `/data` (same as `xero-mcp`) so the token survives redeploys
6. **One-time consent** (do this once, by a human, after deploy):
   - Build the consent URL:
     ```
     https://accounts.google.com/o/oauth2/v2/auth?client_id=<GOOGLE_CLIENT_ID>&redirect_uri=<GOOGLE_REDIRECT_URI>&response_type=code&scope=https://www.googleapis.com/auth/drive.file&access_type=offline&prompt=consent
     ```
   - Open it, sign in as whichever Google account should own the uploaded files,
     approve.
   - Google redirects to `.../callback?code=...` — append `&key=<MCP_AUTH_TOKEN>`
     to that URL before loading it (or the callback will only show you the code
     instead of storing it) — this guards against a stranger hijacking the connection
     by completing their own consent for this client_id.
   - On success you'll see "✅ Google Drive connection stored." The refresh token is
     now persisted on the Railway volume; no `GOOGLE_REFRESH_TOKEN` env var is needed
     going forward (it's only a fallback seed for first boot if the volume is empty).

## Paperclip connector

Paperclip agents cannot consume external MCP servers directly (see
`soapbox-plugin-design` skill) — they need a thin proxy plugin, same pattern as
`plugin-xero` (vendored into `soapboxbuild/paperclip`, `adapters/paperclip-xero`).
Not yet built for this service — follow that same vendoring pattern
(`adapters/paperclip-drive` + an `inject-paperclip-drive.sh` + Dockerfile esbuild step)
when a Paperclip agent needs this, rather than duplicating the OAuth logic there.

## Local dev

```bash
npm install
npm run dev   # tsx src/index.ts, reads env vars from your shell
```

`npm run build && npm start` for the compiled path (what Railway actually runs).
