# Skills

This directory holds specialized skills/workflows that use the `google-drive-mcp`
tools for a specific purpose — the MCP server itself stays generic (upload, share,
export), and each skill here encodes one concrete, opinionated workflow on top of it.

Add a new skill as `skills/<name>/skill.md` (standard Claude Code skill frontmatter:
`name`, `description` with trigger phrases, `version`). Keep skills thin — they should
call this plugin's tools (`upload_and_convert_to_sheet`, `share_file`,
`export_sheet_to_xlsx`, etc.), not reimplement Drive API calls inline.

## Current skills

- **`xlsx-live-review/`** — take an already-filled `.xlsx` and produce a shareable,
  live-editable Google Sheet for human review/sign-off, then pull edits back out.
