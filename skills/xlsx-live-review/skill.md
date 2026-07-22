---
name: xlsx-live-review
description: >
  Turn an already-filled .xlsx into a shareable, live-editable Google Sheet for human
  review and sign-off, then pull any edits back out as the .xlsx of record. Use this
  whenever a filled Excel deliverable (a filled template, a generated register, a
  cost model) needs a human to actually open, comment on, or edit it before it's
  considered final — not for generating a new spreadsheet from scratch.
  Triggers on: "let me review this in Sheets", "make this editable", "share this
  spreadsheet for review", "get sign-off on this xlsx", "live-edit this excel file".
version: 1.0.0
---

# xlsx Live Review — Google Sheets Hand-off

## When to use this

Any workflow that fills a real `.xlsx` (a specific template with formulas, pivot
tables, or a fixed official structure — e.g. a government/utility program form) and
then needs a human to review or edit it before it's final. Do NOT use this to author
a spreadsheet from nothing — that's a from-scratch generation task, not a review
hand-off, and belongs to whatever tool is actually building the workbook.

**The original `.xlsx` bytes are the source of truth throughout.** This skill never
replaces them — it produces a *converted review copy* in Google Sheets, and (if the
human edits it) pulls the edits back out as a new `.xlsx`, but the original file
stays wherever the calling workflow already saved it.

## Step 1: Upload and convert

Call `upload_and_convert_to_sheet` with:
- `name` — a clear, human-facing filename (e.g. `"Cascade IEA Opportunity Register — for review.xlsx"`)
- `contentBase64` — the exact bytes just produced by whatever filled the template
- `shareWithEmail` — the reviewer's email if known up front (skip this and use
  `share_file` separately if you don't have it yet)
- `shareRole` — `"writer"` if they need to edit, `"commenter"` if they should only
  annotate, `"reader"` for view-only

This returns a `webViewLink` — that's the URL to hand the human.

## Step 2: Hand it to the reviewer

Tell the user directly: here's the link, please review/edit and let me know when
you're done (or approve as-is). Do not treat silence or an unrelated message as
approval — if it's ambiguous whether they're actually signing off, ask directly:
"Should I treat this as approved, or are there changes first?"

Note plainly, once, that this is a **converted copy** for review — Excel-specific
details (complex pivot cache internals, some conditional formats) may not carry over
perfectly into Sheets; the original `.xlsx` is unaffected and remains the real file.

## Step 3: If they edited it — pull the edits back out

Call `export_sheet_to_xlsx` with the same `fileId` returned in Step 1. This returns
fresh `.xlsx` bytes reflecting whatever the human changed in Sheets. Hand these back
to whichever workflow/skill owns saving the "real" file (e.g. re-save via
`soapbox-files`' `write_binary_file`) — this skill's job ends at handing back bytes,
it does not know where the calling workflow's file lives.

## Step 4: If they approved as-is

No export needed — just record the sign-off in whatever ledger/gate the calling
workflow uses (e.g. `gates.register_signoff` in an `iea-audit-plan`-style state
ledger). The Sheet can be left in Drive as a record of what was reviewed, or deleted
via `delete_file` if the calling workflow doesn't want it lingering — don't delete it
without checking, since some workflows want the reviewed copy kept as an audit trail.
