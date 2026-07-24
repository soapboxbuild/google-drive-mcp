// ---------------------------------------------------------------------------
// Google Drive API v3 client — thin fetch wrapper, no googleapis SDK
// dependency (keeps the service lean, matching the xero-mcp/costing-mcp
// style of direct REST calls over a heavy client library).
//
// Auth model: this server does NOT hold or refresh its own Google OAuth
// token. Each request arrives with a real Google access token as its Bearer
// token — minted per-portfolio via soapbox-platform's
// /api/oauth/google-drive OAuth flow, registered as an `mcp_oauth` credential
// in Anthropic's credential vault, which auto-refreshes it using Google's own
// token endpoint before every call. This server is a stateless pass-through:
// whatever token arrives IS the Google access token to use for that request.
// ---------------------------------------------------------------------------

const DRIVE_API = "https://www.googleapis.com/drive/v3";
const UPLOAD_API = "https://www.googleapis.com/upload/drive/v3";

async function authedFetch(accessToken: string, url: string, init: RequestInit = {}): Promise<Response> {
  if (!accessToken) {
    throw new Error("No Google access token on this request — the Google Drive connector isn't connected for this portfolio yet.");
  }
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${accessToken}`);
  const resp = await fetch(url, { ...init, headers });
  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`Drive API ${url} failed (${resp.status}): ${text}`);
  }
  return resp;
}

export interface DriveFile {
  id: string;
  name: string;
  mimeType: string;
  webViewLink?: string;
  webContentLink?: string;
  parents?: string[];
  size?: string;
  modifiedTime?: string;
}

const FILE_FIELDS = "id,name,mimeType,webViewLink,webContentLink,parents,size,modifiedTime";

export async function listFiles(accessToken: string, opts: {
  query?: string;
  pageSize?: number;
  folderId?: string;
}): Promise<DriveFile[]> {
  const q = [
    opts.query,
    opts.folderId ? `'${opts.folderId}' in parents` : undefined,
    "trashed = false",
  ].filter(Boolean).join(" and ");
  const url = new URL(`${DRIVE_API}/files`);
  url.searchParams.set("q", q);
  url.searchParams.set("fields", `files(${FILE_FIELDS})`);
  url.searchParams.set("pageSize", String(opts.pageSize ?? 25));
  const resp = await authedFetch(accessToken, url.toString());
  const json = (await resp.json()) as { files: DriveFile[] };
  return json.files;
}

export async function getFile(accessToken: string, fileId: string): Promise<DriveFile> {
  const url = new URL(`${DRIVE_API}/files/${fileId}`);
  url.searchParams.set("fields", FILE_FIELDS);
  const resp = await authedFetch(accessToken, url.toString());
  return resp.json() as Promise<DriveFile>;
}

/**
 * Upload file bytes. If `convertToGoogleFormat` is true and the source is an
 * Office format (xlsx/docx/pptx), Drive auto-converts it into the matching
 * Google Workspace format (Sheets/Docs/Slides) on upload — this is what makes
 * an already-filled .xlsx live-editable/collaborative without rebuilding it
 * from scratch in some other tool.
 */
export async function uploadFile(accessToken: string, opts: {
  name: string;
  contentBase64: string;
  sourceMimeType: string;
  folderId?: string;
  convertToGoogleFormat?: boolean;
}): Promise<DriveFile> {
  const targetMimeType = opts.convertToGoogleFormat
    ? googleFormatFor(opts.sourceMimeType)
    : opts.sourceMimeType;

  const metadata: Record<string, unknown> = {
    name: opts.name,
    mimeType: targetMimeType,
  };
  if (opts.folderId) metadata.parents = [opts.folderId];

  // The media part MUST be the real decoded binary bytes, not the base64 text
  // itself — Google's multipart upload endpoint does not decode
  // Content-Transfer-Encoding on the media part (that header only means
  // anything to email-style MIME parsers), so a body built by string-
  // interpolating contentBase64 sends literal base64 CHARACTERS as the
  // "file", and Drive correctly rejects them as unconvertible ("Conversion
  // of the uploaded content to the requested output type is not supported").
  // Confirmed live: a small enough base64 blob was silently accepted as a
  // blank Sheet (Drive didn't bother validating it), but a real ~230KB
  // xlsx's base64 text failed conversion outright. Build the body as a
  // Buffer so the binary part is genuinely binary.
  const boundary = `drivempcboundary${Date.now()}`;
  const preamble = Buffer.from(
    `--${boundary}\r\n` +
      `Content-Type: application/json; charset=UTF-8\r\n\r\n` +
      `${JSON.stringify(metadata)}\r\n` +
      `--${boundary}\r\n` +
      `Content-Type: ${opts.sourceMimeType}\r\n\r\n`,
    "utf-8",
  );
  const fileBytes = Buffer.from(opts.contentBase64, "base64");
  const trailer = Buffer.from(`\r\n--${boundary}--`, "utf-8");
  const body = Buffer.concat([preamble, fileBytes, trailer]);

  const url = new URL(`${UPLOAD_API}/files`);
  url.searchParams.set("uploadType", "multipart");
  url.searchParams.set("fields", FILE_FIELDS);

  const resp = await authedFetch(accessToken, url.toString(), {
    method: "POST",
    headers: { "Content-Type": `multipart/related; boundary=${boundary}` },
    body,
  });
  return resp.json() as Promise<DriveFile>;
}

function googleFormatFor(sourceMimeType: string): string {
  const map: Record<string, string> = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "application/vnd.google-apps.spreadsheet",
    "application/vnd.ms-excel": "application/vnd.google-apps.spreadsheet",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "application/vnd.google-apps.document",
    "application/msword": "application/vnd.google-apps.document",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": "application/vnd.google-apps.presentation",
  };
  const target = map[sourceMimeType];
  if (!target) {
    throw new Error(
      `No Google Workspace conversion target known for mimeType "${sourceMimeType}". ` +
        `Pass convertToGoogleFormat: false to upload it as a plain (non-editable-in-Sheets/Docs) file instead.`,
    );
  }
  return target;
}

export type DriveRole = "reader" | "commenter" | "writer";

export async function shareFile(accessToken: string, opts: {
  fileId: string;
  role: DriveRole;
  type: "anyone" | "user";
  emailAddress?: string;
}): Promise<{ id: string }> {
  if (opts.type === "user" && !opts.emailAddress) {
    throw new Error("emailAddress is required when type is 'user'");
  }
  const url = new URL(`${DRIVE_API}/files/${opts.fileId}/permissions`);
  url.searchParams.set("fields", "id");
  const body: Record<string, unknown> = { role: opts.role, type: opts.type };
  if (opts.emailAddress) body.emailAddress = opts.emailAddress;
  const resp = await authedFetch(accessToken, url.toString(), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  return resp.json() as Promise<{ id: string }>;
}

export async function deleteFile(accessToken: string, fileId: string): Promise<void> {
  await authedFetch(accessToken, `${DRIVE_API}/files/${fileId}`, { method: "DELETE" });
}

/** Download raw bytes (for a plain uploaded file, not a native Google format). */
export async function downloadFile(accessToken: string, fileId: string): Promise<string> {
  const resp = await authedFetch(accessToken, `${DRIVE_API}/files/${fileId}?alt=media`);
  const buf = Buffer.from(await resp.arrayBuffer());
  return buf.toString("base64");
}

/** Export a native Google format (Sheets/Docs/Slides) back to an Office format. */
export async function exportFile(accessToken: string, fileId: string, exportMimeType: string): Promise<string> {
  const url = new URL(`${DRIVE_API}/files/${fileId}/export`);
  url.searchParams.set("mimeType", exportMimeType);
  const resp = await authedFetch(accessToken, url.toString());
  const buf = Buffer.from(await resp.arrayBuffer());
  return buf.toString("base64");
}

const DOCS_API = "https://docs.googleapis.com/v1/documents";

export interface DocsTextReplacement {
  find: string;
  replaceText: string;
  matchCase?: boolean;
}

/** Fetch a Google Doc's plain text content (concatenated paragraph text runs). */
export async function getDocumentText(accessToken: string, documentId: string): Promise<string> {
  const resp = await authedFetch(accessToken, `${DOCS_API}/${documentId}`);
  const doc = (await resp.json()) as { body?: { content?: Array<{ paragraph?: { elements?: Array<{ textRun?: { content?: string } }> } }> } };
  const parts: string[] = [];
  for (const el of doc.body?.content ?? []) {
    for (const run of el.paragraph?.elements ?? []) {
      if (run.textRun?.content) parts.push(run.textRun.content);
    }
  }
  return parts.join("");
}

/**
 * Replace literal text occurrences in a Google Doc via `replaceAllText`.
 * IMPORTANT: this is a DIRECT edit, not a suggestion/track-changes edit — the
 * Docs API has no facility to author suggestions (only to view existing ones
 * on `documents.get` via SuggestionsViewMode); every write here lands as
 * immediately-accepted content, indistinguishable from the doc owner typing
 * it themselves. Confirmed against Google's own docs (no suggest-mode write
 * path exists as of 2026) — do not represent this as "tracked changes" to a
 * caller or user.
 */
export async function replaceTextInDocument(accessToken: string, documentId: string, replacements: DocsTextReplacement[]): Promise<void> {
  const requests = replacements.map((r) => ({
    replaceAllText: {
      containsText: { text: r.find, matchCase: r.matchCase ?? true },
      replaceText: r.replaceText,
    },
  }));
  await authedFetch(accessToken, `${DOCS_API}/${documentId}:batchUpdate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ requests }),
  });
}

const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";

export interface SheetErrorCell {
  sheet: string;
  cell: string; // A1 notation, e.g. "Opportunity Register!P15"
  errorType: string; // e.g. "DIVIDE_BY_ZERO", "REF", "VALUE", "N_A", "ERROR"
  message: string;
}

function columnLetter(index: number): string {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Checks every formula cell in a live Google Sheet for a computed error value
 * (DIV/0, REF, VALUE, N/A, NAME, NUM, ERROR) -- this is what caught the BC
 * Hydro register's Excel-structured-table-reference incompatibility live
 * (600+ broken cells) before it ever reached a human reviewer. Reads
 * `effectiveValue.errorValue` from the Sheets API, which reflects Sheets'
 * OWN recalculation -- the authoritative source, not a guess from formula
 * text or a raw xlsx's stale cached values.
 */
export async function checkSheetForErrors(accessToken: string, fileId: string): Promise<SheetErrorCell[]> {
  const url = new URL(`${SHEETS_API}/${fileId}`);
  url.searchParams.set("includeGridData", "true");
  url.searchParams.set("fields", "sheets(properties.title,data.rowData.values(effectiveValue))");
  const resp = await authedFetch(accessToken, url.toString());
  const json = (await resp.json()) as {
    sheets: Array<{
      properties: { title: string };
      data?: Array<{ rowData?: Array<{ values?: Array<{ effectiveValue?: { errorValue?: { type: string; message: string } } }> }> }>;
    }>;
  };

  const errors: SheetErrorCell[] = [];
  for (const sheet of json.sheets ?? []) {
    const title = sheet.properties.title;
    for (const grid of sheet.data ?? []) {
      const rowData = grid.rowData ?? [];
      for (let r = 0; r < rowData.length; r++) {
        const values = rowData[r].values ?? [];
        for (let c = 0; c < values.length; c++) {
          const err = values[c].effectiveValue?.errorValue;
          if (err) {
            errors.push({
              sheet: title,
              cell: `${title}!${columnLetter(c)}${r + 1}`,
              errorType: err.type,
              message: err.message,
            });
          }
        }
      }
    }
  }
  return errors;
}
