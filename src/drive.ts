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

  const boundary = `drivempcboundary${Date.now()}`;
  const body =
    `--${boundary}\r\n` +
    `Content-Type: application/json; charset=UTF-8\r\n\r\n` +
    `${JSON.stringify(metadata)}\r\n` +
    `--${boundary}\r\n` +
    `Content-Type: ${opts.sourceMimeType}\r\n` +
    `Content-Transfer-Encoding: base64\r\n\r\n` +
    `${opts.contentBase64}\r\n` +
    `--${boundary}--`;

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
