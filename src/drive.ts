import { readFile, writeFile, mkdir, rename } from "node:fs/promises";
import { dirname } from "node:path";

// ---------------------------------------------------------------------------
// Token store — this service is the SOLE owner of the Google OAuth refresh
// token, mirroring the xero-mcp pattern (see ~/xero-mcp/src/xero.ts). Google
// does not rotate the refresh token on every use the way Xero does, but we
// still persist it durably (a Railway volume, NOT ephemeral container disk)
// so a redeploy never loses the connection.
// ---------------------------------------------------------------------------

const TOKEN_FILE = process.env.DRIVE_TOKEN_FILE ?? "/data/google_token.json";
const SEED_REFRESH = process.env.GOOGLE_REFRESH_TOKEN; // used only to seed the store on first boot
const TOKEN_URL = "https://oauth2.googleapis.com/token";

// Checked lazily (only when a tool actually needs to talk to Google), not at
// module load — so the service can deploy and pass its health check before
// the Google Cloud OAuth app exists yet, and only fails the specific tool
// call that needed credentials, with a clear error, instead of crash-looping.
function requireEnv(name: string): string {
  const v = process.env[name];
  if (!v) throw new Error(`Missing required env var ${name} — OAuth app not configured yet (see README Setup).`);
  return v;
}

interface StoredToken {
  refresh_token: string;
  access_token?: string;
  expires_at?: number; // epoch ms
}

async function readStore(): Promise<StoredToken | null> {
  try {
    const raw = await readFile(TOKEN_FILE, "utf8");
    return JSON.parse(raw) as StoredToken;
  } catch {
    return null;
  }
}

async function writeStore(tok: StoredToken): Promise<void> {
  await mkdir(dirname(TOKEN_FILE), { recursive: true });
  const tmp = `${TOKEN_FILE}.tmp`;
  await writeFile(tmp, JSON.stringify(tok), { mode: 0o600 });
  await rename(tmp, TOKEN_FILE); // atomic replace
}

// Serialize refreshes so two concurrent tool calls never race.
let refreshLock: Promise<string> | null = null;

async function doRefresh(refreshToken: string): Promise<string> {
  const body = new URLSearchParams({
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    client_id: requireEnv("GOOGLE_CLIENT_ID"),
    client_secret: requireEnv("GOOGLE_CLIENT_SECRET"),
  });
  const resp = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(
      `Google token refresh failed (${resp.status}): ${text}. ` +
        `If this is 'invalid_grant', the refresh token was revoked or the OAuth consent screen is still ` +
        `in "Testing" mode (7-day token expiry) — publish the app or re-consent.`,
    );
  }
  const json = (await resp.json()) as {
    access_token: string;
    expires_in: number;
    refresh_token?: string; // Google only returns this on the FIRST exchange, not every refresh
  };
  const existing = await readStore();
  const store: StoredToken = {
    refresh_token: json.refresh_token ?? existing?.refresh_token ?? refreshToken,
    access_token: json.access_token,
    expires_at: Date.now() + json.expires_in * 1000,
  };
  await writeStore(store);
  return json.access_token;
}

/** Return a valid access token, refreshing (and persisting) as needed. */
export async function getAccessToken(): Promise<string> {
  let store = await readStore();

  // First boot: seed the durable store from the env-provided refresh token.
  if (!store) {
    if (!SEED_REFRESH) {
      throw new Error(
        `No token store at ${TOKEN_FILE} and no GOOGLE_REFRESH_TOKEN to seed it. ` +
          `Run the one-time OAuth consent (see README) and set GOOGLE_REFRESH_TOKEN, ` +
          `or hit /callback?code=...&key=<MCP_AUTH_TOKEN> after consenting.`,
      );
    }
    store = { refresh_token: SEED_REFRESH };
    await writeStore(store);
  }

  // Reuse a still-valid access token (60s safety margin).
  if (store.access_token && store.expires_at && store.expires_at > Date.now() + 60_000) {
    return store.access_token;
  }

  if (!refreshLock) {
    const rt = store.refresh_token;
    refreshLock = doRefresh(rt).finally(() => {
      refreshLock = null;
    });
  }
  return refreshLock;
}

// ---------------------------------------------------------------------------
// Google Drive API v3 client — thin fetch wrapper, no googleapis SDK
// dependency (keeps the service lean, matching the xero-mcp/costing-mcp
// style of direct REST calls over a heavy client library).
// ---------------------------------------------------------------------------

const DRIVE_API = "https://www.googleapis.com/drive/v3";
const UPLOAD_API = "https://www.googleapis.com/upload/drive/v3";

async function authedFetch(url: string, init: RequestInit = {}): Promise<Response> {
  const token = await getAccessToken();
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${token}`);
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

export async function listFiles(opts: {
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
  const resp = await authedFetch(url.toString());
  const json = (await resp.json()) as { files: DriveFile[] };
  return json.files;
}

export async function getFile(fileId: string): Promise<DriveFile> {
  const url = new URL(`${DRIVE_API}/files/${fileId}`);
  url.searchParams.set("fields", FILE_FIELDS);
  const resp = await authedFetch(url.toString());
  return resp.json() as Promise<DriveFile>;
}

/**
 * Upload file bytes. If `convertToGoogleFormat` is true and the source is an
 * Office format (xlsx/docx/pptx), Drive auto-converts it into the matching
 * Google Workspace format (Sheets/Docs/Slides) on upload — this is what makes
 * an already-filled .xlsx live-editable/collaborative without rebuilding it
 * from scratch in some other tool.
 */
export async function uploadFile(opts: {
  name: string;
  contentBase64: string;
  sourceMimeType: string;
  folderId?: string;
  convertToGoogleFormat?: boolean;
}): Promise<DriveFile> {
  const bytes = Buffer.from(opts.contentBase64, "base64");

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

  const resp = await authedFetch(url.toString(), {
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

export async function shareFile(opts: {
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
  const resp = await authedFetch(url.toString(), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  return resp.json() as Promise<{ id: string }>;
}

export async function deleteFile(fileId: string): Promise<void> {
  await authedFetch(`${DRIVE_API}/files/${fileId}`, { method: "DELETE" });
}

/** Download raw bytes (for a plain uploaded file, not a native Google format). */
export async function downloadFile(fileId: string): Promise<string> {
  const resp = await authedFetch(`${DRIVE_API}/files/${fileId}?alt=media`);
  const buf = Buffer.from(await resp.arrayBuffer());
  return buf.toString("base64");
}

/** Export a native Google format (Sheets/Docs/Slides) back to an Office format. */
export async function exportFile(fileId: string, exportMimeType: string): Promise<string> {
  const url = new URL(`${DRIVE_API}/files/${fileId}/export`);
  url.searchParams.set("mimeType", exportMimeType);
  const resp = await authedFetch(url.toString());
  const buf = Buffer.from(await resp.arrayBuffer());
  return buf.toString("base64");
}
