import { createServer, IncomingMessage, ServerResponse } from "node:http";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { z } from "zod";
import { writeFile, mkdir, rename } from "node:fs/promises";
import { dirname } from "node:path";
import {
  listFiles, getFile, uploadFile, shareFile, deleteFile, downloadFile, exportFile,
} from "./drive.js";

const VERSION = "0.1.0";
const PORT = parseInt(process.env.PORT ?? "3000", 10);
const MCP_AUTH_TOKEN = process.env.MCP_AUTH_TOKEN ?? "";
const TOKEN_FILE = process.env.DRIVE_TOKEN_FILE ?? "/data/google_token.json";

function ok(data: unknown) {
  return { content: [{ type: "text" as const, text: JSON.stringify(data) }] };
}
function fail(err: unknown) {
  const msg = err instanceof Error ? err.message : String(err);
  return { content: [{ type: "text" as const, text: `ERROR: ${msg}` }], isError: true };
}

function buildMcpServer(): McpServer {
  const server = new McpServer({ name: "google-drive", version: VERSION });

  server.tool(
    "list_files",
    "List Drive files, optionally filtered by a Drive query string (e.g. \"name contains 'IEA'\") and/or a folderId. Returns id, name, mimeType, webViewLink for each.",
    {
      query: z.string().optional().describe("Drive API 'q' search syntax, e.g. \"name contains 'Register'\""),
      folderId: z.string().optional().describe("Restrict to files inside this folder"),
      pageSize: z.number().int().optional(),
    },
    async ({ query, folderId, pageSize }) => {
      try { return ok(await listFiles({ query, folderId, pageSize })); } catch (e) { return fail(e); }
    },
  );

  server.tool(
    "get_file",
    "Get metadata for a single Drive file by id (name, mimeType, webViewLink, size, modifiedTime).",
    { fileId: z.string() },
    async ({ fileId }) => {
      try { return ok(await getFile(fileId)); } catch (e) { return fail(e); }
    },
  );

  server.tool(
    "upload_and_convert_to_sheet",
    "Upload an already-generated .xlsx (base64) and have Drive convert it into a native, live-editable Google Sheet. " +
      "Use this to hand a human a shareable, collaboratively-editable copy of a file another tool already filled " +
      "(e.g. a filled Excel template) — this does NOT rebuild the spreadsheet from scratch, it converts the exact " +
      "bytes you already have. Note: Google's xlsx→Sheets conversion can flatten some Excel-specific features " +
      "(complex pivot table caches, certain conditional formats) — the original .xlsx you already have elsewhere " +
      "remains the source of truth; this is a review/editing copy, not a replacement artifact.",
    {
      name: z.string().describe("Filename to show in Drive, e.g. 'Cascade IEA Opportunity Register.xlsx'"),
      contentBase64: z.string().describe("Base64-encoded .xlsx bytes"),
      folderId: z.string().optional().describe("Drive folder id to upload into; omitted = user's root Drive"),
      shareWithEmail: z.string().optional().describe("If given, share the resulting Sheet with this email as a writer/commenter"),
      shareRole: z.enum(["reader", "commenter", "writer"]).optional().default("writer"),
    },
    async ({ name, contentBase64, folderId, shareWithEmail, shareRole }) => {
      try {
        const file = await uploadFile({
          name,
          contentBase64,
          sourceMimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          folderId,
          convertToGoogleFormat: true,
        });
        if (shareWithEmail) {
          await shareFile({ fileId: file.id, role: shareRole ?? "writer", type: "user", emailAddress: shareWithEmail });
        }
        return ok(file);
      } catch (e) { return fail(e); }
    },
  );

  server.tool(
    "upload_file",
    "Upload arbitrary file bytes (base64) to Drive as-is, without any Google-format conversion (e.g. for a PDF you just want stored/shared, not live-edited).",
    {
      name: z.string(),
      contentBase64: z.string(),
      mimeType: z.string().describe("Source MIME type, e.g. application/pdf"),
      folderId: z.string().optional(),
    },
    async ({ name, contentBase64, mimeType, folderId }) => {
      try {
        return ok(await uploadFile({ name, contentBase64, sourceMimeType: mimeType, folderId, convertToGoogleFormat: false }));
      } catch (e) { return fail(e); }
    },
  );

  server.tool(
    "share_file",
    "Grant access to a Drive file/Sheet — either to a specific email (role: reader/commenter/writer) or to anyone with the link.",
    {
      fileId: z.string(),
      role: z.enum(["reader", "commenter", "writer"]),
      type: z.enum(["anyone", "user"]),
      emailAddress: z.string().optional().describe("Required when type is 'user'"),
    },
    async ({ fileId, role, type, emailAddress }) => {
      try { return ok(await shareFile({ fileId, role, type, emailAddress })); } catch (e) { return fail(e); }
    },
  );

  server.tool(
    "download_file",
    "Download the raw bytes of a plain (non-Google-native) Drive file as base64.",
    { fileId: z.string() },
    async ({ fileId }) => {
      try { return ok({ contentBase64: await downloadFile(fileId) }); } catch (e) { return fail(e); }
    },
  );

  server.tool(
    "export_sheet_to_xlsx",
    "Export a native Google Sheet (or Doc/Slides) back to an Office format as base64 — the return path after a human has live-edited the Sheet, so the edited version can be re-saved as the .xlsx of record.",
    {
      fileId: z.string(),
      exportMimeType: z.string().optional().default("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    },
    async ({ fileId, exportMimeType }) => {
      try { return ok({ contentBase64: await exportFile(fileId, exportMimeType ?? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") }); } catch (e) { return fail(e); }
    },
  );

  server.tool(
    "delete_file",
    "Permanently delete a Drive file by id. Use with care — there is no undo via this tool.",
    { fileId: z.string() },
    async ({ fileId }) => {
      try { await deleteFile(fileId); return ok({ deleted: true, fileId }); } catch (e) { return fail(e); }
    },
  );

  return server;
}

// ---------------------------------------------------------------------------
// OAuth callback (initial consent + re-consent). Guarded: only exchanges +
// persists when the caller supplies ?key=<MCP_AUTH_TOKEN>, so a stranger who
// starts their own consent for our client_id cannot hijack the connection.
// ---------------------------------------------------------------------------
async function handleCallback(url: URL, res: ServerResponse) {
  const code = url.searchParams.get("code");
  const key = url.searchParams.get("key");
  const error = url.searchParams.get("error");

  if (error) {
    res.writeHead(400, { "Content-Type": "text/plain" });
    res.end(`Google returned error: ${error}`);
    return;
  }
  if (!code) {
    res.writeHead(400, { "Content-Type": "text/plain" });
    res.end("Missing ?code");
    return;
  }
  if (!MCP_AUTH_TOKEN || key !== MCP_AUTH_TOKEN) {
    res.writeHead(200, { "Content-Type": "text/plain" });
    res.end(`Authorization code received. To auto-store, re-run consent with &key=<MCP_AUTH_TOKEN> appended to the redirect, or exchange this code manually:\n\ncode=${code}`);
    return;
  }

  try {
    const redirectUri = process.env.GOOGLE_REDIRECT_URI ?? "https://google-drive.mcp.soapbox.build/callback";
    const resp = await fetch("https://oauth2.googleapis.com/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "authorization_code",
        code,
        redirect_uri: redirectUri,
        client_id: process.env.GOOGLE_CLIENT_ID ?? "",
        client_secret: process.env.GOOGLE_CLIENT_SECRET ?? "",
      }),
    });
    const json = (await resp.json()) as { access_token?: string; refresh_token?: string; expires_in?: number; error?: string };
    if (!resp.ok || !json.refresh_token) {
      res.writeHead(502, { "Content-Type": "application/json" });
      res.end(JSON.stringify(json));
      return;
    }
    await mkdir(dirname(TOKEN_FILE), { recursive: true });
    const tmp = `${TOKEN_FILE}.tmp`;
    await writeFile(tmp, JSON.stringify({
      refresh_token: json.refresh_token,
      access_token: json.access_token,
      expires_at: Date.now() + (json.expires_in ?? 3600) * 1000,
    }), { mode: 0o600 });
    await rename(tmp, TOKEN_FILE);
    res.writeHead(200, { "Content-Type": "text/plain" });
    res.end("✅ Google Drive connection stored. You may close this tab.");
  } catch (e) {
    res.writeHead(500, { "Content-Type": "text/plain" });
    res.end(`Exchange failed: ${e instanceof Error ? e.message : String(e)}`);
  }
}

// ---------------------------------------------------------------------------
// HTTP server
// ---------------------------------------------------------------------------
const httpServer = createServer(async (req: IncomingMessage, res: ServerResponse) => {
  const url = new URL(req.url ?? "/", `http://localhost:${PORT}`);

  if (url.pathname === "/health" && req.method === "GET") {
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ status: "ok", service: "google-drive-mcp", version: VERSION }));
    return;
  }

  if (url.pathname === "/callback" && req.method === "GET") {
    await handleCallback(url, res);
    return;
  }

  if (url.pathname === "/mcp" && (req.method === "POST" || req.method === "GET" || req.method === "DELETE")) {
    if (MCP_AUTH_TOKEN) {
      const auth = req.headers["authorization"] ?? "";
      if (auth !== `Bearer ${MCP_AUTH_TOKEN}`) {
        res.writeHead(401, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: "Unauthorized" }));
        return;
      }
    }

    const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
    const server = buildMcpServer();
    await server.connect(transport);

    let body: unknown;
    if (req.method === "POST") {
      body = await new Promise((resolve, reject) => {
        let raw = "";
        req.on("data", (c: Buffer) => (raw += c.toString()));
        req.on("end", () => { try { resolve(JSON.parse(raw)); } catch { resolve(undefined); } });
        req.on("error", reject);
      });
    }
    await transport.handleRequest(req, res, body);
    return;
  }

  res.writeHead(404, { "Content-Type": "application/json" });
  res.end(JSON.stringify({ error: "Not found" }));
});

httpServer.listen(PORT, () => {
  console.log(`Google Drive MCP server v${VERSION} listening on port ${PORT}`);
});
