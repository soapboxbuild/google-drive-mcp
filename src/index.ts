import { createServer, IncomingMessage, ServerResponse } from "node:http";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { z } from "zod";
import {
  listFiles, getFile, uploadFile, shareFile, deleteFile, downloadFile, exportFile,
} from "./drive.js";

const VERSION = "0.2.0";
const PORT = parseInt(process.env.PORT ?? "3000", 10);

function ok(data: unknown) {
  return { content: [{ type: "text" as const, text: JSON.stringify(data) }] };
}
function fail(err: unknown) {
  const msg = err instanceof Error ? err.message : String(err);
  return { content: [{ type: "text" as const, text: `ERROR: ${msg}` }], isError: true };
}

// Each request's bearer token IS the Google access token for that call — see
// drive.ts's header comment. Minted per-portfolio via soapbox-platform's
// /api/oauth/google-drive flow, auto-refreshed by Anthropic's credential
// vault (mcp_oauth), so this server never stores or refreshes its own token.
function buildMcpServer(accessToken: string): McpServer {
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
      try { return ok(await listFiles(accessToken, { query, folderId, pageSize })); } catch (e) { return fail(e); }
    },
  );

  server.tool(
    "get_file",
    "Get metadata for a single Drive file by id (name, mimeType, webViewLink, size, modifiedTime).",
    { fileId: z.string() },
    async ({ fileId }) => {
      try { return ok(await getFile(accessToken, fileId)); } catch (e) { return fail(e); }
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
        const file = await uploadFile(accessToken, {
          name,
          contentBase64,
          sourceMimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          folderId,
          convertToGoogleFormat: true,
        });
        if (shareWithEmail) {
          await shareFile(accessToken, { fileId: file.id, role: shareRole ?? "writer", type: "user", emailAddress: shareWithEmail });
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
        return ok(await uploadFile(accessToken, { name, contentBase64, sourceMimeType: mimeType, folderId, convertToGoogleFormat: false }));
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
      try { return ok(await shareFile(accessToken, { fileId, role, type, emailAddress })); } catch (e) { return fail(e); }
    },
  );

  server.tool(
    "download_file",
    "Download the raw bytes of a plain (non-Google-native) Drive file as base64.",
    { fileId: z.string() },
    async ({ fileId }) => {
      try { return ok({ contentBase64: await downloadFile(accessToken, fileId) }); } catch (e) { return fail(e); }
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
      try { return ok({ contentBase64: await exportFile(accessToken, fileId, exportMimeType ?? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") }); } catch (e) { return fail(e); }
    },
  );

  server.tool(
    "delete_file",
    "Permanently delete a Drive file by id. Use with care — there is no undo via this tool.",
    { fileId: z.string() },
    async ({ fileId }) => {
      try { await deleteFile(accessToken, fileId); return ok({ deleted: true, fileId }); } catch (e) { return fail(e); }
    },
  );

  return server;
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

  if (url.pathname === "/mcp" && (req.method === "POST" || req.method === "GET" || req.method === "DELETE")) {
    // The incoming bearer token IS the Google access token for this call —
    // see drive.ts's header comment. No static shared-secret check here:
    // whichever portfolio's OAuth connection this is, its own token gates it.
    const auth = req.headers["authorization"] ?? "";
    const accessToken = auth.startsWith("Bearer ") ? auth.slice("Bearer ".length) : "";

    const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
    const server = buildMcpServer(accessToken);
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
