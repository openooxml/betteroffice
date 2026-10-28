import { loadConfig, type WopiConfig } from "./config";
import { buildDiscovery, capabilities } from "./discovery";
import { resolveHost } from "./host/allowlist";
import { createHostClient, type HostClient } from "./host/client";
import { WopiHostError } from "./host/errors";
import {
  createProofSigner,
  loadKeyPairFile,
  loadOrCreateKeyPairFile,
  type ProofKeyPair,
} from "./proof";
import { existsSync, statSync } from "node:fs";
import { extname, resolve, sep } from "node:path";
import { CollabRelay, findByHandle } from "./relay/relay";
import { SessionError } from "./session/fileSession";
import { SessionRegistry } from "./session/registry";
import { decodeSaveStateVector, encodeJoinPayload, requestOrderedSave, SaveSyncError, sourceContentType } from "./sessionApi";

export interface WopiServerDeps {
  config: WopiConfig;
  currentKey: ProofKeyPair;
  oldKey?: ProofKeyPair;
  productVersion?: string;
  registry?: SessionRegistry;
  hostClient?: HostClient;
  relay?: CollabRelay;
  webRoot?: string;
  fontRoots?: string[];
  autosave?: boolean;
  saveSyncTimeoutMs?: number;
}

const STATIC_CONTENT_TYPES: Record<string, string> = {
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".wasm": "application/wasm",
  ".ttf": "font/ttf",
  ".otf": "font/otf",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".ico": "image/x-icon",
  ".map": "application/json; charset=utf-8",
};

const IMMUTABLE_CACHE = "public, max-age=31536000, immutable";

/** Resolves a request path inside a root, or null when it escapes. */
export function resolveStaticPath(root: string, requestPath: string): string | null {
  let decoded: string;
  try {
    decoded = decodeURIComponent(requestPath);
  } catch {
    return null;
  }
  if (decoded.includes("\\") || decoded.includes("\0")) return null;
  const segments = decoded.split("/").filter((segment) => segment.length > 0);
  if (segments.some((segment) => segment === ".." || segment === ".")) return null;
  const rootResolved = resolve(root);
  const candidate = resolve(rootResolved, ...segments);
  if (candidate !== rootResolved && !candidate.startsWith(rootResolved + sep)) return null;
  return candidate;
}

function staticFileResponse(filePath: string, cacheControl: string, method: string): Response | null {
  if (!existsSync(filePath) || !statSync(filePath).isFile()) return null;
  const contentType = STATIC_CONTENT_TYPES[extname(filePath).toLowerCase()] ?? "application/octet-stream";
  const file = Bun.file(filePath);
  return new Response(method === "HEAD" ? null : file, {
    status: 200,
    headers: {
      "Content-Type": contentType,
      "Content-Length": String(file.size),
      "Cache-Control": cacheControl,
      ...SECURITY_HEADERS,
    },
  });
}

/** Minimal upgrade surface of a Bun server. */
export interface UpgradeServer {
  upgrade(request: Request, options?: { headers?: Record<string, string>; data?: unknown }): boolean;
}

export interface WopiHandler {
  (request: Request, server?: UpgradeServer): Promise<Response>;
  readonly registry: SessionRegistry;
  readonly relay: CollabRelay;
}

const SECURITY_HEADERS = {
  "X-Content-Type-Options": "nosniff",
  "Referrer-Policy": "no-referrer",
} as const;

export const MAX_FORM_BODY_BYTES = 64 * 1024;
export const MAX_JSON_BODY_BYTES = 64 * 1024;

function respond(body: string, contentType: string, status = 200): Response {
  return new Response(body, {
    status,
    headers: { "Content-Type": contentType, ...SECURITY_HEADERS },
  });
}

function json(body: unknown, status = 200, extra: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8", ...SECURITY_HEADERS, ...extra },
  });
}

function typedError(kind: string, status: number): Response {
  return json({ error: { kind } }, status);
}

async function readLimited(request: Request, maxBytes: number): Promise<Uint8Array | null> {
  const length = request.headers.get("content-length");
  if (length !== null && /^\d+$/.test(length.trim()) && Number(length) > maxBytes) return null;
  if (!request.body) return new Uint8Array(await request.arrayBuffer());
  const reader = request.body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const chunk = await reader.read();
    if (chunk.done) break;
    total += chunk.value.byteLength;
    if (total > maxBytes) {
      try {
        await reader.cancel();
      } catch {}
      return null;
    }
    chunks.push(chunk.value);
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return out;
}

function escapeJsonForScript(value: unknown): string {
  return JSON.stringify(value)
    .replace(/</g, "\\u003c")
    .replace(/>/g, "\\u003e")
    .replace(/&/g, "\\u0026")
    .replace(/\u2028/g, "\\u2028")
    .replace(/\u2029/g, "\\u2029");
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function bearerHandle(request: Request): string | null {
  const header = request.headers.get("authorization");
  if (!header) return null;
  const match = /^Bearer\s+(.+)$/i.exec(header.trim());
  return match ? match[1]!.trim() : null;
}

export function resolvePort(env: Record<string, string | undefined> = process.env): number {
  const raw = env.PORT;
  if (raw === undefined || raw === "") return 3000;
  if (!/^\d+$/.test(raw)) throw new Error("PORT must be an integer between 1 and 65535");
  const value = Number(raw);
  if (value < 1 || value > 65535) throw new Error("PORT must be an integer between 1 and 65535");
  return value;
}

/** Request handler for the WOPI hosting endpoints. */
export function createHandler(deps: WopiServerDeps): WopiHandler {
  const discoveryXml = buildDiscovery({
    publicUrl: deps.config.publicUrl,
    currentKey: deps.currentKey,
    oldKey: deps.oldKey,
  });
  const capabilitiesJson = JSON.stringify(capabilities(deps.productVersion));
  const healthJson = JSON.stringify({ status: "ok" });

  let registry = deps.registry ?? deps.relay?.registry;
  let relay = deps.relay;
  const ensure = (): { registry: SessionRegistry; relay: CollabRelay } => {
    if (!registry) {
      const hostClient =
        deps.hostClient ??
        createHostClient({
          signer: createProofSigner(deps.currentKey, deps.oldKey),
          maxFileBytes: deps.config.maxFileBytes,
        });
      registry = new SessionRegistry({ config: deps.config, hostClient, autosave: deps.autosave });
    }
    if (!relay) relay = new CollabRelay(registry);
    return { registry, relay };
  };

  const handlePage = async (request: Request, url: URL, mode: "edit" | "view"): Promise<Response> => {
    if (request.method !== "POST") {
      return new Response("Method Not Allowed", {
        status: 405,
        headers: { "Content-Type": "text/plain; charset=utf-8", Allow: "POST", ...SECURITY_HEADERS },
      });
    }
    const wopiSrc = url.searchParams.get("WOPISrc");
    if (!wopiSrc) return typedError("badRequest", 400);
    let host;
    try {
      host = resolveHost(deps.config, wopiSrc);
    } catch {
      return typedError("notAllowed", 400);
    }
    const body = await readLimited(request, MAX_FORM_BODY_BYTES);
    if (body === null) return typedError("tooLarge", 413);
    const form = new URLSearchParams(new TextDecoder().decode(body));
    const accessToken = form.get("access_token");
    if (!accessToken) return typedError("badRequest", 400);
    const ttlRaw = form.get("access_token_ttl");
    const accessTokenTtl = ttlRaw !== null && /^-?\d+$/.test(ttlRaw.trim()) ? Number(ttlRaw) : 0;
    const { registry: sessions } = ensure();
    let joined;
    try {
      joined = await sessions.join({ wopiSrc, accessToken, accessTokenTtl, mode, hostId: host.id });
    } catch (err) {
      if (err instanceof SessionError) {
        const status =
          err.kind === "notAllowed" || err.kind === "unknownHost"
            ? 400
            : err.kind === "unsupportedFormat"
              ? 415
              : err.kind === "tooManyEditors"
                ? 429
                : err.kind === "closed"
                  ? 409
                  : 502;
        return typedError(err.kind, status);
      }
      if (err instanceof WopiHostError) {
        const status =
          err.kind === "notAllowed"
            ? 400
            : err.kind === "unauthorized"
              ? 401
              : err.kind === "tooLarge"
                ? 413
                : 502;
        return typedError(err.kind, status);
      }
      return typedError("unknown", 502);
    }
    const session = joined.session;
    const configJson = escapeJsonForScript({
      sessionHandle: joined.user.id,
      fileName: session.fileName || joined.file.name,
      format: session.format,
      canWrite: joined.user.canWrite,
      viewOnlyReason: joined.viewOnlyReason ?? session.viewOnlyReason ?? null,
      postMessageOrigin: joined.user.postMessageOrigin,
      closeUrl: joined.user.closeUrl,
      closePostMessage: joined.user.closePostMessage,
      locale: url.searchParams.get("ui"),
      userFriendlyName: joined.user.friendlyName,
      tokenExpiresAt: joined.user.expiresAt,
    });
    const ancestors = host.frameOrigins.length > 0 ? host.frameOrigins.join(" ") : "'none'";
    let collabOrigin = "";
    try {
      const publicUrl = new URL(deps.config.publicUrl);
      collabOrigin = `${publicUrl.protocol === "https:" ? "wss:" : "ws:"}//${publicUrl.host}`;
    } catch {}
    const csp = [
      `frame-ancestors ${ancestors}`,
      "default-src 'self'",
      collabOrigin ? `connect-src 'self' ${collabOrigin}` : "connect-src 'self'",
      "img-src 'self' blob: data:",
      "font-src 'self' data:",
      "style-src 'self' 'unsafe-inline'",
      "worker-src 'self' blob:",
      "script-src 'self' 'wasm-unsafe-eval'",
    ].join("; ");
    const html = [
      "<!doctype html>",
      '<html lang="en">',
      `<head><meta charset="utf-8"><title>${escapeHtml(session.fileName || "BetterOffice")}</title><link rel="stylesheet" href="/app/main.css"></head>`,
      "<body>",
      '<div id="root"></div>',
      `<script type="application/json" id="wopi-config">${configJson}</script>`,
      '<script type="module" src="/app/main.js"></script>',
      "</body>",
      "</html>",
      "",
    ].join("\n");
    return new Response(html, {
      status: 200,
      headers: {
        "Content-Type": "text/html; charset=utf-8",
        "Content-Security-Policy": csp,
        "Cache-Control": "no-store",
        ...SECURITY_HEADERS,
      },
    });
  };

  const authenticate = (request: Request) => {
    const handle = bearerHandle(request);
    if (!handle) return null;
    return findByHandle(ensure().registry, handle);
  };

  const handleApi = async (request: Request, url: URL): Promise<Response> => {
    const resolved = authenticate(request);
    if (!resolved) return typedError("unauthorized", 401);
    const { session, user } = resolved;
    const path = url.pathname;
    if (path === "/api/session/source") {
      if (request.method !== "GET") return typedError("methodNotAllowed", 405);
      const bytes = session.sourceBytes;
      if (!bytes) return typedError("notFound", 404);
      return new Response(bytes as BodyInit, {
        status: 200,
        headers: {
          "Content-Type": sourceContentType(session.format),
          "Content-Length": String(bytes.byteLength),
          "Cache-Control": "no-store",
          ...SECURITY_HEADERS,
        },
      });
    }
    if (path === "/api/session/join") {
      if (request.method !== "GET") return typedError("methodNotAllowed", 405);
      const replica = session.replica;
      if (!replica) return typedError("closed", 409);
      return json(encodeJoinPayload(replica.joinPayload()), 200, { "Cache-Control": "no-store" });
    }
    if (path === "/api/session/save") {
      if (request.method !== "POST") return typedError("methodNotAllowed", 405);
      const body = await readLimited(request, MAX_JSON_BODY_BYTES);
      if (body === null) return typedError("tooLarge", 413);
      try {
        const stateVector = decodeSaveStateVector(body);
        const result = await requestOrderedSave(session, user.id, stateVector, { timeoutMs: deps.saveSyncTimeoutMs, signal: request.signal });
        if (!result.saved && session.dirty) return typedError("saveFailed", 502);
        return json({ itemVersion: result.itemVersion }, 200, { "Cache-Control": "no-store" });
      } catch (err) {
        if (err instanceof SaveSyncError) return typedError(err.kind, err.kind === "notSynced" ? 409 : 400);
        if (err instanceof SessionError) {
          return typedError(err.kind, err.kind === "closed" ? 409 : err.kind === "unknownUser" ? 401 : 502);
        }
        return typedError("saveFailed", 502);
      }
    }
    if (path === "/api/session/close") {
      if (request.method !== "POST") return typedError("methodNotAllowed", 405);
      const body = await readLimited(request, MAX_JSON_BODY_BYTES);
      if (body === null) return typedError("tooLarge", 413);
      try {
        await session.leave(user.id);
      } catch (err) {
        if (err instanceof SessionError && err.kind !== "unknownUser") return typedError(err.kind, 502);
      } finally {
        ensure().relay.closeUser(user.id, 1000, "session closed");
      }
      return json({ closed: true }, 200, { "Cache-Control": "no-store" });
    }
    if (path === "/api/session/events") {
      if (request.method !== "GET") return typedError("methodNotAllowed", 405);
      return eventsResponse(session, user.id);
    }
    return respond("Not Found", "text/plain; charset=utf-8", 404);
  };

  const handler = async (request: Request, server?: UpgradeServer): Promise<Response> => {
    const url = new URL(request.url);
    if (request.method === "GET" && url.pathname === "/hosting/discovery") {
      return respond(discoveryXml, "text/xml; charset=utf-8");
    }
    if (request.method === "GET" && url.pathname === "/hosting/capabilities") {
      return respond(capabilitiesJson, "application/json; charset=utf-8");
    }
    if (request.method === "GET" && url.pathname === "/healthz") {
      return respond(healthJson, "application/json; charset=utf-8");
    }
    if (url.pathname === "/wopi/edit") return handlePage(request, url, "edit");
    if (url.pathname === "/wopi/view") return handlePage(request, url, "view");
    if ((request.method === "GET" || request.method === "HEAD") && (url.pathname === "/app" || url.pathname.startsWith("/app/"))) {
      const webRoot = deps.webRoot ?? resolve(import.meta.dir, "../web/dist");
      const relative = url.pathname.slice("/app".length) || "/main.js";
      const filePath = resolveStaticPath(webRoot, relative);
      if (filePath === null) return respond("Not Found", "text/plain; charset=utf-8", 404);
      const base = filePath.split(sep).pop() ?? "";
      const cache = base === "main.js" || base === "main.css" || base.endsWith(".html") ? "no-cache" : IMMUTABLE_CACHE;
      return staticFileResponse(filePath, cache, request.method) ?? respond("Not Found", "text/plain; charset=utf-8", 404);
    }
    if ((request.method === "GET" || request.method === "HEAD") && url.pathname.startsWith("/fonts/")) {
      const fontRoots = deps.fontRoots ?? [
        resolve(import.meta.dir, "../../../packages/fonts/assets"),
        resolve(import.meta.dir, "../../../packages/fonts-cjk/assets"),
      ];
      const relative = url.pathname.slice("/fonts".length);
      for (const root of fontRoots) {
        const filePath = resolveStaticPath(root, relative);
        if (filePath === null) return respond("Not Found", "text/plain; charset=utf-8", 404);
        const found = staticFileResponse(filePath, IMMUTABLE_CACHE, request.method);
        if (found) return found;
      }
      return respond("Not Found", "text/plain; charset=utf-8", 404);
    }
    if (url.pathname.startsWith("/api/session/")) return handleApi(request, url);
    if (url.pathname === "/collab") {
      if (request.method !== "GET") return typedError("methodNotAllowed", 405);
      const protocolHeader = request.headers.get("sec-websocket-protocol") ?? "";
      const protocols = protocolHeader.split(",").map((entry) => entry.trim()).filter(Boolean);
      if (protocols[0] !== "bo-wopi.v1" || !protocols[1]) return typedError("unauthorized", 401);
      const resolved = ensure().relay.resolve(protocols[1]);
      if (!resolved) return typedError("unauthorized", 401);
      if (!server) return typedError("upgradeUnavailable", 503);
      const upgraded = server.upgrade(request, {
        headers: { "Sec-WebSocket-Protocol": "bo-wopi.v1" },
        data: { handle: protocols[1] },
      });
      if (!upgraded) return typedError("badRequest", 400);
      return undefined as unknown as Response;
    }
    return respond("Not Found", "text/plain; charset=utf-8", 404);
  };

  return Object.assign(handler, {
    get registry(): SessionRegistry {
      return ensure().registry;
    },
    get relay(): CollabRelay {
      return ensure().relay;
    },
  });
}

function eventsResponse(session: import("./session/fileSession").FileSession, userId: string): Response {
  const encoder = new TextEncoder();
  const unsubscribes: Array<() => void> = [];
  let controller: ReadableStreamDefaultController<Uint8Array> | null = null;
  let heartbeat: ReturnType<typeof setInterval> | null = null;
  let cleaned = false;
  const cleanup = (): void => {
    if (cleaned) return;
    cleaned = true;
    if (heartbeat !== null) clearInterval(heartbeat);
    for (const unsubscribe of unsubscribes) unsubscribe();
  };
  const send = (event: string, data: Record<string, unknown>): void => {
    if (!controller || cleaned) return;
    try {
      controller.enqueue(encoder.encode(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`));
    } catch {}
  };
  const stream = new ReadableStream<Uint8Array>({
    start(ctrl) {
      controller = ctrl;
      try {
        ctrl.enqueue(encoder.encode(": connected\n\n"));
      } catch {}
      heartbeat = setInterval(() => {
        try {
          ctrl.enqueue(encoder.encode(": heartbeat\n\n"));
        } catch {}
      }, 15000);
      (heartbeat as unknown as { unref?: () => void }).unref?.();
      unsubscribes.push(
        session.on("saved", (p) => send("saved", { itemVersion: p.itemVersion, by: p.by === userId ? p.by : null })),
        session.on("saveFailed", (p) => send("saveFailed", { kind: p.kind })),
        session.on("conflict", (p) => send("conflict", { name: p.name, url: p.url, copyName: p.copyName, copyUrl: p.copyUrl })),
        session.on("viewOnly", (p) => send("viewOnly", { reason: p.reason })),
        session.on("tokenExpiring", (p) => {
          if (p.userId === userId) send("tokenExpiring", { expiresAt: p.expiresAt });
        }),
        session.on("leaderChanged", (p) =>
          send("leaderChanged", {
            leaderId: p.leaderId === userId ? p.leaderId : null,
            previousLeaderId: p.previousLeaderId === userId ? p.previousLeaderId : null,
          }),
        ),
        session.on("closed", () => {
          send("closed", {});
          cleanup();
          try {
            controller?.close();
          } catch {}
        }),
      );
    },
    cancel() {
      cleanup();
    },
  });
  return new Response(stream, {
    status: 200,
    headers: {
      "Content-Type": "text/event-stream",
      "Cache-Control": "no-store",
      Connection: "keep-alive",
      ...SECURITY_HEADERS,
    },
  });
}

export function loadServerDeps(config: WopiConfig): WopiServerDeps {
  const oldKey = config.proofOldKeyFile ? loadKeyPairFile(config.proofOldKeyFile) : undefined;
  const currentKey = loadOrCreateKeyPairFile(config.proofKeyFile);
  return { config, currentKey, oldKey };
}

export function startServer(deps: WopiServerDeps, port = resolvePort()) {
  const handler = createHandler(deps);
  return Bun.serve({
    port,
    fetch: (request, server) => handler(request, server as unknown as UpgradeServer),
    websocket: handler.relay.websocket as never,
  });
}

if ((import.meta as { main?: boolean }).main) {
  const config = loadConfig();
  const server = startServer(loadServerDeps(config));
  console.log(`WOPI app listening on ${server.url}`);
}
