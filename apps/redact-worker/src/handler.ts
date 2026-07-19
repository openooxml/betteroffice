export interface Env {
  REDACTED_BUCKET?: R2Bucket;
  PUBLIC_BASE_URL?: string;
}

export type Sanitizer = (
  bytes: Uint8Array,
  format: SupportedFormat,
) => Uint8Array | Promise<Uint8Array>;

export type SupportedFormat = "docx" | "xlsx" | "pptx";

interface FormatSpec {
  contentType: string;
}

interface WorkerOptions {
  rateLimit?: number;
  rateWindowMs?: number;
  now?: () => number;
}

const MAX_FILE_BYTES = 64 * 1024 * 1024;
const DEFAULT_RATE_LIMIT = 10;
const DEFAULT_RATE_WINDOW_MS = 60 * 60 * 1000;
const FORMATS: Record<SupportedFormat, FormatSpec> = {
  docx: {
    contentType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  },
  xlsx: {
    contentType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  },
  pptx: {
    contentType:
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  },
};

export function createWorker(sanitize: Sanitizer, options: WorkerOptions = {}) {
  const attempts = new Map<string, number[]>();
  const limit = options.rateLimit ?? DEFAULT_RATE_LIMIT;
  const windowMs = options.rateWindowMs ?? DEFAULT_RATE_WINDOW_MS;
  const now = options.now ?? Date.now;

  return {
    async fetch(request: Request, env: Env): Promise<Response> {
      const url = new URL(request.url);
      if (request.method === "POST" && url.pathname === "/upload") {
        if (!allowUpload(request, attempts, limit, windowMs, now())) {
          return errorResponse(429, "upload rate limit exceeded");
        }
        return upload(request, env, sanitize);
      }
      const file = url.pathname.match(/^\/f\/([a-f0-9]{32})$/);
      if (request.method === "GET" && file) {
        return retrieve(file[1], env);
      }
      return errorResponse(404, "not found");
    },
  };
}

async function upload(
  request: Request,
  env: Env,
  sanitize: Sanitizer,
): Promise<Response> {
  if (!env.REDACTED_BUCKET) {
    return errorResponse(503, "storage binding is not configured");
  }
  const format = request.headers.get("X-BetterOffice-Format")?.toLowerCase();
  if (!isSupportedFormat(format)) {
    return errorResponse(415, "DOCX, XLSX, or PPTX format header required");
  }
  const contentType = request.headers.get("Content-Type")?.split(";", 1)[0].trim();
  const spec = FORMATS[format];
  if (contentType !== spec.contentType) {
    return errorResponse(415, "content type does not match OOXML format");
  }

  let source: Uint8Array;
  try {
    source = await readBoundedBody(request);
  } catch (error) {
    return errorResponse(413, error instanceof Error ? error.message : "file too large");
  }
  if (source.byteLength === 0) {
    return errorResponse(400, "empty upload");
  }

  let sanitized: Uint8Array;
  try {
    sanitized = await sanitize(source, format);
  } catch {
    return errorResponse(400, "OOXML sanitizer rejected the file");
  }
  if (sanitized.byteLength > MAX_FILE_BYTES) {
    return errorResponse(413, "sanitized file exceeds size limit");
  }

  const id = crypto.randomUUID().replaceAll("-", "");
  try {
    await env.REDACTED_BUCKET.put(id, sanitized, {
      httpMetadata: { contentType: spec.contentType },
      customMetadata: { format },
    });
  } catch {
    return errorResponse(503, "storage is unavailable");
  }
  const base = env.PUBLIC_BASE_URL?.replace(/\/$/, "") ?? new URL(request.url).origin;
  return jsonResponse(
    { id, url: `${base}/f/${id}` },
    200,
    { "Cache-Control": "no-store" },
  );
}

async function retrieve(id: string, env: Env): Promise<Response> {
  if (!env.REDACTED_BUCKET) {
    return errorResponse(503, "storage binding is not configured");
  }
  let object: R2ObjectBody | null;
  try {
    object = await env.REDACTED_BUCKET.get(id);
  } catch {
    return errorResponse(503, "storage is unavailable");
  }
  if (!object) return errorResponse(404, "file not found");
  const format = object.customMetadata?.format;
  const contentType = isSupportedFormat(format)
    ? FORMATS[format].contentType
    : object.httpMetadata?.contentType ?? "application/octet-stream";
  const headers = new Headers({
    "Cache-Control": "public, max-age=31536000, immutable",
    "Content-Disposition": "attachment",
    "Content-Length": object.size.toString(),
    "Content-Type": contentType,
    ETag: object.httpEtag,
    "X-Content-Type-Options": "nosniff",
  });
  return new Response(object.body, { headers });
}

async function readBoundedBody(request: Request): Promise<Uint8Array> {
  const declared = request.headers.get("Content-Length");
  if (declared !== null) {
    const length = Number(declared);
    if (!Number.isSafeInteger(length) || length < 0 || length > MAX_FILE_BYTES) {
      throw new Error("file exceeds 64 MiB limit");
    }
  }
  if (!request.body) return new Uint8Array();
  const reader = request.body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > MAX_FILE_BYTES) {
      await reader.cancel();
      throw new Error("file exceeds 64 MiB limit");
    }
    chunks.push(value);
  }
  const output = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return output;
}

function allowUpload(
  request: Request,
  attempts: Map<string, number[]>,
  limit: number,
  windowMs: number,
  now: number,
): boolean {
  const client = request.headers.get("CF-Connecting-IP") ?? "unknown";
  const recent = (attempts.get(client) ?? []).filter(
    (timestamp) => now - timestamp < windowMs,
  );
  if (recent.length >= limit) {
    attempts.set(client, recent);
    return false;
  }
  recent.push(now);
  attempts.set(client, recent);
  return true;
}

function isSupportedFormat(value: string | undefined): value is SupportedFormat {
  return value === "docx" || value === "xlsx" || value === "pptx";
}

function errorResponse(status: number, error: string): Response {
  return jsonResponse({ error }, status, { "Cache-Control": "no-store" });
}

function jsonResponse(
  body: unknown,
  status: number,
  headers: Record<string, string>,
): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", ...headers },
  });
}
