import { createPublicKey, randomUUID, verify, type KeyObject } from "node:crypto";
import {
  buildProofBytes,
  proofKeyAttributes,
  type ProofKeyAttributes,
  type ProofKeyInput,
  type ProofKeyPair,
  type ProofSigner,
} from "../../src/proof";
import { cancelStream } from "../../src/host/streams";
import { decodeUtf7, encodeUtf7 } from "../../src/host/utf7";

export interface ReferenceToken {
  userId: string;
  userName?: string;
  canWrite?: boolean;
  userCanWrite?: boolean;
  expiresAt?: number;
  fileIds?: Set<string>;
}

export interface ReferenceFile {
  id: string;
  name: string;
  bytes: Uint8Array;
  version: number;
  itemVersion: string;
  lock: string | null;
  lockExpiresAt: number | null;
  lastModifiedTime: string;
}

export interface ReferenceHostQuirks {
  timestampConflict?: boolean;
  tokenExpired?: boolean;
  slowBody?: boolean;
  noContentLength?: boolean;
}

export type VerifyProofInput =
  | ProofKeyInput
  | ProofKeyPair
  | ProofKeyAttributes
  | ProofSigner
  | Array<ProofKeyInput | ProofKeyAttributes>
  | { current?: unknown; old?: unknown; currentKey?: unknown; oldKey?: unknown };

export interface ReferenceHostOptions {
  verifyProof?: VerifyProofInput;
  now?: () => number;
  quirks?: ReferenceHostQuirks;
}

export interface ReferenceHost {
  fetch(request: Request): Promise<Response>;
  files: Map<string, ReferenceFile>;
  tokens: Map<string, ReferenceToken | string>;
  setFile(id: string, bytes: Uint8Array | string, options?: { name?: string; lastModifiedTime?: string }): ReferenceFile;
}

const LOCK_TTL_MS = 30 * 60 * 1000;
const PROOF_WINDOW_MS = 20 * 60 * 1000;
const DOTNET_EPOCH_TICKS = 621355968000000000n;

function publicKeyFromAttributes(attributes: ProofKeyAttributes): KeyObject {
  const toBase64Url = (base64: string) => Buffer.from(base64, "base64").toString("base64url");
  return createPublicKey({
    format: "jwk",
    key: { kty: "RSA", n: toBase64Url(attributes.modulus), e: toBase64Url(attributes.exponent) },
  });
}

function isAttributes(value: unknown): value is ProofKeyAttributes {
  return typeof value === "object" && value !== null && "modulus" in value && "exponent" in value;
}

function verificationKey(input: unknown): KeyObject | undefined {
  if (input === undefined || input === null) return undefined;
  try {
    return publicKeyFromAttributes(isAttributes(input) ? input : proofKeyAttributes(input as ProofKeyInput));
  } catch {
    return undefined;
  }
}

function verificationKeys(input: VerifyProofInput | undefined): { current?: KeyObject; old?: KeyObject } {
  if (input === undefined || input === null) return {};
  if (Array.isArray(input)) {
    return { current: verificationKey(input[0]), old: verificationKey(input[1]) };
  }
  if (typeof input === "object" && "signProof" in input && "proofKeyAttributes" in input) {
    const signer = input as ProofSigner;
    return { current: verificationKey(signer.proofKeyAttributes()), old: verificationKey(signer.oldProofKeyAttributes()) };
  }
  if (typeof input === "object" && ("current" in input || "old" in input || "currentKey" in input || "oldKey" in input)) {
    const record = input as Record<string, unknown>;
    return { current: verificationKey(record.current ?? record.currentKey), old: verificationKey(record.old ?? record.oldKey) };
  }
  return { current: verificationKey(input) };
}

function fileIdFromPath(pathname: string): string {
  const withoutContents = pathname.endsWith("/contents") ? pathname.slice(0, -"/contents".length) : pathname;
  const segments = withoutContents.split("/").filter((segment) => segment.length > 0);
  const last = segments[segments.length - 1] ?? "file1";
  try {
    return decodeURIComponent(last);
  } catch {
    return last;
  }
}

function extensionOf(name: string): string {
  const index = name.lastIndexOf(".");
  return index > 0 ? name.slice(index) : "";
}

function json(body: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", ...headers },
  });
}

function empty(status = 200, headers: Record<string, string> = {}): Response {
  return new Response(null, { status, headers });
}

function lockConflict(currentLock: string | null, reason = "Lock mismatch"): Response {
  return empty(409, {
    "X-WOPI-Lock": currentLock ?? "",
    "X-WOPI-LockFailureReason": reason,
  });
}

/** In process WOPI host for tests. */
export function createReferenceHost(options: ReferenceHostOptions = {}): ReferenceHost {
  const now = options.now ?? (() => Date.now());
  const quirks = options.quirks ?? {};
  const files = new Map<string, ReferenceFile>();
  const tokens = new Map<string, ReferenceToken | string>();
  const keys = verificationKeys(options.verifyProof);
  let nextFileNumber = 1;

  function uniqueFileId(): string {
    let candidate = `generated-${nextFileNumber}`;
    while (files.has(candidate)) {
      nextFileNumber += 1;
      candidate = `generated-${nextFileNumber}`;
    }
    nextFileNumber += 1;
    return candidate;
  }

  function resolveToken(token: string | null): ReferenceToken | null {
    if (quirks.tokenExpired) return null;
    if (!token) return null;
    const entry = tokens.get(token);
    if (entry !== undefined) {
      if (typeof entry === "string") return { userId: entry, canWrite: true };
      if (entry.expiresAt !== undefined && now() >= entry.expiresAt) return null;
      return entry;
    }
    if (tokens.size === 0) return { userId: "user1", userName: "User One", canWrite: true };
    return null;
  }

  function relativeUrl(request: Request, fileId: string, user: ReferenceToken): string {
    const url = new URL(request.url);
    const sourceToken = url.searchParams.get("access_token");
    if (sourceToken !== null && !tokens.has(sourceToken)) tokens.set(sourceToken, user);
    user.fileIds?.add(fileId);
    const token = randomUUID();
    tokens.set(token, user);
    url.pathname = url.pathname.replace(/[^/]+$/, encodeURIComponent(fileId));
    url.search = "";
    url.searchParams.set("access_token", token);
    return url.toString();
  }

  function verifyRequestProof(request: Request, token: string): boolean {
    if (options.verifyProof === undefined) return true;
    const timestampHeader = request.headers.get("X-WOPI-TimeStamp");
    if (!timestampHeader) return false;
    let ticks: bigint;
    try {
      ticks = BigInt(timestampHeader);
    } catch {
      return false;
    }
    const timestampMs = Number((ticks - DOTNET_EPOCH_TICKS) / 10000n);
    if (Math.abs(now() - timestampMs) > PROOF_WINDOW_MS) return false;
    const bytes = buildProofBytes(token, request.url, ticks);
    for (const [headerName, key] of [
      ["X-WOPI-Proof", keys.current],
      ["X-WOPI-Proof", keys.old],
      ["X-WOPI-ProofOld", keys.current],
    ] as const) {
      const proof = request.headers.get(headerName);
      if (!proof || !key) continue;
      const signature = Buffer.from(proof, "base64");
      try {
        if (verify("RSA-SHA256", bytes, key, signature)) return true;
      } catch {}
    }
    return false;
  }

  function liveLock(file: ReferenceFile): string | null {
    if (file.lock !== null && file.lockExpiresAt !== null && now() > file.lockExpiresAt) {
      file.lock = null;
      file.lockExpiresAt = null;
    }
    return file.lock;
  }

  function timestampConflict(request: Request, file: ReferenceFile): Response | null {
    const stamp = request.headers.get("X-COOL-WOPI-Timestamp");
    return quirks.timestampConflict && stamp !== null && stamp !== file.lastModifiedTime
      ? json({ COOLStatusCode: 1010 }, 409) : null;
  }

  function validatePutFile(request: Request, file: ReferenceFile): Response | null {
    const conflict = timestampConflict(request, file);
    if (conflict) return conflict;
    const currentLock = liveLock(file);
    if (currentLock !== null) {
      if (request.headers.get("X-WOPI-Lock") !== currentLock) return lockConflict(currentLock);
    } else if (file.bytes.byteLength !== 0) {
      return lockConflict(null, "File is not locked");
    }
    return null;
  }

  function validateRelativeTarget(request: Request, file: ReferenceFile, overwrite: boolean): Response | null {
    const targetLock = liveLock(file);
    if (targetLock !== null) return lockConflict(targetLock);
    if (!overwrite) return lockConflict(null, "Target exists");
    return timestampConflict(request, file);
  }

  function setFile(id: string, data: Uint8Array | string, opts: { name?: string; lastModifiedTime?: string } = {}): ReferenceFile {
    const fileId = fileIdFromPath(id.startsWith("/") || id.includes("/") ? id : `/files/${id}`);
    const bytes = typeof data === "string" ? new TextEncoder().encode(data) : data;
    const existing = files.get(fileId);
    const file: ReferenceFile = {
      id: fileId,
      name: opts.name ?? existing?.name ?? fileId,
      bytes,
      version: existing ? existing.version : 1,
      itemVersion: existing ? existing.itemVersion : "1",
      lock: existing?.lock ?? null,
      lockExpiresAt: existing?.lockExpiresAt ?? null,
      lastModifiedTime: opts.lastModifiedTime ?? new Date(now()).toISOString(),
    };
    files.set(fileId, file);
    return file;
  }

  async function readBody(request: Request): Promise<Uint8Array | null> {
    const lengthHeader = request.headers.get("Content-Length");
    const announced = lengthHeader !== null && /^\d+$/.test(lengthHeader.trim()) ? Number(lengthHeader.trim()) : null;
    if (!request.body) {
      return announced !== null && Number.isSafeInteger(announced) && announced !== 0 ? null : new Uint8Array(0);
    }
    const reader = request.body.getReader();
    if (announced === null || !Number.isSafeInteger(announced)) {
      const chunks: Uint8Array[] = [];
      let total = 0;
      for (;;) {
        const chunk = await reader.read();
        if (chunk.done) break;
        chunks.push(chunk.value);
        total += chunk.value.byteLength;
      }
      const out = new Uint8Array(total);
      let offset = 0;
      for (const chunk of chunks) {
        out.set(chunk, offset);
        offset += chunk.byteLength;
      }
      return out;
    }
    const out = new Uint8Array(announced);
    let offset = 0;
    for (;;) {
      const chunk = await reader.read();
      if (chunk.done) break;
      if (offset + chunk.value.byteLength > out.length) {
        cancelStream(reader);
        return null;
      }
      out.set(chunk.value, offset);
      offset += chunk.value.byteLength;
    }
    return offset === out.length ? out : null;
  }

  function chunkedBody(bytes: Uint8Array): ReadableStream<Uint8Array> {
    const chunkSize = 16384;
    let offset = 0;
    return new ReadableStream<Uint8Array>({
      async pull(controller) {
        if (quirks.slowBody) await new Promise((resolve) => setTimeout(resolve, 1));
        if (offset >= bytes.length) {
          controller.close();
          return;
        }
        const next = Math.min(bytes.length, offset + chunkSize);
        controller.enqueue(bytes.subarray(offset, next));
        offset = next;
      },
    });
  }

  async function handle(request: Request): Promise<Response> {
    const url = new URL(request.url);
    const token = url.searchParams.get("access_token");
    const user = resolveToken(token);
    if (!user || token === null) return json({ error: "unauthorized" }, 401);
    if (!verifyRequestProof(request, token)) return json({ error: "invalid proof" }, 401);

    const isContents = url.pathname.endsWith("/contents");
    const fileId = fileIdFromPath(url.pathname);
    if (user.fileIds && !user.fileIds.has(fileId)) return json({ error: "unauthorized" }, 401);
    const file = files.get(fileId);
    if (!file) return json({ error: "not found" }, 404);
    const canWrite = user.canWrite ?? user.userCanWrite ?? true;
    const currentLock = liveLock(file);
    const override = (request.headers.get("X-WOPI-Override") ?? "").toUpperCase();

    if (!isContents && request.method === "GET") {
      return json({
        BaseFileName: file.name,
        Size: file.bytes.byteLength,
        UserId: user.userId,
        OwnerId: user.userId,
        Version: file.itemVersion,
        UserFriendlyName: user.userName ?? user.userId,
        UserCanWrite: canWrite,
        ReadOnly: !canWrite,
        SupportsLocks: true,
        SupportsUpdate: true,
        SupportsGetLock: true,
        SupportsExtendedLockLength: false,
        SupportsRename: true,
        UserCanNotWriteRelative: false,
        UserCanRename: canWrite,
        PostMessageOrigin: url.origin,
        CloseUrl: "",
        LastModifiedTime: file.lastModifiedTime,
      });
    }

    if (isContents && request.method === "GET") {
      const maxExpected = request.headers.get("X-WOPI-MaxExpectedSize");
      if (maxExpected !== null && /^\d+$/.test(maxExpected.trim()) && file.bytes.byteLength > Number(maxExpected.trim())) {
        return json({ error: "file exceeds X-WOPI-MaxExpectedSize" }, 412);
      }
      const headers: Record<string, string> = { "X-WOPI-ItemVersion": file.itemVersion };
      if (quirks.noContentLength || quirks.slowBody) {
        return new Response(chunkedBody(file.bytes), { status: 200, headers });
      }
      headers["Content-Length"] = String(file.bytes.byteLength);
      return new Response(file.bytes as BodyInit, { status: 200, headers });
    }

    if (isContents && request.method === "POST") {
      if (override !== "PUT") return json({ error: "X-WOPI-Override PUT is required" }, 400);
      if (!canWrite) return json({ error: "unauthorized" }, 401);
      const initialConflict = validatePutFile(request, file);
      if (initialConflict) return initialConflict;
      const putBody = await readBody(request);
      if (putBody === null) return json({ error: "body does not match Content-Length" }, 413);
      const conflict = validatePutFile(request, file);
      if (conflict) return conflict;
      file.bytes = putBody;
      file.version += 1;
      file.itemVersion = String(file.version);
      file.lastModifiedTime = new Date(now()).toISOString();
      return empty(200, { "X-WOPI-ItemVersion": file.itemVersion });
    }

    if (!isContents && request.method === "POST") {
      if (override === "GET_LOCK") {
        return empty(200, { "X-WOPI-Lock": currentLock ?? "" });
      }
      if (override === "LOCK") {
        if (!canWrite) return json({ error: "unauthorized" }, 401);
        const newLock = request.headers.get("X-WOPI-Lock") ?? "";
        const oldLock = request.headers.get("X-WOPI-OldLock");
        if (newLock === "") return json({ error: "X-WOPI-Lock is required" }, 400);
        if (oldLock !== null && oldLock === "") return json({ error: "X-WOPI-OldLock is required" }, 400);
        if (oldLock !== null) {
          if (currentLock !== oldLock) return lockConflict(currentLock);
          file.lock = newLock;
          file.lockExpiresAt = now() + LOCK_TTL_MS;
          return empty(200);
        }
        if (currentLock !== null && currentLock !== newLock) return lockConflict(currentLock);
        file.lock = newLock;
        file.lockExpiresAt = now() + LOCK_TTL_MS;
        return empty(200);
      }
      if (override === "REFRESH_LOCK") {
        if (!canWrite) return json({ error: "unauthorized" }, 401);
        const requestLock = request.headers.get("X-WOPI-Lock") ?? "";
        if (requestLock === "") return json({ error: "X-WOPI-Lock is required" }, 400);
        if (currentLock === null || requestLock !== currentLock) return lockConflict(currentLock);
        file.lockExpiresAt = now() + LOCK_TTL_MS;
        return empty(200);
      }
      if (override === "UNLOCK") {
        if (!canWrite) return json({ error: "unauthorized" }, 401);
        const requestLock = request.headers.get("X-WOPI-Lock") ?? "";
        if (requestLock === "") return json({ error: "X-WOPI-Lock is required" }, 400);
        if (currentLock === null || requestLock !== currentLock) return lockConflict(currentLock);
        file.lock = null;
        file.lockExpiresAt = null;
        return empty(200);
      }
      if (override === "PUT_RELATIVE") {
        if (!canWrite) return json({ error: "unauthorized" }, 401);
        const suggested = request.headers.get("X-WOPI-SuggestedTarget");
        const relative = request.headers.get("X-WOPI-RelativeTarget");
        if ((suggested === null && relative === null) || (suggested !== null && relative !== null) || suggested === "" || relative === "") {
          return json({ error: "exactly one target header is required" }, 400);
        }
        const overwrite = relative !== null && (request.headers.get("X-WOPI-OverwriteRelativeTarget") ?? "").toLowerCase() === "true";
        let targetName = decodeUtf7(relative ?? suggested ?? "copy");
        if (targetName === "") return json({ error: "target header is required" }, 400);
        if (suggested !== null && targetName.startsWith(".")) {
          const ext = extensionOf(file.name);
          targetName = `${ext ? file.name.slice(0, -ext.length) : file.name}${targetName}`;
        }
        let finalName = targetName;
        const existing = [...files.values()].find((entry) => entry.name === targetName);
        if (existing && relative !== null) {
          if (user.fileIds && !user.fileIds.has(existing.id)) return json({ error: "unauthorized" }, 401);
          const conflict = validateRelativeTarget(request, existing, overwrite);
          if (conflict) return conflict;
        }
        const body = await readBody(request);
        if (body === null) return json({ error: "body does not match Content-Length" }, 413);
        const current = [...files.values()].find((entry) => entry.name === targetName);
        if (current && relative !== null) {
          if (user.fileIds && !user.fileIds.has(current.id)) return json({ error: "unauthorized" }, 401);
          const conflict = validateRelativeTarget(request, current, overwrite);
          if (conflict) return conflict;
          current.bytes = body;
          current.version += 1;
          current.itemVersion = String(current.version);
          current.lastModifiedTime = new Date(now()).toISOString();
          return json({ Name: finalName, Url: relativeUrl(request, current.id, user) }, 200, { "X-WOPI-ItemVersion": current.itemVersion });
        }
        if (current && suggested !== null) {
          const ext = extensionOf(targetName);
          const stem = ext ? targetName.slice(0, -ext.length) : targetName;
          let counter = 1;
          do {
            finalName = `${stem} (${counter})${ext}`;
            counter += 1;
          } while ([...files.values()].some((entry) => entry.name === finalName));
        }
        const newId = uniqueFileId();
        const created: ReferenceFile = {
          id: newId,
          name: finalName,
          bytes: body,
          version: 1,
          itemVersion: "1",
          lock: null,
          lockExpiresAt: null,
          lastModifiedTime: new Date(now()).toISOString(),
        };
        files.set(newId, created);
        const headers: Record<string, string> = { "X-WOPI-ItemVersion": "1" };
        if (finalName !== targetName) headers["X-WOPI-ValidRelativeTarget"] = encodeUtf7(finalName);
        return json({ Name: finalName, Url: relativeUrl(request, newId, user) }, 200, headers);
      }
      if (override === "RENAME_FILE") {
        if (!canWrite) return json({ error: "unauthorized" }, 401);
        const requestedHeader = request.headers.get("X-WOPI-RequestedName");
        if (requestedHeader === null || requestedHeader === "") return json({ error: "X-WOPI-RequestedName is required" }, 400);
        const requestLock = request.headers.get("X-WOPI-Lock");
        if (currentLock !== null && requestLock !== currentLock) return lockConflict(currentLock);
        const requested = decodeUtf7(requestedHeader);
        if (requested === "") return json({ error: "X-WOPI-RequestedName is required" }, 400);
        const ext = extensionOf(file.name);
        const newName = `${requested}${ext}`;
        if ([...files.values()].some((entry) => entry.id !== file.id && entry.name === newName)) {
          return empty(400, { "X-WOPI-InvalidFileNameError": "A file with that name already exists" });
        }
        file.name = newName;
        return json({ Name: requested });
      }
      return json({ error: "unknown override" }, 400);
    }

    return json({ error: "not found" }, 404);
  }

  return { fetch: handle, files, tokens, setFile };
}
