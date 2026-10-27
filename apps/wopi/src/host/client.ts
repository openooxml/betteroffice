import { dotnetTicks, type ProofSigner } from "../proof";
import { WopiHostError } from "./errors";
import { redactText } from "./redact";
import { decodeUtf7, encodeUtf7 } from "./utf7";

type ChunkResult = Awaited<ReturnType<ReadableStreamDefaultReader<Uint8Array>["read"]>>;

export type HostFetch = (input: string | URL | Request, init?: RequestInit) => Promise<Response>;

export interface HostClientOptions {
  signer: ProofSigner;
  maxFileBytes: number;
  fetch?: HostFetch;
  now?: () => number | Date;
  timeoutMs?: number;
  bodyIdleTimeoutMs?: number;
}

export interface WopiRequest {
  wopiSrc: string;
  accessToken: string;
}

/** CheckFileInfo payload returned by a WOPI host. */
export interface CheckFileInfo {
  BaseFileName: string;
  Size: number;
  UserId: string;
  Version?: string;
  OwnerId?: string;
  UserFriendlyName?: string;
  UserCanWrite?: boolean;
  UserCanNotWriteRelative?: boolean;
  UserCanRename?: boolean;
  ReadOnly?: boolean;
  SupportsLocks?: boolean;
  SupportsUpdate?: boolean;
  SupportsGetLock?: boolean;
  SupportsExtendedLockLength?: boolean;
  SupportsRename?: boolean;
  PostMessageOrigin?: string;
  CloseUrl?: string;
  LastModifiedTime?: string;
  [key: string]: unknown;
}

export interface GetFileResult {
  stream: ReadableStream<Uint8Array>;
  size: number | null;
  itemVersion: string | null;
}

export interface PutFileResult {
  itemVersion: string | null;
}

export interface PutRelativeFileResult {
  name: string;
  url: string;
  validRelativeTarget?: string;
  itemVersion: string | null;
}

export interface RenameFileResult {
  name: string;
}

export interface GetLockResult {
  lock: string | null;
}

export type HostBody = Uint8Array | ReadableStream<Uint8Array>;

export interface PutFileArgs extends WopiRequest {
  body: HostBody;
  lock?: string;
  size?: number;
  extraHeaders?: HeadersInit;
}

export interface PutRelativeFileArgs extends WopiRequest {
  body: HostBody;
  suggestedTarget?: string;
  relativeTarget?: string;
  overwrite?: boolean;
  size?: number;
  lock?: string;
  extraHeaders?: HeadersInit;
}

export interface LockArgs extends WopiRequest {
  lock: string;
}

export interface UnlockAndRelockArgs extends WopiRequest {
  lock: string;
  oldLock: string;
}

export interface RenameFileArgs extends WopiRequest {
  requestedName: string;
  lock?: string;
}

export interface HostClient {
  checkFileInfo(input: WopiRequest): Promise<CheckFileInfo>;
  getFile(input: WopiRequest): Promise<GetFileResult>;
  readFile(input: WopiRequest): Promise<Uint8Array>;
  putFile(input: PutFileArgs): Promise<PutFileResult>;
  lock(input: LockArgs): Promise<void>;
  refreshLock(input: LockArgs): Promise<void>;
  unlock(input: LockArgs): Promise<void>;
  unlockAndRelock(input: UnlockAndRelockArgs): Promise<void>;
  getLock(input: WopiRequest): Promise<GetLockResult>;
  putRelativeFile(input: PutRelativeFileArgs): Promise<PutRelativeFileResult>;
  renameFile(input: RenameFileArgs): Promise<RenameFileResult>;
}

const RETRYABLE_STATUSES = new Set([502, 503, 504]);
const DEFAULT_TIMEOUT_MS = 30000;
const DEFAULT_BODY_IDLE_TIMEOUT_MS = 60000;

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function scrub(text: string, accessToken: string): string {
  const redacted = redactText(text);
  return accessToken.length > 0 ? redacted.split(accessToken).join("***") : redacted;
}

function isStream(body: unknown): body is ReadableStream<Uint8Array> {
  return typeof body === "object" && body !== null && "getReader" in body;
}

function parseContentLength(header: string | null): number | null {
  if (header === null) return null;
  const trimmed = header.trim();
  if (!/^\d+$/.test(trimmed)) return null;
  const value = Number(trimmed);
  return Number.isSafeInteger(value) ? value : null;
}

interface RequestSpec {
  op: string;
  contents: boolean;
  method: "GET" | "POST";
  override?: string;
  headers?: HeadersInit;
  body?: HostBody;
  size?: number;
  idempotent: boolean;
  exceeded?: () => boolean;
}

/** Creates a proof signing WOPI host client with streamed transfers. */
export function createHostClient(options: HostClientOptions): HostClient {
  const fetchFn: HostFetch = options.fetch ?? ((input, init) => fetch(input, init));
  const maxFileBytes = options.maxFileBytes;
  const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  const bodyIdleTimeoutMs = options.bodyIdleTimeoutMs ?? DEFAULT_BODY_IDLE_TIMEOUT_MS;
  if (!Number.isSafeInteger(maxFileBytes) || maxFileBytes <= 0) {
    throw new WopiHostError("invalidResponse", "maxFileBytes must be a positive integer");
  }

  function nowMs(): number {
    const value = options.now ? options.now() : Date.now();
    return value instanceof Date ? value.getTime() : value;
  }

  function buildUrl(wopiSrc: string, contents: boolean, accessToken: string): string {
    let url: URL;
    try {
      url = new URL(wopiSrc);
    } catch {
      throw new WopiHostError("invalidResponse", `Invalid WOPISrc: ${scrub(wopiSrc, accessToken)}`);
    }
    if (url.hash !== "") {
      throw new WopiHostError("notAllowed", "WOPISrc must not contain a fragment");
    }
    if (contents) url.pathname = `${url.pathname.replace(/\/+$/, "")}/contents`;
    url.searchParams.set("access_token", accessToken);
    return url.toString();
  }

  function tooLarge(op: string): WopiHostError {
    return new WopiHostError("tooLarge", `${op} exceeds the maximum file size of ${maxFileBytes} bytes`);
  }

  function bodyError(err: unknown, op: string, accessToken: string): WopiHostError {
    if (err instanceof WopiHostError) return err;
    return new WopiHostError("network", `${op} network error: ${scrub(err instanceof Error ? err.message : String(err), accessToken)}`);
  }

  async function withTotalTimeout<T>(op: string, accessToken: string, run: () => Promise<T>, cancel?: () => Promise<void> | undefined): Promise<T> {
    let timedOut = false;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timeout = new Promise<never>((_resolve, reject) => {
      timer = setTimeout(() => {
        timedOut = true;
        reject(new WopiHostError("timeout", `${op} timed out after ${timeoutMs} ms`));
      }, timeoutMs);
    });
    try {
      return await Promise.race([run(), timeout]);
    } catch (err) {
      if (timedOut) {
        try {
          const pending = cancel?.();
          if (pending) pending.catch(() => {});
        } catch {}
        throw new WopiHostError("timeout", `${op} timed out after ${timeoutMs} ms`);
      }
      throw bodyError(err, op, accessToken);
    } finally {
      if (timer) clearTimeout(timer);
    }
  }

  async function readResponseText(response: Response, op: string, accessToken: string): Promise<string> {
    return withTotalTimeout(op, accessToken, () => response.text(), () => response.body?.cancel());
  }

  async function idleRead(reader: ReadableStreamDefaultReader<Uint8Array>, op: string, accessToken: string): Promise<ChunkResult> {
    let timedOut = false;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timeout = new Promise<never>((_resolve, reject) => {
      timer = setTimeout(() => {
        timedOut = true;
        reject(new WopiHostError("timeout", `${op} timed out after ${bodyIdleTimeoutMs} ms without data`));
      }, bodyIdleTimeoutMs);
    });
    try {
      return await Promise.race([reader.read(), timeout]);
    } catch (err) {
      if (timedOut) {
        try {
          const pending = reader.cancel();
          pending.catch(() => {});
        } catch {}
        throw new WopiHostError("timeout", `${op} timed out after ${bodyIdleTimeoutMs} ms without data`);
      }
      throw bodyError(err, op, accessToken);
    } finally {
      if (timer) clearTimeout(timer);
    }
  }

  async function responseError(response: Response, op: string, accessToken: string): Promise<WopiHostError> {
    const status = response.status;
    const fullText = scrub(await readResponseText(response, op, accessToken), accessToken);
    const bodyText = fullText.slice(0, 200);
    const detail = bodyText.length > 0 ? `: ${bodyText}` : "";
    if (status >= 300 && status < 400) {
      return new WopiHostError("redirect", `${op} failed with redirect status ${status}`, { status });
    }
    if (status === 401) return new WopiHostError("unauthorized", `${op} failed with status 401${detail}`, { status });
    if (status === 404) return new WopiHostError("notFound", `${op} failed with status 404${detail}`, { status });
    if (status === 409) {
      const rawLock = response.headers.get("X-WOPI-Lock");
      const currentLock = rawLock === null || rawLock === "" ? null : scrub(rawLock, accessToken);
      const rawReason = response.headers.get("X-WOPI-LockFailureReason");
      const reason = rawReason === null ? undefined : scrub(rawReason, accessToken);
      if (/COOLStatusCode/i.test(fullText)) {
        return new WopiHostError("conflict", `${op} failed with status 409${detail}`, { status, currentLock, reason });
      }
      if (rawLock !== null || reason !== undefined) {
        return new WopiHostError("lockMismatch", `${op} failed with status 409${detail}`, { status, currentLock, reason });
      }
      return new WopiHostError("conflict", `${op} failed with status 409${detail}`, { status });
    }
    if (status === 412) return new WopiHostError("preconditionFailed", `${op} failed with status 412${detail}`, { status });
    if (status === 413) return new WopiHostError("tooLarge", `${op} failed with status 413${detail}`, { status });
    if (status === 501) return new WopiHostError("notImplemented", `${op} failed with status 501${detail}`, { status });
    if (status >= 500) return new WopiHostError("server", `${op} failed with status ${status}${detail}`, { status });
    return new WopiHostError("invalidResponse", `${op} failed with status ${status}${detail}`, { status });
  }

  function wrapDownload(source: ReadableStream<Uint8Array>, op: string, accessToken: string): ReadableStream<Uint8Array> {
    const reader = source.getReader();
    let count = 0;
    return new ReadableStream<Uint8Array>({
      async pull(controller) {
        let chunk: ChunkResult;
        try {
          chunk = await idleRead(reader, op, accessToken);
        } catch (err) {
          controller.error(bodyError(err, op, accessToken));
          return;
        }
        if (chunk.done) {
          controller.close();
          return;
        }
        count += chunk.value.byteLength;
        if (count > maxFileBytes) {
          try {
            await reader.cancel();
          } catch {}
          controller.error(tooLarge(op));
          return;
        }
        controller.enqueue(chunk.value);
      },
      async cancel(reason) {
        await reader.cancel(reason);
      },
    });
  }

  function wrapUpload(source: ReadableStream<Uint8Array>, state: { exceeded: boolean; error?: WopiHostError }, op: string, accessToken: string): ReadableStream<Uint8Array> {
    const reader = source.getReader();
    let count = 0;
    return new ReadableStream<Uint8Array>({
      async pull(controller) {
        let chunk: ChunkResult;
        try {
          chunk = await idleRead(reader, op, accessToken);
        } catch (err) {
          state.error = bodyError(err, op, accessToken);
          controller.error(state.error);
          return;
        }
        if (chunk.done) {
          controller.close();
          return;
        }
        count += chunk.value.byteLength;
        if (count > maxFileBytes) {
          state.exceeded = true;
          try {
            await reader.cancel();
          } catch {}
          controller.error(tooLarge(op));
          return;
        }
        controller.enqueue(chunk.value);
      },
      async cancel(reason) {
        await reader.cancel(reason);
      },
    });
  }

  async function execute(input: WopiRequest, spec: RequestSpec): Promise<Response> {
    const url = buildUrl(input.wopiSrc, spec.contents, input.accessToken);
    let lastError: WopiHostError | undefined;
    for (let attempt = 0; attempt < 3; attempt += 1) {
      if (attempt > 0) await delay(100 * 2 ** (attempt - 1));
      const ticks = dotnetTicks(new Date(nowMs()));
      const signed = options.signer.signProof({ accessToken: input.accessToken, url, ticks });
      const headers = new Headers(spec.headers);
      headers.set("X-WOPI-Proof", signed.proof);
      headers.set("X-WOPI-ProofOld", signed.proofOld);
      headers.set("X-WOPI-TimeStamp", signed.timestamp);
      if (spec.override) headers.set("X-WOPI-Override", spec.override);
      if (spec.op === "GetFile") headers.set("X-WOPI-MaxExpectedSize", String(maxFileBytes));
      const uploadState: { exceeded: boolean; error?: WopiHostError } = { exceeded: false };
      let body: BodyInit | undefined;
      if (spec.body instanceof Uint8Array) {
        body = spec.body as BodyInit;
        headers.set("Content-Length", String(spec.body.byteLength));
      } else if (isStream(spec.body)) {
        if (spec.size !== undefined) headers.set("Content-Length", String(spec.size));
        body = wrapUpload(spec.body, uploadState, spec.op, input.accessToken) as BodyInit;
      }
      const exceeded = () => uploadState.exceeded || (spec.exceeded ? spec.exceeded() : false);
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs);
      let response: Response;
      try {
        const init: RequestInit & { duplex?: string } = {
          method: spec.method,
          headers,
          body,
          redirect: "manual",
          signal: controller.signal,
        };
        if (isStream(spec.body)) init.duplex = "half";
        response = await fetchFn(url, init);
      } catch (err) {
        clearTimeout(timer);
        if (exceeded()) throw tooLarge(spec.op);
        if (uploadState.error) throw uploadState.error;
        const timedOut = controller.signal.aborted;
        const mapped = timedOut
          ? new WopiHostError("timeout", `${spec.op} timed out after ${timeoutMs} ms`)
          : new WopiHostError("network", `${spec.op} network error: ${scrub(err instanceof Error ? err.message : String(err), input.accessToken)}`);
        lastError = mapped;
        if (spec.idempotent && mapped.kind === "network" && attempt < 2) continue;
        throw mapped;
      }
      clearTimeout(timer);
      if (exceeded()) {
        try {
          await response.body?.cancel();
        } catch {}
        throw tooLarge(spec.op);
      }
      if (uploadState.error) throw uploadState.error;
      if (spec.idempotent && RETRYABLE_STATUSES.has(response.status) && attempt < 2) {
        try {
          await response.body?.cancel();
        } catch {}
        continue;
      }
      if (response.status < 200 || response.status >= 300) {
        throw await responseError(response, spec.op, input.accessToken);
      }
      return response;
    }
    throw lastError ?? new WopiHostError("network", `${spec.op} network error`);
  }

  async function checkFileInfo(input: WopiRequest): Promise<CheckFileInfo> {
    const response = await execute(input, { op: "CheckFileInfo", contents: false, method: "GET", idempotent: true });
    const text = await readResponseText(response, "CheckFileInfo", input.accessToken);
    let parsed: unknown;
    try {
      parsed = JSON.parse(text) as unknown;
    } catch {
      throw new WopiHostError("invalidResponse", "CheckFileInfo returned invalid JSON");
    }
    if (typeof parsed !== "object" || parsed === null) {
      throw new WopiHostError("invalidResponse", "CheckFileInfo returned an invalid payload");
    }
    const info = parsed as CheckFileInfo;
    if (typeof info.BaseFileName !== "string" || info.BaseFileName.length === 0) {
      throw new WopiHostError("invalidResponse", "CheckFileInfo is missing BaseFileName");
    }
    if (typeof info.Size !== "number" || !Number.isFinite(info.Size) || info.Size < 0) {
      throw new WopiHostError("invalidResponse", "CheckFileInfo is missing Size");
    }
    if (typeof info.UserId !== "string" || info.UserId.length === 0) {
      throw new WopiHostError("invalidResponse", "CheckFileInfo is missing UserId");
    }
    if (info.Size > maxFileBytes) throw tooLarge("CheckFileInfo");
    return info;
  }

  async function getFile(input: WopiRequest): Promise<GetFileResult> {
    const response = await execute(input, { op: "GetFile", contents: true, method: "GET", idempotent: true });
    const announced = parseContentLength(response.headers.get("Content-Length"));
    if (announced !== null && announced > maxFileBytes) {
      try {
        await response.body?.cancel();
      } catch {}
      throw tooLarge("GetFile");
    }
    const source = response.body ?? new ReadableStream<Uint8Array>({ start(controller) { controller.close(); } });
    const stream = wrapDownload(source, "GetFile", input.accessToken);
    return { stream, size: announced, itemVersion: response.headers.get("X-WOPI-ItemVersion") };
  }

  async function readFile(input: WopiRequest): Promise<Uint8Array> {
    const result = await getFile(input);
    const reader = result.stream.getReader();
    if (result.size !== null) {
      const buffer = new Uint8Array(result.size);
      let offset = 0;
      for (;;) {
        const chunk = await reader.read();
        if (chunk.done) break;
        if (offset + chunk.value.byteLength > buffer.length) {
          try {
            await reader.cancel();
          } catch {}
          throw new WopiHostError("invalidResponse", "GetFile returned more bytes than announced");
        }
        buffer.set(chunk.value, offset);
        offset += chunk.value.byteLength;
      }
      if (offset < buffer.length) {
        throw new WopiHostError("invalidResponse", "GetFile returned fewer bytes than announced");
      }
      return buffer;
    }
    const chunks: Uint8Array[] = [];
    let total = 0;
    for (;;) {
      const chunk = await reader.read();
      if (chunk.done) break;
      total += chunk.value.byteLength;
      chunks.push(chunk.value);
    }
    const buffer = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) {
      buffer.set(chunk, offset);
      offset += chunk.byteLength;
    }
    return buffer;
  }

  async function sendPut(input: WopiRequest, args: PutFileArgs, contents: boolean, override: string, extra: HeadersInit | undefined, op: string): Promise<Response> {
    const declaredSize = args.body instanceof Uint8Array ? args.body.byteLength : args.size;
    if (declaredSize !== undefined && declaredSize > maxFileBytes) {
      if (isStream(args.body)) {
        try {
          await args.body.cancel();
        } catch {}
      }
      throw tooLarge(op);
    }
    const headers = new Headers(extra);
    if (args.lock) headers.set("X-WOPI-Lock", args.lock);
    return execute(input, {
      op,
      contents,
      method: "POST",
      override,
      headers,
      body: args.body,
      size: declaredSize,
      idempotent: false,
    });
  }

  async function putFile(args: PutFileArgs): Promise<PutFileResult> {
    const response = await sendPut(args, args, true, "PUT", args.extraHeaders, "PutFile");
    try {
      await response.body?.cancel();
    } catch {}
    return { itemVersion: response.headers.get("X-WOPI-ItemVersion") };
  }

  async function discard(response: Response): Promise<void> {
    try {
      await response.body?.cancel();
    } catch {}
  }

  async function lock(args: LockArgs): Promise<void> {
    await discard(await execute(args, { op: "Lock", contents: false, method: "POST", override: "LOCK", headers: { "X-WOPI-Lock": args.lock }, idempotent: false }));
  }

  async function refreshLock(args: LockArgs): Promise<void> {
    await discard(await execute(args, { op: "RefreshLock", contents: false, method: "POST", override: "REFRESH_LOCK", headers: { "X-WOPI-Lock": args.lock }, idempotent: true }));
  }

  async function unlock(args: LockArgs): Promise<void> {
    await discard(await execute(args, { op: "Unlock", contents: false, method: "POST", override: "UNLOCK", headers: { "X-WOPI-Lock": args.lock }, idempotent: false }));
  }

  async function unlockAndRelock(args: UnlockAndRelockArgs): Promise<void> {
    await discard(await execute(args, {
      op: "UnlockAndRelock",
      contents: false,
      method: "POST",
      override: "LOCK",
      headers: { "X-WOPI-Lock": args.lock, "X-WOPI-OldLock": args.oldLock },
      idempotent: false,
    }));
  }

  async function getLock(input: WopiRequest): Promise<GetLockResult> {
    const response = await execute(input, { op: "GetLock", contents: false, method: "POST", override: "GET_LOCK", idempotent: true });
    const raw = response.headers.get("X-WOPI-Lock");
    const currentLock = raw === null || raw === "" ? null : raw;
    await discard(response);
    return { lock: currentLock };
  }

  async function readJsonObject(response: Response, op: string, accessToken: string): Promise<Record<string, unknown>> {
    const text = await readResponseText(response, op, accessToken);
    let parsed: unknown;
    try {
      parsed = JSON.parse(text) as unknown;
    } catch {
      throw new WopiHostError("invalidResponse", `${op} returned invalid JSON`);
    }
    if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
      throw new WopiHostError("invalidResponse", `${op} returned an invalid payload`);
    }
    return parsed as Record<string, unknown>;
  }

  function requiredString(record: Record<string, unknown>, key: string, op: string): string {
    const value = record[key];
    if (typeof value !== "string" || value.length === 0) {
      throw new WopiHostError("invalidResponse", `${op} response is missing ${key}`);
    }
    return value;
  }

  async function putRelativeFile(args: PutRelativeFileArgs): Promise<PutRelativeFileResult> {
    if (args.suggestedTarget && args.relativeTarget) {
      throw new WopiHostError("invalidResponse", "PutRelativeFile needs only one of suggestedTarget and relativeTarget");
    }
    if (!args.suggestedTarget && !args.relativeTarget) {
      throw new WopiHostError("invalidResponse", "PutRelativeFile needs a suggestedTarget or relativeTarget");
    }
    const headers = new Headers(args.extraHeaders);
    if (args.suggestedTarget) headers.set("X-WOPI-SuggestedTarget", encodeUtf7(args.suggestedTarget));
    if (args.relativeTarget) headers.set("X-WOPI-RelativeTarget", encodeUtf7(args.relativeTarget));
    if (args.overwrite !== undefined) headers.set("X-WOPI-OverwriteRelativeTarget", args.overwrite ? "true" : "false");
    const response = await sendPut(args, args, false, "PUT_RELATIVE", headers, "PutRelativeFile");
    const validHeader = response.headers.get("X-WOPI-ValidRelativeTarget");
    const validRelativeTarget = validHeader ? decodeUtf7(validHeader) : undefined;
    const parsed = await readJsonObject(response, "PutRelativeFile", args.accessToken);
    const name = requiredString(parsed, "Name", "PutRelativeFile");
    const url = requiredString(parsed, "Url", "PutRelativeFile");
    return { name, url, validRelativeTarget, itemVersion: response.headers.get("X-WOPI-ItemVersion") };
  }

  async function renameFile(args: RenameFileArgs): Promise<RenameFileResult> {
    const headers = new Headers({ "X-WOPI-RequestedName": encodeUtf7(args.requestedName) });
    if (args.lock) headers.set("X-WOPI-Lock", args.lock);
    const response = await execute(args, { op: "RenameFile", contents: false, method: "POST", override: "RENAME_FILE", headers, idempotent: false });
    const parsed = await readJsonObject(response, "RenameFile", args.accessToken);
    return { name: requiredString(parsed, "Name", "RenameFile") };
  }

  const client: HostClient = {
    checkFileInfo,
    getFile,
    readFile,
    putFile,
    lock,
    refreshLock,
    unlock,
    unlockAndRelock,
    getLock,
    putRelativeFile,
    renameFile,
  };
  return client;
}
