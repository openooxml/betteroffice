import { dotnetTicks, type ProofSigner } from "../proof";
import { WopiHostError } from "./errors";
import { redactText } from "./redact";
import { cancelStream } from "./streams";
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

interface HostResponse {
  response: Response;
  controller: AbortController;
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

  async function withTotalTimeout<T>(op: string, accessToken: string, run: () => Promise<T>, cancel?: () => void): Promise<T> {
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
          cancel?.();
        } catch {}
        throw new WopiHostError("timeout", `${op} timed out after ${timeoutMs} ms`);
      }
      throw bodyError(err, op, accessToken);
    } finally {
      if (timer) clearTimeout(timer);
    }
  }

  async function readResponseText(result: HostResponse, op: string, accessToken: string, maxBytes = 1024 * 1024, throwOnOverCap = true): Promise<string> {
    const { response, controller } = result;
    if (!response.body) return "";
    const reader = response.body.getReader();
    return withTotalTimeout(op, accessToken, async () => {
      const kept: Uint8Array[] = [];
      let keptBytes = 0;
      let total = 0;
      try {
        for (;;) {
          const chunk = await reader.read();
          if (chunk.done) break;
          total += chunk.value.byteLength;
          if (total > maxBytes) {
            cancelStream(reader, () => controller.abort());
            if (throwOnOverCap) {
              throw new WopiHostError("invalidResponse", `${op} response exceeds ${maxBytes} bytes`);
            }
            const room = maxBytes - keptBytes;
            if (room > 0) {
              kept.push(chunk.value.subarray(0, room));
              keptBytes += Math.min(room, chunk.value.byteLength);
            }
            break;
          }
          kept.push(chunk.value);
          keptBytes += chunk.value.byteLength;
        }
        const out = new Uint8Array(keptBytes);
        let offset = 0;
        for (const part of kept) {
          out.set(part, offset);
          offset += part.byteLength;
        }
        return new TextDecoder().decode(out);
      } finally {
        reader.releaseLock();
      }
    }, () => {
      cancelStream(reader, () => controller.abort());
    });
  }

  async function idleRead(reader: ReadableStreamDefaultReader<Uint8Array>, op: string, accessToken: string, abort?: () => void): Promise<ChunkResult> {
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
        cancelStream(reader, abort);
        throw new WopiHostError("timeout", `${op} timed out after ${bodyIdleTimeoutMs} ms without data`);
      }
      throw bodyError(err, op, accessToken);
    } finally {
      if (timer) clearTimeout(timer);
    }
  }

  async function responseError(result: HostResponse, op: string, accessToken: string): Promise<WopiHostError> {
    const { response } = result;
    const status = response.status;
    const fullText = scrub(await readResponseText(result, op, accessToken, 64 * 1024, false), accessToken);
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

  function wrapDownload(source: ReadableStream<Uint8Array>, op: string, accessToken: string, abort: () => void): ReadableStream<Uint8Array> {
    const reader = source.getReader();
    let count = 0;
    return new ReadableStream<Uint8Array>({
      async pull(controller) {
        let chunk: ChunkResult;
        try {
          chunk = await idleRead(reader, op, accessToken, abort);
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
          controller.error(tooLarge(op));
          cancelStream(reader, abort);
          return;
        }
        controller.enqueue(chunk.value);
      },
      cancel(reason) {
        cancelStream(reader, abort, reason);
      },
    });
  }

  function streamFromBytes(bytes: Uint8Array): ReadableStream<Uint8Array> {
    const chunkSize = 64 * 1024;
    let offset = 0;
    return new ReadableStream<Uint8Array>({
      pull(controller) {
        if (offset >= bytes.byteLength) {
          controller.close();
          return;
        }
        const next = Math.min(bytes.byteLength, offset + chunkSize);
        controller.enqueue(bytes.subarray(offset, next));
        offset = next;
      },
    });
  }

  interface UploadState {
    exceeded: boolean;
    error?: WopiHostError;
    onDone?: () => void;
    pulled?: boolean;
    abort?: () => void;
    clearIdleTimer?: () => void;
  }

  function wrapUpload(source: ReadableStream<Uint8Array>, state: UploadState, op: string, accessToken: string): ReadableStream<Uint8Array> {
    const reader = source.getReader();
    let count = 0;
    let idleTimer: ReturnType<typeof setTimeout> | undefined;
    state.clearIdleTimer = () => {
      if (idleTimer) clearTimeout(idleTimer);
      idleTimer = undefined;
    };
    return new ReadableStream<Uint8Array>({
      async pull(controller) {
        state.pulled = true;
        state.clearIdleTimer?.();
        let chunk: ChunkResult;
        try {
          chunk = await idleRead(reader, op, accessToken, state.abort);
        } catch (err) {
          state.error = bodyError(err, op, accessToken);
          controller.error(state.error);
          return;
        }
        if (chunk.done) {
          state.onDone?.();
          controller.close();
          return;
        }
        count += chunk.value.byteLength;
        if (count > maxFileBytes) {
          state.exceeded = true;
          controller.error(tooLarge(op));
          cancelStream(reader, state.abort);
          return;
        }
        controller.enqueue(chunk.value);
        idleTimer = setTimeout(() => {
          idleTimer = undefined;
          state.error = new WopiHostError("timeout", `${op} timed out after ${bodyIdleTimeoutMs} ms without data`);
          state.abort?.();
        }, bodyIdleTimeoutMs);
      },
      cancel(reason) {
        state.clearIdleTimer?.();
        cancelStream(reader, state.abort, reason);
      },
    }, { highWaterMark: 0 });
  }

  async function execute(input: WopiRequest, spec: RequestSpec): Promise<HostResponse> {
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
      const controller = new AbortController();
      const uploadState: UploadState = { exceeded: false, abort: () => controller.abort() };
      let body: BodyInit | undefined;
      let streamed = false;
      if (spec.body instanceof Uint8Array && spec.body.byteLength > 1024 * 1024) {
        streamed = true;
        headers.set("Content-Length", String(spec.body.byteLength));
        body = wrapUpload(streamFromBytes(spec.body), uploadState, spec.op, input.accessToken) as BodyInit;
      } else if (spec.body instanceof Uint8Array) {
        body = spec.body as BodyInit;
        headers.set("Content-Length", String(spec.body.byteLength));
      } else if (isStream(spec.body)) {
        streamed = true;
        if (spec.size !== undefined) headers.set("Content-Length", String(spec.size));
        body = wrapUpload(spec.body, uploadState, spec.op, input.accessToken) as BodyInit;
      }
      const exceeded = () => uploadState.exceeded || (spec.exceeded ? spec.exceeded() : false);
      let timer: ReturnType<typeof setTimeout> | undefined;
      let connectTimer: ReturnType<typeof setTimeout> | undefined;
      const armHeaderTimeout = () => {
        if (timer === undefined) timer = setTimeout(() => controller.abort(), timeoutMs);
      };
      if (streamed) {
        uploadState.onDone = armHeaderTimeout;
        connectTimer = setTimeout(() => {
          if (!uploadState.pulled) controller.abort();
        }, timeoutMs);
      } else {
        armHeaderTimeout();
      }
      let response: Response;
      try {
        const init: RequestInit & { duplex?: string } = {
          method: spec.method,
          headers,
          body,
          redirect: "manual",
          signal: controller.signal,
        };
        if (streamed) init.duplex = "half";
        response = await fetchFn(url, init);
      } catch (err) {
        if (timer) clearTimeout(timer);
        if (connectTimer) clearTimeout(connectTimer);
        uploadState.clearIdleTimer?.();
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
      if (timer) clearTimeout(timer);
      if (connectTimer) clearTimeout(connectTimer);
      uploadState.clearIdleTimer?.();
      if (exceeded()) {
        cancelStream(response.body, () => controller.abort());
        throw tooLarge(spec.op);
      }
      if (uploadState.error) throw uploadState.error;
      if (spec.idempotent && RETRYABLE_STATUSES.has(response.status) && attempt < 2) {
        cancelStream(response.body, () => controller.abort());
        continue;
      }
      if (response.status < 200 || response.status >= 300) {
        throw await responseError({ response, controller }, spec.op, input.accessToken);
      }
      return { response, controller };
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
    const { response, controller } = await execute(input, { op: "GetFile", contents: true, method: "GET", idempotent: true });
    const announced = parseContentLength(response.headers.get("Content-Length"));
    if (announced !== null && announced > maxFileBytes) {
      cancelStream(response.body, () => controller.abort());
      throw tooLarge("GetFile");
    }
    const source = response.body ?? new ReadableStream<Uint8Array>({ start(controller) { controller.close(); } });
    const stream = wrapDownload(source, "GetFile", input.accessToken, () => controller.abort());
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
          cancelStream(reader);
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

  async function sendPut(input: WopiRequest, args: PutFileArgs, contents: boolean, override: string, extra: HeadersInit | undefined, op: string): Promise<HostResponse> {
    const declaredSize = args.body instanceof Uint8Array ? args.body.byteLength : args.size;
    if (declaredSize !== undefined && declaredSize > maxFileBytes) {
      if (isStream(args.body)) {
        cancelStream(args.body);
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
    const result = await sendPut(args, args, true, "PUT", args.extraHeaders, "PutFile");
    discard(result);
    return { itemVersion: result.response.headers.get("X-WOPI-ItemVersion") };
  }

  function discard({ response, controller }: HostResponse): void {
    cancelStream(response.body, () => controller.abort());
  }

  async function lock(args: LockArgs): Promise<void> {
    discard(await execute(args, { op: "Lock", contents: false, method: "POST", override: "LOCK", headers: { "X-WOPI-Lock": args.lock }, idempotent: false }));
  }

  async function refreshLock(args: LockArgs): Promise<void> {
    discard(await execute(args, { op: "RefreshLock", contents: false, method: "POST", override: "REFRESH_LOCK", headers: { "X-WOPI-Lock": args.lock }, idempotent: true }));
  }

  async function unlock(args: LockArgs): Promise<void> {
    discard(await execute(args, { op: "Unlock", contents: false, method: "POST", override: "UNLOCK", headers: { "X-WOPI-Lock": args.lock }, idempotent: false }));
  }

  async function unlockAndRelock(args: UnlockAndRelockArgs): Promise<void> {
    discard(await execute(args, {
      op: "UnlockAndRelock",
      contents: false,
      method: "POST",
      override: "LOCK",
      headers: { "X-WOPI-Lock": args.lock, "X-WOPI-OldLock": args.oldLock },
      idempotent: false,
    }));
  }

  async function getLock(input: WopiRequest): Promise<GetLockResult> {
    const result = await execute(input, { op: "GetLock", contents: false, method: "POST", override: "GET_LOCK", idempotent: true });
    const { response } = result;
    const raw = response.headers.get("X-WOPI-Lock");
    const currentLock = raw === null || raw === "" ? null : raw;
    discard(result);
    return { lock: currentLock };
  }

  async function readJsonObject(response: HostResponse, op: string, accessToken: string): Promise<Record<string, unknown>> {
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
    const result = await sendPut(args, args, false, "PUT_RELATIVE", headers, "PutRelativeFile");
    const { response } = result;
    const validHeader = response.headers.get("X-WOPI-ValidRelativeTarget");
    const validRelativeTarget = validHeader ? decodeUtf7(validHeader) : undefined;
    const parsed = await readJsonObject(result, "PutRelativeFile", args.accessToken);
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
