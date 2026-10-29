import { SseParser, toSessionEvent, type SessionEvent } from "./sse";
import type { SaveBarrier } from "../../src/relay/sequence";
import { BASELINE_GENERATION_HEADER, parseBaselineGeneration } from "../../src/baseline";

/** Join payload as served by /api/session/join. */
export interface JoinResponse {
  format: "docx" | "xlsx" | "pptx";
  initialUpdate: Uint8Array;
  baselineGeneration: number;
}

function fromBase64(value: string): Uint8Array {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/** Decodes the base64 join payload from the session API. */
export function decodeJoinResponse(raw: unknown): JoinResponse {
  if (typeof raw !== "object" || raw === null) throw new Error("Invalid join response.");
  const record = raw as Record<string, unknown>;
  if (record.format !== "docx" && record.format !== "xlsx" && record.format !== "pptx") {
    throw new Error("Invalid join response format.");
  }
  if (typeof record.initialUpdate !== "string") throw new Error("Invalid join response update.");
  if (typeof record.baselineGeneration !== "number" || !Number.isSafeInteger(record.baselineGeneration) || record.baselineGeneration < 0) {
    throw new Error("Invalid baseline generation.");
  }
  return { format: record.format, initialUpdate: fromBase64(record.initialUpdate), baselineGeneration: record.baselineGeneration };
}

function defaultFetch(input: RequestInfo | URL, init?: RequestInit): Promise<Response> {
  return globalThis.fetch(input, init);
}

export class SessionEndedError extends Error {
  constructor() {
    super("sessionEnded");
    this.name = "SessionEndedError";
  }
}

/** Session API client. The handle is only ever sent as a Bearer header. */
export class SessionApi {
  readonly #handle: string;
  readonly #fetch: (input: RequestInfo | URL, init?: RequestInit) => Promise<Response>;

  constructor(handle: string, fetchFn: (input: RequestInfo | URL, init?: RequestInit) => Promise<Response> = defaultFetch) {
    this.#handle = handle;
    this.#fetch = fetchFn;
  }

  #headers(): Record<string, string> {
    return { Authorization: `Bearer ${this.#handle}` };
  }

  async source(signal?: AbortSignal): Promise<{ bytes: Uint8Array; baselineGeneration: number }> {
    const response = await this.#fetch("/api/session/source", { headers: this.#headers(), signal });
    if (!response.ok) throw new Error(`Could not load the document (${response.status}).`);
    const baselineGeneration = parseBaselineGeneration(response.headers.get(BASELINE_GENERATION_HEADER));
    if (baselineGeneration === null) throw new Error("Invalid baseline generation.");
    return { bytes: new Uint8Array(await response.arrayBuffer()), baselineGeneration };
  }

  async join(signal?: AbortSignal): Promise<JoinResponse> {
    const response = await this.#fetch("/api/session/join", { headers: this.#headers(), signal });
    if (!response.ok) throw new Error(`Could not join the session (${response.status}).`);
    return decodeJoinResponse(await response.json());
  }

  async save(barrier: SaveBarrier, signal?: AbortSignal): Promise<{ itemVersion: string | null }> {
    const response = await this.#fetch("/api/session/save", {
      method: "POST",
      headers: { ...this.#headers(), "Content-Type": "application/json" },
      body: JSON.stringify(barrier),
      signal,
    });
    if (!response.ok) {
      const failure = await response.json().catch(() => null) as { error?: { kind?: unknown } } | null;
      const reason = failure?.error?.kind;
      if ([401, 404, 410].includes(response.status) || reason === "closed") throw new SessionEndedError();
      throw new Error(typeof reason === "string" ? reason : `Save failed (${response.status}).`);
    }
    const result = (await response.json()) as { itemVersion?: unknown };
    return { itemVersion: typeof result.itemVersion === "string" ? result.itemVersion : null };
  }

  async isActive(signal?: AbortSignal): Promise<boolean> {
    const response = await this.#fetch("/api/session/join", { headers: this.#headers(), signal });
    if ([401, 404, 410].includes(response.status)) return false;
    if (response.status === 409) {
      const failure = await response.json().catch(() => null) as { error?: { kind?: string } } | null;
      if (failure?.error?.kind === "closed") return false;
    }
    if (!response.ok) throw new Error(`Could not check the session (${response.status}).`);
    await response.arrayBuffer();
    return true;
  }

  async recover(bytes: Uint8Array, signal?: AbortSignal): Promise<{ copyName: string }> {
    const response = await this.#fetch("/api/session/recovery", {
      method: "POST", headers: { ...this.#headers(), "Content-Type": "application/octet-stream" }, body: bytes as BodyInit, signal,
    });
    if (!response.ok) {
      const failure = await response.json().catch(() => null) as { error?: { kind?: unknown } } | null;
      if ([401, 404, 410].includes(response.status) || failure?.error?.kind === "closed") throw new SessionEndedError();
      throw new Error(`Recovery copy failed (${response.status}).`);
    }
    const result = await response.json() as { copyName?: unknown };
    if (typeof result.copyName !== "string" || !result.copyName) throw new Error("Invalid recovery copy response.");
    return { copyName: result.copyName };
  }

  async close(signal?: AbortSignal): Promise<void> {
    const response = await this.#fetch("/api/session/close", { method: "POST", headers: this.#headers(), signal });
    if ([401, 404, 410].includes(response.status)) return;
    if (!response.ok) throw new Error(`Close failed (${response.status}).`);
  }

  /** Streams session events, reconnecting until aborted or closed. */
  async events(signal: AbortSignal, onEvent: (event: SessionEvent) => void, options: { baselineGeneration: number; retryMs?: number }): Promise<void> {
    const retryMs = options.retryMs ?? 1000;
    while (!signal.aborted) {
      let closed = false;
      try {
        const response = await this.#fetch("/api/session/events", {
          headers: { ...this.#headers(), [BASELINE_GENERATION_HEADER]: String(options.baselineGeneration) }, signal,
        });
        if ([401, 404, 410].includes(response.status)) {
          onEvent({ type: "closed" });
          return;
        }
        if (response.ok && response.body) {
          const reader = response.body.getReader();
          const decoder = new TextDecoder();
          const parser = new SseParser();
          for (;;) {
            const chunk = await reader.read();
            if (chunk.done) break;
            for (const raw of parser.feed(decoder.decode(chunk.value, { stream: true }))) {
              if (signal.aborted) return;
              const event = toSessionEvent(raw);
              if (!event) continue;
              onEvent(event);
              if (event.type === "closed" || event.type === "rejoinRequired") {
                closed = true;
                break;
              }
            }
            if (closed) break;
          }
        }
      } catch {
        if (signal.aborted) return;
      }
      if (closed || signal.aborted) return;
      await new Promise((resolve) => setTimeout(resolve, retryMs));
    }
  }
}
