import { SseParser, toSessionEvent, type SessionEvent } from "./sse";

/** Join payload as served by /api/session/join. */
export interface JoinResponse {
  format: "docx" | "xlsx" | "pptx";
  initialUpdate: Uint8Array;
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
  return { format: record.format, initialUpdate: fromBase64(record.initialUpdate) };
}

function defaultFetch(input: RequestInfo | URL, init?: RequestInit): Promise<Response> {
  return globalThis.fetch(input, init);
}

/** Session API client. The handle is only ever sent as a Bearer header. */
export class SessionApi {
  readonly #handle: string;
  readonly #fetch: typeof fetch;

  constructor(handle: string, fetchFn: typeof fetch = defaultFetch) {
    this.#handle = handle;
    this.#fetch = fetchFn;
  }

  #headers(): Record<string, string> {
    return { Authorization: `Bearer ${this.#handle}` };
  }

  async source(signal?: AbortSignal): Promise<Uint8Array> {
    const response = await this.#fetch("/api/session/source", { headers: this.#headers(), signal });
    if (!response.ok) throw new Error(`Could not load the document (${response.status}).`);
    return new Uint8Array(await response.arrayBuffer());
  }

  async join(signal?: AbortSignal): Promise<JoinResponse> {
    const response = await this.#fetch("/api/session/join", { headers: this.#headers(), signal });
    if (!response.ok) throw new Error(`Could not join the session (${response.status}).`);
    return decodeJoinResponse(await response.json());
  }

  async save(): Promise<{ itemVersion: string | null }> {
    const response = await this.#fetch("/api/session/save", { method: "POST", headers: this.#headers() });
    if (!response.ok) throw new Error(`Save failed (${response.status}).`);
    const body = (await response.json()) as { itemVersion?: unknown };
    return { itemVersion: typeof body.itemVersion === "string" ? body.itemVersion : null };
  }

  async close(): Promise<void> {
    await this.#fetch("/api/session/close", { method: "POST", headers: this.#headers() }).catch(() => undefined);
  }

  /** Streams session events, reconnecting until aborted or closed. */
  async events(signal: AbortSignal, onEvent: (event: SessionEvent) => void, options: { retryMs?: number } = {}): Promise<void> {
    const retryMs = options.retryMs ?? 1000;
    while (!signal.aborted) {
      let closed = false;
      try {
        const response = await this.#fetch("/api/session/events", { headers: this.#headers(), signal });
        if (response.ok && response.body) {
          const reader = response.body.getReader();
          const decoder = new TextDecoder();
          const parser = new SseParser();
          for (;;) {
            const chunk = await reader.read();
            if (chunk.done) break;
            for (const raw of parser.feed(decoder.decode(chunk.value, { stream: true }))) {
              const event = toSessionEvent(raw);
              if (!event) continue;
              onEvent(event);
              if (event.type === "closed") closed = true;
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
