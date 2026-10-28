/** One parsed server-sent event. */
export interface SseEvent {
  event: string;
  data: string;
}

/** Incremental SSE parser for a fetched event stream. */
export class SseParser {
  #buffer = "";

  /** Feeds one chunk and returns the events it completed. */
  feed(chunk: string): SseEvent[] {
    this.#buffer += chunk;
    const events: SseEvent[] = [];
    for (;;) {
      const end = this.#buffer.indexOf("\n\n");
      const crEnd = this.#buffer.indexOf("\r\n\r\n");
      let index = -1;
      let skip = 0;
      if (end !== -1 && (crEnd === -1 || end < crEnd)) {
        index = end;
        skip = 2;
      } else if (crEnd !== -1) {
        index = crEnd;
        skip = 4;
      }
      if (index === -1) break;
      const block = this.#buffer.slice(0, index);
      this.#buffer = this.#buffer.slice(index + skip);
      let name = "message";
      const dataLines: string[] = [];
      for (const line of block.split(/\r?\n/)) {
        if (line.startsWith("event:")) name = line.slice("event:".length).trim();
        else if (line.startsWith("data:")) dataLines.push(line.slice("data:".length).trim());
      }
      if (dataLines.length > 0) events.push({ event: name, data: dataLines.join("\n") });
    }
    return events;
  }
}

/** Session events understood by the page. */
export type SessionEvent =
  | { type: "saved"; itemVersion: string | null }
  | { type: "saveFailed"; kind: string }
  | { type: "conflict"; copyName: string }
  | { type: "viewOnly"; reason: string }
  | { type: "tokenExpiring"; expiresAt: number | null }
  | { type: "closed" };

/** Maps a raw SSE event to a typed session event, or null. */
export function toSessionEvent(raw: SseEvent): SessionEvent | null {
  let data: Record<string, unknown> = {};
  try {
    data = JSON.parse(raw.data) as Record<string, unknown>;
  } catch {
    return null;
  }
  switch (raw.event) {
    case "saved":
      return { type: "saved", itemVersion: typeof data.itemVersion === "string" ? data.itemVersion : null };
    case "saveFailed":
      return { type: "saveFailed", kind: typeof data.kind === "string" ? data.kind : "unknown" };
    case "conflict":
      return { type: "conflict", copyName: typeof data.copyName === "string" ? data.copyName : typeof data.name === "string" ? data.name : "" };
    case "viewOnly":
      return { type: "viewOnly", reason: typeof data.reason === "string" ? data.reason : "" };
    case "tokenExpiring":
      return { type: "tokenExpiring", expiresAt: typeof data.expiresAt === "number" ? data.expiresAt : null };
    case "closed":
      return { type: "closed" };
    default:
      return null;
  }
}
