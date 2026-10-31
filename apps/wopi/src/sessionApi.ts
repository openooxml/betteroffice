import type { FileSession } from "./session/fileSession";
import type { JoinPayload } from "./replica/replica";
import type { CollabRelay } from "./relay/relay";
import { isSocketSaveBarrier, type SaveBarrier } from "./relay/sequence";

/** JoinPayload over JSON: the binary initialUpdate is a base64 string. */
export interface EncodedJoinPayload {
  format: JoinPayload["format"];
  initialUpdate: string;
}

function toBase64(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString("base64");
}

function fromBase64(value: string): Uint8Array {
  return new Uint8Array(Buffer.from(value, "base64"));
}

/** Encodes a join payload for JSON transport. */
export function encodeJoinPayload(payload: JoinPayload): EncodedJoinPayload {
  return { format: payload.format, initialUpdate: toBase64(payload.initialUpdate) };
}

/** Decodes a join payload from JSON transport. */
export function decodeJoinPayload(payload: EncodedJoinPayload): JoinPayload {
  return { format: payload.format, initialUpdate: fromBase64(payload.initialUpdate) };
}

const CONTENT_TYPES: Record<string, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

/** Source content type for a format. */
export function sourceContentType(format: string | null): string {
  return (format && CONTENT_TYPES[format]) || "application/octet-stream";
}

export class SaveSyncError extends Error {
  constructor(readonly kind: "notSynced" | "invalidSaveBarrier") {
    super(kind);
    this.name = "SaveSyncError";
  }
}

export function decodeSaveBarrier(body: Uint8Array): SaveBarrier {
  let value: unknown;
  try {
    value = JSON.parse(new TextDecoder().decode(body));
  } catch {
    throw new SaveSyncError("invalidSaveBarrier");
  }
  if (!isSocketSaveBarrier(value)) throw new SaveSyncError("invalidSaveBarrier");
  return { connectionId: value.connectionId, seq: value.seq };
}

export async function requestOrderedSave(
  session: FileSession,
  userId: string,
  barrier: SaveBarrier,
  options: { timeoutMs?: number; signal?: AbortSignal; relay?: Pick<CollabRelay, "appliedSequence"> } = {},
) {
  const deadline = Date.now() + (options.timeoutMs ?? 5000);
  for (;;) {
    if (options.signal?.aborted) throw new SaveSyncError("notSynced");
    const applied = options.relay?.appliedSequence(session, userId, barrier.connectionId);
    if (applied !== undefined && applied !== null && applied >= barrier.seq) break;
    const remaining = deadline - Date.now();
    if (remaining <= 0) throw new SaveSyncError("notSynced");
    await new Promise((resolve) => setTimeout(resolve, Math.min(50, remaining)));
  }
  return session.requestSave(userId);
}
