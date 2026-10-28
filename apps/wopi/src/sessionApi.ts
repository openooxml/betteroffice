import { decodeStateVector } from "yjs";
import type { FileSession } from "./session/fileSession";
import type { JoinPayload } from "./replica/replica";

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
  constructor(readonly kind: "notSynced" | "invalidStateVector") {
    super(kind);
    this.name = "SaveSyncError";
  }
}

export function decodeSaveStateVector(body: Uint8Array): Uint8Array {
  try {
    const value = JSON.parse(new TextDecoder().decode(body)) as { stateVector?: unknown } | null;
    const encoded = value?.stateVector;
    if (typeof encoded !== "string" || encoded === "" || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(encoded)) {
      throw new SaveSyncError("invalidStateVector");
    }
    const vector = fromBase64(encoded);
    decodeStateVector(vector.slice());
    return vector;
  } catch {
    throw new SaveSyncError("invalidStateVector");
  }
}

export async function waitForStateVector(
  readVector: () => Uint8Array,
  wanted: Uint8Array,
  options: { timeoutMs?: number; signal?: AbortSignal } = {},
): Promise<void> {
  const required = decodeStateVector(wanted.slice());
  const deadline = Date.now() + (options.timeoutMs ?? 5000);
  for (;;) {
    if (options.signal?.aborted) throw new SaveSyncError("notSynced");
    const current = decodeStateVector(readVector().slice());
    if ([...required].every(([client, clock]) => (current.get(client) ?? 0) >= clock)) return;
    const remaining = deadline - Date.now();
    if (remaining <= 0) throw new SaveSyncError("notSynced");
    await new Promise((resolve) => setTimeout(resolve, Math.min(50, remaining)));
  }
}

export async function requestOrderedSave(
  session: FileSession,
  userId: string,
  stateVector: Uint8Array,
  options: { timeoutMs?: number; signal?: AbortSignal } = {},
) {
  await waitForStateVector(() => {
    if (!session.replica) throw new SaveSyncError("notSynced");
    return session.replica.encodeStateVector();
  }, stateVector, options);
  return session.requestSave(userId);
}
