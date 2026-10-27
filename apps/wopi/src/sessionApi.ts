import type { JoinPayload } from "./replica/replica";

/** JoinPayload over JSON: binary fields are base64 strings. */
export type EncodedJoinPayload =
  | { format: "docx"; state: string; peerMetadata: string }
  | { format: "xlsx"; update: string }
  | { format: "pptx"; initialUpdate: string };

function toBase64(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString("base64");
}

function fromBase64(value: string): Uint8Array {
  return new Uint8Array(Buffer.from(value, "base64"));
}

/** Encodes a join payload for JSON transport. */
export function encodeJoinPayload(payload: JoinPayload): EncodedJoinPayload {
  if (payload.format === "docx") {
    return { format: "docx", state: toBase64(payload.state), peerMetadata: toBase64(payload.peerMetadata) };
  }
  if (payload.format === "xlsx") return { format: "xlsx", update: toBase64(payload.update) };
  return { format: "pptx", initialUpdate: toBase64(payload.initialUpdate) };
}

/** Decodes a join payload from JSON transport. */
export function decodeJoinPayload(payload: EncodedJoinPayload): JoinPayload {
  if (payload.format === "docx") {
    return { format: "docx", state: fromBase64(payload.state), peerMetadata: fromBase64(payload.peerMetadata) };
  }
  if (payload.format === "xlsx") return { format: "xlsx", update: fromBase64(payload.update) };
  return { format: "pptx", initialUpdate: fromBase64(payload.initialUpdate) };
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
