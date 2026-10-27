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
