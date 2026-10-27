import { MAX_AWARENESS_PAYLOAD_BYTES } from "../../../../shared/collaboration-limits";

export { MAX_AWARENESS_PAYLOAD_BYTES, MAX_COLLABORATION_FRAME_BYTES } from "../../../../shared/collaboration-limits";

export const TOP_LEVEL_SYNC = 0;
export const TOP_LEVEL_AWARENESS = 1;
export const TOP_LEVEL_AUTH = 2;
export const TOP_LEVEL_QUERY_AWARENESS = 3;
export const SYNC_STEP_1 = 0;
export const SYNC_STEP_2 = 1;
export const SYNC_UPDATE = 2;

const MAX_MESSAGES_PER_FRAME = 4096;
const MAX_VAR_UINT = Number.MAX_SAFE_INTEGER;

/** One decoded message in a collaboration frame. */
export type FrameMessage =
  | { type: "sync-step-1"; stateVector: Uint8Array }
  | { type: "sync-step-2"; update: Uint8Array }
  | { type: "update"; update: Uint8Array }
  | { type: "awareness"; update: Uint8Array; clientIds: number[] }
  | { type: "auth"; reason: string }
  | { type: "query-awareness" };

/** Classification matching the relay retention codec. */
export type FrameKind = "document" | "transient" | "auth" | "invalid" | "oversize-awareness";

export interface ClassifiedFrame {
  kind: FrameKind;
  hasAwareness: boolean;
  awarenessBytes: number;
  messages: FrameMessage[];
}

class Decoder {
  offset = 0;
  constructor(private bytes: Uint8Array) {}
  get done(): boolean {
    return this.offset === this.bytes.byteLength;
  }
  readVarUint(): number | null {
    let value = 0;
    let multiplier = 1;
    let count = 0;
    while (true) {
      if (this.offset >= this.bytes.byteLength) return null;
      const byte = this.bytes[this.offset++]!;
      const digit = byte & 0x7f;
      if (digit > Math.floor((MAX_VAR_UINT - value) / multiplier)) return null;
      value += digit * multiplier;
      count += 1;
      if ((byte & 0x80) === 0) {
        if (count > 1 && digit === 0) return null;
        return value;
      }
      if (count >= 8) return null;
      multiplier *= 128;
    }
  }
  readBytes(): Uint8Array | null {
    const length = this.readVarUint();
    if (length === null || length > this.bytes.byteLength - this.offset) return null;
    const value = this.bytes.slice(this.offset, this.offset + length);
    this.offset += length;
    return value;
  }
}

/** Encodes one varUint. */
export function encodeVarUint(value: number): Uint8Array {
  const bytes: number[] = [];
  let remaining = value;
  while (remaining >= 128) {
    bytes.push((remaining % 128) | 0x80);
    remaining = Math.floor(remaining / 128);
  }
  bytes.push(remaining);
  return Uint8Array.from(bytes);
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.byteLength, 0));
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.byteLength;
  }
  return out;
}

function encodeSync(subtype: number, payload: Uint8Array): Uint8Array {
  return concat([encodeVarUint(TOP_LEVEL_SYNC), encodeVarUint(subtype), encodeVarUint(payload.byteLength), payload]);
}

/** Encodes a sync step 1 frame. */
export function encodeSyncStep1(stateVector: Uint8Array): Uint8Array {
  return encodeSync(SYNC_STEP_1, stateVector);
}

/** Encodes a sync step 2 frame. */
export function encodeSyncStep2(update: Uint8Array): Uint8Array {
  return encodeSync(SYNC_STEP_2, update);
}

/** Encodes a sync update frame. */
export function encodeUpdateFrame(update: Uint8Array): Uint8Array {
  return encodeSync(SYNC_UPDATE, update);
}

function encodeVarString(value: string): Uint8Array {
  const bytes = new TextEncoder().encode(value);
  return concat([encodeVarUint(bytes.byteLength), bytes]);
}

/** One awareness entry, state null means removal. */
export interface AwarenessEntry {
  clientId: number;
  clock: number;
  state: unknown;
}

/** Encodes an awareness payload. */
export function encodeAwarenessUpdate(entries: readonly AwarenessEntry[]): Uint8Array {
  const parts: Uint8Array[] = [encodeVarUint(entries.length)];
  for (const entry of entries) {
    parts.push(
      encodeVarUint(entry.clientId),
      encodeVarUint(entry.clock),
      encodeVarString(entry.state === null ? "null" : JSON.stringify(entry.state)),
    );
  }
  return concat(parts);
}

/** Encodes an awareness frame. */
export function encodeAwarenessFrame(entries: readonly AwarenessEntry[]): Uint8Array {
  const update = encodeAwarenessUpdate(entries);
  return concat([encodeVarUint(TOP_LEVEL_AWARENESS), encodeVarUint(update.byteLength), update]);
}

/** Encodes an awareness removal frame for the given clients. */
export function encodeAwarenessRemoval(entries: readonly { clientId: number; clock: number }[]): Uint8Array {
  return encodeAwarenessFrame(entries.map((e) => ({ clientId: e.clientId, clock: e.clock + 1, state: null })));
}

/** Client ids and clocks in an awareness payload, empty when malformed. */
export function awarenessEntries(update: Uint8Array): { clientId: number; clock: number }[] {
  const decoder = new Decoder(update);
  const count = decoder.readVarUint();
  if (count === null || count > 4096) return [];
  const entries: { clientId: number; clock: number }[] = [];
  for (let i = 0; i < count; i += 1) {
    const clientId = decoder.readVarUint();
    const clock = decoder.readVarUint();
    const state = decoder.readBytes();
    if (clientId === null || clock === null || state === null) return [];
    entries.push({ clientId, clock });
  }
  return entries;
}

function isValidUtf8(bytes: Uint8Array): boolean {
  try {
    new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    return true;
  } catch {
    return false;
  }
}

const OPEN_SYNC_CLIENT = BigInt(Number.MAX_SAFE_INTEGER - 1);

/** Removes reserved sync request entries from a state vector. */
export function stripReservedSyncEntries(stateVector: Uint8Array): Uint8Array {
  let offset = 0;
  const read = (): bigint | null => {
    let value = 0n;
    let shift = 0n;
    for (let i = 0; i < 10 && offset < stateVector.byteLength; i += 1) {
      const byte = stateVector[offset++]!;
      value |= BigInt(byte & 0x7f) << shift;
      if (byte < 0x80) return value;
      shift += 7n;
    }
    return null;
  };
  const count = read();
  if (count === null) return stateVector;
  const kept: Uint8Array[] = [];
  let removed = 0;
  for (let left = count; left > 0n; left -= 1n) {
    const start = offset;
    const client = read();
    const clock = read();
    if (client === null || clock === null) return stateVector;
    if (client >= OPEN_SYNC_CLIENT) removed += 1;
    else kept.push(stateVector.subarray(start, offset));
  }
  if (offset !== stateVector.byteLength || removed === 0) return stateVector;
  const parts = [encodeVarUint(Number(count) - removed), ...kept];
  const out = new Uint8Array(parts.reduce((n, p) => n + p.byteLength, 0));
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.byteLength;
  }
  return out;
}

/** Decodes every message in a frame, throwing on malformed input. */
export function decodeMessages(frame: Uint8Array): FrameMessage[] {
  const classified = classifyFrame(frame);
  if (classified.kind === "invalid") throw new Error("malformed collaboration frame");
  return classified.messages;
}

/** Classifies a frame with the retention codec semantics. */
export function classifyFrame(frame: Uint8Array): ClassifiedFrame {
  const invalid: ClassifiedFrame = { kind: "invalid", hasAwareness: false, awarenessBytes: 0, messages: [] };
  if (frame.byteLength === 0) return invalid;
  const decoder = new Decoder(frame);
  const messages: FrameMessage[] = [];
  let hasAwareness = false;
  let awarenessBytes = 0;
  let hasAuth = false;
  let hasDocument = false;
  let count = 0;
  while (!decoder.done) {
    if (count >= MAX_MESSAGES_PER_FRAME) return invalid;
    count += 1;
    const type = decoder.readVarUint();
    if (type === null) return invalid;
    if (type === TOP_LEVEL_SYNC) {
      const subtype = decoder.readVarUint();
      const payload = decoder.readBytes();
      if (subtype === null || subtype > SYNC_UPDATE || payload === null) return invalid;
      if (subtype === SYNC_STEP_1) messages.push({ type: "sync-step-1", stateVector: payload });
      else if (subtype === SYNC_STEP_2) {
        hasDocument = true;
        messages.push({ type: "sync-step-2", update: payload });
      } else {
        hasDocument = true;
        messages.push({ type: "update", update: payload });
      }
    } else if (type === TOP_LEVEL_AWARENESS) {
      const payload = decoder.readBytes();
      if (payload === null) return invalid;
      hasAwareness = true;
      awarenessBytes += payload.byteLength;
      messages.push({ type: "awareness", update: payload, clientIds: awarenessEntries(payload).map((e) => e.clientId) });
    } else if (type === TOP_LEVEL_AUTH) {
      const subtype = decoder.readVarUint();
      const reason = decoder.readBytes();
      if (subtype !== 0 || reason === null || !isValidUtf8(reason)) return invalid;
      hasAuth = true;
      messages.push({ type: "auth", reason: new TextDecoder().decode(reason) });
    } else if (type === TOP_LEVEL_QUERY_AWARENESS) {
      messages.push({ type: "query-awareness" });
    } else {
      return invalid;
    }
  }
  let kind: FrameKind = hasDocument ? "document" : "transient";
  if (hasAuth) kind = "auth";
  if (awarenessBytes > MAX_AWARENESS_PAYLOAD_BYTES) kind = "oversize-awareness";
  return { kind, hasAwareness, awarenessBytes, messages };
}
