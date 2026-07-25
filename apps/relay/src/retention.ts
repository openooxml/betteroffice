const TOP_LEVEL_SYNC = 0;
const TOP_LEVEL_AWARENESS = 1;
const TOP_LEVEL_AUTH = 2;
const TOP_LEVEL_QUERY_AWARENESS = 3;
/** A state-vector query: valid to relay, never worth replaying to joiners. */
const SYNC_STEP_1 = 0;
const MAX_SYNC_SUBTYPE = 2;
const AUTH_PERMISSION_DENIED = 0;
const MAX_MESSAGES_PER_FRAME = 4096;
const MAX_VAR_UINT = Number.MAX_SAFE_INTEGER;

/** `document` frames carry state worth retaining, `transient` ones do not. */
export type FrameKind = "document" | "transient" | "auth" | "invalid";

interface DocumentMessage {
  subtype: number;
  payload: Uint8Array;
}

interface DecodedFrame {
  documents: DocumentMessage[];
  hasAuth: boolean;
}

class FrameDecoder {
  private offset = 0;

  constructor(private readonly bytes: Uint8Array) {}

  get done(): boolean {
    return this.offset === this.bytes.byteLength;
  }

  readVarUint(): number | null {
    let value = 0;
    let multiplier = 1;
    let count = 0;

    while (true) {
      if (this.offset >= this.bytes.byteLength) return null;
      const byte = this.bytes[this.offset++];
      const digit = byte & 0x7f;
      if (digit > Math.floor((MAX_VAR_UINT - value) / multiplier)) {
        return null;
      }

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

  readVarUint8Array(): Uint8Array | null {
    const length = this.readVarUint();
    if (length === null || length > this.bytes.byteLength - this.offset) {
      return null;
    }
    const value = this.bytes.subarray(this.offset, this.offset + length);
    this.offset += length;
    return value;
  }
}

function encodeVarUint(value: number): Uint8Array {
  const bytes: number[] = [];
  let remaining = value;
  while (remaining >= 128) {
    bytes.push((remaining % 128) | 0x80);
    remaining = Math.floor(remaining / 128);
  }
  bytes.push(remaining);
  return Uint8Array.from(bytes);
}

function encodeFrame(parts: readonly Uint8Array[]): Uint8Array {
  const frame = new Uint8Array(
    parts.reduce((length, part) => length + part.byteLength, 0),
  );
  let offset = 0;
  for (const part of parts) {
    frame.set(part, offset);
    offset += part.byteLength;
  }
  return frame;
}

function decodeFrame(frame: Uint8Array): DecodedFrame | null {
  if (frame.byteLength === 0) return null;
  const decoder = new FrameDecoder(frame);
  const documents: DocumentMessage[] = [];
  let hasAuth = false;
  let messageCount = 0;

  while (!decoder.done) {
    if (messageCount >= MAX_MESSAGES_PER_FRAME) return null;
    messageCount += 1;
    const type = decoder.readVarUint();
    if (type === null) return null;

    if (type === TOP_LEVEL_SYNC) {
      const subtype = decoder.readVarUint();
      const payload = decoder.readVarUint8Array();
      if (
        subtype === null ||
        subtype > MAX_SYNC_SUBTYPE ||
        payload === null
      ) {
        return null;
      }
      if (subtype !== SYNC_STEP_1) documents.push({ subtype, payload });
    } else if (type === TOP_LEVEL_AWARENESS) {
      if (decoder.readVarUint8Array() === null) return null;
    } else if (type === TOP_LEVEL_AUTH) {
      const subtype = decoder.readVarUint();
      const reason = decoder.readVarUint8Array();
      if (
        subtype !== AUTH_PERMISSION_DENIED ||
        reason === null ||
        !isValidUtf8(reason)
      ) {
        return null;
      }
      hasAuth = true;
    } else if (type !== TOP_LEVEL_QUERY_AWARENESS) {
      return null;
    }
  }

  return { documents, hasAuth };
}

export function classifyFrame(frame: Uint8Array): FrameKind {
  const decoded = decodeFrame(frame);
  if (!decoded) return "invalid";
  if (decoded.hasAuth) return "auth";
  return decoded.documents.length > 0 ? "document" : "transient";
}

function isValidUtf8(bytes: Uint8Array): boolean {
  try {
    new TextDecoder("utf-8", { fatal: true, ignoreBOM: false }).decode(bytes);
    return true;
  } catch {
    return false;
  }
}

function retainDocumentMessages(frame: Uint8Array): Uint8Array | null {
  const decoded = decodeFrame(frame);
  if (!decoded || decoded.hasAuth || decoded.documents.length === 0) return null;

  const parts: Uint8Array[] = [];
  for (const document of decoded.documents) {
    parts.push(
      encodeVarUint(TOP_LEVEL_SYNC),
      encodeVarUint(document.subtype),
      encodeVarUint(document.payload.byteLength),
      document.payload,
    );
  }
  return encodeFrame(parts);
}

function bytesEqual(first: Uint8Array, second: Uint8Array): boolean {
  if (first.byteLength !== second.byteLength) return false;
  return first.every((byte, index) => byte === second[index]);
}

export interface RetainedEntry {
  seq: number;
  bytes: Uint8Array;
}

/** The storage writes that bring the persisted log back in line with memory. */
export interface LogMutation {
  puts: readonly RetainedEntry[];
  deletes: readonly number[];
}

export class RetainedUpdateLog {
  private updates: RetainedEntry[] = [];
  private retainedBytes = 0;
  private nextSeq = 0;

  constructor(
    private readonly maxCount: number,
    private readonly maxBytes: number,
  ) {}

  restore(stored: readonly RetainedEntry[]): LogMutation | null {
    this.clear();
    const puts: RetainedEntry[] = [];
    const deletes: number[] = [];
    for (const entry of [...stored].sort((first, second) => first.seq - second.seq)) {
      this.nextSeq = Math.max(this.nextSeq, entry.seq + 1);
      const retained =
        entry.bytes.byteLength <= this.maxBytes
          ? retainDocumentMessages(entry.bytes)
          : null;
      if (!retained || retained.byteLength > this.maxBytes) {
        deletes.push(entry.seq);
        continue;
      }
      if (!bytesEqual(retained, entry.bytes)) {
        puts.push({ seq: entry.seq, bytes: retained });
      }
      this.updates.push({ seq: entry.seq, bytes: retained });
      this.retainedBytes += retained.byteLength;
      deletes.push(...this.trim());
    }

    const evicted = new Set(deletes);
    const rewrites = puts.filter((entry) => !evicted.has(entry.seq));
    if (rewrites.length === 0 && deletes.length === 0) return null;
    return { puts: rewrites, deletes };
  }

  retain(update: Uint8Array): LogMutation | null {
    if (update.byteLength > this.maxBytes) return null;
    const retained = retainDocumentMessages(update);
    if (!retained || retained.byteLength > this.maxBytes) return null;
    const entry: RetainedEntry = { seq: this.nextSeq++, bytes: retained };
    this.updates.push(entry);
    this.retainedBytes += retained.byteLength;
    return { puts: [entry], deletes: this.trim() };
  }

  replay(send: (update: Uint8Array) => void): void {
    for (const entry of this.updates) send(entry.bytes.slice());
  }

  snapshot(): Uint8Array[] {
    return this.updates.map((entry) => entry.bytes.slice());
  }

  clear(): void {
    this.updates = [];
    this.retainedBytes = 0;
    this.nextSeq = 0;
  }

  private trim(): number[] {
    const evicted: number[] = [];
    while (
      this.updates.length > this.maxCount ||
      this.retainedBytes > this.maxBytes
    ) {
      const removed = this.updates.shift();
      if (!removed) break;
      this.retainedBytes -= removed.bytes.byteLength;
      evicted.push(removed.seq);
    }
    return evicted;
  }
}
