import { decodeStateVector, decodeUpdate, diffUpdate, encodeStateVectorFromUpdate, mergeUpdates } from "yjs";

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
  queries: Uint8Array[];
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
  const queries: Uint8Array[] = [];
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
      if (subtype === SYNC_STEP_1) queries.push(payload);
      else documents.push({ subtype, payload });
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

  return { documents, queries, hasAuth };
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

export interface RetainedEntry {
  seq: number;
  bytes: Uint8Array;
}

export interface Checkpoint extends RetainedEntry {
  version: 1;
}

export interface LogMutation {
  puts: readonly RetainedEntry[];
  deletes: readonly number[];
  checkpoint?: Checkpoint;
}

export class RoomCapacityError extends Error {}

function syncFrame(subtype: number, payload: Uint8Array): Uint8Array {
  return encodeFrame([Uint8Array.of(0, subtype), encodeVarUint(payload.length), payload]);
}

function payloads(frame: Uint8Array): Uint8Array[] {
  const decoded = decodeFrame(frame);
  if (!decoded || decoded.hasAuth) throw new Error("Invalid stored collaboration frame");
  for (const message of decoded.documents) decodeUpdate(message.payload);
  for (const vector of decoded.queries) decodeStateVector(vector);
  return decoded.documents.map(message => message.payload);
}

export class RetainedUpdateLog {
  private updates: RetainedEntry[] = [];
  private base: Checkpoint | undefined;
  private retainedBytes = 0;
  private nextSeq = 0;
  private mergedUpdate: Uint8Array | undefined;

  constructor(
    private readonly maxCount: number,
    private readonly maxBytes: number,
    private readonly maxEntryBytes = 64 * 1024,
  ) {}

  restore(stored: readonly RetainedEntry[], checkpoint?: Checkpoint): LogMutation | null {
    this.clear();
    if (checkpoint) {
      if (checkpoint.version !== 1 || !Number.isSafeInteger(checkpoint.seq) || checkpoint.seq < 0 ||
        checkpoint.bytes.byteLength > this.maxBytes || payloads(checkpoint.bytes).length !== 1) {
        throw new Error("Invalid relay checkpoint");
      }
      this.base = { ...checkpoint, bytes: checkpoint.bytes.slice() };
      this.nextSeq = checkpoint.seq + 1;
    }
    const puts: RetainedEntry[] = [];
    const deletes: number[] = [];
    for (const entry of [...stored].sort((a, b) => a.seq - b.seq)) {
      if (!Number.isSafeInteger(entry.seq) || entry.seq < 0) throw new Error("Invalid update sequence");
      this.nextSeq = Math.max(this.nextSeq, entry.seq + 1);
      if (this.base && entry.seq <= this.base.seq) { deletes.push(entry.seq); continue; }
      try { payloads(entry.bytes); } catch { deletes.push(entry.seq); continue; }
      const retained = retainDocumentMessages(entry.bytes);
      if (!retained) { deletes.push(entry.seq); continue; }
      if (retained.byteLength > this.maxBytes) throw new RoomCapacityError("Stored update exceeds room capacity");
      if (retained.length !== entry.bytes.length) puts.push({ seq: entry.seq, bytes: retained });
      this.updates.push({ seq: entry.seq, bytes: retained });
      this.retainedBytes += retained.length;
    }
    if (this.needsCheckpoint()) {
      const compacted = this.checkpoint()!;
      return { ...compacted, deletes: [...deletes, ...compacted.deletes] };
    }
    return puts.length || deletes.length ? { puts, deletes } : null;
  }

  retain(update: Uint8Array): LogMutation | null {
    payloads(update);
    const retained = retainDocumentMessages(update);
    if (!retained) return null;
    if (retained.byteLength > this.maxBytes) throw new RoomCapacityError("Update exceeds room capacity");
    if (!Number.isSafeInteger(this.nextSeq + 1)) throw new RoomCapacityError("Room sequence exhausted");
    const entry = { seq: this.nextSeq, bytes: retained };
    const candidate = [...this.updates, entry];
    if (candidate.length >= this.maxCount || retained.length + this.retainedBytes + (this.base?.bytes.length ?? 0) > this.maxBytes || retained.length > this.maxEntryBytes) {
      const mutation = this.compact(candidate);
      this.adoptCheckpoint(mutation.checkpoint!);
      return mutation;
    }
    this.nextSeq++;
    this.updates.push(entry);
    this.mergedUpdate = undefined;
    this.retainedBytes += retained.length;
    return { puts: [entry], deletes: [] };
  }

  checkpoint(): LogMutation | null {
    if (this.updates.length === 0) return null;
    const mutation = this.compact(this.updates);
    this.adoptCheckpoint(mutation.checkpoint!);
    return mutation;
  }

  syncRequest(): Uint8Array {
    return syncFrame(0, encodeStateVectorFromUpdate(this.merged()));
  }

  responses(frame: Uint8Array): Uint8Array[] {
    const queries = decodeFrame(frame)?.queries ?? [];
    if (queries.length === 0) return [];
    if (queries.length > 1) throw new Error("Only one state-vector query is allowed per frame");
    decodeStateVector(queries[0]);
    const response = syncFrame(1, diffUpdate(this.merged(), queries[0]));
    if (response.length > this.maxBytes) throw new RoomCapacityError("Sync response exceeds room capacity");
    return [response];
  }

  replay(send: (update: Uint8Array) => void): void {
    if (this.base) send(this.base.bytes.slice());
    for (const entry of this.updates) send(entry.bytes.slice());
  }

  snapshot(): Uint8Array[] {
    const frames: Uint8Array[] = [];
    this.replay(frame => frames.push(frame));
    return frames;
  }

  clear(): void {
    this.updates = [];
    this.base = undefined;
    this.retainedBytes = 0;
    this.nextSeq = 0;
    this.mergedUpdate = undefined;
  }

  private needsCheckpoint(): boolean {
    return this.updates.length >= this.maxCount ||
      this.retainedBytes + (this.base?.bytes.length ?? 0) > this.maxBytes ||
      this.updates.some(entry => entry.bytes.length > this.maxEntryBytes);
  }

  private merged(entries: readonly RetainedEntry[] = this.updates): Uint8Array {
    if (entries === this.updates && this.mergedUpdate) return this.mergedUpdate;
    const updates = [...(this.base ? payloads(this.base.bytes) : []), ...entries.flatMap(entry => payloads(entry.bytes))];
    const merged = updates.length ? mergeUpdates(updates) : Uint8Array.of(0, 0);
    if (entries === this.updates) this.mergedUpdate = merged;
    return merged;
  }

  private compact(entries: readonly RetainedEntry[]): LogMutation {
    const bytes = syncFrame(1, this.merged(entries));
    if (bytes.length > this.maxBytes) throw new RoomCapacityError("Checkpoint exceeds room capacity; export the document before continuing");
    return { puts: [], deletes: entries.map(entry => entry.seq), checkpoint: { version: 1, seq: entries.at(-1)!.seq, bytes } };
  }

  private adoptCheckpoint(checkpoint: Checkpoint): void {
    this.base = checkpoint;
    this.mergedUpdate = undefined;
    this.updates = [];
    this.retainedBytes = 0;
    this.nextSeq = checkpoint.seq + 1;
  }
}
