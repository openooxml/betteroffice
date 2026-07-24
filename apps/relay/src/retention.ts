const TOP_LEVEL_SYNC = 0;
const TOP_LEVEL_AWARENESS = 1;
const TOP_LEVEL_AUTH = 2;
const TOP_LEVEL_QUERY_AWARENESS = 3;
const MAX_SYNC_SUBTYPE = 2;
const AUTH_PERMISSION_DENIED = 0;
const MAX_MESSAGES_PER_FRAME = 4096;
const MAX_VAR_UINT = Number.MAX_SAFE_INTEGER;

interface DocumentMessage {
  subtype: number;
  payload: Uint8Array;
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

function decodeDocumentMessages(
  frame: Uint8Array,
): DocumentMessage[] | null {
  if (frame.byteLength === 0) return null;
  const decoder = new FrameDecoder(frame);
  const documents: DocumentMessage[] = [];
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
      documents.push({ subtype, payload });
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
    } else if (type !== TOP_LEVEL_QUERY_AWARENESS) {
      return null;
    }
  }

  return documents;
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
  const documents = decodeDocumentMessages(frame);
  if (!documents || documents.length === 0) return null;

  const parts: Uint8Array[] = [];
  for (const document of documents) {
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

export function isDocumentFrame(frame: Uint8Array): boolean {
  const retained = retainDocumentMessages(frame);
  return retained !== null && bytesEqual(retained, frame);
}

export class RetainedUpdateLog {
  private updates: Uint8Array[] = [];
  private retainedBytes = 0;

  constructor(
    private readonly maxCount: number,
    private readonly maxBytes: number,
  ) {}

  restore(updates: readonly Uint8Array[]): boolean {
    this.updates = [];
    this.retainedBytes = 0;
    let changed = false;
    for (const update of updates) {
      const retained =
        update.byteLength <= this.maxBytes
          ? retainDocumentMessages(update)
          : null;
      if (!retained || retained.byteLength > this.maxBytes) {
        changed = true;
        continue;
      }
      this.updates.push(retained);
      this.retainedBytes += retained.byteLength;
      changed = this.trim() || !bytesEqual(retained, update) || changed;
    }
    return changed;
  }

  retain(update: Uint8Array): boolean {
    if (update.byteLength > this.maxBytes) return false;
    const retained = retainDocumentMessages(update);
    if (!retained || retained.byteLength > this.maxBytes) return false;
    this.updates.push(retained);
    this.retainedBytes += retained.byteLength;
    this.trim();
    return true;
  }

  replay(send: (update: Uint8Array) => void): void {
    for (const update of this.updates) send(update.slice());
  }

  snapshot(): Uint8Array[] {
    return this.updates.map((update) => update.slice());
  }

  private trim(): boolean {
    let changed = false;
    while (
      this.updates.length > this.maxCount ||
      this.retainedBytes > this.maxBytes
    ) {
      const removed = this.updates.shift();
      if (removed) this.retainedBytes -= removed.byteLength;
      changed = true;
    }
    return changed;
  }
}
