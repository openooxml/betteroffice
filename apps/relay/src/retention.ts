const TOP_LEVEL_SYNC = 0;
const MAX_SYNC_SUBTYPE = 2;
const MAX_VAR_UINT = Number.MAX_SAFE_INTEGER;

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

  skip(length: number): boolean {
    if (length > this.bytes.byteLength - this.offset) return false;
    this.offset += length;
    return true;
  }
}

export function isDocumentFrame(frame: Uint8Array): boolean {
  if (frame.byteLength === 0) return false;
  const decoder = new FrameDecoder(frame);
  while (!decoder.done) {
    const type = decoder.readVarUint();
    if (type !== TOP_LEVEL_SYNC) return false;
    const subtype = decoder.readVarUint();
    if (subtype === null || subtype > MAX_SYNC_SUBTYPE) return false;
    const length = decoder.readVarUint();
    if (length === null || !decoder.skip(length)) return false;
  }
  return true;
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
      if (!isDocumentFrame(update) || update.byteLength > this.maxBytes) {
        changed = true;
        continue;
      }
      this.updates.push(update.slice());
      this.retainedBytes += update.byteLength;
      changed = this.trim() || changed;
    }
    return changed;
  }

  retain(update: Uint8Array): boolean {
    if (!isDocumentFrame(update) || update.byteLength > this.maxBytes) {
      return false;
    }
    this.updates.push(update.slice());
    this.retainedBytes += update.byteLength;
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
