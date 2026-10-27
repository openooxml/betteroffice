import { createHash } from "node:crypto";
import { closeSync, existsSync, fsyncSync, mkdirSync, openSync, readFileSync, unlinkSync, writeFileSync, writeSync } from "node:fs";
import { join } from "node:path";

/** Journal contents read back from disk. */
export interface JournalData {
  baseItemVersion: string | null;
  updates: Uint8Array[];
}

const MAGIC = Buffer.from("BOJ1\n", "utf8");

function record(prefix: string, payload: Buffer): Buffer {
  const body = Buffer.concat([Buffer.from(prefix, "utf8"), payload]);
  const header = Buffer.alloc(4);
  header.writeUInt32BE(body.length, 0);
  return Buffer.concat([header, body]);
}

/** Append only journal for one file key. */
export class FileJournal {
  readonly path: string;
  #fd: number | null = null;
  #begun = false;

  constructor(dir: string, fileKey: string) {
    const name = `${createHash("sha256").update(fileKey).digest("hex")}.journal`;
    this.path = join(dir, name);
    this.dir = dir;
  }

  readonly dir: string;

  /** Path for a file key inside a journal directory. */
  static pathFor(dir: string, fileKey: string): string {
    return new FileJournal(dir, fileKey).path;
  }

  #ensureFd(): number {
    if (this.#fd === null) {
      mkdirSync(this.dir, { recursive: true });
      this.#fd = openSync(this.path, "a");
    }
    return this.#fd;
  }

  /** Creates the journal with its base version when missing. */
  begin(baseItemVersion: string | null): void {
    if (this.#begun) return;
    this.#begun = true;
    if (existsSync(this.path) && readFileSync(this.path).length > 0) return;
    const fd = this.#ensureFd();
    writeSync(fd, MAGIC);
    writeSync(fd, record("B", Buffer.from(baseItemVersion ?? "", "utf8")));
  }

  /** Appends one accepted update. */
  appendUpdate(update: Uint8Array): void {
    const fd = this.#ensureFd();
    if (!this.#begun) this.begin(null);
    writeSync(fd, record("U", Buffer.from(update)));
  }

  /** Appends a new base version record, keeping every update. */
  setBaseVersion(baseItemVersion: string | null): void {
    const fd = this.#ensureFd();
    if (!this.#begun) this.begin(null);
    writeSync(fd, record("B", Buffer.from(baseItemVersion ?? "", "utf8")));
  }

  /** Starts a fresh segment at a new base, keeping only updates after the first dropCount. */
  startSegment(baseItemVersion: string | null, dropCount: number): void {
    const existing = this.read();
    const kept = existing ? existing.updates.slice(Math.max(0, dropCount)) : [];
    this.close();
    mkdirSync(this.dir, { recursive: true });
    const parts: Buffer[] = [MAGIC, record("B", Buffer.from(baseItemVersion ?? "", "utf8"))];
    for (const update of kept) parts.push(record("U", Buffer.from(update)));
    writeFileSync(this.path, Buffer.concat(parts));
    this.#begun = true;
    this.#fd = openSync(this.path, "a");
  }

  /** Flushes buffered journal bytes to disk. */
  flush(): void {
    if (this.#fd !== null) fsyncSync(this.#fd);
  }

  /** Reads the base version and every complete update. */
  read(): JournalData | null {
    if (!existsSync(this.path)) return null;
    const data = readFileSync(this.path);
    if (data.length === 0) return { baseItemVersion: null, updates: [] };
    if (!data.subarray(0, MAGIC.length).equals(MAGIC)) return { baseItemVersion: null, updates: [] };
    let offset = MAGIC.length;
    let baseItemVersion: string | null = null;
    const updates: Uint8Array[] = [];
    while (offset + 4 <= data.length) {
      const length = data.readUInt32BE(offset);
      offset += 4;
      if (length === 0 || offset + length > data.length) break;
      const body = data.subarray(offset, offset + length);
      offset += length;
      const prefix = String.fromCharCode(body[0] ?? 0);
      if (prefix === "B") baseItemVersion = body.subarray(1).toString("utf8") || null;
      else if (prefix === "U") updates.push(new Uint8Array(body.subarray(1)));
    }
    return { baseItemVersion, updates };
  }

  /** Deletes the journal file. */
  delete(): void {
    this.close();
    try {
      unlinkSync(this.path);
    } catch {}
    this.#begun = false;
  }

  /** Closes the underlying file handle. */
  close(): void {
    if (this.#fd !== null) {
      try {
        closeSync(this.#fd);
      } catch {}
      this.#fd = null;
    }
  }
}
