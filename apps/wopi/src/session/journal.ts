import { createHash } from "node:crypto";
import { closeSync, existsSync, fsyncSync, mkdirSync, openSync, readFileSync, unlinkSync, writeSync } from "node:fs";
import { join } from "node:path";

/** Journal contents read back from disk. The journal holds the full source at rest. */
export interface JournalData {
  fileKey: string | null;
  itemVersion: string | null;
  sha256: string | null;
  clientId: number | null;
  initialState: Uint8Array | null;
  source: Uint8Array | null;
  updates: Uint8Array[];
  writtenVersions: string[];
}

const MAGIC = Buffer.from("BOJ2\n", "utf8");

function record(prefix: string, payload: Buffer): Buffer {
  const body = Buffer.concat([Buffer.from(prefix, "utf8"), payload]);
  const header = Buffer.alloc(4);
  header.writeUInt32BE(body.length, 0);
  return Buffer.concat([header, body]);
}

function sha256Hex(bytes: Uint8Array): string {
  return createHash("sha256").update(bytes).digest("hex");
}

/** Append only journal for one file key. It holds the full source at rest. */
export class FileJournal {
  readonly path: string;
  readonly dir: string;
  readonly fileKey: string;
  #fd: number | null = null;
  #begun = false;

  constructor(dir: string, fileKey: string) {
    const name = `${createHash("sha256").update(fileKey).digest("hex")}.journal`;
    this.path = join(dir, name);
    this.dir = dir;
    this.fileKey = fileKey;
  }

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

  /** Creates the journal with its source header when missing. */
  begin(source: Uint8Array, itemVersion: string | null, clientId?: number | null, initialState?: Uint8Array | null): void {
    if (this.#begun) return;
    this.#begun = true;
    if (existsSync(this.path) && readFileSync(this.path).length > 0) return;
    const fd = this.#ensureFd();
    const header = JSON.stringify({ fileKey: this.fileKey, itemVersion, sha256: sha256Hex(source), clientId: clientId ?? null });
    writeSync(fd, MAGIC);
    writeSync(fd, record("H", Buffer.from(header, "utf8")));
    writeSync(fd, record("S", Buffer.from(source)));
    if (initialState) writeSync(fd, record("I", Buffer.from(initialState)));
  }

  /** Appends one accepted or replica local update. */
  appendUpdate(update: Uint8Array): void {
    const fd = this.#ensureFd();
    writeSync(fd, record("U", Buffer.from(update)));
  }

  /** Records an itemVersion this session wrote with PutFile. */
  recordWrittenVersion(itemVersion: string | null): void {
    if (itemVersion === null) return;
    const fd = this.#ensureFd();
    writeSync(fd, record("V", Buffer.from(itemVersion, "utf8")));
  }

  /** Flushes buffered journal bytes to disk. */
  flush(): void {
    if (this.#fd !== null) fsyncSync(this.#fd);
  }

  /** Reads the header, source, updates and written versions. */
  read(): JournalData | null {
    if (!existsSync(this.path)) return null;
    const data = readFileSync(this.path);
    const empty: JournalData = { fileKey: null, itemVersion: null, sha256: null, clientId: null, initialState: null, source: null, updates: [], writtenVersions: [] };
    if (data.length === 0) return empty;
    if (!data.subarray(0, MAGIC.length).equals(MAGIC)) return empty;
    let offset = MAGIC.length;
    const out: JournalData = { fileKey: null, itemVersion: null, sha256: null, clientId: null, initialState: null, source: null, updates: [], writtenVersions: [] };
    while (offset + 4 <= data.length) {
      const length = data.readUInt32BE(offset);
      offset += 4;
      if (length === 0 || offset + length > data.length) break;
      const body = data.subarray(offset, offset + length);
      offset += length;
      const prefix = String.fromCharCode(body[0] ?? 0);
      if (prefix === "H") {
        try {
          const header = JSON.parse(body.subarray(1).toString("utf8")) as Record<string, unknown>;
          out.fileKey = typeof header.fileKey === "string" ? header.fileKey : null;
          out.itemVersion = typeof header.itemVersion === "string" ? header.itemVersion : null;
          out.sha256 = typeof header.sha256 === "string" ? header.sha256 : null;
          out.clientId = typeof header.clientId === "number" ? header.clientId : null;
        } catch {}
      } else if (prefix === "I") {
        out.initialState = new Uint8Array(body.subarray(1));
      } else if (prefix === "S") {
        out.source = new Uint8Array(body.subarray(1));
      } else if (prefix === "U") {
        out.updates.push(new Uint8Array(body.subarray(1)));
      } else if (prefix === "V") {
        const version = body.subarray(1).toString("utf8");
        if (version) out.writtenVersions.push(version);
      }
    }
    if (out.source !== null && out.sha256 !== null && sha256Hex(out.source) !== out.sha256) {
      out.source = null;
    }
    return out;
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
