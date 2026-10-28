import { createHash } from "node:crypto";
import { closeSync, existsSync, fstatSync, fsyncSync, ftruncateSync, mkdirSync, openSync, readFileSync, unlinkSync, writeSync } from "node:fs";
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
type JournalWrite = (fd: number, buffer: Uint8Array, offset: number, length: number) => number;

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
  #broken = false;
  #write: JournalWrite;

  constructor(dir: string, fileKey: string, options: { write?: JournalWrite } = {}) {
    const name = `${createHash("sha256").update(fileKey).digest("hex")}.journal`;
    this.path = join(dir, name);
    this.dir = dir;
    this.fileKey = fileKey;
    this.#write = options.write ?? ((fd, buffer, offset, length) => writeSync(fd, buffer, offset, length));
  }

  /** Path for a file key inside a journal directory. */
  static pathFor(dir: string, fileKey: string): string {
    return new FileJournal(dir, fileKey).path;
  }

  #ensureFd(): number {
    if (this.#broken) throw new Error("journal is broken");
    if (this.#fd === null) {
      mkdirSync(this.dir, { recursive: true });
      this.#fd = openSync(this.path, "a");
      try {
        const data = readFileSync(this.path);
        if (data.length > 0) {
          if (!data.subarray(0, MAGIC.length).equals(MAGIC)) throw new Error("invalid journal header");
          let end = MAGIC.length;
          while (end + 4 <= data.length) {
            const length = data.readUInt32BE(end);
            if (length === 0 || end + 4 + length > data.length) break;
            end += 4 + length;
          }
          if (end < data.length) ftruncateSync(this.#fd, end);
        }
      } catch (err) {
        this.#broken = true;
        this.close();
        throw err;
      }
    }
    return this.#fd;
  }

  #writeFully(data: Uint8Array): void {
    const fd = this.#ensureFd();
    const start = fstatSync(fd).size;
    try {
      let offset = 0;
      while (offset < data.length) {
        const written = this.#write(fd, data, offset, data.length - offset);
        if (!Number.isInteger(written) || written <= 0 || written > data.length - offset) throw new Error("incomplete journal write");
        offset += written;
      }
    } catch (err) {
      this.#broken = true;
      try {
        ftruncateSync(fd, start);
      } catch {}
      throw err;
    }
  }

  /** Creates the journal with its source header when missing. */
  begin(source: Uint8Array, itemVersion: string | null, clientId?: number | null, initialState?: Uint8Array | null): void {
    if (this.#begun) return;
    this.#ensureFd();
    if (readFileSync(this.path).length > 0) {
      if (this.read()?.source === null) throw new Error("incomplete journal source");
      this.#begun = true;
      return;
    }
    const header = JSON.stringify({ fileKey: this.fileKey, itemVersion, sha256: sha256Hex(source), clientId: clientId ?? null });
    const records = [MAGIC, record("H", Buffer.from(header, "utf8")), record("S", Buffer.from(source))];
    if (initialState) records.push(record("I", Buffer.from(initialState)));
    this.#writeFully(Buffer.concat(records));
    this.#begun = true;
  }

  /** Appends one accepted or replica local update. */
  appendUpdate(update: Uint8Array): void {
    this.#writeFully(record("U", Buffer.from(update)));
  }

  /** Records an itemVersion this session wrote with PutFile. */
  recordWrittenVersion(itemVersion: string | null): void {
    if (itemVersion === null) return;
    this.#writeFully(record("V", Buffer.from(itemVersion, "utf8")));
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
    this.#broken = false;
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
