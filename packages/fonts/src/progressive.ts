import type { FontDecoder } from './decode';
import { chunkEntries, parseSkeleton, type SliceSkeleton } from './slices';
import { sha256 } from './sha256';

/**
 * A sliced face that loads its outlines as they are needed. Its layout tables
 * are complete from the start, so shaping, metrics and line breaks equal the
 * whole face's; a glyph whose outline chunk is not loaded yet draws nothing.
 * @public
 */
export interface ProgressiveFontFace {
  /** Increases each time outlines are added. */
  readonly revision: number;
  /** The face with every outline loaded so far, in a new buffer on each call. */
  bytes(): ArrayBuffer;
  /** Loads the outlines of these glyph ids; resolves with a revision that includes them. */
  ensureGlyphs(glyphIds: Iterable<number>): Promise<number>;
  /** Starts loading the outlines of these characters' glyphs; `ensureGlyphs` is authoritative. */
  prefetch(text: string): void;
  /** Calls `listener` with each new revision; returns an unsubscribe. */
  onUpdate(listener: (revision: number) => void): () => void;
  /** Whether every outline is loaded. */
  complete(): boolean;
}

/**
 * Where a sliced face's files are and how to decode them: what a font loader
 * carries as `slices`, which `openSlicedFace` from `@betteroffice/fonts/progressive` opens.
 * @public
 */
export interface SliceLocation {
  ref: SliceRef;
  /** The asset base the files resolve against; part of the shared face's identity. */
  base: string;
  /** The URL of a pack file by its manifest path. */
  url(file: string): Promise<URL>;
  timeoutMs?: number;
  /** Decompresses a zstd file to at most `maxOutput` bytes. */
  decode(encoded: ArrayBuffer, maxOutput: number): Promise<ArrayBuffer>;
  /** The host's decoder behind `decode`, when it has one; part of the shared face's identity. */
  decoder?: FontDecoder;
  /** @internal Drop the opened face from the shared cache once it settles, as `retain: false` does for whole faces. */
  retain?: boolean;
}

/** @internal A face row's pointer to its slice index. */
export interface SliceRef {
  format: number;
  index: string;
  /** Pins the index, which pins every file it lists by length and SHA-256. */
  indexSha256: string;
}

interface SliceFile {
  id?: number;
  file: string;
  encodedLength: number;
  byteLength: number;
  sha256: string;
}

interface SliceIndex {
  format: number;
  skeleton: SliceFile;
  core: SliceFile;
  chunks: SliceFile[];
  byteLength: number;
  sha256: string;
}

/** @internal */
export interface SliceTransport {
  /** The bytes of a pack file by its manifest path. */
  read(file: string): Promise<ArrayBuffer>;
  /** Decompresses a zstd file to at most `maxOutput` bytes. */
  decode(encoded: ArrayBuffer, maxOutput: number): Promise<ArrayBuffer>;
}

async function readFile(transport: SliceTransport, entry: SliceFile): Promise<Uint8Array> {
  const encoded = await transport.read(entry.file);
  if (encoded.byteLength !== entry.encodedLength) {
    throw new Error(`Font slice ${entry.file} has ${encoded.byteLength} encoded bytes; expected ${entry.encodedLength}`);
  }
  const decoded = await transport.decode(encoded, entry.byteLength);
  if (decoded.byteLength !== entry.byteLength) {
    throw new Error(`Font slice ${entry.file} has ${decoded.byteLength} bytes; expected ${entry.byteLength}`);
  }
  if (await sha256(decoded) !== entry.sha256) throw new Error(`Font slice ${entry.file} SHA-256 mismatch`);
  return new Uint8Array(decoded);
}

/** @internal Loads a sliced face's index, skeleton and core; chunks follow on demand. */
export async function openProgressiveFace(ref: SliceRef, transport: SliceTransport): Promise<ProgressiveFontFace> {
  if (ref.format !== 1) throw new Error(`Unsupported font slice format ${ref.format}`);
  const indexBytes = await transport.read(ref.index);
  if (await sha256(indexBytes) !== ref.indexSha256) throw new Error(`Font slice index ${ref.index} SHA-256 mismatch`);
  const index = JSON.parse(new TextDecoder().decode(indexBytes)) as SliceIndex;
  if (index.format !== 1) throw new Error(`Unsupported font slice index format in ${ref.index}`);
  const [skeletonBytes, core] = await Promise.all([readFile(transport, index.skeleton), readFile(transport, index.core)]);
  const skeleton: SliceSkeleton = parseSkeleton(skeletonBytes);
  const files = new Map<number, SliceFile>(index.chunks.map((chunk) => [chunk.id!, chunk]));
  const outlines = new Map<number, Uint8Array>();
  const loaded = new Set<number>([0]);
  const pending = new Map<number, Promise<void>>();
  const listeners = new Set<(revision: number) => void>();
  let revision = 1;
  let dirty = false;
  const install = (bytes: Uint8Array) => {
    for (const [glyph, outline] of chunkEntries(bytes)) outlines.set(glyph, outline);
  };
  install(core);

  const load = (chunk: number): Promise<void> => {
    if (loaded.has(chunk)) return Promise.resolve();
    let promise = pending.get(chunk);
    if (promise === undefined) {
      const entry = files.get(chunk);
      if (entry === undefined) return Promise.reject(new Error(`Unknown font slice chunk ${chunk} in ${ref.index}`));
      promise = readFile(transport, entry).then((bytes) => {
        install(bytes);
        loaded.add(chunk);
        pending.delete(chunk);
        dirty = true;
      }, (error: unknown) => {
        pending.delete(chunk);
        throw error;
      });
      pending.set(chunk, promise);
    }
    return promise;
  };

  const publish = (): number => {
    if (!dirty) return revision;
    dirty = false;
    revision += 1;
    for (const listener of [...listeners]) {
      try {
        listener(revision);
      } catch {}
    }
    return revision;
  };

  const ensure = async (chunks: Iterable<number>): Promise<number> => {
    const wanted = [...new Set(chunks)].filter((chunk) => !loaded.has(chunk));
    if (wanted.length === 0) return publish();
    const results = await Promise.allSettled(wanted.map(load));
    const failure = results.find((result) => result.status === 'rejected');
    const next = publish();
    if (failure) throw (failure as PromiseRejectedResult).reason;
    return next;
  };

  return {
    get revision() {
      return revision;
    },
    bytes() {
      return skeleton.reassemble(outlines);
    },
    ensureGlyphs(glyphIds) {
      const chunks: number[] = [];
      for (const glyph of glyphIds) {
        const chunk = skeleton.chunkOfGlyph(glyph);
        if (chunk !== undefined) chunks.push(chunk);
      }
      return ensure(chunks);
    },
    prefetch(text) {
      const chunks: number[] = [];
      for (const character of text) {
        const chunk = skeleton.chunkOfCodepoint(character.codePointAt(0)!);
        if (chunk !== undefined) chunks.push(chunk);
      }
      void ensure(chunks).catch(() => {});
    },
    onUpdate(listener) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    complete() {
      return loaded.size === files.size + 1;
    },
  };
}
