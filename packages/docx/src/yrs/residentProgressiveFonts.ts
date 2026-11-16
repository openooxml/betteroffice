import { isEngineTrap } from '../../../../shared/office-session/trap';

/** A sliced face's pack files, as `sliceSourceOf` from `@betteroffice/fonts/progressive` describes them. */
export interface ProgressiveFontSource {
  ref: { format: number; index: string; indexSha256: string };
  /** Absolute URL of the index, skeleton, core and every chunk by manifest path. */
  urls: Record<string, string>;
  timeoutMs?: number;
}

/** A command font slot registered from a sliced face's skeleton and core. */
export interface ProgressiveFontSpec {
  slot: number;
  source: ProgressiveFontSource;
  /** Load every outline of the face once the editor is idle, so any character in it draws without waiting. */
  fill?: boolean;
}

/** What the worker uses of a face `@betteroffice/fonts/progressive` opens. */
interface OutlineFace {
  readonly revision: number;
  bytes(): ArrayBuffer;
  ensureGlyphs(glyphIds: Iterable<number>): Promise<number>;
  complete(): boolean;
  /** Absent from a `@betteroffice/fonts` release older than filling; such a face only loads as frames need it. */
  chunks?(): readonly number[];
  loadedChunks?(): ReadonlySet<number>;
  ensureChunks?(chunks: Iterable<number>): Promise<number>;
}

/**
 * Opened faces by identity, kept by the worker across font environments: a face a fill loaded, or one frames
 * loaded, is not fetched and decoded again when the host sends the fonts anew.
 */
export type SharedOutlineFaces = Map<string, Promise<OutlineFace>>;

/** How long a frame waits for missing outlines before it paints without them and repaints later. */
export const PROGRESSIVE_OUTLINE_WAIT_MS = 10_000;

/** Pauses before each retry of a chunk that failed to load. */
const CHUNK_RETRY_DELAYS_MS = [250, 1_000, 4_000];

/** A fill starts this long after the fonts land, so opening and the first frames go first. */
export const FILL_DELAY_MS = 1_000;
/** A fill loads its next chunks only once no frame has drawn for this long. */
export const FILL_QUIET_MS = 300;
/** Chunks a fill loads in one step, between which frames go first. */
const FILL_BATCH = 8;

interface Owner {
  /** The font id of every slot registered with these bytes; the engine may share one face among them. */
  ids: number[];
  open: () => Promise<OutlineFace>;
  face?: OutlineFace;
  installed: number;
  complete: boolean;
  refused: boolean;
  fill: boolean;
}

interface OutlineEngine {
  updateFontOutlines(fontId: number, bytes: Uint8Array): void;
}

/** Glyph ids per font id drawn by a display page: its own runs and those of its regions and note areas. */
export function glyphsByFont(page: unknown, into = new Map<number, Set<number>>()): Map<number, Set<number>> {
  const visit = (node: unknown, depth: number): void => {
    if (depth > 12 || node === null || typeof node !== 'object') return;
    if (Array.isArray(node)) {
      for (const child of node) visit(child, depth + 1);
      return;
    }
    const value = node as { kind?: unknown; fontId?: unknown; glyphs?: unknown };
    if (value.kind === 'glyphRun' && typeof value.fontId === 'number' && Array.isArray(value.glyphs)) {
      let glyphs = into.get(value.fontId);
      if (!glyphs) into.set(value.fontId, glyphs = new Set());
      for (const glyph of value.glyphs as Array<{ id?: unknown }>) {
        if (typeof glyph.id === 'number') glyphs.add(glyph.id);
      }
      return;
    }
    for (const child of Object.values(node)) visit(child, depth + 1);
  };
  visit(page, 0);
  return into;
}

/** What the worker uses of `@betteroffice/fonts/progressive`. */
interface ProgressiveModule {
  openProgressiveFace(ref: ProgressiveFontSource['ref'], transport: {
    read(file: string): Promise<ArrayBuffer>;
    decode(encoded: ArrayBuffer, maxOutput: number): Promise<ArrayBuffer>;
  }): Promise<OutlineFace>;
  readFontAsset(file: string, url: URL, timeoutMs?: number): Promise<ArrayBuffer>;
}

/** The progressive font module could not be loaded in this worker, as in a host that serves the packages without a bundler. */
class ProgressiveModuleUnavailable extends Error {}

/**
 * Sliced CJK faces registered from their skeleton and core. Before a page is
 * rasterized, the outlines of the glyphs it draws are fetched (each chunk
 * once, shared across frames) and installed in place; layout never changes.
 */
export class ResidentProgressiveFonts {
  private readonly owners = new Map<number, Owner>();
  private readonly ownerOf = new Map<number, number>();
  private readonly pending = new Set<Promise<unknown>>();
  private readonly lateListeners = new Set<() => void>();
  private readonly refusalListeners = new Set<(owner: number) => void>();
  private readonly fillListeners = new Set<(owners: ReadonlySet<number>) => void>();
  private warned = false;
  private late = 0;
  private disposed = false;
  /** When a frame last asked for outlines; a fill waits for a quiet stretch after it. */
  private lastFrame = -Infinity;
  private fillTimer: ReturnType<typeof setTimeout> | undefined;
  private filling: Promise<void> | undefined;

  /** Outline installs that arrived after the frame that needed them had painted. */
  get lateInstalls(): number {
    return this.late;
  }

  constructor(
    specs: readonly ProgressiveFontSpec[],
    ids: readonly (number | null)[],
    substituteOf: readonly (number | undefined)[],
    private readonly engine: OutlineEngine,
    decode: (encoded: ArrayBuffer, maxOutput: number) => Promise<ArrayBuffer>,
    private readonly log: (message: string) => void = (message) => console.warn(message),
    loadModule: () => Promise<ProgressiveModule> = () => import('@betteroffice/fonts/progressive'),
    shared: SharedOutlineFaces = new Map(),
  ) {
    const rootSlot = (slot: number): number => {
      let current = slot;
      for (let hops = 0; substituteOf[current] !== undefined && hops < substituteOf.length; hops += 1) {
        current = substituteOf[current]!;
      }
      return current;
    };
    const byFace = new Map<string, number>();
    for (const spec of specs) {
      const id = ids[spec.slot];
      if (id === null || id === undefined) continue;
      const face = `${spec.source.ref.indexSha256}|${spec.source.ref.index}`;
      const known = byFace.get(face);
      const owner = known ?? id;
      if (known === undefined) byFace.set(face, id);
      else {
        const state = this.owners.get(known)!;
        if (!state.ids.includes(id)) state.ids.push(id);
        state.fill ||= spec.fill === true;
      }
      for (let slot = 0; slot < ids.length; slot += 1) {
        const view = ids[slot];
        if (view !== null && view !== undefined && rootSlot(slot) === spec.slot) this.ownerOf.set(view, owner);
      }
      if (known !== undefined) continue;
      let opening: Promise<OutlineFace> | undefined;
      const open = (): Promise<OutlineFace> => {
        if (opening) return opening;
        let created = shared.get(face);
        if (created === undefined) {
          created = loadModule().then(undefined, (error: unknown) => {
            throw new ProgressiveModuleUnavailable(error instanceof Error ? error.message : String(error));
          }).then(({ openProgressiveFace, readFontAsset }) =>
            openProgressiveFace(spec.source.ref, {
              read: (file) => {
                const url = spec.source.urls[file];
                if (url === undefined) return Promise.reject(new Error(`Font slice ${file} has no URL`));
                return readFontAsset(file, new URL(url), spec.source.timeoutMs);
              },
              decode,
            }));
          shared.set(face, created);
        }
        const attempt = created;
        opening = attempt.catch((error: unknown) => {
          opening = undefined;
          if (shared.get(face) === attempt) shared.delete(face);
          throw error;
        });
        return opening;
      };
      this.owners.set(owner, { ids: [id], open, installed: 1, complete: false, refused: false, fill: spec.fill === true });
    }
  }

  /** Stops installing outlines, for fonts that were cleared or replaced; pending loads finish unused. */
  dispose(): void {
    this.disposed = true;
    clearTimeout(this.fillTimer);
    this.fillTimer = undefined;
    this.lateListeners.clear();
    this.refusalListeners.clear();
    this.fillListeners.clear();
  }

  get size(): number {
    return this.owners.size;
  }

  /** Every font id drawing with the bytes of `owners`: the owners and their substitute views. */
  fontIdsOf(owners: ReadonlySet<number>): Set<number> {
    const ids = new Set<number>();
    for (const [id, owner] of this.ownerOf) if (owners.has(owner)) ids.add(id);
    return ids;
  }

  /** Calls `listener` when outlines a frame stopped waiting for are installed. */
  onLateInstall(listener: () => void): () => void {
    this.lateListeners.add(listener);
    return () => {
      this.lateListeners.delete(listener);
    };
  }

  /**
   * Installs the outlines of `glyphs` (per font id) before a raster. Resolves
   * with the owner font ids whose outlines changed once they are installed,
   * or after `waitMs` with those installed so far; the rest install later.
   */
  async ensure(glyphs: Map<number, Set<number>>, waitMs = PROGRESSIVE_OUTLINE_WAIT_MS): Promise<Set<number>> {
    this.lastFrame = performance.now();
    const wanted = new Map<number, Set<number>>();
    for (const [fontId, ids] of glyphs) {
      const owner = this.ownerOf.get(fontId);
      if (owner === undefined || ids.size === 0) continue;
      const set = wanted.get(owner) ?? new Set<number>();
      for (const id of ids) set.add(id);
      wanted.set(owner, set);
    }
    if (wanted.size === 0 || this.disposed) return new Set();
    const changed = new Set<number>();
    let waiting = true;
    const work = Promise.all([...wanted].map(async ([owner, ids]) => {
      const before = this.owners.get(owner)!.installed;
      await this.load(owner, ids);
      // A fill may have installed these outlines meanwhile; they changed for this frame all the same.
      if (this.install(owner) || this.owners.get(owner)!.installed > before) {
        changed.add(owner);
        if (!waiting) {
          this.late += 1;
          for (const listener of [...this.lateListeners]) listener();
        }
      }
    }));
    this.track(work);
    let timer: ReturnType<typeof setTimeout> | undefined;
    const timedOut = await Promise.race([
      work.then(() => false),
      new Promise<boolean>((resolve) => {
        timer = setTimeout(() => resolve(true), waitMs);
      }),
    ]).finally(() => clearTimeout(timer));
    waiting = false;
    if (timedOut && !this.warned) {
      this.warned = true;
      this.log(`[docx] font outlines still loading after ${waitMs} ms; painting with the outlines installed and repainting when the rest arrive`);
    }
    return new Set(changed);
  }

  /** Settles once every outline requested so far is installed or has failed. */
  async whenInstalled(): Promise<void> {
    while (this.pending.size) await Promise.allSettled([...this.pending]);
  }

  /** Calls `listener` with the owner font ids a fill installed; glyph images cached from their earlier bytes go stale. */
  onFilled(listener: (owners: ReadonlySet<number>) => void): () => void {
    this.fillListeners.add(listener);
    return () => {
      this.fillListeners.delete(listener);
    };
  }

  /**
   * Loads every outline of the faces marked `fill`, a few chunks at a time while no frame draws, and installs each
   * face once it is whole, so later text in it never waits for a chunk. Starts `FILL_DELAY_MS` after the first call;
   * later calls while one is scheduled or running do nothing. A face that cannot fill keeps loading as frames need it.
   */
  scheduleFill(): void {
    if (this.disposed || this.fillTimer !== undefined || this.filling) return;
    if (![...this.owners.values()].some((state) => state.fill && !state.complete && !state.refused)) return;
    this.fillTimer = setTimeout(() => {
      this.fillTimer = undefined;
      const filling: Promise<void> = this.fill().catch((error: unknown) => {
        this.log(`[docx] filling sliced faces stopped: ${error instanceof Error ? error.message : String(error)}`);
      }).finally(() => {
        if (this.filling === filling) this.filling = undefined;
      });
      this.filling = filling;
    }, FILL_DELAY_MS);
  }

  /** Settles once a scheduled fill has run, or at once without one. */
  async whenFilled(): Promise<void> {
    while (!this.disposed && (this.fillTimer !== undefined || this.filling)) {
      await (this.filling ?? new Promise((resolve) => setTimeout(resolve, 10)));
    }
  }

  private async fill(): Promise<void> {
    for (const [owner, state] of this.owners) {
      if (!state.fill || state.complete || state.refused) continue;
      try {
        await this.quiet();
        if (this.disposed || state.refused) return;
        const face = state.face ??= await state.open();
        const { chunks, loadedChunks, ensureChunks } = face;
        if (!chunks || !loadedChunks || !ensureChunks) continue;
        const missing = chunks.call(face).filter((chunk) => !loadedChunks.call(face).has(chunk));
        for (let start = 0; start < missing.length; start += FILL_BATCH) {
          await this.quiet();
          if (this.disposed || state.refused) return;
          await ensureChunks.call(face, missing.slice(start, start + FILL_BATCH));
        }
        await this.quiet();
        if (this.disposed || state.refused) return;
        if (this.install(owner)) {
          const owners = new Set([owner]);
          for (const listener of [...this.fillListeners]) listener(owners);
        }
      } catch (error) {
        if (isEngineTrap(error)) throw error;
        if (this.disposed) return;
        if (error instanceof ProgressiveModuleUnavailable) {
          this.refuse(owner, `the progressive font module is unavailable here: ${error.message}`);
          continue;
        }
        this.log(`[docx] a sliced face could not load whole in the background; its outlines keep loading as frames need them: ${error instanceof Error ? error.message : String(error)}`);
      }
    }
  }

  /** Resolves once no frame has asked for outlines for `FILL_QUIET_MS` and no frame's outlines are loading. */
  private async quiet(): Promise<void> {
    for (;;) {
      if (this.disposed) return;
      const idle = performance.now() - this.lastFrame;
      if (idle >= FILL_QUIET_MS && this.pending.size === 0) return;
      await new Promise((resolve) => setTimeout(resolve, Math.max(FILL_QUIET_MS - idle, 20)));
    }
  }

  private track(work: Promise<unknown>): void {
    const settled = work.then(() => {}, () => {});
    this.pending.add(settled);
    void settled.finally(() => this.pending.delete(settled));
  }

  /**
   * Loads the chunks of `ids`, retrying a chunk that fails. Chunks that keep
   * failing leave their glyphs blank; the next frame drawing them asks again.
   */
  private async load(owner: number, ids: Set<number>): Promise<void> {
    const state = this.owners.get(owner)!;
    if (state.complete || this.disposed) return;
    for (let attempt = 0; ; attempt += 1) {
      try {
        state.face ??= await state.open();
        await state.face.ensureGlyphs(ids);
        return;
      } catch (error) {
        if (error instanceof ProgressiveModuleUnavailable) {
          this.refuse(owner, `the progressive font module is unavailable here: ${error.message}`);
          return;
        }
        const pause = CHUNK_RETRY_DELAYS_MS[attempt];
        if (pause === undefined) {
          this.log(`[docx] font outline chunks failed to load after ${attempt + 1} attempts; their glyphs stay blank until a frame needs them again: ${error instanceof Error ? error.message : String(error)}`);
          return;
        }
        await new Promise((resolve) => setTimeout(resolve, pause));
      }
    }
  }

  /** Installs the face's current outlines; on refusal stops progressive loading for it. True when outlines changed. */
  private install(owner: number): boolean {
    const state = this.owners.get(owner)!;
    const face = state.face;
    if (!face || this.disposed || state.refused || face.revision <= state.installed) return false;
    try {
      const bytes = new Uint8Array(face.bytes());
      for (const id of state.ids) this.engine.updateFontOutlines(id, bytes);
      state.installed = face.revision;
      state.complete = face.complete();
      return true;
    } catch (error) {
      if (isEngineTrap(error)) throw error;
      this.refuse(owner, `the engine refused the outlines of a sliced face: ${error instanceof Error ? error.message : String(error)}`);
      return false;
    }
  }

  /** Stops progressive loading for `owner` and asks the host to reload it whole. */
  private refuse(owner: number, reason: string): void {
    const state = this.owners.get(owner)!;
    if (state.refused) return;
    state.complete = true;
    state.refused = true;
    this.log(`[docx] ${reason}`);
    for (const listener of [...this.refusalListeners]) listener(owner);
  }

  /** Calls `listener` with a font id that cannot load progressively here (the engine refused its outlines or the module is unavailable); the host reloads it whole. */
  onRefused(listener: (owner: number) => void): () => void {
    this.refusalListeners.add(listener);
    return () => {
      this.refusalListeners.delete(listener);
    };
  }
}
