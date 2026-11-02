import {
  openProgressiveFace,
  readFontAsset,
  type ProgressiveFontFace,
  type SliceSource,
} from '@betteroffice/fonts/progressive';

/** A command font slot registered from a sliced face's skeleton and core. */
export interface ProgressiveFontSpec {
  slot: number;
  source: SliceSource;
}

/** How long a frame waits for missing outlines before it paints without them and repaints later. */
export const PROGRESSIVE_OUTLINE_WAIT_MS = 10_000;

/** Pauses before each retry of a chunk that failed to load. */
const CHUNK_RETRY_DELAYS_MS = [250, 1_000, 4_000];

interface Owner {
  /** The font id of every slot registered with these bytes; the engine may share one face among them. */
  ids: number[];
  open: () => Promise<ProgressiveFontFace>;
  face?: ProgressiveFontFace;
  installed: number;
  complete: boolean;
  refused: boolean;
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
  private warned = false;
  private late = 0;
  private disposed = false;

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
      const face = `${spec.source.ref.sha256}|${spec.source.ref.index}`;
      const shared = byFace.get(face);
      const owner = shared ?? id;
      if (shared === undefined) byFace.set(face, id);
      else if (!this.owners.get(shared)!.ids.includes(id)) this.owners.get(shared)!.ids.push(id);
      for (let slot = 0; slot < ids.length; slot += 1) {
        const view = ids[slot];
        if (view !== null && view !== undefined && rootSlot(slot) === spec.slot) this.ownerOf.set(view, owner);
      }
      if (shared !== undefined) continue;
      let opening: Promise<ProgressiveFontFace> | undefined;
      const open = () => opening ??= openProgressiveFace(spec.source.ref, {
        read: (file) => {
          const url = spec.source.urls[file];
          if (url === undefined) return Promise.reject(new Error(`Font slice ${file} has no URL`));
          return readFontAsset(file, new URL(url), spec.source.timeoutMs);
        },
        decode,
      }).catch((error: unknown) => {
        opening = undefined;
        throw error;
      });
      this.owners.set(owner, { ids: [id], open, installed: 1, complete: false, refused: false });
    }
  }

  /** Stops installing outlines, for fonts that were cleared or replaced; pending loads finish unused. */
  dispose(): void {
    this.disposed = true;
    this.lateListeners.clear();
    this.refusalListeners.clear();
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
      await this.load(owner, ids);
      if (this.install(owner)) {
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
      if (error instanceof WebAssembly.RuntimeError) throw error;
      state.complete = true;
      state.refused = true;
      this.log(`[docx] the engine refused the outlines of a sliced face: ${error instanceof Error ? error.message : String(error)}`);
      for (const listener of [...this.refusalListeners]) listener(owner);
      return false;
    }
  }

  /** Calls `listener` with a font id whose sliced outlines the engine refused; the host reloads it whole. */
  onRefused(listener: (owner: number) => void): () => void {
    this.refusalListeners.add(listener);
    return () => {
      this.refusalListeners.delete(listener);
    };
  }
}
