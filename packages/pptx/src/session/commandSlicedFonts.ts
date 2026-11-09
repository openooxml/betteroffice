import { readFontTraits } from '../../../../shared/fontTraits';
import type { PaintFontCoverage } from '../render/canvas';

/** What the session needs from a progressive face. */
export interface SlicedFontFace {
  readonly revision: number;
  bytes(): ArrayBuffer;
  ensureText(text: string): Promise<number>;
  prefetch(text: string): void;
  chunksOf(text: string): number[];
  loadedChunks(): ReadonlySet<number>;
}

interface ExportFaceHandle {
  replaceExportFace(family: string, bytes: Uint8Array, keep?: (face: { id: number; bold?: boolean; italic?: boolean }) => boolean): void;
}

/** A face's drawable bytes for a realm that keeps its own copy; `bytes` is absent when `have` is still current. */
export interface SlicedFontRead {
  revision: number;
  bytes?: ArrayBuffer;
}

const WHOLE = Number.MAX_SAFE_INTEGER;
const KEPT_REVISIONS = 2;
const FILL_BATCH = 48;
const FILL_QUIET_MS = 300;
const FILL_DELAY_MS = 1000;

/**
 * Sliced faces are verified in Chromium and Firefox. Safari, iOS browsers (all WebKit) and unknown engines
 * use whole faces.
 */
export function slicedFacesSupported(
  userAgent = (globalThis as { navigator?: { userAgent?: string } }).navigator?.userAgent ?? '',
): boolean {
  return /(?:Chrome|Chromium|Firefox)\//.test(userAgent) && !/\b(?:iPhone|iPad|iPod|CriOS|FxiOS)\b/.test(userAgent);
}

interface Entry {
  family: string;
  face: SlicedFontFace;
  whole(): Promise<ArrayBuffer>;
  wholeBytes?: ArrayBuffer;
  wholeLoad?: Promise<ArrayBuffer>;
  /** `chunks` is what the revision's bytes hold; a whole face holds everything. */
  installed: { revision: number; family: string; face: FontFace; chunks: ReadonlySet<number> | null }[];
  installing?: { revision: number; done: Promise<string> };
}

const keyOf = (family: string) => family.trim().toLowerCase();
const styleOf = (bold: boolean, italic: boolean) => `${bold ? 1 : 0}|${italic ? 1 : 0}`;

/**
 * The sliced faces of one session. A paint prepares its text first: the chunks it needs load, and when that
 * changed the face a new revision is installed under its own family name, because a canvas in a worker never sees
 * a replaced face under the same family. If a chunk cannot load the whole face takes over, so no glyph draws blank.
 */
export function createSlicedFontFaces(quietMs = FILL_QUIET_MS) {
  const entries = new Map<string, Entry>();
  const retired = new Map<string, { entry: Entry; hostFaces: number[] }>();
  const plain = new Map<string, Set<string>>();
  let disposed = false;
  let lastPaint = -Infinity;
  let fillTimer: ReturnType<typeof setTimeout> | undefined;
  const now = () => (globalThis as { performance?: { now(): number } }).performance?.now() ?? Date.now();
  const quiet = async () => {
    while (!disposed && now() - lastPaint < quietMs) await new Promise((resolve) => setTimeout(resolve, quietMs));
  };

  const descriptors = (bytes: ArrayBuffer): FontFaceDescriptors => {
    const traits = readFontTraits(new Uint8Array(bytes));
    return traits ? { weight: String(traits.weight), style: traits.italic ? 'italic' : traits.oblique ? 'oblique' : 'normal' } : {};
  };

  /** Revisions only grow, so a parse that finishes behind a newer revision gives way to it instead of waiting in line. */
  async function install(entry: Entry, revision: number, bytes: ArrayBuffer, chunks: ReadonlySet<number> | null): Promise<string> {
    const fonts = (globalThis as typeof globalThis & { fonts?: FontFaceSet }).fonts;
    if (!fonts) return entry.family;
    const family = revision === WHOLE ? `${entry.family} whole` : `${entry.family} r${revision}`;
    const face = await new FontFace(family, bytes, descriptors(bytes)).load();
    if (disposed || !entries.has(keyOf(entry.family))) return entry.family;
    const newest = latestOf(entry);
    if (newest && newest.revision >= revision) return newest.family;
    fonts.add(face);
    entry.installed.push({ revision, family, face, chunks });
    while (entry.installed.length > KEPT_REVISIONS) fonts.delete(entry.installed.shift()!.face);
    return family;
  }

  const wholeBytes = (entry: Entry): Promise<ArrayBuffer> => entry.wholeLoad ??= entry.whole().then((bytes) => {
    entry.wholeBytes = bytes;
    return bytes;
  }, (error: unknown) => {
    entry.wholeLoad = undefined;
    throw error;
  });
  const latestOf = (entry: Entry) => entry.installed[entry.installed.length - 1];

  /** Installs what has loaded unless the latest revision already has it; the chunks and bytes are read together. */
  function publish(entry: Entry): Promise<string> {
    const latest = latestOf(entry);
    if (latest?.revision === WHOLE) return Promise.resolve(latest.family);
    const revision = entry.face.revision;
    if (latest?.revision === revision) return Promise.resolve(latest.family);
    if (entry.installing?.revision === revision) return entry.installing.done;
    const chunks = new Set(entry.face.loadedChunks());
    const installing = { revision, done: install(entry, revision, entry.face.bytes(), chunks) };
    entry.installing = installing;
    void installing.done.catch(() => {}).finally(() => { if (entry.installing === installing) entry.installing = undefined; });
    return installing.done;
  }

  async function present(entry: Entry, text: string): Promise<string> {
    try {
      await entry.face.ensureText(text);
    } catch {
      let bytes: ArrayBuffer;
      try { bytes = await wholeBytes(entry); } catch { return entry.family; }
      const latest = latestOf(entry);
      if (latest?.revision === WHOLE) return latest.family;
      try { return await install(entry, WHOLE, bytes, null); } catch { return entry.family; }
    }
    return publish(entry);
  }

  const coverage: PaintFontCoverage = {
    async prepare(family, text) {
      const entry = entries.get(keyOf(family));
      if (!entry || disposed) return family;
      lastPaint = now();
      const latest = latestOf(entry);
      if (latest && (latest.chunks === null || entry.face.chunksOf(text).every((chunk) => latest.chunks!.has(chunk)))) return latest.family;
      const drawn = await present(entry, text);
      return drawn === entry.family ? family : drawn;
    },
    latest(family) {
      const entry = entries.get(keyOf(family));
      return entry && latestOf(entry)?.family;
    },
    covers: (family, bold, italic) => !plain.get(keyOf(family))?.has(styleOf(bold, italic)),
  };

  /**
   * Loads `text` into every face while nothing is painting, then installs one revision, so later paints of it need
   * no fetch or new face. It fetches in small batches and waits whenever a paint is recent,
   * so a paint only ever waits for its own outlines.
   */
  async function fill(text: string): Promise<void> {
    const characters = [...new Set(`${text}${text.normalize('NFC')}`)];
    for (const entry of [...entries.values()]) {
      const loaded = entry.face.loadedChunks();
      const missing = characters.filter((character) => entry.face.chunksOf(character).some((chunk) => !loaded.has(chunk)));
      for (let start = 0; start < missing.length; start += FILL_BATCH) {
        await quiet();
        if (disposed || entry.wholeBytes) return;
        const batch = missing.slice(start, start + FILL_BATCH);
        await entry.face.ensureText(batch.join(''));
      }
      await quiet();
      if (disposed || entry.wholeBytes) return;
      await publish(entry);
    }
  }

  /** The complete bytes of `family` if they have already loaded. */
  function loadedWhole(family: string): ArrayBuffer | undefined {
    return (entries.get(keyOf(family)) ?? retired.get(keyOf(family))?.entry)?.wholeBytes;
  }

  /** Every sliced family with its complete bytes, for paints that do not prepare their text. Rejects if one cannot load. */
  function wholeFaces(names: readonly string[]): Promise<{ family: string; bytes: ArrayBuffer }[]> {
    return Promise.all(names.flatMap((family) => {
      const entry = entries.get(keyOf(family)) ?? retired.get(keyOf(family))?.entry;
      return entry ? [wholeBytes(entry).then((bytes) => ({ family: entry.family, bytes: bytes.slice(0) }))] : [];
    }));
  }

  /** Whether an export keeps this face of `family`: a host's, or a style that loaded whole. */
  function keepsExportFace(family: string, face: { id: number; bold?: boolean; italic?: boolean }): boolean {
    return retired.get(keyOf(family))?.hostFaces.includes(face.id) === true ||
      plain.get(keyOf(family))?.has(styleOf(face.bold === true, face.italic === true)) === true;
  }

  const families = () => [...entries.values()].map((entry) => entry.family);
  const exportFamilies = () => [...entries.values(), ...[...retired.values()].map(({ entry }) => entry)].map((entry) => entry.family);
  const plainStyles = (): string[] => [...entries.keys()].flatMap((key) => [...plain.get(key) ?? []].map((style) => `${key}|${style}`));

  return {
    coverage,
    fill,
    loadedWhole,
    wholeFaces,
    keepsExportFace,
    families,
    exportFamilies,
    plainStyles,
    supported: () => slicedFacesSupported(),
    /** What an acknowledgement carries, so a realm that paints too knows which families are sliced. */
    ack: () => ({ slicedFamilies: families(), slicedPlainStyles: plainStyles() }),
    /** Fills from `read` a moment after fonts land, once for any number of calls. */
    scheduleFill(read: () => string | undefined): void {
      if (fillTimer !== undefined || disposed || !entries.size) return;
      fillTimer = setTimeout(() => {
        fillTimer = undefined;
        if (disposed) return;
        try {
          const text = read();
          if (text !== undefined) void fill(text).catch(() => {});
        } catch {}
      }, FILL_DELAY_MS);
    },
    /** Upgrades a face that registered after its family's whole bytes loaded, whatever the ordering. */
    upgradeRegistered(handle: ExportFaceHandle, family: string): void {
      const whole = loadedWhole(family);
      if (whole) handle.replaceExportFace(family, new Uint8Array(whole), (kept) => keepsExportFace(family, kept));
    },
    /** Gives export whole faces by font id, and by family for text measured with browser fonts. Rejects if one cannot load. */
    async completeExport(handle: ExportFaceHandle) {
      const complete = new Set<string>();
      const byFamily: { family: string; bold: boolean; italic: boolean; bytes: Uint8Array }[] = [];
      for (;;) {
        const missing = exportFamilies().filter((family) => !complete.has(family));
        if (!missing.length) return byFamily;
        for (const { family, bytes } of await wholeFaces(missing)) {
          const whole = new Uint8Array(bytes);
          handle.replaceExportFace(family, whole, (face) => keepsExportFace(family, face));
          if (entries.has(keyOf(family))) byFamily.push({ family, bold: false, italic: false, bytes: whole });
        }
        for (const family of missing) complete.add(family);
      }
    },
    /** Whole faces of the families `have` lacks; one that cannot load is left out. */
    async wholeFonts(have: unknown) {
      const known = new Set(Array.isArray(have) ? have : []);
      const faces = await Promise.all(families().filter((family) => !known.has(family)).map(async (family) => {
        const entry = entries.get(keyOf(family));
        const bytes = await (entry ? wholeBytes(entry).then((whole) => whole.slice(0)) : undefined)?.catch(() => undefined);
        return bytes ? { family, bold: false, italic: false, bytes, kind: 'coverage' as const } : undefined;
      }));
      return faces.filter((face): face is NonNullable<typeof face> => face !== undefined);
    },
    add(family: string, face: SlicedFontFace, whole: () => Promise<ArrayBuffer>): void {
      if (!disposed && !entries.has(keyOf(family))) {
        entries.set(keyOf(family), { family, face, whole, installed: [] });
      }
    },
    has: (family: string) => entries.has(keyOf(family)),
    prefetch(text: string): void { for (const entry of entries.values()) entry.face.prefetch(text); },
    /** The face with `text` loaded, for a realm that keeps its own copy at revision `have`. */
    read(family: string, text: string, have: number | null): Promise<SlicedFontRead | undefined> {
      const entry = entries.get(keyOf(family));
      if (!entry || disposed) return Promise.resolve(undefined);
      lastPaint = now();
      return (async () => {
        if (!entry.wholeBytes) {
          try {
            const revision = await entry.face.ensureText(text);
            return revision === have ? { revision } : { revision, bytes: entry.face.bytes() };
          } catch {}
        }
        const bytes = await wholeBytes(entry);
        return have === WHOLE ? { revision: WHOLE } : { revision: WHOLE, bytes: bytes.slice(0) };
      })();
    },
    /** The complete face, for paints that cannot wait on coverage. */
    whole(family: string): Promise<ArrayBuffer> | undefined {
      const entry = entries.get(keyOf(family));
      return entry && wholeBytes(entry).then((bytes) => bytes.slice(0));
    },
    /** Records a style of `family` that loaded as its own whole face, so it draws under its own name. */
    plain(family: string, bold: boolean, italic: boolean): void {
      const styles = plain.get(keyOf(family)) ?? new Set<string>();
      styles.add(styleOf(bold, italic));
      plain.set(keyOf(family), styles);
    },
    /** Stops slicing a family a host registered a face for; its older faces still need whole bytes to export. */
    forget(family: string, fontId: number): void {
      const key = keyOf(family);
      const replaced = retired.get(key);
      if (replaced) { replaced.hostFaces.push(fontId); return; }
      const entry = entries.get(key);
      if (!entry) return;
      entries.delete(key);
      retired.set(key, { entry, hostFaces: [fontId] });
      const fonts = (globalThis as typeof globalThis & { fonts?: FontFaceSet }).fonts;
      for (const { face } of entry.installed) fonts?.delete(face);
    },
    dispose(): void {
      disposed = true;
      if (fillTimer !== undefined) clearTimeout(fillTimer);
      const fonts = (globalThis as typeof globalThis & { fonts?: FontFaceSet }).fonts;
      for (const entry of entries.values()) for (const { face } of entry.installed) fonts?.delete(face);
      entries.clear();
      retired.clear();
      plain.clear();
    },
  };
}
