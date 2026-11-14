import { PngExportError, type DisplayList, type WorkbookSession, type XlsxFontFace, type XlsxPngFont } from '@betteroffice/xlsx';
import type { BundledFontFace } from '@betteroffice/fonts';

export const xlsxDefaultFontManager = {
  load: () => import('@betteroffice/fonts'),
};

const PRELOAD_MAX_BYTES = 1.5 * 1024 * 1024;
const registeredFaces = new Map<string, Promise<boolean>>();
const pendingFaces = new Map<string, Set<symbol>>();
const absentFaces = new Set<string>();
let activeFontSet: FontFaceSet | undefined;

function registerFace(
  key: string, face: BundledFontFace, family: string, load: () => Promise<ArrayBuffer>,
  fontSet: FontFaceSet, owners: Set<symbol>,
): Promise<boolean> {
  const current = () => activeFontSet === fontSet && owners.size > 0;
  return (async () => {
    try {
      if (!current()) return false;
      const bytes = await load();
      if (bytes.byteLength !== face.byteLength) {
        absentFaces.add(key);
        return false;
      }
      if (!current()) return false;
      const loaded = await new FontFace(family, bytes, {
        weight: String(face.weight), style: face.style,
      }).load();
      if (!current()) return false;
      fontSet.add(loaded);
      return true;
    } catch { return false; }
  })();
}

export function createDefaultFontPainter(
  decoder: WorkbookSession['call']['decodeFontFace'] | null | undefined,
  repaint: () => void, options: { defaultFonts?: boolean; onSettled?: () => void } = {},
) {
  const owner = Symbol();
  const interests = new Set<Set<symbol>>();
  const loaded = new Set<string>();
  const settled = new Set<string>();
  const pending = new Map<string, Promise<boolean>>();
  let disposed = false;
  let changed = false;
  let unsettledPaint = false;
  let lastPainted: DisplayList | undefined;
  let provider: Promise<{
    fonts: Awaited<ReturnType<typeof xlsxDefaultFontManager.load>>;
    source: ReturnType<Awaited<ReturnType<typeof xlsxDefaultFontManager.load>>['createFontProvider']>;
  }> | undefined;

  async function loadFace(
    fontFamily: string, bold: boolean, italic: boolean, retryShared = true, maxBytes = Infinity,
  ): Promise<boolean> {
    let owners: Set<symbol> | undefined;
    try {
      if (!provider) {
        const loading = xlsxDefaultFontManager.load().then((fonts) => ({
          fonts, source: fonts.createFontProvider(undefined, decoder ? { decoder, retain: false } : undefined),
        }));
        provider = loading;
        void loading.catch(() => { if (provider === loading) provider = undefined; });
      }
      const { fonts, source } = await provider;
      if (disposed || typeof document === 'undefined' || !document.fonts) return false;
      const face = fonts.resolveMetricCompatFace(fontFamily, bold, italic)
        ?? fonts.resolveBundledFamilyFace(fontFamily, bold, italic);
      if (!face || face.byteLength > maxBytes) return false;
      const fontSet = document.fonts;
      if (activeFontSet !== fontSet) {
        registeredFaces.clear();
        pendingFaces.clear();
        absentFaces.clear();
        activeFontSet = fontSet;
      }
      const key = JSON.stringify([fontFamily.trim().toLowerCase(), face.weight, face.style]);
      if (absentFaces.has(key)) return false;
      let registration = registeredFaces.get(key);
      const shared = registration !== undefined;
      owners = pendingFaces.get(key);
      if (!registration) {
        const load = source.resolve(fontFamily, bold, italic) ?? source.resolveFamily(fontFamily, bold, italic);
        if (!load) return false;
        owners = new Set([owner]);
        pendingFaces.set(key, owners);
        registration = registerFace(key, face, fontFamily, load, fontSet, owners);
        registeredFaces.set(key, registration);
        const task = registration;
        const group = owners;
        void task.then((success) => {
          if (!success && registeredFaces.get(key) === task) registeredFaces.delete(key);
          if (pendingFaces.get(key) === group) pendingFaces.delete(key);
        });
      }
      if (owners) {
        owners.add(owner);
        interests.add(owners);
      }
      const success = await registration;
      if (!success && shared && retryShared && !disposed) return await loadFace(fontFamily, bold, italic, false, maxBytes);
      return success;
    } catch { return false; }
    finally {
      if (owners) {
        owners.delete(owner);
        interests.delete(owners);
      }
    }
  }

  const inactive = () => disposed || options.defaultFonts === false || decoder === undefined
    || typeof document === 'undefined' || typeof FontFace === 'undefined' || !document.fonts;

  function* faceKeys(displayList: DisplayList) {
    for (const command of displayList.commands) {
      if (command.op !== 'text' || !command.fontFamily || command.chart) continue;
      const { fontFamily, bold = false, italic = false } = command;
      yield { fontFamily, bold, italic, key: JSON.stringify([fontFamily.trim().toLowerCase(), bold, italic]) };
    }
  }

  function isSettled(displayList: DisplayList): boolean {
    if (inactive()) return true;
    for (const { key } of faceKeys(displayList)) if (!settled.has(key)) return false;
    return true;
  }

  async function load(displayList: DisplayList): Promise<void> {
    if (inactive()) return;
    const batch: Promise<boolean>[] = [];
    for (const { fontFamily, bold, italic, key } of faceKeys(displayList)) {
      if (loaded.has(key) || pending.has(key)) continue;
      const task = loadFace(fontFamily, bold, italic).then((success) => {
        pending.delete(key);
        settled.add(key);
        if (success) {
          loaded.add(key);
          changed = true;
        }
        if (!disposed && !pending.size && changed) {
          changed = false;
          unsettledPaint = false;
          repaint();
        } else if (!disposed && !pending.size && unsettledPaint && lastPainted && isSettled(lastPainted)) {
          unsettledPaint = false;
          options.onSettled?.();
        }
        return success;
      });
      pending.set(key, task);
      batch.push(task);
    }
    await Promise.all(batch);
  }

  return {
    load,
    /**
     * Loads the faces of a frame that is not painted yet, one at a time and only while
     * `idle()` holds, skipping faces larger than 1.5 MB. The current frame is not repainted.
     */
    preload(displayList: DisplayList, idle: () => boolean): void {
      void (async () => {
        for (const { fontFamily, bold, italic, key } of faceKeys(displayList)) {
          if (inactive() || !idle()) return;
          if (loaded.has(key) || pending.has(key) || settled.has(key)) continue;
          if (await loadFace(fontFamily, bold, italic, true, PRELOAD_MAX_BYTES)) {
            loaded.add(key);
            settled.add(key);
          }
        }
      })().catch(() => {});
    },
    /** Takes over a frame painted before this painter existed. */
    adopt(displayList: DisplayList): void {
      lastPainted = displayList;
      if (isSettled(displayList)) {
        if (!disposed) options.onSettled?.();
        return;
      }
      unsettledPaint = true;
      void load(displayList).catch(() => {});
    },
    /** Paints the frame and reports whether its faces had settled before it was painted. */
    paint(displayList: DisplayList, paint: () => void): boolean {
      lastPainted = displayList;
      const ready = isSettled(displayList);
      if (!ready) unsettledPaint = true;
      paint();
      void load(displayList).catch(() => {});
      return ready;
    },
    dispose(): void {
      disposed = true;
      for (const owners of interests) owners.delete(owner);
      interests.clear();
      pending.clear();
    },
  };
}

/**
 * The faces the editor paints `fonts` with, for a PNG export to paint the same:
 * the bundled metric-compatible or same-family face, fetched as the editor
 * fetches it. A face the editor would use but cannot fetch rejects with
 * `font-unavailable`, so the export never paints a substitute for it.
 */
export async function exportFaces(
  fonts: readonly XlsxPngFont[], decoder: WorkbookSession['call']['decodeFontFace'] | null | undefined,
): Promise<XlsxFontFace[]> {
  if (fonts.length === 0) return [];
  const bundled = await xlsxDefaultFontManager.load();
  const source = bundled.createFontProvider(undefined, decoder ? { decoder, retain: false } : undefined);
  const faces: XlsxFontFace[] = [];
  for (const font of fonts) {
    const face = bundled.resolveMetricCompatFace(font.family, font.bold, font.italic)
      ?? bundled.resolveBundledFamilyFace(font.family, font.bold, font.italic);
    const load = face && (source.resolve(font.family, font.bold, font.italic) ?? source.resolveFamily(font.family, font.bold, font.italic));
    if (!face || !load) continue;
    let bytes: ArrayBuffer;
    try { bytes = await load(); } catch {
      throw new PngExportError('font-unavailable', `The font "${font.family}" could not load, so the sheet was not exported with a substitute.`);
    }
    if (bytes.byteLength !== face.byteLength) continue;
    faces.push({ family: font.family, bold: face.weight >= 600, italic: face.style === 'italic', bytes: new Uint8Array(bytes) });
  }
  return faces;
}
