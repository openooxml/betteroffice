import type { DisplayList, WorkbookSession } from '@betteroffice/xlsx';
import type { BundledFontFace } from '@betteroffice/fonts';

export const xlsxDefaultFontManager = {
  load: () => import('@betteroffice/fonts'),
};

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

  async function loadFace(fontFamily: string, bold: boolean, italic: boolean, retryShared = true): Promise<boolean> {
    let owners: Set<symbol> | undefined;
    try {
      if (!provider) {
        const loading = xlsxDefaultFontManager.load().then((fonts) => ({
          fonts, source: fonts.createFontProvider(undefined, decoder ? { decoder } : undefined),
        }));
        provider = loading;
        void loading.catch(() => { if (provider === loading) provider = undefined; });
      }
      const { fonts, source } = await provider;
      if (disposed || typeof document === 'undefined' || !document.fonts) return false;
      const face = fonts.resolveMetricCompatFace(fontFamily, bold, italic)
        ?? fonts.resolveBundledFamilyFace(fontFamily, bold, italic);
      if (!face) return false;
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
      if (!success && shared && retryShared && !disposed) return await loadFace(fontFamily, bold, italic, false);
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
