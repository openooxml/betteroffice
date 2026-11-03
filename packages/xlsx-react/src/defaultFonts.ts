import type { DisplayList, WorkbookSession } from '@betteroffice/xlsx';
import type { BundledFontFace } from '@betteroffice/fonts';
import type { RemoteFontProvider } from '@betteroffice/fonts/remote';

export const xlsxDefaultFontManager = {
  load: () => import('@betteroffice/fonts'),
  loadRemote: () => import('@betteroffice/fonts/remote'),
};

interface RemoteTask {
  done: Promise<{ success: boolean; partial: boolean }>;
  owners: Set<symbol>;
  controller: AbortController;
}

const remoteFaces = new Map<string, RemoteTask>();
let remoteFontSet: FontFaceSet | undefined;

function installedLocally(family: string): boolean {
  let context: CanvasRenderingContext2D | null = null;
  try { context = document.createElement('canvas').getContext('2d'); } catch {}
  if (!context) return false;
  const sample = 'mmmmmmmmmmwwwwwlliI0O&@';
  return (['monospace', 'serif', 'sans-serif'] as const).some((generic) => {
    context.font = `72px ${generic}`;
    const fallback = context.measureText(sample).width;
    context.font = `72px ${JSON.stringify(family)}, ${generic}`;
    return context.measureText(sample).width !== fallback;
  });
}

function providedElsewhere(fontSet: FontFaceSet, family: string, isRemote: (face: FontFace) => boolean): boolean {
  const wanted = family.toLowerCase();
  const registered = typeof (fontSet as Partial<Iterable<FontFace>>)[Symbol.iterator] === 'function'
    ? [...(fontSet as unknown as Iterable<FontFace>)].filter((face) => face.family.replace(/^(['"])(.*)\1$/, '$2').toLowerCase() === wanted)
    : [];
  return registered.length > 0 ? registered.some((face) => !isRemote(face)) : installedLocally(family);
}

function releaseRemote(key: string, task: RemoteTask, owner: symbol): void {
  task.owners.delete(owner);
  if (task.owners.size > 0) return;
  task.controller.abort();
  if (remoteFaces.get(key) === task) remoteFaces.delete(key);
}

async function loadRemoteFace(
  fontFamily: string, bold: boolean, italic: boolean, provider: RemoteFontProvider, owner: symbol,
  held: Map<string, RemoteTask>, isLive: () => boolean, text: string,
): Promise<{ success: boolean; partial: boolean }> {
  const fontSet = document.fonts;
  if (remoteFontSet !== fontSet) {
    for (const task of remoteFaces.values()) task.controller.abort();
    remoteFaces.clear();
    remoteFontSet = fontSet;
  }
  const key = JSON.stringify([provider, fontFamily.trim().toLowerCase(), bold, italic, text]);
  let task = remoteFaces.get(key);
  if (!task) {
    const controller = new AbortController();
    const done = (async () => {
      const remote = await xlsxDefaultFontManager.loadRemote();
      controller.signal.throwIfAborted();
      if (!remote.resolveWordFamilySubstitute(fontFamily)?.remote) return { success: false, partial: false };
      if (providedElsewhere(fontSet, fontFamily.trim(), (face) => remote.isRemoteFontFace?.(face) ?? false)) {
        return { success: false, partial: false };
      }
      await remote.loadRemoteFontFaces(fontFamily, fontFamily,
        { weight: bold ? 700 : 400, style: italic ? 'italic' : 'normal' }, fontSet,
        { provider, signal: controller.signal, text });
      return { success: true, partial: false };
    })().catch((error: unknown) => ({
      success: false, partial: ((error as { faces?: unknown[] } | null)?.faces?.length ?? 0) > 0,
    }));
    const created: RemoteTask = { done, owners: new Set(), controller };
    task = created;
    remoteFaces.set(key, created);
    void done.then(() => { if (remoteFaces.get(key) === created) remoteFaces.delete(key); });
  }
  const current = task;
  current.owners.add(owner);
  held.set(key, current);
  try {
    const { success, partial } = await current.done;
    return { success: success && isLive(), partial: partial && isLive() };
  } finally {
    if (held.get(key) === current) held.delete(key);
    current.owners.delete(owner);
  }
}

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
  repaint: () => void,
  options: { defaultFonts?: boolean; remoteFonts?: RemoteFontProvider; onSettled?: () => void } = {},
) {
  const owner = Symbol();
  const remoteHeld = new Map<string, RemoteTask>();
  const interests = new Set<Set<symbol>>();
  const loaded = new Set<string>();
  const settled = new Set<string>();
  const pending = new Map<string, Promise<boolean>>();
  const remote = new Set<string>();
  const requested = new Map<string, Set<string>>();
  const attempted = new Map<string, Set<string>>();
  let disposed = false;
  let changed = false;
  let unsettledPaint = false;
  let lastPainted: DisplayList | undefined;
  let provider: Promise<{
    fonts: Awaited<ReturnType<typeof xlsxDefaultFontManager.load>>;
    source: ReturnType<Awaited<ReturnType<typeof xlsxDefaultFontManager.load>>['createFontProvider']>;
  }> | undefined;

  async function loadFace(
    fontFamily: string, bold: boolean, italic: boolean, text: string, onRemote: () => void, onPartial: () => void,
    retryShared = true,
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
      if (!face) {
        const provider = options.remoteFonts ?? 'bunny';
        if (provider === 'off') return false;
        onRemote();
        const { success, partial } = await loadRemoteFace(fontFamily, bold, italic, provider, owner, remoteHeld, () => !disposed, text);
        if (partial) onPartial();
        return success;
      }
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
      if (!success && shared && retryShared && !disposed) return await loadFace(fontFamily, bold, italic, text, onRemote, onPartial, false);
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

  function faceKeys(displayList: DisplayList) {
    const faces = new Map<string, { fontFamily: string; bold: boolean; italic: boolean; characters: Set<string> }>();
    for (const command of displayList.commands) {
      if (command.op !== 'text' || !command.fontFamily || command.chart) continue;
      const { fontFamily, bold = false, italic = false } = command;
      const key = JSON.stringify([fontFamily.trim().toLowerCase(), bold, italic]);
      let face = faces.get(key);
      if (!face) faces.set(key, face = { fontFamily, bold, italic, characters: new Set() });
      for (const character of command.text ?? '') face.characters.add(character);
      if (/[\t\n\f\r]/.test(command.text ?? '')) face.characters.add(' ');
      if (command.numberFormat) face.characters.add('#');
      if (command.numberFormat?.general !== undefined) for (const character of '0123456789.-E+') face.characters.add(character);
      for (const part of command.numberFormat?.parts ?? []) for (const character of part.character) face.characters.add(character);
    }
    return faces;
  }

  function missing(record: Map<string, Set<string>>, key: string, characters: Set<string>): string {
    const known = record.get(key);
    let text = '';
    for (const character of characters) if (!known?.has(character)) text += character;
    return text;
  }

  function isSettled(displayList: DisplayList): boolean {
    if (inactive()) return true;
    for (const [key, { characters }] of faceKeys(displayList)) {
      if (!settled.has(key) || (remote.has(key) && missing(attempted, key, characters))) return false;
    }
    return true;
  }

  async function load(displayList: DisplayList): Promise<void> {
    if (inactive()) return;
    const batch: Promise<boolean>[] = [];
    for (const [key, { fontFamily, bold, italic, characters }] of faceKeys(displayList)) {
      if (pending.has(key)) continue;
      const text = loaded.has(key) ? missing(requested, key, characters) : [...characters].join('');
      if (loaded.has(key) && !(remote.has(key) && text)) continue;
      let partial = false;
      const task = loadFace(fontFamily, bold, italic, text, () => { remote.add(key); }, () => { partial = true; }).then((success) => {
        pending.delete(key);
        for (const record of success ? [attempted, requested] : [attempted]) {
          let known = record.get(key);
          if (!known) record.set(key, known = new Set());
          for (const character of text) known.add(character);
        }
        settled.add(key);
        if (success) loaded.add(key);
        if (success || partial) changed = true;
        const latest = !disposed && remote.has(key) && lastPainted ? faceKeys(lastPainted).get(key) : undefined;
        if (latest && missing(attempted, key, latest.characters)) void load(lastPainted!).catch(() => {});
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
      for (const [key, task] of remoteHeld) releaseRemote(key, task, owner);
      remoteHeld.clear();
      for (const owners of interests) owners.delete(owner);
      interests.clear();
      pending.clear();
    },
  };
}
