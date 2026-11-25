import { fontCdnCacheKey } from '../../../shared/office-session/fontCdnRuntime';
import { isImplicitFontCdnError, reportImplicitCdnFailure } from '../../../shared/fontCdnDiagnostics';
import { importFonts } from '../../../shared/office-session/fontModules';
import { nonCjkScriptUnicodeRange } from '../../../shared/office-session/textScripts';
import {
  cellFontKey, cellScriptFamilies, cellScriptFontKey, cellScriptFontRequests,
  type CellScriptFamilies, type EditorPngFont, type PngScriptFallbackFace, type ScriptFontDisplayList,
} from '../../../shared/office-session/xlsxScriptFonts';
import { PngExportError, type DisplayList, type WorkbookSession, type XlsxFontFace } from '@betteroffice/xlsx';
import type { BundledFontFace, BundledFontScript } from '@betteroffice/fonts';
import type { RemoteFontProvider } from '@betteroffice/fonts/remote';

export const xlsxDefaultFontManager = {
  load: importFonts,
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

const PRELOAD_MAX_BYTES = 1.5 * 1024 * 1024;
const registeredFaces = new Map<string, Promise<boolean>>();
const pendingFaces = new Map<string, Set<symbol>>();
let activeFontSet: FontFaceSet | undefined;

interface FrameFontRequest {
  fontFamily: string;
  bold: boolean;
  italic: boolean;
  characters: Set<string>;
  script?: BundledFontScript;
}

interface FrameFontRequests {
  cdn: string;
  faces: Map<string, FrameFontRequest>;
}

function registerFace(
  face: BundledFontFace, family: string, load: () => Promise<ArrayBuffer>,
  fontSet: FontFaceSet, owners: Set<symbol>, unicodeRange?: string,
): Promise<boolean> {
  const current = () => activeFontSet === fontSet && owners.size > 0;
  return (async () => {
    try {
      if (!current()) return false;
      const bytes = await load();
      if (bytes.byteLength !== face.byteLength) return false;
      if (!current()) return false;
      const loaded = await new FontFace(family, bytes, {
        weight: String(face.weight), style: face.style, ...(unicodeRange ? { unicodeRange } : {}),
      }).load();
      if (!current()) return false;
      fontSet.add(loaded);
      return true;
    } catch (error) {
      if (isImplicitFontCdnError(error)) reportImplicitCdnFailure(error);
      return false;
    }
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
  const scriptFamilies = new Map<string, string>();
  const resolveScriptFamilies: CellScriptFamilies = (cmd) => cellScriptFamilies(cmd, scriptFamilies);
  let disposed = false;
  let changed = false;
  let unsettledPaint = false;
  let lastPainted: FrameFontRequests | undefined;
  let provider: Promise<{
    fonts: Awaited<ReturnType<typeof xlsxDefaultFontManager.load>>;
    source: ReturnType<Awaited<ReturnType<typeof xlsxDefaultFontManager.load>>['createFontProvider']>;
  }> | undefined;

  async function loadFace(
    request: FrameFontRequest, text: string, onRemote: () => void, onPartial: () => void,
    retryShared = true, maxBytes = Infinity,
  ): Promise<boolean> {
    const { fontFamily, bold, italic, script } = request;
    let owners: Set<symbol> | undefined;
    try {
      if (!provider) {
        const loading = xlsxDefaultFontManager.load().then((fonts) => ({
          fonts, source: fonts.createFontProvider(undefined, { ...(decoder ? { decoder, retain: false } : {}), strict: true }),
        }));
        provider = loading;
        void loading.catch(() => { if (provider === loading) provider = undefined; });
      }
      const { fonts, source } = await provider;
      if (disposed || typeof document === 'undefined' || !document.fonts) return false;
      const face = script ? fonts.resolveScriptFallbackFace(script, bold, italic, fontFamily)
        : fonts.resolveMetricCompatFace(fontFamily, bold, italic) ?? fonts.resolveBundledFamilyFace(fontFamily, bold, italic);
      if (!face) {
        const provider = options.remoteFonts ?? 'bunny';
        if (script || provider === 'off' || maxBytes !== Infinity) return false;
        onRemote();
        const { success, partial } = await loadRemoteFace(fontFamily, bold, italic, provider, owner, remoteHeld, () => !disposed, text);
        if (partial) onPartial();
        return success;
      }
      if (face.byteLength > maxBytes) return false;
      const cdn = fontCdnCacheKey();
      if (source.faceUnavailableUntil && await source.faceUnavailableUntil(face) !== undefined) return false;
      if (disposed || !document.fonts || cdn !== fontCdnCacheKey()) return false;
      const fontSet = document.fonts;
      if (activeFontSet !== fontSet) {
        registeredFaces.clear();
        pendingFaces.clear();
        activeFontSet = fontSet;
      }
      const family = script ? `bo-xlsx-editor-${face.family}` : fontFamily;
      const unicodeRange = script ? nonCjkScriptUnicodeRange(script) : undefined;
      const key = JSON.stringify([family.trim().toLowerCase(), face.weight, face.style, cdn, unicodeRange]);
      let registration = registeredFaces.get(key);
      const shared = registration !== undefined;
      owners = pendingFaces.get(key);
      if (!registration) {
        const load = script ? source.resolveScriptFallback(script, bold, italic, fontFamily)
          : source.resolve(fontFamily, bold, italic) ?? source.resolveFamily(fontFamily, bold, italic);
        if (!load) return false;
        owners = new Set([owner]);
        pendingFaces.set(key, owners);
        registration = registerFace(face, family, load, fontSet, owners, unicodeRange);
        registeredFaces.set(key, registration);
        const task = registration;
        const group = owners;
        void task.then((success) => {
          if (success !== true && registeredFaces.get(key) === task) registeredFaces.delete(key);
          if (pendingFaces.get(key) === group) pendingFaces.delete(key);
        });
      }
      if (owners) {
        owners.add(owner);
        interests.add(owners);
      }
      const success = await registration;
      if (success === false && shared && retryShared && !disposed) return await loadFace(request, text, onRemote, onPartial, false, maxBytes);
      if (success && script && !disposed) scriptFamilies.set(cellScriptFontKey({ script, fontFamily, bold, italic }), family);
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

  function faceKeys(displayList: DisplayList, includeScriptFallbacks = true): FrameFontRequests {
    const cdn = fontCdnCacheKey();
    const faces = new Map<string, FrameFontRequest>();
    for (const command of displayList.commands) {
      if (command.op !== 'text' || !command.fontFamily || command.chart) continue;
      const { fontFamily, bold = false, italic = false } = command;
      const key = `${cdn}\0${cellFontKey(fontFamily, bold, italic)}`;
      let face = faces.get(key);
      if (!face) faces.set(key, face = { fontFamily, bold, italic, characters: new Set() });
      for (const character of command.text ?? '') face.characters.add(character);
      if (command.numberFormat) face.characters.add('#');
      if (command.numberFormat?.general !== undefined) for (const character of '0123456789.-E+') face.characters.add(character);
      for (const part of command.numberFormat?.parts ?? []) for (const character of part.character) face.characters.add(character);
    }
    for (const face of faces.values()) {
      if (['\t', '\n', '\f', '\r'].some((character) => face.characters.has(character))) face.characters.add(' ');
    }
    if (includeScriptFallbacks) {
      for (const request of cellScriptFontRequests(displayList)) {
        faces.set(`${cdn}\0${cellScriptFontKey(request)}`, { ...request, characters: new Set() });
      }
    }
    return { cdn, faces };
  }

  function missing(record: Map<string, Set<string>>, key: string, characters: Set<string>): string {
    const known = record.get(key);
    let text = '';
    for (const character of characters) if (!known?.has(character)) text += character;
    return text;
  }

  function isSettled(frame: FrameFontRequests): boolean {
    for (const [key, { characters }] of frame.faces) {
      if (!settled.has(key) || (remote.has(key) && missing(attempted, key, characters))) return false;
    }
    return true;
  }

  async function loadRequests(frame: FrameFontRequests): Promise<void> {
    if (inactive() || frame.cdn !== fontCdnCacheKey()) return;
    const batch: Promise<boolean>[] = [];
    for (const [key, request] of frame.faces) {
      if (pending.has(key)) continue;
      const { characters } = request;
      const ready = loaded.has(key);
      if (ready && !remote.has(key)) continue;
      const text = ready ? missing(requested, key, characters) : [...characters].join('');
      if (ready && !text) continue;
      let partial = false;
      const task = loadFace(request, text, () => { remote.add(key); }, () => { partial = true; }).then((success) => {
        pending.delete(key);
        for (const record of success ? [attempted, requested] : [attempted]) {
          let known = record.get(key);
          if (!known) record.set(key, known = new Set());
          for (const character of text) known.add(character);
        }
        settled.add(key);
        if (success) loaded.add(key);
        if (success || partial) changed = true;
        const current = !disposed && lastPainted?.cdn === fontCdnCacheKey() ? lastPainted : undefined;
        const latest = current && remote.has(key) ? current.faces.get(key) : undefined;
        if (current && latest && missing(attempted, key, latest.characters)) void loadRequests(current).catch(() => {});
        if (!disposed && !pending.size && changed) {
          changed = false;
          unsettledPaint = false;
          repaint();
        } else if (!pending.size && unsettledPaint && current && isSettled(current)) {
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

  function withScriptFonts(displayList: DisplayList): ScriptFontDisplayList {
    return !inactive() && scriptFamilies.size ? { ...displayList, cellScriptFamilies: resolveScriptFamilies } : displayList;
  }

  return {
    async load(displayList: DisplayList): Promise<void> {
      if (!inactive()) await loadRequests(faceKeys(displayList));
    },
    withScriptFonts,
    /**
     * Loads the faces of a frame that is not painted yet, one at a time and only while
     * `idle()` holds, skipping faces larger than 1.5 MB. The current frame is not repainted.
     */
    preload(displayList: DisplayList, idle: () => boolean): void {
      void (async () => {
        for (const [key, request] of faceKeys(displayList, false).faces) {
          if (inactive() || !idle()) return;
          if (loaded.has(key) || pending.has(key) || settled.has(key)) continue;
          if (await loadFace(request, '', () => {}, () => {}, true, PRELOAD_MAX_BYTES) === true) {
            loaded.add(key);
            settled.add(key);
          }
        }
      })().catch(() => {});
    },
    /** Takes over a frame painted before this painter existed. */
    adopt(displayList: DisplayList): void {
      const frame = lastPainted = inactive() ? undefined : faceKeys(displayList);
      if (!frame || isSettled(frame)) {
        if (!disposed) options.onSettled?.();
        return;
      }
      unsettledPaint = true;
      void loadRequests(frame).catch(() => {});
    },
    /** Paints the frame and reports whether its faces had settled before it was painted. */
    paint(displayList: DisplayList, paint: (list: DisplayList) => void): boolean {
      const frame = lastPainted = inactive() ? undefined : faceKeys(displayList);
      const ready = !frame || isSettled(frame);
      if (!ready) unsettledPaint = true;
      paint(withScriptFonts(displayList));
      if (frame) void loadRequests(frame).catch(() => {});
      return ready;
    },
    dispose(): void {
      disposed = true;
      for (const [key, task] of remoteHeld) releaseRemote(key, task, owner);
      remoteHeld.clear();
      for (const owners of interests) owners.delete(owner);
      interests.clear();
      pending.clear();
      scriptFamilies.clear();
      lastPainted = undefined;
    },
  };
}

/**
 * The faces the editor paints `fonts` with, for a PNG export to paint the same:
 * the bundled metric-compatible or same-family face, fetched as the editor fetches it.
 */
export async function exportFaces(
  fonts: readonly EditorPngFont[], decoder: WorkbookSession['call']['decodeFontFace'] | null | undefined,
  reported = new Set<Error>(),
): Promise<XlsxFontFace[]> {
  if (fonts.length === 0) return [];
  const bundled = await xlsxDefaultFontManager.load();
  const source = bundled.createFontProvider(undefined, { ...(decoder ? { decoder, retain: false } : {}), strict: true });
  const faces: XlsxFontFace[] = [];
  const scriptFaces = new Map<string, PngScriptFallbackFace>();

  async function loadExportFace(face: BundledFontFace, family: string, load: (() => Promise<ArrayBuffer>) | undefined): Promise<XlsxFontFace | undefined> {
    if (!load) return undefined;
    let bytes: ArrayBuffer;
    try { bytes = await load(); }
    catch (error) {
      if (face.script && face.script !== 'arabic' && face.script !== 'hebrew' && bundled.isMissingFontAddon(error)) return undefined;
      if (isImplicitFontCdnError(error) && (error.diagnostic.pack === 'cjk' || error.diagnostic.pack === 'intl')) {
        if (!reported.has(error)) {
          reported.add(error);
          reportImplicitCdnFailure(error);
        }
        return undefined;
      }
      throw new PngExportError('font-unavailable', `The font "${family}" could not load, so the sheet was not exported with a substitute.`);
    }
    if (bytes.byteLength !== face.byteLength) return undefined;
    return { family, bold: face.weight >= 600, italic: face.style === 'italic', bytes: new Uint8Array(bytes) };
  }

  for (const { family, bold, italic, scripts } of fonts) {
    const primary = family ? bundled.resolveMetricCompatFace(family, bold, italic) ?? bundled.resolveBundledFamilyFace(family, bold, italic) : undefined;
    if (primary) {
      const face = await loadExportFace(primary, family, source.resolve(family, bold, italic) ?? source.resolveFamily(family, bold, italic));
      if (face) faces.push(face);
    }
    for (const script of scripts ?? []) {
      const unicodeRange = nonCjkScriptUnicodeRange(script);
      if (!unicodeRange) continue;
      const fallback = bundled.resolveScriptFallbackFace(script, bold, italic, family);
      if (!fallback) continue;
      const key = `${script}\0${fallback.file}`;
      const request = { script, fontFamily: family, bold, italic };
      const known = scriptFaces.get(key);
      if (known) known.scriptFallbacks.push(request);
      else {
        const face = await loadExportFace(fallback, fallback.family, source.resolveScriptFallback(script, bold, italic, family));
        if (face) scriptFaces.set(key, { ...face, unicodeRange, scriptFallbacks: [request] });
      }
    }
  }
  return [...faces, ...scriptFaces.values()];
}
