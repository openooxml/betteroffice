import { afterEach, expect, mock, spyOn, test } from 'bun:test';
import { configureFontCdn, type BundledFontFace } from '@betteroffice/fonts';
import * as fonts from '@betteroffice/fonts';
import { initWasm, openWorkbook } from '@betteroffice/xlsx';
import JSZip from 'jszip';
import { fontAssetResponse, fontRuntimeWithoutPacks, withBrowserFontGlobals } from '../../fonts/src/fontAssets.test-support';
import { spyOnFontFetch } from '../../fonts/src/fontFetch.test-support';
import type { FontDecoder } from '../../../shared/office-session/fontDecoder';
import { getFontCdnConfiguration, setFontCdnConfiguration } from '../../../shared/office-session/fontCdn';
import { withinTestBudget } from '../../../shared/office-session/testing/loadBudget';
import * as scripts from '../../../shared/office-session/textScripts';
import { configureFontCdn as configureTestFontCdn } from '../../fonts/src/fontCdn';
import { createSessionHost } from '../../../shared/office-session/host';
import { createInProcessPair } from '../../../shared/office-session/testing/inProcessTransport';
import type { DisplayList, TextCmd } from '@betteroffice/xlsx';
import { createDefaultFontPainter, xlsxDefaultFontManager } from './defaultFonts';
import { createWorkbookSession } from '../../xlsx/src/session/client';
import type { WorkbookSessionMethods } from '../../xlsx/src/session/methods';
import { paintDisplayList } from '../../xlsx/src/render/canvas2d';
import { renderPng } from '../../xlsx/src/wasm/enginePng';
import { createXlsxCommandController } from './commands/createXlsxCommandStore';
import { createInThreadAsyncEditorApi } from './createInThreadAsyncEditorApi';
import type { XlsxEditorApi } from './XlsxEditor.types';

const descriptors = ['document', 'FontFace', 'OffscreenCanvas'].map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const);
const restorers: (() => void)[] = [];

afterEach(() => {
  for (const restore of restorers.splice(0).reverse()) restore();
  for (const [key, descriptor] of descriptors) {
    if (descriptor) Object.defineProperty(globalThis, key, descriptor);
    else Reflect.deleteProperty(globalThis, key);
  }
});

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function browser(gate: Promise<void> | ((family: string, descriptors: FontFaceDescriptors) => Promise<void>) = Promise.resolve()) {
  const added: Face[] = [];
  class Face {
    constructor(readonly family: string, readonly bytes: ArrayBuffer, readonly descriptors: FontFaceDescriptors) {}
    async load() { await (typeof gate === 'function' ? gate(this.family, this.descriptors) : gate); return this; }
  }
  const remove = mock((face: Face) => { added.splice(added.indexOf(face), 1); });
  Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: Face });
  Object.defineProperty(globalThis, 'document', { configurable: true,
    value: { fonts: { add: (face: Face) => added.push(face), delete: remove, [Symbol.iterator]: () => added[Symbol.iterator]() } } });
  return { added, remove };
}

function display(...faces: { family: string; bold?: boolean; italic?: boolean }[]): DisplayList {
  return { width: 200, height: 100, commands: faces.map(({ family, bold = false, italic = false }) => ({
    op: 'text', fontFamily: family, bold, italic, text: 'Example',
  } as TextCmd)) };
}

function provider(bytes: (face: BundledFontFace) => Promise<ArrayBuffer> = async (face) => new ArrayBuffer(face.byteLength)) {
  const face = (family: string, bold: boolean, italic: boolean): BundledFontFace => ({
    family, weight: bold ? 700 : 400, style: italic ? 'italic' : 'normal',
    file: `${family}-${bold}-${italic}.ttf.zst`, byteLength: family === 'Noto Sans JP' ? 12 : 4,
  });
  const metric = mock((family: string, bold: boolean, italic: boolean) =>
    family === 'Calibri' ? face('Carlito', bold, italic) : family === 'Arial' ? face('Liberation Sans', bold, italic)
      : family === 'MS Gothic' ? face('Noto Sans JP', bold, italic) : undefined);
  const bundled = mock((family: string, bold: boolean, italic: boolean) =>
    family === 'Carlito' || family === 'Noto Sans Thai' ? face(family, bold, italic) : undefined);
  const scriptFace = mock(fonts.resolveScriptFallbackFace);
  const scriptLoads = mock(bytes);
  const create = mock((_options?: unknown, runtime?: { decoder?: FontDecoder }) => {
    const load = (resolved: BundledFontFace | undefined, fetch: typeof bytes) => {
      return resolved ? async () => {
        if (runtime?.decoder) {
          const decoded = await runtime.decoder(new Uint8Array([1]), { codec: 'zstd' });
          expect(new Uint8Array(decoded.font)).toEqual(new Uint8Array([0, 1, 0, 0]));
        }
        return fetch(resolved);
      } : undefined;
    };
    const resolve = (resolveFace: typeof metric) => (family: string, bold: boolean, italic: boolean) =>
      load(resolveFace(family, bold, italic), bytes);
    return {
      resolve: resolve(metric), resolveFamily: resolve(bundled),
      resolveScriptFallback: (...args: Parameters<typeof fonts.resolveScriptFallbackFace>) =>
        load(fonts.resolveScriptFallbackFace(...args), scriptLoads),
    };
  });
  const module = {
    resolveMetricCompatFace: metric, resolveBundledFamilyFace: bundled,
    resolveScriptFallbackFace: scriptFace, createFontProvider: create,
  } as unknown as Awaited<ReturnType<typeof xlsxDefaultFontManager.load>>;
  const imported = spyOn(xlsxDefaultFontManager, 'load').mockResolvedValue(module);
  const remoteGate = { wait: Promise.resolve() as Promise<void> };
  const remoteLoads = mock(async (family: string, alias: string, face: { weight: number; style: string }, fontSet: FontFaceSet,
    options: { provider?: unknown; signal?: AbortSignal }) => {
    if (family !== 'Tahoma') return [];
    await remoteGate.wait;
    options.signal?.throwIfAborted();
    const font = { family: alias, remote: 'Open Sans', weight: face.weight, style: face.style, provider: options.provider };
    fontSet.add(font as unknown as FontFace);
    return [font as unknown as FontFace];
  });
  const remote = {
    resolveWordFamilySubstitute: (family: string) => family === 'Tahoma' ? { remote: 'Open Sans', generic: 'sans-serif' } : undefined,
    loadRemoteFontFaces: remoteLoads,
    isRemoteFontFace: (face: object) => 'remote' in face,
  } as unknown as Awaited<ReturnType<typeof xlsxDefaultFontManager.loadRemote>>;
  const remoteImported = spyOn(xlsxDefaultFontManager, 'loadRemote').mockResolvedValue(remote);
  restorers.push(() => imported.mockRestore(), () => remoteImported.mockRestore());
  return { imported, create, metric, bundled, scriptFace, scriptLoads, module, remoteImported, remoteLoads, remoteGate };
}

function decoder() { return mock<FontDecoder>(async () => ({ font: new Uint8Array([0, 1, 0, 0]).buffer })); }

function cellText(text: string, options: Partial<TextCmd> = {}): TextCmd {
  return { op: 'text', x: 0, y: 12, text, color: '#000000', fontSize: 12, fontFamily: 'Calibri', ...options };
}

function cells(...commands: TextCmd[]): DisplayList {
  return { width: 200, height: 100, commands };
}

function canvasText() {
  const painted: { text: string; font: string }[] = [];
  const context = new Proxy({ font: '', canvas: { width: 200, height: 100 } }, {
    get(target, property) {
      if (property === 'measureText') return (text: string) => ({ width: text.length });
      if (property === 'fillText') return (text: string) => painted.push({ text, font: target.font });
      return Reflect.has(target, property) ? Reflect.get(target, property) : () => {};
    },
  }) as unknown as CanvasRenderingContext2D;
  Object.defineProperty(globalThis, 'OffscreenCanvas', { configurable: true, value: class {
    getContext() { return context; }
    async convertToBlob() { return new Blob(['png']); }
  } });
  return { context, painted };
}

test('registers Carlito regular and bold as Calibri and repaints once after the whole batch loads', async () => {
  const gate = deferred<void>();
  const { added } = browser(gate.promise);
  const source = provider();
  const decode = decoder();
  const repaint = mock(() => {});
  const fonts = createDefaultFontPainter(decode, repaint);
  const list = display({ family: 'Calibri' }, { family: 'Calibri', bold: true }, { family: 'Calibri' });
  delete (list.commands[2] as TextCmd).bold;
  delete (list.commands[2] as TextCmd).italic;
  const loading = fonts.load(list);
  expect(added).toHaveLength(0);
  expect(repaint).not.toHaveBeenCalled();
  gate.resolve();
  await loading;
  expect(added.map(({ family, descriptors }) => ({ family, ...descriptors }))).toEqual([
    { family: 'Calibri', weight: '400', style: 'normal' },
    { family: 'Calibri', weight: '700', style: 'normal' },
  ]);
  expect(source.create).toHaveBeenCalledWith(undefined, { decoder: decode, retain: false, strict: true });
  expect(decode).toHaveBeenCalledTimes(2);
  expect(repaint).toHaveBeenCalledTimes(1);
  await fonts.load(list);
  expect(repaint).toHaveBeenCalledTimes(1);
  fonts.dispose();
  expect(added).toHaveLength(2);
});

test('requests script-restricted pack faces for Hebrew and Devanagari cells and uses device fallback for Arabic', async () => {
  const { added } = browser();
  const { context, painted } = canvasText();
  const { scriptFace, scriptLoads } = provider();
  const repaint = mock(() => {});
  const painter = createDefaultFontPainter(decoder(), repaint);
  const cells = (...texts: string[]): DisplayList => ({ width: 200, height: 100, commands: texts.map<TextCmd>((text) => ({
    op: 'text', x: 0, y: 0, text, color: '#000000', fontSize: 12, fontFamily: 'Arial',
  })) });
  try {
    await painter.load(cells('Résumé'));
    expect(scriptFace).not.toHaveBeenCalled();
    expect(scriptLoads).not.toHaveBeenCalled();
    expect(added.map(({ family }) => family)).toEqual(['Arial']);

    const frame = cells('שלום', 'नमस्ते', 'مرحبا', 'שלום');
    const loading = painter.load(frame);
    const ready = painter.paint(frame, () => {});
    await loading;
    expect(scriptLoads.mock.calls.map(([face]) => face.family).sort()).toEqual(['Noto Sans Devanagari', 'Noto Sans Hebrew']);
    expect(added.map(({ family }) => family).sort()).toEqual(['Arial', 'bo-xlsx-editor-Noto Sans Devanagari', 'bo-xlsx-editor-Noto Sans Hebrew']);
    expect(added.find(({ family }) => family.endsWith('Noto Sans Hebrew'))?.descriptors.unicodeRange ?? '').toMatch(/U\+0?590-0?5FF/i);
    expect(added.find(({ family }) => family.endsWith('Noto Sans Devanagari'))?.descriptors.unicodeRange ?? '').toMatch(/U\+0?900-0?97F/i);
    expect(ready).toBe(false);
    expect(painter.paint(frame, (list = frame) => paintDisplayList(context, list, 1))).toBe(true);
    expect(painted).toEqual([
      { text: 'שלום', font: '16px "Arial", "bo-xlsx-editor-Noto Sans Hebrew", sans-serif' },
      { text: 'नमस्ते', font: '16px "Arial", "bo-xlsx-editor-Noto Sans Devanagari", sans-serif' },
      { text: 'مرحبا', font: '16px "Arial", sans-serif' },
      { text: 'שלום', font: '16px "Arial", "bo-xlsx-editor-Noto Sans Hebrew", sans-serif' },
    ]);
    expect(scriptFace).toHaveBeenCalledTimes(2);
    expect(repaint).toHaveBeenCalledTimes(2);
  } finally { painter.dispose(); }
});

test('only requests and paints the script face for the rich-text piece that uses it', async () => {
  browser();
  const { context, painted } = canvasText();
  const { scriptLoads } = provider();
  const painter = createDefaultFontPainter(decoder(), () => {});
  const frame = cells(cellText('Bold नमस्ते', { runs: [{ start: 0, end: 5, bold: true }] }));
  try {
    await painter.load(frame);
    expect(scriptLoads.mock.calls.map(([face]) => [face.family, face.weight]))
      .toEqual([['Noto Sans Devanagari', 400]]);
    painter.paint(frame, (list = frame) => paintDisplayList(context, list, 1));
    expect(painted).toEqual([
      { text: 'Bold ', font: 'bold 16px "Calibri", sans-serif' },
      { text: 'नमस्ते', font: '16px "Calibri", "bo-xlsx-editor-Noto Sans Devanagari", sans-serif' },
    ]);
  } finally { painter.dispose(); }
});

test('another editor and an explicit-face PNG export do not pick up a disposed editor’s script fallback', async () => {
  const { added } = browser();
  const { context, painted } = canvasText();
  provider();
  const first = createDefaultFontPainter(decoder(), () => {});
  const second = createDefaultFontPainter(decoder(), () => {}, { defaultFonts: false });
  const named = cellText('שלום', { fontFamily: 'Unbundled Sans' });
  const unnamed = cellText('שלום', { fontFamily: undefined });
  const frame = cells(named, unnamed);
  try {
    await first.load(frame);
    first.dispose();
    const readFrame = () => JSON.stringify(frame);
    await renderPng({
      onSheet: (_sheet, read) => read(),
      doc: { pngDisplayListJson: readFrame, rangePngDisplayListJson: readFrame },
    }, { sheet: 0, range: 'A1:B1' }, { fonts: [] });
    second.paint(frame, (list = frame) => paintDisplayList(context, list, 1));
    expect(painted.map(({ font }) => font)).toEqual([
      '16px "Unbundled Sans", sans-serif', '16px sans-serif',
      '16px "Unbundled Sans", sans-serif', '16px sans-serif',
    ]);
    expect(added.map(({ family }) => family)).toEqual(['bo-xlsx-editor-Noto Sans Hebrew']);
  } finally { first.dispose(); second.dispose(); }
});

test('an unsettled paint collects the frame’s font requests once', async () => {
  browser();
  provider();
  const painter = createDefaultFontPainter(decoder(), () => {});
  const read = mock(() => 'नमस्ते');
  const cmd = cellText('');
  Object.defineProperty(cmd, 'text', { get: read });
  try {
    expect(painter.paint(cells(cmd), () => {})).toBe(false);
    expect(read).toHaveBeenCalledTimes(2);
  } finally { painter.dispose(); }
});

test('recreated rich-text pieces reuse script detection for their text', async () => {
  browser();
  provider();
  const detect = spyOn(scripts, 'textScripts');
  const painter = createDefaultFontPainter(decoder(), () => {});
  const { context } = canvasText();
  try {
    const frame = () => cells(cellText('Report हेतु फ़ॉन्ट', { runs: [{ start: 0, end: 7, bold: true }] }));
    const render = () => {
      const list = frame();
      painter.paint(list, (paintable = list) => paintDisplayList(context, paintable, 1));
    };
    await painter.load(frame());
    detect.mockClear();
    render();
    render();
    expect(detect).not.toHaveBeenCalled();
  } finally { painter.dispose(); detect.mockRestore(); }
});

test('painting more distinct script cells than the text cache scans each cell at most once', async () => {
  browser();
  provider();
  const detect = spyOn(scripts, 'textScripts');
  const painter = createDefaultFontPainter(decoder(), () => {});
  const { context, painted } = canvasText();
  try {
    await painter.load(cells(cellText('नमस्ते')));
    detect.mockClear();
    const frame = cells(...Array.from({ length: 1_100 }, (_, i) => cellText(`नमस्ते ${i}`)));
    expect(painter.paint(frame, (list) => paintDisplayList(context, list, 1))).toBe(true);
    expect(painted).toHaveLength(frame.commands.length);
    expect(painted.every(({ font }) => font.includes('bo-xlsx-editor-Noto Sans Devanagari'))).toBe(true);
    const scans = detect.mock.calls.map(([text]) => text);
    expect(scans).toHaveLength(frame.commands.length);
    expect(scans.length).toBe(new Set(scans).size);
  } finally { painter.dispose(); detect.mockRestore(); }
});

test('an echo paint scans repeated text once even when it is too long to cache across paints', async () => {
  browser();
  provider();
  const detect = spyOn(scripts, 'textScripts');
  const painter = createDefaultFontPainter(decoder(), () => {});
  const { context, painted } = canvasText();
  const text = `${'Résumé '.repeat(80)}שלום`;
  const frame = cells(cellText(text), cellText(text, { y: 24 }));
  try {
    await painter.load(cells(cellText('שלום')));
    detect.mockClear();
    paintDisplayList(context, painter.withScriptFonts(frame), 1);
    expect(painted.map(({ font }) => font)).toEqual(Array(2).fill('16px "Calibri", "bo-xlsx-editor-Noto Sans Hebrew", sans-serif'));
    expect(detect.mock.calls).toEqual([[text]]);
  } finally { painter.dispose(); detect.mockRestore(); }
});

async function singleCellWorkbook() {
  await initWasm();
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>');
  return openWorkbook(await zip.generateAsync({ type: 'uint8array' }));
}

test.each([
  ['Latin', 'Résumé', 'Calibri', 4],
  ['CJK', '日本語', 'MS Gothic', 12],
] as const)('an editor PNG keeps the editor-provided face for a %s-only sheet', async (_script, text, family, byteLength) => {
  const handle = await singleCellWorkbook();
  const { remove } = browser();
  const { painted } = canvasText();
  const { create } = provider();
  const target = { sheet: 0, range: 'A1' };
  try {
    handle.editCell(0, 0, 0, text);
    handle.patchRangeStyle(0, 'A1', { fontFamily: family });
    expect(handle.renderPngFonts(target)).toEqual([{ family, bold: false, italic: false }]);

    const api = createInThreadAsyncEditorApi({ handle, commands: createXlsxCommandController().store,
      flushPendingInput: async () => {} } as unknown as XlsxEditorApi, false, () => {});
    await expect(api.renderPng(target)).resolves.toEqual(new TextEncoder().encode('png'));
    expect(create).toHaveBeenCalledTimes(1);
    expect(create).toHaveBeenCalledWith(undefined, { strict: true });
    expect(remove).toHaveBeenCalledTimes(1);
    const [face] = remove.mock.calls[0]!;
    expect(face.bytes.byteLength).toBe(byteLength);
    expect(painted).toContainEqual({ text, font: expect.stringContaining(`"${face.family}"`) });
  } finally { handle.dispose(); }
}, 30_000);

test.each([
  ['Hebrew', 'שלום', 'hebrew', 'Noto Sans Hebrew'],
  ['Arabic', 'مرحبا', 'arabic', undefined],
] as const)('an editor PNG uses the engine fallback for %s in Calibri', async (_name, text, script, canvasFamily) => {
  const handle = await singleCellWorkbook();
  const { added, remove } = browser();
  const { context, painted } = canvasText();
  const bytes = async (face: BundledFontFace) => {
    const data = new Uint8Array(face.byteLength);
    new TextEncoder().encodeInto(face.file, data);
    return data.buffer;
  };
  const { scriptLoads } = provider(bytes);
  const load = spyOn(fonts, 'loadBundledFontBytes').mockImplementation(bytes);
  const installed = spyOn(fonts, 'bundledFaceInstalled').mockResolvedValue(true);
  restorers.push(() => load.mockRestore(), () => installed.mockRestore());
  const painter = createDefaultFontPainter(decoder(), () => {});
  const target = { sheet: 0, range: 'A1' };
  try {
    handle.editCell(0, 0, 0, text);
    handle.patchRangeStyle(0, 'A1', { fontFamily: 'Calibri' });
    expect(handle.renderPngFonts(target)).toEqual([]);
    const frame = handle.displayList({ x: 0, y: 0, width: 200, height: 100 });
    await painter.load(frame);
    painter.paint(frame, (list) => paintDisplayList(context, list, 1));
    const engineFace = fonts.resolveScriptFallbackFace(script, false, false, 'Calibri')!;
    expect(scriptLoads.mock.calls.map(([face]) => face.family)).toEqual(canvasFamily ? [canvasFamily] : []);
    const editorFamily = canvasFamily ? `bo-xlsx-editor-${canvasFamily}` : undefined;
    expect(painted).toContainEqual({ text, font: editorFamily
      ? expect.stringContaining(`"${editorFamily}"`) : expect.stringMatching(/px "Calibri", sans-serif$/) });
    painted.length = 0;

    const api = createInThreadAsyncEditorApi({ handle, commands: createXlsxCommandController().store,
      flushPendingInput: async () => {} } as unknown as XlsxEditorApi, false, () => {});
    await expect(api.renderPng(target)).resolves.toEqual(new TextEncoder().encode('png'));
    expect(load.mock.calls.filter(([face]) => face.file === engineFace.file)).toHaveLength(1);
    const exported = remove.mock.calls.map(([face]) => face).filter(({ bytes }) =>
      new TextDecoder().decode(bytes.slice(0, engineFace.file.length)) === engineFace.file);
    expect(exported).toHaveLength(1);
    if (editorFamily) expect(exported[0]!.bytes).toEqual(added.find(({ family }) => family === editorFamily)?.bytes);
    expect(painted).toContainEqual({ text, font: expect.stringContaining(`"${exported[0]!.family}"`) });
  } finally { painter.dispose(); handle.dispose(); }
}, 30_000);

for (const cdn of ['disabled', 'implicit', 'explicit'] as const) {
  test(`paint settles with an unavailable intl pack and ${cdn} CDN`, async () => {
    browser();
    const { context, painted } = canvasText();
    const previous = getFontCdnConfiguration();
    const restore = withBrowserFontGlobals();
    const runtime = fontRuntimeWithoutPacks();
    const manager = spyOn(xlsxDefaultFontManager, 'load').mockResolvedValue({ ...fonts, default: fonts, ...runtime });
    const fetch = spyOnFontFetch(async () => new Response(null, { status: 403 }));
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    const decode = decoder();
    const painter = createDefaultFontPainter(decode, () => {}, { remoteFonts: 'off' });
    const text = 'Total ภาษาไทย';
    const frame = cells(cellText(text, { fontFamily: 'Aptos' }));
    try {
      if (cdn === 'implicit') setFontCdnConfiguration(undefined);
      else configureFontCdn({ baseUrl: cdn === 'disabled' ? false : 'https://blocked-export.example/' });
      await painter.load(frame);
      expect(painter.paint(frame, (list) => paintDisplayList(context, list, 1))).toBe(true);
      const editor = painted.splice(0);
      expect(editor).toEqual([{ text, font: '16px "Aptos", sans-serif' }]);

      if (cdn === 'disabled') expect(fetch).not.toHaveBeenCalled();
      else expect(fetch).toHaveBeenCalled();
    } finally {
      painter.dispose();
      warn.mockRestore();
      fetch.mockRestore();
      manager.mockRestore();
      restore();
      setFontCdnConfiguration(previous);
    }
  });
}

test('CJK punctuation in Calibri makes no script pack request', async () => {
  const { added } = browser();
  const { scriptFace, scriptLoads } = provider();
  const painter = createDefaultFontPainter(decoder(), () => {});
  const frame = display({ family: 'Calibri' }, { family: 'Calibri' });
  (frame.commands[0] as TextCmd).text = '。';
  (frame.commands[1] as TextCmd).text = '（注）';
  try {
    await painter.load(frame);
    expect(scriptLoads).not.toHaveBeenCalled();
    expect(scriptFace).not.toHaveBeenCalled();
    expect(added.map(({ family }) => family)).toEqual(['Calibri']);
  } finally { painter.dispose(); }
});

test('preloading a frame that is not painted registers its faces without repainting, and its later paint is settled', async () => {
  const { added } = browser();
  provider();
  const repaint = mock(() => {});
  const fonts = createDefaultFontPainter(decoder(), repaint);
  const painted = display({ family: 'Carlito' });
  fonts.paint(painted, () => {});
  for (let tick = 0; tick < 50 && repaint.mock.calls.length === 0; tick++) await new Promise((resolve) => setTimeout(resolve, 0));
  expect(repaint).toHaveBeenCalledTimes(1);
  const neighbour = display({ family: 'Calibri' });
  fonts.preload(neighbour, () => true);
  for (let tick = 0; tick < 50 && added.length < 2; tick++) await new Promise((resolve) => setTimeout(resolve, 0));
  expect(added.map((face) => face.family)).toEqual(['Carlito', 'Calibri']);
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(repaint).toHaveBeenCalledTimes(1);
  expect(fonts.paint(neighbour, () => {})).toBe(true);
  fonts.dispose();
});

test('preloading loads one face at a time and stops once the editor is busy', async () => {
  const { added } = browser();
  provider();
  const decode = decoder();
  const fonts = createDefaultFontPainter(decode, mock(() => {}));
  fonts.preload(display({ family: 'Calibri' }, { family: 'Carlito', bold: true }), () => false);
  await new Promise((resolve) => setTimeout(resolve, 10));
  expect(added).toEqual([]);
  expect(decode).not.toHaveBeenCalled();
  let budget = 1;
  fonts.preload(display({ family: 'Calibri' }, { family: 'Carlito', bold: true }), () => budget-- > 0);
  for (let tick = 0; tick < 50 && added.length < 1; tick++) await new Promise((resolve) => setTimeout(resolve, 0));
  await new Promise((resolve) => setTimeout(resolve, 10));
  expect(added.map((face) => face.family)).toEqual(['Calibri']);
  fonts.dispose();
});

test('a preload with a byte budget makes no remote font request; an unbudgeted load does', async () => {
  const { added } = browser();
  const { remoteLoads, remoteImported } = provider();
  const fonts = createDefaultFontPainter(decoder(), mock(() => {}));
  let checks = 0;
  fonts.preload(display({ family: 'Tahoma' }), () => { checks++; return true; });
  for (let tick = 0; tick < 50 && checks === 0; tick++) await new Promise((resolve) => setTimeout(resolve, 0));
  await new Promise((resolve) => setTimeout(resolve, 10));
  expect(checks).toBeGreaterThan(0);
  expect(remoteImported).not.toHaveBeenCalled();
  expect(remoteLoads).not.toHaveBeenCalled();
  expect(added).toEqual([]);
  await fonts.load(display({ family: 'Tahoma' }));
  expect(remoteLoads).toHaveBeenCalledTimes(1);
  expect(added).toEqual([expect.objectContaining({ family: 'Tahoma', remote: 'Open Sans' })]);
  fonts.dispose();
});

test('uses bundled family names and italic descriptors', async () => {
  const { added } = browser();
  provider();
  const fonts = createDefaultFontPainter(decoder(), () => {});
  await fonts.load(display({ family: 'Carlito', italic: true }));
  expect(added.map(({ family, descriptors }) => ({ family, ...descriptors }))).toEqual([
    { family: 'Carlito', weight: '400', style: 'italic' },
  ]);
  fonts.dispose();
});

test('packed faces decode through the workbook worker RPC', async () => {
  const { added } = browser();
  provider();
  const pair = createInProcessPair();
  const decode = mock((_bytes: ArrayBuffer, _options: Parameters<FontDecoder>[1]) => ({
    font: new Uint8Array([0, 1, 0, 0]).buffer,
  }));
  createSessionHost<Pick<WorkbookSessionMethods, 'open' | 'decodeFontFace'>, {}, undefined>(pair.host, {
    context: undefined,
    policies: { open: { lane: 'interactive' }, decodeFontFace: { lane: 'interactive' } },
    handlers: {
      open: () => ({ format: 'xlsx', stage: 'ready', version: 0, dirty: false, activeSheet: 0, sheets: [] }),
      decodeFontFace: (_, bytes, options) => decode(bytes, options),
    },
  });
  const session = await createWorkbookSession(new Uint8Array([80, 75]), {}, pair.client);
  const fonts = createDefaultFontPainter(session.call.decodeFontFace, () => {});
  try {
    await fonts.load(display({ family: 'Calibri' }));
    expect(added).toHaveLength(1);
    expect(decode).toHaveBeenCalledTimes(1);
    expect(new Uint8Array(decode.mock.calls[0][0])).toEqual(new Uint8Array([1]));
    expect(decode.mock.calls[0][1]).toEqual({ codec: 'zstd', dict: undefined });
  } finally { fonts.dispose(); await session.dispose(); }
});

test('a frame reports settled faces only once its families have loaded', async () => {
  const { added } = browser();
  provider();
  const repainted = deferred<void>();
  const repaint = mock(() => repainted.resolve());
  const fonts = createDefaultFontPainter(decoder(), repaint);
  const list = display({ family: 'Calibri' });
  const paint = mock(() => {});
  expect(fonts.paint(list, paint)).toBe(false);
  expect(paint).toHaveBeenCalledTimes(1);
  await repainted.promise;
  expect(added).toHaveLength(1);
  expect(fonts.paint(list, paint)).toBe(true);
  expect(repaint).toHaveBeenCalledTimes(1);
  fonts.dispose();
});

test('a frame whose faces all fail settles without a repaint', async () => {
  const { added } = browser();
  provider(async () => { throw new Error('Missing optional font pack'); });
  const settled = deferred<void>();
  const onSettled = mock(() => settled.resolve());
  const repaint = mock(() => {});
  const fonts = createDefaultFontPainter(decoder(), repaint, { onSettled });
  const list = display({ family: 'Noto Sans Thai' }, { family: 'Unknown Face' });
  expect(fonts.paint(list, () => {})).toBe(false);
  await settled.promise;
  expect(added).toHaveLength(0);
  expect(repaint).not.toHaveBeenCalled();
  expect(fonts.paint(list, () => {})).toBe(true);
  expect(onSettled).toHaveBeenCalledTimes(1);
  fonts.dispose();
});

test('a newer unsettled frame is not marked settled by an older batch', async () => {
  browser();
  const gate = deferred<void>();
  provider(async (face) => {
    if (face.family === 'Noto Sans Thai') throw new Error('Missing optional font pack');
    await gate.promise;
    return new ArrayBuffer(face.byteLength);
  });
  const onSettled = mock(() => {});
  const repainted = deferred<void>();
  const fonts = createDefaultFontPainter(decoder(), () => repainted.resolve(), { onSettled });
  expect(fonts.paint(display({ family: 'Noto Sans Thai' }), () => {})).toBe(false);
  expect(fonts.paint(display({ family: 'Calibri' }), () => {})).toBe(false);
  await Promise.resolve();
  expect(onSettled).not.toHaveBeenCalled();
  gate.resolve();
  await repainted.promise;
  expect(onSettled).not.toHaveBeenCalled();
  fonts.dispose();
});

test('an opted-out painter reports every frame as settled', () => {
  browser();
  const { imported } = provider();
  const fonts = createDefaultFontPainter(decoder(), () => {}, { defaultFonts: false });
  expect(fonts.paint(display({ family: 'Calibri' }), () => {})).toBe(true);
  expect(imported).not.toHaveBeenCalled();
});

test('unknown families register nothing', async () => {
  const { added } = browser();
  provider();
  const decode = decoder();
  const repaint = mock(() => {});
  await createDefaultFontPainter(decode, repaint).load(display({ family: 'Unknown Face' }));
  expect(added).toHaveLength(0);
  expect(decode).not.toHaveBeenCalled();
  expect(repaint).not.toHaveBeenCalled();
});

test('a family without a bundled face loads its look-alike from the remote provider under its own name', async () => {
  const { added } = browser();
  const { remoteLoads } = provider();
  const repaint = mock(() => {});
  await createDefaultFontPainter(decoder(), repaint).load(display({ family: 'Tahoma', bold: true }));
  expect(added).toEqual([expect.objectContaining({ family: 'Tahoma', remote: 'Open Sans', weight: 700, style: 'normal', provider: 'bunny' })]);
  expect(repaint).toHaveBeenCalledTimes(1);
  await createDefaultFontPainter(decoder(), repaint, { remoteFonts: 'google' }).load(display({ family: 'Tahoma', bold: true }));
  expect(remoteLoads.mock.calls.map((call) => call[4].provider)).toEqual(['bunny', 'google']);
});

test('remote look-alikes load for the painted characters and again only for newly painted ones', async () => {
  browser();
  const { remoteLoads } = provider();
  const repaint = mock(() => {});
  const painter = createDefaultFontPainter(decoder(), repaint);
  const cell = (text: string): DisplayList => ({ width: 200, height: 100, commands: [
    { op: 'text', fontFamily: 'Tahoma', bold: false, italic: false, text } as TextCmd] });
  await painter.load(cell('Abc'));
  await painter.load(cell('bcA'));
  expect(painter.paint(cell('Abc'), () => {})).toBe(true);
  await painter.load(cell('Abcδ'));
  expect(painter.paint(cell('Abcδ'), () => {})).toBe(true);
  expect(remoteLoads.mock.calls.map((call) => (call[4] as { text?: string }).text)).toEqual(['Abc', 'δ']);
  expect(repaint).toHaveBeenCalledTimes(2);
});

test('a registered look-alike does not count as a locally installed family for later characters', async () => {
  const { added } = browser();
  const { fonts } = (globalThis as unknown as { document: { fonts: FontFaceSet } }).document;
  Object.defineProperty(globalThis, 'document', { configurable: true, value: { fonts, createElement: () => ({ getContext: () => {
    let font = '';
    return { set font(value: string) { font = value; }, get font() { return font; },
      measureText: () => ({ width: font.includes('"Tahoma"') && added.some((face) => face.family === 'Tahoma') ? 2 : 1 }) };
  } }) } });
  const { remoteLoads } = provider();
  const painter = createDefaultFontPainter(decoder(), mock(() => {}));
  const cell = (text: string): DisplayList => ({ width: 200, height: 100, commands: [
    { op: 'text', fontFamily: 'Tahoma', bold: false, italic: false, text } as TextCmd] });
  await painter.load(cell('A'));
  await painter.load(cell('AŁ'));
  expect(remoteLoads.mock.calls.map((call) => (call[4] as { text?: string }).text)).toEqual(['A', 'Ł']);
  expect(painter.paint(cell('AŁ'), () => {})).toBe(true);
  await createDefaultFontPainter(decoder(), mock(() => {})).load(cell('A'));
  expect(remoteLoads).toHaveBeenCalledTimes(3);
});

test('a registered regular look-alike does not stop the bold and italic faces of the family from loading', async () => {
  const { added } = browser();
  const { fonts } = (globalThis as unknown as { document: { fonts: FontFaceSet } }).document;
  Object.defineProperty(globalThis, 'document', { configurable: true, value: { fonts, createElement: () => ({ getContext: () => {
    let font = '';
    return { set font(value: string) { font = value; }, get font() { return font; },
      measureText: () => ({ width: font.includes('"Tahoma"') && added.some((face) => face.family === 'Tahoma') ? 2 : 1 }) };
  } }) } });
  provider();
  const painter = createDefaultFontPainter(decoder(), mock(() => {}));
  const cell = (bold: boolean, italic: boolean): DisplayList => ({ width: 200, height: 100, commands: [
    { op: 'text', fontFamily: 'Tahoma', bold, italic, text: 'A' } as TextCmd] });
  await painter.load(cell(false, false));
  await painter.load(cell(true, false));
  await painter.load(cell(false, true));
  expect((added as unknown as { family: string; weight: number; style: string }[]).map((face) => [face.family, face.weight, face.style]))
    .toEqual([['Tahoma', 400, 'normal'], ['Tahoma', 700, 'normal'], ['Tahoma', 400, 'italic']]);
  painter.dispose();
});

test('characters painted while a failing remote load is pending are requested once it settles', async () => {
  browser();
  const { remoteLoads } = provider();
  const gate = deferred<void>();
  remoteLoads.mockImplementationOnce(async () => { await gate.promise; throw new Error('404'); });
  const repaint = mock(() => {});
  const painter = createDefaultFontPainter(decoder(), repaint);
  const cell = (text: string): DisplayList => ({ width: 200, height: 100, commands: [
    { op: 'text', fontFamily: 'Tahoma', bold: false, italic: false, text } as TextCmd] });
  const first = painter.load(cell('A'));
  for (let tick = 0; tick < 50 && remoteLoads.mock.calls.length === 0; tick++) await Promise.resolve();
  expect(painter.paint(cell('AŁ'), () => {})).toBe(false);
  gate.resolve();
  await first;
  for (let wait = 0; wait < 100 && repaint.mock.calls.length === 0; wait++) await new Promise((resolve) => setTimeout(resolve, 5));
  expect(remoteLoads.mock.calls.map((call) => (call[4] as { text?: string }).text)).toEqual(['A', 'AŁ']);
  expect(repaint).toHaveBeenCalledTimes(1);
  expect(painter.paint(cell('AŁ'), () => {})).toBe(true);
});

test('a failed load of newly painted characters settles and is retried only by a later paint', async () => {
  browser();
  const { remoteLoads } = provider();
  const settled = mock(() => {});
  const painter = createDefaultFontPainter(decoder(), mock(() => {}), { onSettled: settled });
  const cell = (text: string): DisplayList => ({ width: 200, height: 100, commands: [
    { op: 'text', fontFamily: 'Tahoma', bold: false, italic: false, text } as TextCmd] });
  await painter.load(cell('A'));
  remoteLoads.mockImplementationOnce(async () => { throw new Error('404'); });
  expect(painter.paint(cell('AŁ'), () => {})).toBe(false);
  await new Promise((resolve) => setTimeout(resolve, 30));
  expect(remoteLoads.mock.calls.map((call) => (call[4] as { text?: string }).text)).toEqual(['A', 'Ł']);
  expect(settled).toHaveBeenCalledTimes(1);
  expect(painter.paint(cell('Ł'), () => {})).toBe(true);
  await new Promise((resolve) => setTimeout(resolve, 30));
  expect(remoteLoads.mock.calls.map((call) => (call[4] as { text?: string }).text)).toEqual(['A', 'Ł', 'Ł']);
});

test('number format fill and padding characters count as painted characters', async () => {
  browser();
  const { remoteLoads } = provider();
  await createDefaultFontPainter(decoder(), mock(() => {})).load({ width: 200, height: 100, commands: [{
    op: 'text', fontFamily: 'Tahoma', bold: false, italic: false, text: '7',
    numberFormat: { left: 0, width: 80, parts: [{ offset: 0, kind: 'fill', character: 'Ł' }] },
  } as TextCmd] });
  expect(remoteLoads.mock.calls.map((call) => (call[4] as { text?: string }).text)).toEqual(['7#Ł']);
});

test('a partly failed remote load repaints for the slices it registered and retries only on a later paint', async () => {
  browser();
  const { remoteLoads } = provider();
  remoteLoads.mockImplementationOnce(async () => {
    throw Object.assign(new Error('Some remote font slices failed to load'), { faces: [{ family: 'Tahoma', remote: 'Open Sans' }] });
  });
  const repaint = mock(() => {});
  const painter = createDefaultFontPainter(decoder(), repaint);
  const cell = (text: string): DisplayList => ({ width: 200, height: 100, commands: [
    { op: 'text', fontFamily: 'Tahoma', bold: false, italic: false, text } as TextCmd] });
  await painter.load(cell('AŁ'));
  expect(repaint).toHaveBeenCalledTimes(1);
  remoteLoads.mockImplementationOnce(async () => { throw Object.assign(new Error('Some remote font slices failed to load'), { faces: [] }); });
  await painter.load(cell('AŁ'));
  expect(remoteLoads.mock.calls.map((call) => (call[4] as { text?: string }).text)).toEqual(['AŁ', 'AŁ']);
  expect(repaint).toHaveBeenCalledTimes(1);
});

test('characters the General format generates when it shortens a number are requested', async () => {
  browser();
  const { remoteLoads } = provider();
  await createDefaultFontPainter(decoder(), mock(() => {})).load({ width: 200, height: 100, commands: [{
    op: 'text', fontFamily: 'Tahoma', bold: false, italic: false, text: '7',
    numberFormat: { left: 0, width: 80, parts: [], general: 7 },
  } as unknown as TextCmd] });
  expect(remoteLoads.mock.calls.map((call) => (call[4] as { text?: string }).text)).toEqual(['7#012345689.-E+']);
});

test('whitespace the canvas paints as a space requests the space character', async () => {
  browser();
  const { remoteLoads } = provider();
  await createDefaultFontPainter(decoder(), mock(() => {})).load({ width: 200, height: 100, commands: [
    { op: 'text', fontFamily: 'Tahoma', bold: false, italic: false, text: 'A\tA' } as TextCmd] });
  expect(remoteLoads.mock.calls.map((call) => (call[4] as { text?: string }).text)).toEqual(['A\t ']);
  await createDefaultFontPainter(decoder(), mock(() => {})).load({ width: 200, height: 100, commands: [{
    op: 'text', fontFamily: 'Tahoma', bold: false, italic: false, text: '7',
    numberFormat: { left: 0, width: 80, parts: [{ offset: 1, kind: 'fill', character: '\t' }] },
  } as TextCmd] });
  expect(remoteLoads.mock.calls.map((call) => (call[4] as { text?: string }).text)).toEqual(['A\t ', '7#\t ']);
});

test('a face the host registered under the family keeps the remote provider away', async () => {
  const { added } = browser();
  const { remoteLoads } = provider();
  (globalThis as unknown as { document: { fonts: FontFaceSet } }).document.fonts.add(new FontFace('Tahoma', new ArrayBuffer(4)) as never);
  await createDefaultFontPainter(decoder(), mock(() => {})).load(display({ family: 'Tahoma' }));
  expect(remoteLoads).not.toHaveBeenCalled();
  expect(added).toHaveLength(1);
});

test('disposing the painter, as switching remote fonts off does, aborts its remote load before anything registers', async () => {
  const { added } = browser();
  const { remoteLoads, remoteGate } = provider();
  const gate = deferred<void>();
  remoteGate.wait = gate.promise;
  const repaint = mock(() => {});
  const painter = createDefaultFontPainter(decoder(), repaint);
  const loading = painter.load(display({ family: 'Tahoma' }));
  for (let tick = 0; tick < 20 && remoteLoads.mock.calls.length === 0; tick++) await Promise.resolve();
  expect(remoteLoads).toHaveBeenCalledTimes(1);
  const signal = remoteLoads.mock.calls[0]![4].signal!;
  painter.dispose();
  expect(signal.aborted).toBe(true);
  gate.resolve();
  await loading;
  expect(added).toEqual([]);
  expect(repaint).not.toHaveBeenCalled();
});

test('remote fonts off, bundled faces and locally installed families never reach the provider', async () => {
  const { added } = browser();
  const { remoteLoads, remoteImported } = provider();
  const repaint = mock(() => {});
  await createDefaultFontPainter(decoder(), repaint, { remoteFonts: 'off' }).load(display({ family: 'Tahoma' }));
  await createDefaultFontPainter(decoder(), repaint).load(display({ family: 'Calibri' }));
  expect(remoteImported).not.toHaveBeenCalled();
  Object.defineProperty(globalThis, 'document', { configurable: true, value: {
    fonts: (globalThis.document as Document).fonts,
    createElement: () => ({ getContext: () => {
      let font = '';
      return { set font(value: string) { font = value; }, get font() { return font; },
        measureText: () => ({ width: font.includes('Tahoma') ? 2 : 1 }) };
    } }),
  } });
  await createDefaultFontPainter(decoder(), repaint).load(display({ family: 'Tahoma', italic: true }));
  expect(remoteLoads).not.toHaveBeenCalled();
  expect(added.map((face) => (face as unknown as { family: string }).family)).toEqual(['Calibri']);
});

test('a painter retries a missing CJK face without registering its Latin fallback', async () => {
  const { added } = browser();
  const fallback = mock(async () => new ArrayBuffer(4));
  provider(fallback);
  const repaint = mock(() => {});
  const painter = createDefaultFontPainter(decoder(), repaint);
  const frame = display({ family: 'MS Gothic' });
  try {
    await painter.load(frame);
    await painter.load(frame);
    expect(fallback).toHaveBeenCalledTimes(2);
    expect(added).toHaveLength(0);
    expect(repaint).not.toHaveBeenCalled();
  } finally { painter.dispose(); }
});

test('a missing intl pack leaves browser text untouched', async () => {
  const { added } = browser();
  provider(async () => { throw new Error('Missing optional font pack'); });
  const repaint = mock(() => {});
  await createDefaultFontPainter(decoder(), repaint).load(display({ family: 'Noto Sans Thai' }));
  expect(added).toHaveLength(0);
  expect(repaint).not.toHaveBeenCalled();
});

test('opt-out and non-DOM environments never import the font provider', async () => {
  browser();
  const source = provider();
  await createDefaultFontPainter(decoder(), () => {}, { defaultFonts: false }).load(display({ family: 'Calibri' }));
  Reflect.deleteProperty(globalThis, 'document');
  await createDefaultFontPainter(decoder(), () => {}).load(display({ family: 'Calibri' }));
  expect(source.imported).not.toHaveBeenCalled();
});

test('generic browser text and chart families do not request default faces', async () => {
  browser();
  const source = provider();
  const list = display({ family: 'Calibri' }, { family: 'Calibri' });
  delete (list.commands[0] as TextCmd).fontFamily;
  (list.commands[1] as TextCmd).chart = true;
  await createDefaultFontPainter(decoder(), () => {}).load(list);
  expect(source.imported).not.toHaveBeenCalled();
});

test('disposal during loading prevents late registration and repaint', async () => {
  const gate = deferred<void>();
  const { added } = browser(gate.promise);
  provider();
  const repaint = mock(() => {});
  const fonts = createDefaultFontPainter(decoder(), repaint);
  const loading = fonts.load(display({ family: 'Calibri' }));
  await new Promise((resolve) => setTimeout(resolve, 0));
  fonts.dispose();
  gate.resolve();
  await loading;
  expect(added).toHaveLength(0);
  expect(repaint).not.toHaveBeenCalled();
});

test('first paint precedes importing and resolving default fonts', async () => {
  const gate = deferred<ArrayBuffer>();
  const { added } = browser();
  const source = provider(async () => gate.promise);
  const events: string[] = [];
  source.imported.mockImplementation(async () => {
    events.push('import');
    return source.module;
  });
  const repaint = mock(() => {});
  const fonts = createDefaultFontPainter(decoder(), repaint);
  fonts.paint(display({ family: 'Calibri' }), () => { events.push('paint'); });
  expect(events).toEqual(['paint', 'import']);
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(added).toHaveLength(0);
  expect(repaint).not.toHaveBeenCalled();
  gate.resolve(new ArrayBuffer(4));
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(added).toHaveLength(1);
  expect(repaint).toHaveBeenCalledTimes(1);
  fonts.dispose();
});

test('defaultFonts false paints without importing the font manager', async () => {
  browser();
  const source = provider();
  const paint = mock(() => {});
  const fonts = createDefaultFontPainter(decoder(), () => {}, { defaultFonts: false });
  fonts.paint(display({ family: 'Calibri' }), paint);
  await Promise.resolve();
  expect(paint).toHaveBeenCalledTimes(1);
  expect(source.imported).not.toHaveBeenCalled();
  fonts.dispose();
});

test('disposal before the provider arrives prevents decoding and registration', async () => {
  const { added } = browser();
  const source = provider();
  const gate = deferred<typeof source.module>();
  source.imported.mockImplementation(() => gate.promise);
  const decode = decoder();
  const repaint = mock(() => {});
  const fonts = createDefaultFontPainter(decode, repaint);
  const loading = fonts.load(display({ family: 'Calibri' }));
  fonts.dispose();
  gate.resolve(source.module);
  await loading;
  expect(decode).not.toHaveBeenCalled();
  expect(added).toHaveLength(0);
  expect(repaint).not.toHaveBeenCalled();
});

test('frames arriving during the provider import retain their faces and share a repaint', async () => {
  const { added } = browser();
  const source = provider();
  const gate = deferred<typeof source.module>();
  source.imported.mockImplementation(() => gate.promise);
  const repaint = mock(() => {});
  const fonts = createDefaultFontPainter(decoder(), repaint);
  const first = fonts.load(display({ family: 'Calibri' }));
  const second = fonts.load(display({ family: 'Carlito' }));
  expect(source.imported).toHaveBeenCalledTimes(1);
  gate.resolve(source.module);
  await Promise.all([first, second]);
  expect(added.map((face) => face.family)).toEqual(['Calibri', 'Carlito']);
  expect(repaint).toHaveBeenCalledTimes(1);
  fonts.dispose();
});

test('two painters share a face and one disposal does not cancel the other repaint', async () => {
  const gate = deferred<void>();
  const { added, remove } = browser(gate.promise);
  provider();
  const firstDecoder = decoder();
  const secondDecoder = decoder();
  const firstRepaint = mock(() => {});
  const secondRepaint = mock(() => {});
  const first = createDefaultFontPainter(firstDecoder, firstRepaint);
  const second = createDefaultFontPainter(secondDecoder, secondRepaint);
  const list = display({ family: 'Calibri' });
  const firstLoading = first.load(list);
  const secondLoading = second.load(list);
  await new Promise((resolve) => setTimeout(resolve, 0));
  first.dispose();
  gate.resolve();
  await Promise.all([firstLoading, secondLoading]);
  expect(added).toHaveLength(1);
  expect(firstDecoder.mock.calls.length + secondDecoder.mock.calls.length).toBe(1);
  expect(firstRepaint).not.toHaveBeenCalled();
  expect(secondRepaint).toHaveBeenCalledTimes(1);
  await second.load(list);
  expect(secondRepaint).toHaveBeenCalledTimes(1);
  second.dispose();
  expect(remove).not.toHaveBeenCalled();
  expect(added).toHaveLength(1);
});

test('a failed shared registration is evicted and can be retried', async () => {
  const { added } = browser();
  let attempts = 0;
  provider(async () => {
    if (++attempts === 1) throw new Error('Font unavailable');
    return new ArrayBuffer(4);
  });
  const repaint = mock(() => {});
  const fonts = createDefaultFontPainter(decoder(), repaint);
  const list = display({ family: 'Calibri' });
  await fonts.load(list);
  expect(added).toHaveLength(0);
  expect(repaint).not.toHaveBeenCalled();
  await fonts.load(list);
  expect(attempts).toBe(2);
  expect(added).toHaveLength(1);
  expect(repaint).toHaveBeenCalledTimes(1);
  fonts.dispose();
});

test('a live painter retries through its own decoder when the shared worker stops', async () => {
  const { added } = browser();
  provider();
  const gate = deferred<void>();
  const stoppedDecoder = mock<FontDecoder>(async () => {
    await gate.promise;
    throw new Error('Worker stopped');
  });
  const liveDecoder = decoder();
  const stoppedRepaint = mock(() => {});
  const liveRepaint = mock(() => {});
  const first = createDefaultFontPainter(stoppedDecoder, stoppedRepaint);
  const second = createDefaultFontPainter(liveDecoder, liveRepaint);
  const list = display({ family: 'Calibri' });
  const firstLoading = first.load(list);
  const secondLoading = second.load(list);
  await new Promise((resolve) => setTimeout(resolve, 0));
  first.dispose();
  gate.resolve();
  await Promise.all([firstLoading, secondLoading]);
  expect(stoppedDecoder).toHaveBeenCalledTimes(1);
  expect(liveDecoder).toHaveBeenCalledTimes(1);
  expect(added).toHaveLength(1);
  expect(stoppedRepaint).not.toHaveBeenCalled();
  expect(liveRepaint).toHaveBeenCalledTimes(1);
  second.dispose();
});

test('the in-thread editor lets the font manager decode with its own decoder', async () => {
  browser();
  const { create } = provider();
  await createDefaultFontPainter(null, () => {}).load(display({ family: 'Calibri' }));
  expect(create).toHaveBeenCalledWith(undefined, { strict: true });
});

test('no painter work happens before a worker decoder exists', async () => {
  browser();
  const { imported } = provider();
  await createDefaultFontPainter(undefined, () => {}).load(display({ family: 'Calibri' }));
  expect(imported).not.toHaveBeenCalled();
});

test('an already absent editor face is tried again after changing the CDN', async () => {
  const { added } = browser();
  const previous = getFontCdnConfiguration();
  const load = mock(async (face: BundledFontFace) =>
    new ArrayBuffer(getFontCdnConfiguration()?.baseUrl ? face.byteLength : 4));
  provider(load);
  const painter = createDefaultFontPainter(decoder(), () => {});
  const frame = display({ family: 'MS Gothic' });
  try {
    configureFontCdn({ baseUrl: false });
    await painter.load(frame);
    expect(added).toHaveLength(0);
    configureFontCdn({ baseUrl: 'https://editor-fonts.example/' });
    await painter.load(frame);
    expect(added).toHaveLength(1);
    expect(load).toHaveBeenCalledTimes(2);
  } finally { painter.dispose(); setFontCdnConfiguration(previous); }
});

test('repaints leave an uninstalled editor face settled until the CDN changes', async () => {
  const { added } = browser();
  const previous = getFontCdnConfiguration();
  const runtime = fontRuntimeWithoutPacks();
  const source = runtime.createFontProvider(undefined, { strict: true });
  const resolve = spyOn(source, 'resolve');
  const manager = spyOn(xlsxDefaultFontManager, 'load').mockResolvedValue({
    ...fonts, default: fonts, createFontProvider: () => source,
  });
  const fetch = spyOnFontFetch(fontAssetResponse);
  const repainted = deferred<void>();
  const repaint = mock(() => repainted.resolve());
  const onSettled = mock(() => {});
  const painter = createDefaultFontPainter(null, repaint, { onSettled });
  const frame = display({ family: 'Calibri' });
  try {
    configureFontCdn({ baseUrl: false });
    await painter.load(frame);
    for (let i = 0; i < 2; i++) {
      expect(painter.paint(frame, () => {})).toBe(true);
      await new Promise((done) => setTimeout(done, 0));
    }
    expect(fetch.mock.calls.filter(([input]) => !String(input).endsWith('.wasm'))).toHaveLength(0);
    expect(resolve).not.toHaveBeenCalled();
    expect(repaint).not.toHaveBeenCalled();
    expect(onSettled).not.toHaveBeenCalled();

    configureFontCdn({ baseUrl: 'https://editor-fonts.example/' });
    expect(painter.paint(frame, () => {})).toBe(false);
    await withinTestBudget(repainted.promise, 'the repaint after CDN configuration');
    expect(fetch.mock.calls.filter(([input]) => !String(input).endsWith('.wasm'))).toHaveLength(1);
    expect(resolve).toHaveBeenCalledTimes(1);
    expect(added.map(({ family }) => family)).toEqual(['Calibri']);
    expect(repaint).toHaveBeenCalledTimes(1);
  } finally {
    painter.dispose();
    fetch.mockRestore();
    manager.mockRestore();
    resolve.mockRestore();
    setFontCdnConfiguration(previous);
  }
});

test.each([
  ['a blocked explicit CDN backs off script face requests across repaints', 'http'],
  ['an explicit CDN network failure backs off script face requests across repaints', 'network'],
] as const)('%s', async (_name, failure) => {
  browser();
  const previous = getFontCdnConfiguration();
  const runtime = fontRuntimeWithoutPacks();
  const manager = spyOn(xlsxDefaultFontManager, 'load').mockResolvedValue({ ...fonts, default: fonts, ...runtime });
  const fetch = spyOnFontFetch(async () => {
    if (failure === 'network') throw new TypeError('Offline');
    return new Response(null, { status: 403 });
  });
  const now = Date.now();
  const clock = spyOn(Date, 'now').mockReturnValue(now);
  const repaint = mock(() => {});
  const painter = createDefaultFontPainter(decoder(), repaint, { remoteFonts: 'off' });
  const frame = cells(cellText('ภาษาไทย', { fontFamily: 'Aptos' }));
  try {
    configureFontCdn({ baseUrl: `https://${failure}-blocked-editor.example/` });
    await painter.load(frame);
    const initialRequests = fetch.mock.calls.length;
    expect(initialRequests).toBeGreaterThan(0);
    clock.mockReturnValue(now + 59_999);
    for (let i = 0; i < 2; i++) {
      expect(painter.paint(frame, () => {})).toBe(true);
      await new Promise((done) => setTimeout(done, 0));
    }
    expect(fetch).toHaveBeenCalledTimes(initialRequests);
    expect(repaint).not.toHaveBeenCalled();

    clock.mockReturnValue(now + 60_000);
    await painter.load(frame);
    expect(fetch).toHaveBeenCalledTimes(initialRequests * 2);

    const baseUrl = `https://${failure}-replacement-editor.example/`;
    configureFontCdn({ baseUrl });
    await painter.load(frame);
    expect(fetch.mock.calls.some(([url]) => String(url).startsWith(baseUrl))).toBe(true);
  } finally {
    painter.dispose();
    clock.mockRestore();
    fetch.mockRestore();
    manager.mockRestore();
    setFontCdnConfiguration(previous);
  }
});

test('disposing a worker during script font decode leaves its face available to another editor', async () => {
  const { added } = browser();
  provider();
  const previous = getFontCdnConfiguration();
  const pair = createInProcessPair();
  const decoding = deferred<void>();
  const decoded = deferred<Awaited<ReturnType<FontDecoder>>>();
  createSessionHost<Pick<WorkbookSessionMethods, 'open' | 'decodeFontFace'>, {}, undefined>(pair.host, {
    context: undefined,
    policies: { open: { lane: 'interactive' }, decodeFontFace: { lane: 'interactive' } },
    handlers: {
      open: () => ({ format: 'xlsx', stage: 'ready', version: 0, dirty: false, activeSheet: 0, sheets: [] }),
      decodeFontFace() { decoding.resolve(); return decoded.promise; },
    },
  });
  const session = await createWorkbookSession(new Uint8Array([80, 75]), {}, pair.client);
  const first = createDefaultFontPainter(session.call.decodeFontFace, () => {}, { remoteFonts: 'off' });
  const liveDecoder = decoder();
  const repaint = mock(() => {});
  const second = createDefaultFontPainter(liveDecoder, repaint, { remoteFonts: 'off' });
  const frame = cells(cellText('ภาษาไทย', { fontFamily: 'Aptos' }));
  try {
    configureFontCdn({ baseUrl: 'https://disposed-editor.example/' });
    const loading = first.load(frame);
    await withinTestBudget(decoding.promise, 'the worker font decode');
    first.dispose();
    await session.dispose();
    await loading;
    expect(added).toHaveLength(0);

    await second.load(frame);
    expect(liveDecoder).toHaveBeenCalledTimes(1);
    expect(added.map(({ family }) => family)).toEqual(['bo-xlsx-editor-Noto Sans Thai']);
    expect(repaint).toHaveBeenCalledTimes(1);
  } finally {
    decoded.resolve({ font: new ArrayBuffer(0) });
    first.dispose();
    second.dispose();
    await session.dispose();
    setFontCdnConfiguration(previous);
  }
});

test('families that share a script file share the explicit CDN cooldown', async () => {
  browser();
  const previous = getFontCdnConfiguration();
  const runtime = fontRuntimeWithoutPacks();
  const manager = spyOn(xlsxDefaultFontManager, 'load').mockResolvedValue({ ...fonts, default: fonts, ...runtime });
  const fetch = spyOnFontFetch(async () => new Response(null, { status: 403 }));
  const now = Date.now();
  const clock = spyOn(Date, 'now').mockReturnValue(now);
  const first = createDefaultFontPainter(decoder(), () => {}, { remoteFonts: 'off' });
  const second = createDefaultFontPainter(decoder(), () => {}, { remoteFonts: 'off' });
  const text = 'ภาษาไทย';
  const file = fonts.resolveScriptFallbackFace('thai', false, false, 'Aptos')!.file;
  const requests = () => fetch.mock.calls.filter(([url]) => String(url).endsWith(`/${file}`));
  try {
    configureFontCdn({ baseUrl: 'https://shared-blocked-editor.example/' });
    await first.load(cells(cellText(text, { fontFamily: 'Aptos' }), cellText(text, { fontFamily: 'Arial' })));
    expect(requests()).toHaveLength(1);

    clock.mockReturnValue(now + 59_999);
    first.dispose();
    await second.load(cells(cellText(text, { fontFamily: 'Calibri' }), cellText('Example', { fontFamily: 'Noto Sans Thai' })));
    expect(requests()).toHaveLength(1);

    clock.mockReturnValue(now + 60_000);
    await second.load(cells(cellText(text, { fontFamily: 'Calibri' }), cellText(text, { fontFamily: 'Open Sans' })));
    expect(requests()).toHaveLength(2);
  } finally {
    first.dispose();
    second.dispose();
    clock.mockRestore();
    fetch.mockRestore();
    manager.mockRestore();
    setFontCdnConfiguration(previous);
  }
});

test('T14: repaint, editing another face and reopening share the CDN cooldown', async () => {
  const { added } = browser();
  const previous = getFontCdnConfiguration();
  configureTestFontCdn({ baseUrl: false });
  const restore = withBrowserFontGlobals();
  const runtime = fontRuntimeWithoutPacks();
  const source = runtime.createFontProvider(undefined, { strict: true });
  const resolve = spyOn(source, 'resolve');
  const load = spyOn(xlsxDefaultFontManager, 'load').mockResolvedValue({ ...fonts, default: fonts, createFontProvider: () => source });
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const now = Date.now();
  const clock = spyOn(Date, 'now').mockReturnValue(now);
  const requests: string[] = [];
  let online = false;
  const fetch = spyOnFontFetch(async (input) => {
    if (String(input).includes('cdn.jsdelivr.net')) {
      requests.push(String(input));
      if (!online) throw new TypeError('Offline');
    }
    return fontAssetResponse(input);
  });
  const repainted = deferred<void>();
  const repaint = mock(() => repainted.resolve());
  const first = createDefaultFontPainter(null, repaint);
  const second = createDefaultFontPainter(null, () => {});
  const frame = display({ family: 'Calibri' });
  const edited = display({ family: 'Calibri' }, { family: 'Times New Roman' });
  try {
    setFontCdnConfiguration(undefined);
    await first.load(frame);
    expect(added).toHaveLength(0);
    expect(requests).toHaveLength(1);
    clock.mockReturnValue(now + 59_999);
    expect(first.paint(frame, () => {})).toBe(true);
    await new Promise((done) => setTimeout(done, 0));
    expect(requests).toHaveLength(1);
    expect(resolve).toHaveBeenCalledTimes(1);
    expect(repaint).not.toHaveBeenCalled();
    await first.load(edited);
    for (let i = 0; i < 2; i++) {
      second.paint(frame, () => {});
      await new Promise((done) => setTimeout(done, 0));
    }
    expect(requests).toHaveLength(1);
    expect(resolve).toHaveBeenCalledTimes(1);

    online = true;
    clock.mockReturnValue(now + 60_000);
    first.paint(frame, () => {});
    first.paint(frame, () => {});
    await withinTestBudget(repainted.promise, 'the repaint after CDN cooldown');
    expect(requests).toHaveLength(2);
    expect(resolve).toHaveBeenCalledTimes(2);
    expect(added.map(({ family }) => family)).toEqual(['Calibri']);
    expect(repaint).toHaveBeenCalledTimes(1);
  } finally {
    first.dispose();
    second.dispose();
    load.mockRestore();
    resolve.mockRestore();
    fetch.mockRestore();
    warn.mockRestore();
    clock.mockRestore();
    configureTestFontCdn({ baseUrl: false });
    restore();
    setFontCdnConfiguration(previous);
  }
});
