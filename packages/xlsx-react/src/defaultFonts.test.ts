import { afterEach, expect, mock, spyOn, test } from 'bun:test';
import type { BundledFontFace } from '@betteroffice/fonts';
import type { FontDecoder } from '../../../shared/office-session/fontDecoder';
import { createSessionHost } from '../../../shared/office-session/host';
import { createInProcessPair } from '../../../shared/office-session/testing/inProcessTransport';
import type { DisplayList, TextCmd } from '@betteroffice/xlsx';
import { createDefaultFontPainter, xlsxDefaultFontManager } from './defaultFonts';
import { createWorkbookSession } from '../../xlsx/src/session/client';
import type { WorkbookSessionMethods } from '../../xlsx/src/session/methods';

const descriptors = ['document', 'FontFace'].map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const);
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

function browser(gate: Promise<void> = Promise.resolve()) {
  const added: Face[] = [];
  class Face {
    constructor(readonly family: string, readonly bytes: ArrayBuffer, readonly descriptors: FontFaceDescriptors) {}
    async load() { await gate; return this; }
  }
  const remove = mock((face: Face) => { added.splice(added.indexOf(face), 1); });
  Object.defineProperty(globalThis, 'FontFace', { configurable: true, value: Face });
  Object.defineProperty(globalThis, 'document', { configurable: true,
    value: { fonts: { add: (face: Face) => added.push(face), delete: remove } } });
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
    family === 'Calibri' ? face('Carlito', bold, italic) : family === 'MS Gothic' ? face('Noto Sans JP', bold, italic) : undefined);
  const bundled = mock((family: string, bold: boolean, italic: boolean) =>
    family === 'Carlito' || family === 'Noto Sans Thai' ? face(family, bold, italic) : undefined);
  const create = mock((_options?: unknown, runtime?: { decoder?: FontDecoder }) => {
    const resolve = (resolveFace: typeof metric) => (family: string, bold: boolean, italic: boolean) => {
      const resolved = resolveFace(family, bold, italic);
      return resolved ? async () => {
        const decoded = await runtime!.decoder!(new Uint8Array([1]), { codec: 'zstd' });
        expect(new Uint8Array(decoded.font)).toEqual(new Uint8Array([0, 1, 0, 0]));
        return bytes(resolved);
      } : undefined;
    };
    return { resolve: resolve(metric), resolveFamily: resolve(bundled) };
  });
  const module = {
    resolveMetricCompatFace: metric, resolveBundledFamilyFace: bundled,
    createFontProvider: create,
  } as unknown as Awaited<ReturnType<typeof xlsxDefaultFontManager.load>>;
  const imported = spyOn(xlsxDefaultFontManager, 'load').mockResolvedValue(module);
  restorers.push(() => imported.mockRestore());
  return { imported, create, metric, bundled, module };
}

function decoder() { return mock<FontDecoder>(async () => ({ font: new Uint8Array([0, 1, 0, 0]).buffer })); }

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
  expect(source.create).toHaveBeenCalledWith(undefined, { decoder: decode });
  expect(decode).toHaveBeenCalledTimes(2);
  expect(repaint).toHaveBeenCalledTimes(1);
  await fonts.load(list);
  expect(repaint).toHaveBeenCalledTimes(1);
  fonts.dispose();
  expect(added).toHaveLength(2);
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

test('a frame painted before a face fails still repaints once as settled', async () => {
  const { added } = browser();
  provider(async () => { throw new Error('Missing optional font pack'); });
  const repainted = deferred<void>();
  const repaint = mock(() => repainted.resolve());
  const fonts = createDefaultFontPainter(decoder(), repaint);
  const list = display({ family: 'Noto Sans Thai' }, { family: 'Unknown Face' });
  expect(fonts.paint(list, () => {})).toBe(false);
  await repainted.promise;
  expect(added).toHaveLength(0);
  expect(fonts.paint(list, () => {})).toBe(true);
  await Promise.resolve();
  expect(repaint).toHaveBeenCalledTimes(1);
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

test('a missing CJK pack cannot install its Latin fallback under MS Gothic', async () => {
  const { added } = browser();
  const fallback = mock(async () => new ArrayBuffer(4));
  provider(fallback);
  const repaint = mock(() => {});
  await createDefaultFontPainter(decoder(), repaint).load(display({ family: 'MS Gothic' }));
  await createDefaultFontPainter(decoder(), repaint).load(display({ family: 'MS Gothic' }));
  expect(fallback).toHaveBeenCalledTimes(1);
  expect(added).toHaveLength(0);
  expect(repaint).not.toHaveBeenCalled();
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
  expect(create).toHaveBeenCalledWith(undefined, undefined);
});

test('no painter work happens before a worker decoder exists', async () => {
  browser();
  const { imported } = provider();
  await createDefaultFontPainter(undefined, () => {}).load(display({ family: 'Calibri' }));
  expect(imported).not.toHaveBeenCalled();
});
