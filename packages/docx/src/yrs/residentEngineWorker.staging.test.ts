import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { inflateSync } from 'node:zlib';
import { decodeFrameDelta, FRAME_DELTA_VERSION, retainedFramePageById, type DecodedFrameDelta, type RetainedFrame } from '../layout/render/frameDelta';
import { buildCommandFrame } from './residentCommandFrame';
import { closeCommandBitmap, echoBackgroundPage, stageCommandPresentation } from './residentCommandPresentation';
import type { DocxCommandEchoContext, DocxCommandFrame, DocxEchoGlyphSet } from './residentCommandProtocol';
import type { DisplayPrimitive, GlyphRunPrimitive } from '../layout/render/displayList';
import { drawGlyphRunPrimitive } from '../layout/render/canvasBackend';
import { GlyphCache } from '../layout/render/glyphCache';
import { indexEchoGlyphSet, leadingWordWidth, placeEchoGlyphs } from './residentEchoGlyphs';
import { readCellSelection } from './residentCommandTableSelection';
import { readResidentCellSelectionRects } from './residentObjectGeometry';

const echoAlphabet = 'abxy';
const echoGlyphSet: DocxEchoGlyphSet = { key: 'set', fontId: 1, alphabet: echoAlphabet, horizontal: 0.01,
  glyphIds: [...echoAlphabet].map((char) => char.charCodeAt(0)), advances: [1000, 1000, 1000, 1000],
  kerning: [], unsafePairs: [], outlines: Object.fromEntries([...echoAlphabet].map((char) => [char.charCodeAt(0),
    JSON.stringify({ upem: 1000, cmds: [{ t: 'M', x: char.charCodeAt(0), y: 0 }, { t: 'L', x: 0, y: 700 }, { t: 'Z' }] })])) };

test('package_image_fixture_has_complete_colored_scanlines', () => {
  const png = Buffer.from(readFileSync(new URL('./__fixtures__/seed-parity/images/word/media/image1.png.base64', import.meta.url), 'utf8'), 'base64');
  expect(png.subarray(0, 8)).toEqual(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
  expect([png[24], png[25], png[28]]).toEqual([8, 2, 0]);
  const width = png.readUInt32BE(16);
  const height = png.readUInt32BE(20);
  const chunks: Buffer[] = [];
  for (let offset = 8; offset < png.length;) {
    const size = png.readUInt32BE(offset);
    if (png.toString('ascii', offset + 4, offset + 8) === 'IDAT') chunks.push(png.subarray(offset + 8, offset + 8 + size));
    offset += size + 12;
  }
  const pixels = inflateSync(Buffer.concat(chunks));
  expect(pixels).toHaveLength(height * (1 + width * 3));
  for (let row = 0; row < height; row += 1) {
    const offset = row * (1 + width * 3);
    expect(pixels[offset]).toBe(0);
    expect(pixels.subarray(offset + 1, offset + 1 + width * 3)).toEqual(Buffer.from([255, 0, 0, 255, 0, 0]));
  }
});

async function stageRequestedPages(requested: number[], damaged: number[] = [], echoContext?: DocxCommandEchoContext, imageSource?: string,
  primitives?: DisplayPrimitive[]) {
  const pages = Array.from({ length: 4 }, (_, pageIndex) => ({ pageIndex, width: 600, height: 800,
    primitives: primitives ?? (imageSource ? [{ kind: 'image' as const, relId: imageSource, x: 0, y: 0, w: 10, h: 10 }] : []) }));
  const retainedFrame: RetainedFrame = {
    protocolVersion: FRAME_DELTA_VERSION, docEpoch: 1, layoutEpoch: 1, frameEpoch: 7,
    pages: pages.map((page) => ({ page, pageIndex: page.pageIndex, pageId: BigInt(page.pageIndex + 1),
      fingerprint: 0n, primitiveIds: new BigUint64Array(page.primitives.map((_, index) => BigInt(index + 1))) })),
    damagedPageIds: new Set(damaged.map((index) => BigInt(index + 1))), removedPageIds: new Set(), displayList: { pages },
  };
  const bytes = new Uint8Array();
  const delta: DecodedFrameDelta = {
    protocolVersion: FRAME_DELTA_VERSION, docEpoch: 1, layoutEpoch: 1, frameEpoch: 7,
    full: false, baseFrameEpoch: 7, pageCount: pages.length, operations: [], bytes,
  };
  const builds: number[][] = [];
  const rasters: number[] = [];
  const rasterCalls: { primitives: number; clip: unknown }[] = [];
  const decoded: Blob[] = [];
  const resolved: unknown[] = [];
  const mediaReads: string[] = [];
  let closes = 0;
  const fetches: string[] = [];
  const image = { close() { closes += 1; } } as ImageBitmap;
  const bitmap = { width: 600, height: 800, close() {} } as ImageBitmap;
  const background = { width: 100, height: 20, close() {} } as ImageBitmap;
  const source = readFileSync(new URL('./residentEngineWorker.ts', import.meta.url), 'utf8');
  const method = source.slice(source.indexOf('      async frame({ coversThroughSeq, full }) {'), source.indexOf('      snapshot: () => engine.encodeState(),'));
  const pending = new Set(damaged.map((index) => String(index + 1)));
  const dependencies = {
    changes: { pages: new Set(requested), updates: [] },
    engine: {
      partGeometry: () => [],
      mediaDataUrl(token: string) {
        mediaReads.push(token);
        return token === 'media:0' ? 'data:image/png;base64,' + readFileSync(new URL('./__fixtures__/seed-parity/images/word/media/image1.png.base64', import.meta.url), 'utf8').replace(/\s/g, '') : null;
      },
      cellSelection: () => null,
      setDisplayWindow() {}, retainedHeadersFootersJson: () => 'null', proposalEngine: { version: () => '0' },
      buildDisplayPagesFrame: (indices: number[]) => { builds.push([...indices]); return bytes; },
      buildDisplayListFrame: () => bytes,
    },
    commandDisplayWindow: null, layoutDirty: false, revision: 7, layoutPreviewVersion: 0,
    readResidentFrameGeometry: async () => undefined,
    init: { extras: '{}' }, retainedFrame, layoutRevision: 0,
    fonts: { snapshot: undefined },
    finalizeResidentFrame: async () => ({ bytes, delta, caret: { frameEpoch: 7, caretRect: null } }),
    hooks: { editState: () => ({ storedMarks: null }), selection: () => null },
    projectCommandEchoContext: () => echoContext,
    commandEchoGlyphs: () => ({ provider: () => undefined,
      unsent: (context?: DocxCommandEchoContext) => context?.glyphs ? [echoGlyphSet] : [], sent() {} }),
    pendingOffscreenPageIds: pending, activeOffscreenPageIds: new Set(['1', '2', '3', '4']),
    glyphCache: {}, offscreenDpr: 1, offscreenZoom: 1,
    retainedPageByKey: (frame: RetainedFrame, pageId: string) => retainedFramePageById(frame, BigInt(pageId)),
    readCellSelection, readResidentCellSelectionRects,
    stageCommandPresentation, echoBackgroundPage, buildCommandFrame, pendingUpdates: [],
    OffscreenCanvas: class {
      constructor(public width: number, public height: number) {}
      getContext() { return { drawImage() {} }; }
      transferToImageBitmap() { return this.width === 1 ? bitmap : background; }
    },
    fetch: async (source: string) => { fetches.push(source); return fetch(source); },
    createImageBitmap: async (blob: Blob) => { decoded.push(blob); return image; },
    rasterizeDisplayPageToBackBuffer: async (_buffer: unknown, page: RetainedFrame['displayList']['pages'][number], options: { resolveImage(source: string): Promise<unknown> },
      _pixelRatio?: number, _zoom?: number, clip?: unknown) => {
      rasters.push(page.pageIndex);
      rasterCalls.push({ primitives: page.primitives.length, clip });
      if (imageSource) resolved.push(await options.resolveImage(imageSource));
    },
  };
  const frame = await new Function(...Object.keys(dependencies),
    new Bun.Transpiler({ loader: 'ts' }).transformSync('const frameHooks = {\n' + method + '\n};') +
    '\nreturn frameHooks.frame({ coversThroughSeq: 9, full: false });'
  )(...Object.values(dependencies)) as { value: DocxCommandFrame; transfer: Transferable[] };
  return { frame, builds, rasters, rasterCalls, bitmap, background, pending, delta, decoded, resolved, mediaReads, image, closes, fetches };
}

test('remount_request_with_empty_native_delta_stages_requested_page', async () => {
  const result = await stageRequestedPages([2]);
  expect(result.delta.operations).toEqual([]);
  expect(result.builds).toEqual([[2]]);
  expect(result.frame.value.payload?.stagedPresentation?.pages).toEqual([
    { page: { pageId: '3', pageIndex: 2 }, frameEpoch: 7, bitmap: result.bitmap },
  ]);
  expect(result.frame.transfer).toContain(result.bitmap);
  expect(result.frame.value.coversThroughSeq).toBe(9);
  expect(result.frame.value.baseRevision).toBe(7);
  const damaged = await stageRequestedPages([2, 2], [2]);
  expect(damaged.rasters).toEqual([2]);
  expect(damaged.pending.size).toBe(0);
});

test('unrequested_undamaged_pages_stay_unstaged', async () => {
  const result = await stageRequestedPages([1], [3]);
  expect(result.frame.value.payload?.stagedPresentation?.pages.map(({ page }) => page.pageIndex)).toEqual([1, 3]);
  expect(result.rasters).toEqual([1, 3]);
  expect(result.frame.value.payload?.stagedPresentation?.pages.some(({ page }) => page.pageIndex === 0 || page.pageIndex === 2)).toBe(false);
  expect(result.pending.size).toBe(0);
});

test('remount_request_with_pending_echo_restages_context_and_background', async () => {
  const page = { pageId: '3', pageIndex: 2 };
  const rect = { x: 0, y: 0, width: 100, height: 20 };
  const echoContext: DocxCommandEchoContext = {
    page, frameEpoch: 7, layoutRevision: 0, serializedBytes: 900, text: 'ab',
    styleRuns: [{ startUtf16: 0, endUtf16: 2, fontFamily: 'Arial', fontSizePx: 12,
      color: '#000', bold: false, italic: false, underline: false, strike: false,
      transform: [1, 0, 0, 1, 0, 0], baseline: { x: 0, y: 14 } }],
    boundaries: [0, 1, 2].map((utf16) => ({ utf16, grapheme: utf16,
      caret: { x: utf16 * 10, y: 0, width: 1, height: 20 } })),
    selection: { anchorGrapheme: 2, focusGrapheme: 2 }, transform: [1, 0, 0, 1, 0, 0],
    damage: [{ page, rect, backgroundId: 'background', omittedPrimitiveIds: [] }], shaping: 'simple-ltr',
    glyphs: { glyphSet: 'set', fontId: 1, size: 12, color: '#000', x: 0, y: 14, next: '', right: 100,
      device: { x: 0, y: 0, width: 100, height: 20 }, pixelRatio: 1 },
  };
  const result = await stageRequestedPages([2], [], echoContext);
  expect(result.delta.operations).toEqual([]);
  expect(result.frame.value.payload.echoContext).toEqual(echoContext);
  expect(result.frame.value.payload.echoGlyphSets).toEqual([echoGlyphSet]);
  expect(result.frame.value.payload.stagedPresentation?.pages[0]?.backgrounds).toEqual([
    { id: 'background', rect, bitmap: result.background },
  ]);
  expect(result.frame.transfer).toContain(result.background);
  expect(result.rasters).toEqual([2, 2]);
  const texts: string[] = [];
  const scope = globalThis as { Path2D?: unknown };
  const previousPath = scope.Path2D;
  scope.Path2D = class { glyph?: number; moveTo(x: number) { this.glyph ??= x; } lineTo() {} closePath() {} };
  const createCanvas = () => ({ width: 600, height: 800, getContext: () => ({
    save() {}, restore() {}, setTransform() {}, transform() {}, beginPath() {}, rect() {}, clip() {},
    clearRect() {}, fillRect() {}, translate() {}, scale() {},
    drawImage() { texts.push(''); },
    fill(path: { glyph: number }) { texts.push(texts.pop()! + String.fromCharCode(path.glyph)); },
  }) }) as unknown as HTMLCanvasElement;
  let canvas = createCanvas();
  const peerless = (name: string) => readFileSync(new URL(`../../../docx-react/src/components/DocxEditor/peerless/${name}`, import.meta.url), 'utf8');
  const ownershipSource = peerless('bitmapOwnership.ts').replace(/^import .*$/gm, '').replaceAll('export function ', 'function ');
  const ownership = new Function('closeCommandBitmap',
    new Bun.Transpiler({ loader: 'ts' }).transformSync(ownershipSource) + '\nreturn { retainBitmap, releaseBitmap };'
  )(closeCommandBitmap) as Record<'retainBitmap' | 'releaseBitmap', (bitmap: ImageBitmap, holder: object) => void>;
  const deltasSource = peerless('frameDeltas.ts').replace(/^import .*$/gm, '').replaceAll('export function ', 'function ');
  const { decodeCommandFrameDeltas } = new Function('decodeFrameDelta',
    new Bun.Transpiler({ loader: 'ts' }).transformSync(deltasSource) + '\nreturn { decodeCommandFrameDeltas };'
  )(decodeFrameDelta) as { decodeCommandFrameDeltas: unknown };
  const source = peerless('echoPainter.ts');
  const painterSource = source.slice(source.indexOf('export function createEchoPainter('))
    .replace('export function createEchoPainter(', 'function createEchoPainter(');
  const painter = new Function('decodeCommandFrameDeltas', 'canvasForPage', 'retainBitmap', 'releaseBitmap', 'drawGlyphRunPrimitive',
    'GlyphCache', 'indexEchoGlyphSet', 'leadingWordWidth', 'placeEchoGlyphs',
    new Bun.Transpiler({ loader: 'ts' }).transformSync(painterSource) + '\nreturn createEchoPainter({ canvasForPage });'
  )(decodeCommandFrameDeltas, () => canvas, ownership.retainBitmap, ownership.releaseBitmap, drawGlyphRunPrimitive, GlyphCache,
    indexEchoGlyphSet, leadingWordWidth, placeEchoGlyphs) as {
    preview(intent: { kind: 'insertText'; payload: { text: string } }): object | undefined;
    reconcile(frame: DocxCommandFrame, retired: readonly unknown[], remaining: readonly unknown[]): void;
    paint(entries: readonly unknown[]): void;
    dispose(): void;
  };
  try {
    painter.reconcile({ revision: 6, coversThroughSeq: 9, payload: {
      frame: new ArrayBuffer(0), layoutRevision: 0, caret: { frameEpoch: 6, caretRect: null },
      echoContext: { ...echoContext, frameEpoch: 6 }, echoGlyphSets: [echoGlyphSet],
      stagedPresentation: { frameEpoch: 6, pages: [{ page, frameEpoch: 6,
        bitmap: { width: 600, height: 800, close() {} } as ImageBitmap,
        backgrounds: [{ id: 'background', rect, bitmap: { width: 100, height: 20, close() {} } as ImageBitmap }],
      }] },
    } }, [], []);
    const value = painter.preview({ kind: 'insertText', payload: { text: 'x' } });
    expect(value).toBeDefined();
    if (!value) throw new Error('Echo preview is unavailable');
    const pending = { seq: 10, value,
      ticket: { id: 10, seq: 10, ack: Promise.resolve({ seq: 10, status: 'applied' }),
        presented: Promise.resolve('presented') } };
    painter.paint([pending]);
    expect(texts.at(-1)).toBe('abx');
    texts.length = 0;
    canvas = createCanvas();
    painter.reconcile(result.frame.value, [], [pending]);
    expect(result.frame.value.coversThroughSeq).toBe(9);
    expect(texts.at(-1)).toBe('abx');
    expect(painter.preview({ kind: 'insertText', payload: { text: 'y' } })).toBeDefined();
  } finally {
    painter.dispose();
    scope.Path2D = previousPath;
  }
});

test('edited_image_page_resolves_package_media_for_authoritative_pixels', async () => {
  const result = await stageRequestedPages([2], [2], undefined, 'media:0');
  expect(result.frame.value.coversThroughSeq).toBe(9);
  expect(result.resolved).toEqual([result.image]);
  expect(result.mediaReads).toEqual(['media:0']);
  expect(result.decoded[0]?.type).toBe('image/png');
  expect(result.closes).toBe(1);
  expect(new Uint8Array(await result.decoded[0]!.arrayBuffer())).toEqual(
    Uint8Array.from(Buffer.from(readFileSync(new URL('./__fixtures__/seed-parity/images/word/media/image1.png.base64', import.meta.url), 'utf8'), 'base64'))
  );
});

test('package_media_is_decoded_once_per_frame_and_closed_after_rasterization', async () => {
  const result = await stageRequestedPages([1, 2], [1, 2], undefined, 'media:0');
  expect(result.resolved).toEqual([result.image, result.image]);
  expect(result.mediaReads).toEqual(['media:0']);
  expect(result.decoded).toHaveLength(1);
  expect(result.closes).toBe(1);
});

test('unavailable_media_and_external_sources_do_not_fetch', async () => {
  for (const source of ['media:99', 'media:01', 'https://example.invalid/image.png']) {
    const result = await stageRequestedPages([2], [2], undefined, source);
    expect(result.resolved).toEqual([null]);
    expect(result.fetches).toEqual([]);
    expect(result.decoded).toEqual([]);
  }
});

test('the echo background rasters what reaches its damage rect, not the page', async () => {
  const page = { pageId: '3', pageIndex: 2 };
  const rect = { x: 0, y: 0, width: 100, height: 20 };
  const run = (y: number): GlyphRunPrimitive => ({ kind: 'glyphRun', text: 'a', size: 12, fontId: 1, color: '#000',
    glyphs: [{ id: 97, cluster: 0, x: 10, y, advance: 6 }] });
  const echoContext: DocxCommandEchoContext = {
    page, frameEpoch: 7, layoutRevision: 0, serializedBytes: 900, text: 'ab',
    styleRuns: [{ startUtf16: 0, endUtf16: 2, fontFamily: 'Arial', fontSizePx: 12,
      color: '#000', bold: false, italic: false, underline: false, strike: false,
      transform: [1, 0, 0, 1, 0, 0], baseline: { x: 0, y: 14 } }],
    boundaries: [0, 1, 2].map((utf16) => ({ utf16, grapheme: utf16, caret: { x: utf16 * 10, y: 0, width: 1, height: 20 } })),
    selection: { anchorGrapheme: 2, focusGrapheme: 2 }, transform: [1, 0, 0, 1, 0, 0],
    damage: [{ page, rect, backgroundId: 'background', omittedPrimitiveIds: [] }], shaping: 'simple-ltr',
    glyphs: { glyphSet: 'set', fontId: 1, size: 12, color: '#000', x: 0, y: 14, next: '', right: 100,
      device: { x: 0, y: 0, width: 100, height: 20 }, pixelRatio: 1 },
  };
  const result = await stageRequestedPages([2], [], echoContext, undefined,
    [run(14), ...Array.from({ length: 200 }, (_, index) => run(300 + index))]);
  expect(result.rasterCalls.map(({ primitives }) => primitives)).toEqual([201, 1]);
  expect(result.rasterCalls[1]?.clip).toEqual({ x: 0, y: 0, width: 100, height: 20 });
});
