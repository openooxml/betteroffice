import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { CANVAS_PAGE_GAP_PX } from '@betteroffice/docx/layout/render';
import type { DocxCommandAck, DocxCommandEchoContext, DocxCommandEnvelope, DocxCommandFrame, ResidentCommandConnection } from '@betteroffice/docx/peerless';
import type { PeerlessEditSessionBinding, UsePeerlessEditSessionOptions } from '../hooks/usePeerlessEditSession';
import type { PeerlessFrameGeometry } from './viewStore';
import { PeerlessPagedEditor } from './PeerlessPagedEditor';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
afterEach(() => {
  cleanup();
  for (const restore of restorers.splice(0).reverse()) restore();
});
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

const identity = { session: 'viewport', incarnation: 1, epoch: 1 };
const stride = 100 + CANVAS_PAGE_GAP_PX;

type TestBitmap = ImageBitmap & { closes: number; pixels: string };
const damageRect = { x: 0, y: 0, width: 200, height: 30 };

function frame(revision: number, indices: number[], caretPage = 0): DocxCommandFrame {
  const bytes = new ArrayBuffer(88);
  const data = new DataView(bytes);
  new Uint8Array(bytes).set([0x46, 0x44, 0x56, 0x31]);
  data.setUint16(4, 1, true);
  data.setUint16(6, 80, true);
  data.setUint32(8, 88, true);
  data.setUint32(12, 1, true);
  data.setBigUint64(16, 1n, true);
  data.setBigUint64(24, 1n, true);
  data.setBigUint64(32, BigInt(revision), true);
  for (const offset of [56, 60]) data.setUint32(offset, 80, true);
  data.setUint32(64, 4, true);
  data.setUint32(68, 88, true);
  const page = { pageId: String(caretPage + 1), pageIndex: caretPage };
  const echoContext: DocxCommandEchoContext = { page, frameEpoch: revision, layoutRevision: 1, serializedBytes: 900,
    text: 'ab', selection: { anchorGrapheme: 2, focusGrapheme: 2 }, transform: [1, 0, 0, 1, 0, 0],
    styleRuns: [{ startUtf16: 0, endUtf16: 2, fontFamily: 'Arial', fontSizePx: 12, color: '#000',
      bold: false, italic: false, underline: false, strike: false, transform: [1, 0, 0, 1, 0, 0], baseline: { x: 0, y: 14 } }],
    boundaries: [0, 1, 2].map((utf16) => ({ utf16, grapheme: utf16, caret: { x: utf16 * 10, y: 0, width: 1, height: 20 } })),
    damage: [{ page, rect: damageRect, backgroundId: 'background', primitiveIds: [] }], shaping: 'simple-ltr',
  };
  const bitmap = (pixels: string) => ({ width: 200, height: 100, pixels, closes: 0,
    close(this: TestBitmap) { this.closes += 1; } }) as TestBitmap;
  return { revision, coversThroughSeq: 0, payload: {
    frame: bytes, layoutRevision: 1, echoContext,
    caret: { frameEpoch: revision, caretRect: { ...page, x: 10, y: 0, height: 20 } },
    stagedPresentation: { frameEpoch: revision, pages: indices.map((pageIndex) => ({
      page: { pageId: String(pageIndex + 1), pageIndex }, frameEpoch: revision,
      bitmap: bitmap('ab'), backgrounds: [{ id: 'background', rect: damageRect, bitmap: bitmap('') }],
    })) },
  } };
}

async function mount(zoom = 1, caretPage = 0) {
  const scroller = document.createElement('div');
  scroller.style.overflowY = 'auto';
  Object.defineProperties(scroller, { clientHeight: { value: 60 }, scrollHeight: { value: stride * 20 * zoom } });
  document.body.appendChild(scroller);
  restorers.push(() => scroller.remove());
  const rect = spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (this: Element) {
    return { x: 0, y: this === scroller ? 0 : -scroller.scrollTop,
      top: this === scroller ? 0 : -scroller.scrollTop, bottom: 60,
      left: 0, right: 200 * zoom, width: 200 * zoom, height: 60, toJSON() {} };
  });
  const pixels = new Map<HTMLCanvasElement, string>();
  const readbacks = { count: 0 };
  const drawnAfterClose: TestBitmap[] = [];
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (this: HTMLCanvasElement) {
    const canvas = this;
    return { save() {}, restore() {}, setTransform() {}, transform() {}, beginPath() {}, rect() {}, clip() {}, fillRect() {},
      clearRect() { pixels.set(canvas, ''); },
      drawImage(bitmap: TestBitmap) {
        if (bitmap.closes > 0) drawnAfterClose.push(bitmap);
        pixels.set(canvas, bitmap.pixels);
      },
      fillText(text: string) { pixels.set(canvas, text); }, measureText: (text: string) => ({ width: text.length * 10 }),
      getImageData: () => {
        readbacks.count += 1;
        return { width: canvas.width, height: canvas.height, pixels: pixels.get(canvas) ?? '' };
      },
      putImageData(image: ImageData & { pixels: string }) { pixels.set(canvas, image.pixels); },
    } as unknown as CanvasRenderingContext2D;
  } as unknown as typeof HTMLCanvasElement.prototype.getContext);
  restorers.push(() => rect.mockRestore(), () => context.mockRestore());
  let revision = 1;
  let receive: ((candidate: DocxCommandFrame) => void) | undefined;
  const built = new Set([0, 1, 2, caretPage]);
  const frames: DocxCommandFrame[] = [];
  const nextFrame = (indices: number[]) => {
    const candidate = frame(revision, indices, caretPage);
    frames.push(candidate);
    return candidate;
  };
  const requests: { start: number; end: number; built: number[]; released: number[]; rasterScale?: number }[] = [];
  const initialScales: (number | undefined)[] = [];
  const edits: DocxCommandEnvelope[] = [];
  const initial: ResidentCommandConnection['initial'] = { ...identity, version: 0, throughSeq: 0,
    outcomes: [], selection: null,
    editState: { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
      canUndo: false, canRedo: false, hyperlink: null, tableContext: null },
    frame: nextFrame([...built]),
  };
  const connection: ResidentCommandConnection & {
    requestViewport(request: { start: number; end: number; frameEpoch: number; rasterScale?: number }): Promise<void>;
  } = {
    initial, requestFullFrame: async () => initial,
    async requestViewport({ start, end, rasterScale }) {
      const entering = Array.from({ length: end - start }, (_, index) => start + index);
      const released = [...built].filter((index) => index !== caretPage && (index < start || index >= end));
      for (const index of released) built.delete(index);
      for (const index of entering) built.add(index);
      requests.push({ start, end, built: entering, released, rasterScale });
      revision += 1;
      receive?.(nextFrame([...new Set([caretPage, ...entering])]));
    },
    client: { failure: undefined, onFailure: () => () => {}, dispose: async () => {},
      onFrame(listener) { receive = listener; return () => { receive = undefined; }; },
      call: { checkpoint: async () => ({ state: new Uint8Array(), epoch: 1, version: 0, throughSeq: 0 }),
        edit(command) { edits.push(command); return new Promise<DocxCommandAck>(() => {}); } },
    },
  };
  const options: UsePeerlessEditSessionOptions = { source: Uint8Array.of(1), openConnection: async (...args: unknown[]) => {
    initialScales.push(typeof args[2] === 'function' ? (args[2] as () => number)() : undefined);
    return connection;
  },
    canvasForPage: () => null,
    prepareGeometry(candidate) {
      return { identity, coversThroughSeq: candidate.coversThroughSeq, frameEpoch: candidate.revision, layoutRevision: 1,
        pages: Array.from({ length: 20 }, (_, pageIndex): PeerlessFrameGeometry['pages'][number] => ({
          page: { pageId: String(pageIndex + 1), pageIndex },
          bounds: { x: 0, y: pageIndex * stride, width: 200, height: 100 }, transform: [1, 0, 0, 1, 0, 0],
        })), caret: candidate.payload.caret.caretRect ? { x: 10,
          y: candidate.payload.caret.caretRect.pageIndex * stride, width: 1, height: 20 } : null, selectionRects: [],
      };
    },
  };
  let binding: PeerlessEditSessionBinding | undefined;
  const view = render(<PeerlessPagedEditor sessionOptions={options} zoom={zoom} onReady={(ready) => { binding = ready; }} />, { container: scroller });
  await waitFor(() => expect(binding).toBeDefined());
  return { ...view, scroller, requests, edits, frames, pixels, readbacks, drawnAfterClose, initialScales, binding: binding!,
    setZoom(value: number) {
      view.rerender(<PeerlessPagedEditor sessionOptions={options} zoom={value} />);
    },
    async scrollTo(pageIndex: number) {
      const before = requests.length;
      act(() => { scroller.scrollTop = (pageIndex * stride + 1) * zoom; fireEvent.scroll(scroller); });
      await waitFor(() => expect(requests.slice(before).some(({ built }) => built.includes(pageIndex))).toBe(true));
    },
  };
}

test('raster scale starts at device pixels and settles zoom and rearmed DPR changes', async () => {
  const originalDpr = Object.getOwnPropertyDescriptor(window, 'devicePixelRatio');
  Object.defineProperty(window, 'devicePixelRatio', { configurable: true, value: 2 });
  restorers.push(() => {
    if (originalDpr) Object.defineProperty(window, 'devicePixelRatio', originalDpr);
    else Reflect.deleteProperty(window, 'devicePixelRatio');
  });
  const media: { query: string; listener?: () => void }[] = [];
  const matchMedia = spyOn(window, 'matchMedia').mockImplementation((query) => {
    const item: typeof media[number] = { query };
    media.push(item);
    return { matches: true, media: query,
      addEventListener(_type: string, listener: () => void) { item.listener = listener; },
      removeEventListener() { item.listener = undefined; },
    } as unknown as MediaQueryList;
  });
  restorers.push(() => matchMedia.mockRestore());
  const editor = await mount(1.5);
  expect(editor.initialScales).toEqual([3]);
  expect(media.at(-1)?.query).toBe('(resolution: 2dppx)');
  let now = 0;
  let timerId = 0;
  const pending = new Map<number, { at: number; run: () => void }>();
  const timeout = spyOn(globalThis, 'setTimeout').mockImplementation(((run: () => void, delay: number) => {
    const id = ++timerId;
    pending.set(id, { at: now + delay, run });
    return id as unknown as ReturnType<typeof setTimeout>;
  }) as typeof setTimeout);
  const cancel = spyOn(globalThis, 'clearTimeout').mockImplementation(((id: number) => { pending.delete(id); }) as typeof clearTimeout);
  restorers.push(() => timeout.mockRestore(), () => cancel.mockRestore());
  const advance = async (ms: number) => {
    now += ms;
    await act(async () => {
      for (const [id, timer] of [...pending]) if (timer.at <= now) { pending.delete(id); timer.run(); }
    });
  };
  const scales = () => editor.requests.filter(({ rasterScale }) => rasterScale !== undefined).map(({ rasterScale }) => rasterScale);
  const before = scales().length;
  const canvas = editor.container.querySelector<HTMLCanvasElement>('[data-page-id="1"] canvas')!;
  const pixels = editor.pixels.get(canvas);
  const caret = editor.binding.viewStore.getSnapshot().geometry!.caret;
  act(() => editor.setZoom(1.25));
  await advance(100);
  act(() => editor.setZoom(2));
  expect(editor.container.querySelector('div[style*="transform: scale"]')?.parentElement?.style.width).toBe('400px');
  await advance(149);
  expect(scales().slice(before)).toEqual([]);
  expect(canvas.width).toBe(200);
  expect(editor.pixels.get(canvas)).toBe(pixels);
  expect(editor.binding.viewStore.getSnapshot().geometry!.caret).toEqual(caret);
  await advance(1);
  expect(scales().slice(before)).toEqual([4]);
  Object.defineProperty(window, 'devicePixelRatio', { configurable: true, value: 1 });
  act(() => media.at(-1)?.listener?.());
  expect(media.at(-1)?.query).toBe('(resolution: 1dppx)');
  await advance(100);
  Object.defineProperty(window, 'devicePixelRatio', { configurable: true, value: 1.5 });
  act(() => media.at(-1)?.listener?.());
  expect(media.at(-1)?.query).toBe('(resolution: 1.5dppx)');
  await advance(149);
  expect(scales().slice(before)).toEqual([4]);
  await advance(1);
  expect(scales().slice(before)).toEqual([4, 3]);
  act(() => editor.setZoom(1.75));
  editor.unmount();
  await advance(150);
  expect(scales().slice(before)).toEqual([4, 3]);
  expect(media.every(({ listener }) => listener === undefined)).toBe(true);
});

for (const zoom of [0, -1, Number.NaN]) {
  test(`raster scale falls back to device pixels for zoom ${zoom}`, async () => {
    const originalDpr = Object.getOwnPropertyDescriptor(window, 'devicePixelRatio');
    Object.defineProperty(window, 'devicePixelRatio', { configurable: true, value: 2 });
    restorers.push(() => {
      if (originalDpr) Object.defineProperty(window, 'devicePixelRatio', originalDpr);
      else Reflect.deleteProperty(window, 'devicePixelRatio');
    });
    const editor = await mount(zoom);
    expect(editor.initialScales).toEqual([2]);
  });
}

test('peerless_page_offsets_are_monotonic_and_non_overlapping', () => {
  const source = readFileSync(`${import.meta.dir}/PeerlessDocxEditor.tsx`, 'utf8');
  const start = source.indexOf('    if (!identity.current)', source.indexOf('const prepareGeometry ='));
  const end = source.indexOf('\n  }, []);', start);
  const pages = [100, 220, 80].map((height, pageIndex) => ({ pageId: BigInt(pageIndex + 1), pageIndex,
    page: { width: pageIndex === 1 ? 400 : 200, height } }));
  const dependencies = { identity: { current: identity }, retained: { current: null }, prepared: { current: new WeakMap() },
    pageIndexes: { current: new Map() }, applyFrameDelta: () => ({ frameEpoch: 1, pages }), decodeFrameDelta: () => ({}),
    residentCaretSnapshotForFrame: () => ({ caretRect: { pageIndex: 2, x: 10, y: 7, height: 20 } }), CANVAS_PAGE_GAP_PX,
  };
  const body = new Bun.Transpiler({ loader: 'tsx' }).transformSync(`function prepare(frame) {${source.slice(start, end)} }`);
  const prepare = new Function(...Object.keys(dependencies), `${body}\nreturn prepare;`)(...Object.values(dependencies)) as
    (candidate: DocxCommandFrame) => PeerlessFrameGeometry;
  const geometry = prepare(frame(1, []));
  expect(geometry.pages.map(({ bounds }) => bounds.y)).toEqual([0, 100 + CANVAS_PAGE_GAP_PX, 320 + 2 * CANVAS_PAGE_GAP_PX]);
  for (let index = 1; index < geometry.pages.length; index += 1) {
    const previous = geometry.pages[index - 1]!.bounds;
    expect(geometry.pages[index]!.bounds.y).toBeGreaterThanOrEqual(previous.y + previous.height + CANVAS_PAGE_GAP_PX);
  }
  expect(geometry.caret?.y).toBe(327 + 2 * CANVAS_PAGE_GAP_PX);
  expect(geometry.pages.map(({ bounds }) => bounds.x)).toEqual([100, 0, 100]);
  expect(geometry.caret?.x).toBe(110);
});

test('scrolling_to_page_n_builds_n_and_releases_far_pages', async () => {
  const editor = await mount(1.5);
  const farCanvas = editor.container.querySelector<HTMLCanvasElement>('[data-page-id="3"] canvas');
  expect(farCanvas?.width).toBe(200);
  await editor.scrollTo(10);
  expect(editor.requests.at(-1)?.released).toContain(2);
  expect(farCanvas?.width).toBe(0);
  expect(editor.container.querySelector('[data-page-id="11"] canvas')?.getAttribute('width')).toBe('200');
  expect(editor.container.querySelector('[data-page-id="3"] canvas')?.getAttribute('width') ?? '0').toBe('0');
  expect(editor.container.querySelector('div[style*="transform: scale"]')?.parentElement?.style.height).toBe(`${(19 * stride + 100) * 1.5}px`);
  await editor.scrollTo(2);
  expect(editor.requests.at(-1)?.released).toContain(10);
  expect(editor.container.querySelector('[data-page-id="3"] canvas')?.getAttribute('width')).toBe('200');
  expect(editor.container.querySelector('[data-page-id="3"] canvas')).not.toBe(farCanvas);
});

test('scroll_during_pending_typing_keeps_caret_page_staged', async () => {
  const editor = await mount(1, 4);
  const caretCanvas = editor.container.querySelector<HTMLCanvasElement>('[data-page-id="5"] canvas')!;
  const input = editor.getByRole('textbox', { name: 'Document input' });
  act(() => input.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'insertText', data: 'x' })));
  await waitFor(() => expect(editor.edits).toHaveLength(1));
  expect(editor.binding.admission.getSnapshot().originalUnackedIntents).toBe(1);
  expect(editor.pixels.get(caretCanvas)).toBe('abx');
  await editor.scrollTo(15);
  expect(editor.requests.at(-1)?.released).not.toContain(4);
  expect(editor.frames.at(-1)?.payload.stagedPresentation?.pages.map(({ page }) => page.pageIndex)).toContain(4);
  expect(editor.container.querySelector('[data-page-id="5"] canvas')).toBe(caretCanvas);
  expect(caretCanvas.width).toBe(200);
  expect(editor.pixels.get(caretCanvas)).toBe('abx');
  expect(editor.container.querySelector('[data-page-id="16"] canvas')?.getAttribute('width')).toBe('200');
  expect(editor.binding.admission.getSnapshot().originalUnackedIntents).toBe(1);
  expect(editor.edits.map(({ kind }) => kind)).toEqual(['insertText']);
});

test('evicted_pages_are_not_recreated_by_pruning_and_bitmaps_are_closed', async () => {
  const editor = await mount();
  const staged = editor.frames.at(-1)!.payload.stagedPresentation!.pages.find(({ page }) => page.pageIndex === 2)!;
  const evicted = editor.container.querySelector<HTMLCanvasElement>('[data-page-id="3"] canvas')!;
  await editor.scrollTo(15);
  expect(evicted.width).toBe(0);
  expect((staged.bitmap as TestBitmap).closes).toBe(1);
  expect((staged.backgrounds![0]!.bitmap as TestBitmap).closes).toBe(1);
  const allocations = spyOn(document, 'createElement');
  try {
    act(() => editor.binding.presentation.hooks.paint?.([]));
    expect(allocations.mock.calls.filter(([tag]) => tag === 'canvas')).toHaveLength(0);
  } finally { allocations.mockRestore(); }
  await editor.scrollTo(2);
  expect(editor.container.querySelector('[data-page-id="3"] canvas')).not.toBe(evicted);
  expect((staged.bitmap as TestBitmap).closes).toBe(1);
  expect(editor.drawnAfterClose).toEqual([]);
  expect(editor.readbacks.count).toBe(0);
});

test('worker page column centers at the scaled width', async () => {
  const editor = await mount(1.5);
  const column = editor.container.querySelector('div[style*="transform: scale"]')!.parentElement!;
  expect(column.style.width).toBe('300px');
  expect(column.style.marginLeft).toBe('auto');
  expect(column.style.marginRight).toBe('auto');
});
