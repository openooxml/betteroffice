import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, describe, expect, it, mock, spyOn } from 'bun:test';
import { readFile } from 'node:fs/promises';
import * as pptx from '@betteroffice/pptx';
import type { PresentationFrame, PresentationSession, SlideDisplayList } from '@betteroffice/pptx';
import { createPresentationSession } from '../../../pptx/src/session/client';
import { createPresentationSessionHost } from '../../../pptx/src/session/host';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { PptxEditor } from '../PptxEditor';
import type { PptxWorkerViewerApi } from '../PptxEditor';
import { EditorToolbar, ToolbarCommandButton } from '../index';
import { presentationSessionOpener, ViewerSession } from './useSessionPresentation';
import { frameImages } from './sessionPaint';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const file = new Uint8Array([1, 2, 3]);
const restorers: (() => void)[] = [];
let observed: HTMLElement[] = [];
let visibility: IntersectionObserverCallback;
let observerRoot: Element | Document | null | undefined;
let painted: { canvas: HTMLCanvasElement; list: SlideDisplayList; scale: number }[];

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

function frame(index: number): PresentationFrame {
  return { slideIndex: index, version: 'v1', epoch: index + 1, media: new Map(),
    displayList: { contractVersion: 1, width: 960 + index, height: 540, primitives: [] } };
}

function session(count = 4) {
  const slides = Array.from({ length: count }, (_, index) => ({
    id: `s${index}`, index, name: null, layoutPartPath: null,
  }));
  const call = {
    frame: mock(async (index: number) => frame(index)),
    version: mock(async () => 'v1'),
    readContent: mock(async (request?: pptx.PptxReadRequest): Promise<pptx.PptxReadResult> => ({
      ok: true, version: 'v1', stories: [], slides: slides
        .filter((slide) => !request?.slideIds || request.slideIds.includes(slide.id))
        .map((slide) => ({ ...slide, sourcePartPath: null, shapes: [], notes: `Notes ${slide.index}` })),
    })),
    findText: mock(async (): Promise<pptx.PptxFindResult> => ({ ok: true, version: 'v1', matches: [], truncated: false })),
    validateEdits: mock(async (): Promise<pptx.PptxValidationResult> => { throw new Error('Unexpected mutation'); }),
    applyEdits: mock(async (): Promise<pptx.PptxEditResult> => { throw new Error('Unexpected mutation'); }),
    slides: async () => slides, slideSize: async () => ({ width: 960, height: 540 }),
    save: async () => new Uint8Array([8, 9]).buffer,
  };
  const viewer: PresentationSession = {
    state: { format: 'pptx', stage: 'ready', version: 0, dirty: false, slides, size: { width: 960, height: 540 } },
    call, save: mock(async () => new Uint8Array([8, 9])), on: () => () => {},
    onFailure: () => () => {}, failure: undefined, dispose: mock(async () => {}),
  };
  return { viewer, call };
}

function open(viewer: PresentationSession) {
  const spy = spyOn(presentationSessionOpener, 'open').mockResolvedValue(viewer);
  restorers.push(() => spy.mockRestore());
  return spy;
}

function visible(indices: number[], isIntersecting = true) {
  visibility(indices.map((index) => ({ target: observed[index], isIntersecting }) as IntersectionObserverEntry),
    {} as IntersectionObserver);
}

beforeEach(() => {
  observed = [];
  painted = [];
  const original = globalThis.IntersectionObserver;
  globalThis.IntersectionObserver = class {
    constructor(callback: IntersectionObserverCallback, options?: IntersectionObserverInit) {
      visibility = callback;
      observerRoot = options?.root;
    }
    observe(target: HTMLElement) { observed.push(target); }
    disconnect() {}
  } as unknown as typeof IntersectionObserver;
  restorers.push(() => { globalThis.IntersectionObserver = original; });
  const getContext = function (this: HTMLCanvasElement) {
    return { canvas: this, fillRect() {} } as unknown as CanvasRenderingContext2D;
  };
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
    getContext as unknown as HTMLCanvasElement['getContext']);
  restorers.push(() => context.mockRestore());
  const paint = spyOn(pptx, 'paintSlide').mockImplementation(async (ctx, list, _dpr, scale = 1) => {
    painted.push({ canvas: ctx.canvas, list, scale });
  });
  restorers.push(() => paint.mockRestore());
});
afterEach(() => {
  cleanup();
  for (const restore of restorers.reverse()) restore();
  restorers.length = 0;
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

describe('session viewer', () => {
  it('opens through the seam, paints active first and requests only visible thumbnails', async () => {
    const { viewer, call } = session();
    const opener = open(viewer);
    const init = spyOn(pptx, 'initWasm').mockRejectedValue(new Error('Local wasm forbidden'));
    const local = spyOn(pptx, 'openPresentation').mockImplementation(() => { throw new Error('Local open forbidden'); });
    restorers.push(() => init.mockRestore(), () => local.mockRestore());
    const changed = mock(() => {});
    const view = render(<PptxEditor file={file} fonts={[]} readOnly experimentalWorkerOpen initialSlide={2} onChange={changed} />);
    await waitFor(() => expect(observed).toHaveLength(4));
    expect(observerRoot).toBe(view.container.querySelector('aside'));
    await act(async () => visible([0, 1, 2]));
    await waitFor(() => expect(call.frame.mock.calls.map(([index]) => index)).toEqual([1, 0, 2]));
    await waitFor(() => expect(view.container.querySelectorAll('aside canvas')).toHaveLength(3));
    expect(painted[0].canvas.closest('aside')).toBeNull();
    expect(painted[0].list).toEqual(frame(1).displayList);
    expect(painted.filter(({ canvas }) => canvas.closest('aside')).every(({ list, scale }) => scale === 128 / list.width)).toBe(true);
    expect(opener).toHaveBeenCalledWith(file, { fonts: [], clientId: undefined });
    expect(init).not.toHaveBeenCalled();
    expect(local).not.toHaveBeenCalled();
    expect(changed).not.toHaveBeenCalled();
    await act(async () => visible([2], false));
    expect(view.container.querySelectorAll('aside canvas')).toHaveLength(2);
  });

  it('publishes the async contract after paint and resolves navigation after target paint', async () => {
    const { viewer, call } = session();
    open(viewer);
    const first = deferred<void>();
    const target = deferred<void>();
    const paint = spyOn(pptx, 'paintSlide').mockImplementation((_ctx, list) =>
      list.width === 960 ? first.promise : target.promise);
    restorers.push(() => paint.mockRestore());
    let api: PptxWorkerViewerApi | undefined;
    const ready = mock((value: PptxWorkerViewerApi) => { api = value; });
    const view = render(<PptxEditor file={file} fonts={[]} readOnly experimentalWorkerOpen onReady={ready} />);
    await waitFor(() => expect(paint).toHaveBeenCalled());
    expect(ready).not.toHaveBeenCalled();
    await act(async () => first.resolve());
    await waitFor(() => expect(api).toBeDefined());
    expect(api!.handle).toBeNull();
    expect(api!.save()).toBeNull();
    expect(api!.getPositionAtPoint(0, 0)).toBeNull();
    expect(api!.selectText({ slide: 1, shapeId: 'shape', storyId: 'story', start: 0, end: 1 })).toBe(false);
    expect(await api!.saveAsync()).toEqual(new Uint8Array([8, 9]));
    expect(await api!.version()).toBe('v1');
    await api!.readContent({ slideIds: ['s1'] });
    await api!.findText({ text: 'word' });
    expect(call.readContent).toHaveBeenCalledWith({ slideIds: ['s1'] });
    expect(call.findText).toHaveBeenCalledWith({ text: 'word' });
    for (const method of [api!.validateEdits, api!.applyEdits]) {
      expect(await method({ expectVersion: 'v1', steps: [] })).toMatchObject({ ok: false, failure: { code: 'read-only' } });
    }
    expect(call.validateEdits).not.toHaveBeenCalled();
    expect(call.applyEdits).not.toHaveBeenCalled();
    let navigation!: Promise<boolean>;
    let completed = false;
    act(() => { navigation = api!.goToSlideAsync(3).then((value) => { completed = true; return value; }); });
    await waitFor(() => expect(paint.mock.calls.some(([, list]) => list.width === 962)).toBe(true));
    expect(completed).toBe(false);
    await act(async () => target.resolve());
    expect(await navigation).toBe(true);
    expect(ready).toHaveBeenCalledTimes(1);
    await waitFor(() => expect((view.getByTestId('pptx-notes-textarea') as HTMLTextAreaElement).value).toBe('Notes 2'));
    expect(await api!.goToSlideAsync(0)).toBe(false);
    fireEvent.keyDown(view.container.querySelector('[tabindex="0"]')!, { key: 'ArrowLeft' });
    await waitFor(() => expect(view.container.querySelector('[aria-current="page"]')?.getAttribute('data-slide-index')).toBe('1'));
    fireEvent.click(view.container.querySelector('aside button')!);
    await waitFor(() => expect(view.container.querySelector('[aria-current="page"]')?.getAttribute('data-slide-index')).toBe('0'));
  });

  it('requeues superseded visible work with only one frame request in flight', async () => {
    const { viewer, call } = session();
    open(viewer);
    let inFlight = 0;
    let maxInFlight = 0;
    let superseded = false;
    call.frame.mockImplementation(async (index) => {
      inFlight += 1;
      maxInFlight = Math.max(inFlight, maxInFlight);
      try {
        if (index === 1 && !superseded) {
          superseded = true;
          const error = new Error('Replaced');
          error.name = 'SessionSuperseded';
          throw error;
        }
        await Promise.resolve();
        return frame(index);
      } finally { inFlight -= 1; }
    });
    render(<PptxEditor file={file} fonts={[]} readOnly experimentalWorkerOpen />);
    await waitFor(() => expect(observed).toHaveLength(4));
    await act(async () => visible([1]));
    await waitFor(() => expect(call.frame.mock.calls.map(([index]) => index)).toEqual([0, 1, 1]));
    expect(maxInFlight).toBe(1);
  });

  it('prioritizes navigation over queued thumbnails without overlapping frame calls', async () => {
    const { viewer, call } = session(5);
    const thumbnail = deferred<PresentationFrame>();
    call.frame.mockImplementation(async (index) => index === 1 ? thumbnail.promise : frame(index));
    const run = new ViewerSession(viewer, 1, () => {}, () => {}, () => {});
    run.start();
    await waitFor(() => expect(run.frame(0)).toBeDefined());
    run.didPaint(run.frame(0)!);
    run.visibility(1, true);
    run.visibility(2, true);
    const navigation = run.showAsync(5);
    expect(call.frame.mock.calls.map(([index]) => index)).toEqual([0, 1]);
    thumbnail.resolve(frame(1));
    await waitFor(() => expect(run.frame(4)).toBeDefined());
    expect(call.frame.mock.calls.map(([index]) => index)).toEqual([0, 1, 4]);
    run.didPaint(run.frame(4)!);
    expect(await navigation).toBe(true);
    await waitFor(() => expect(call.frame.mock.calls.map(([index]) => index)).toEqual([0, 1, 4, 2]));
    const interrupted = run.showAsync(4);
    run.show(3);
    expect(await interrupted).toBe(false);
    run.dispose();
  });

  it('installs browser fonts, forwards engine fonts and preserves equivalent inline fonts', async () => {
    const originalFace = globalThis.FontFace;
    const originalFonts = Object.getOwnPropertyDescriptor(document, 'fonts');
    const add = mock(() => {});
    const remove = mock(() => true);
    const faces: { family: string; source: ArrayBuffer; descriptors: FontFaceDescriptors }[] = [];
    globalThis.FontFace = class {
      constructor(family: string, source: ArrayBuffer, descriptors: FontFaceDescriptors) {
        faces.push({ family, source, descriptors });
      }
      async load() { return this; }
    } as unknown as typeof FontFace;
    Object.defineProperty(document, 'fonts', { configurable: true, value: { add, delete: remove } });
    restorers.push(() => {
      globalThis.FontFace = originalFace;
      if (originalFonts) Object.defineProperty(document, 'fonts', originalFonts);
      else Reflect.deleteProperty(document, 'fonts');
    });
    const opener = open(session().viewer);
    const fonts = () => [{ family: 'Viewer Font', bold: true, italic: true, bytes: new Uint8Array([3, 4]) }];
    const view = render(<PptxEditor file={file} fonts={fonts()} readOnly experimentalWorkerOpen />);
    await waitFor(() => expect(painted).toHaveLength(1));
    expect(faces).toHaveLength(1);
    expect(faces[0].descriptors).toEqual({ style: 'italic', weight: '700' });
    expect(new Uint8Array(faces[0].source)).toEqual(new Uint8Array([3, 4]));
    expect(add).toHaveBeenCalledTimes(1);
    expect(opener.mock.calls[0][1]?.fonts).toEqual(fonts());
    view.rerender(<PptxEditor file={file} fonts={fonts()} readOnly experimentalWorkerOpen />);
    expect(opener).toHaveBeenCalledTimes(1);
    view.unmount();
    expect(remove).toHaveBeenCalledTimes(1);
  });

  it('disposes replaced and unmounted sessions and ignores their late frames', async () => {
    const old = session();
    const next = session();
    const late = deferred<PresentationFrame>();
    old.call.frame.mockImplementation(() => late.promise);
    const opener = open(old.viewer);
    const view = render(<PptxEditor file={file} fonts={[]} readOnly experimentalWorkerOpen />);
    await waitFor(() => expect(old.call.frame).toHaveBeenCalled());
    opener.mockResolvedValue(next.viewer);
    view.rerender(<PptxEditor file={new Uint8Array([4])} fonts={[]} readOnly experimentalWorkerOpen />);
    await waitFor(() => expect(next.call.frame).toHaveBeenCalled());
    await waitFor(() => expect(painted).toHaveLength(1));
    await act(async () => late.resolve(frame(3)));
    expect(painted).toHaveLength(1);
    expect(old.viewer.dispose).toHaveBeenCalledTimes(1);
    view.unmount();
    expect(next.viewer.dispose).toHaveBeenCalledTimes(1);
  });

  it('blocks late painting into a disposed canvas', async () => {
    const { viewer } = session();
    open(viewer);
    const resume = deferred<void>();
    let context: CanvasRenderingContext2D | undefined;
    const paint = spyOn(pptx, 'paintSlide').mockImplementation(async (ctx) => {
      context = ctx;
      await resume.promise;
      ctx.fillRect(0, 0, 1, 1);
    });
    restorers.push(() => paint.mockRestore());
    const view = render(<PptxEditor file={file} fonts={[]} readOnly experimentalWorkerOpen />);
    await waitFor(() => expect(context).toBeDefined());
    view.unmount();
    expect(() => context!.fillRect(0, 0, 1, 1)).toThrow('Paint superseded');
    await act(async () => resume.resolve());
  });

  it('surfaces worker failure without falling back locally', async () => {
    const opener = open(session().viewer);
    const failure = new Error('Worker unavailable');
    opener.mockRejectedValue(failure);
    const local = spyOn(pptx, 'openPresentation');
    const init = spyOn(pptx, 'initWasm');
    restorers.push(() => local.mockRestore(), () => init.mockRestore());
    const errors = mock(() => {});
    const view = render(<PptxEditor file={file} fonts={[]} readOnly experimentalWorkerOpen onError={errors} />);
    await waitFor(() => expect(view.getByText(failure.message)).toBeDefined());
    expect(errors).toHaveBeenCalledWith(failure);
    expect(local).not.toHaveBeenCalled();
    expect(init).not.toHaveBeenCalled();
  });

  it('binds host save and zoom while disabling edits and exports', async () => {
    const { viewer } = session();
    open(viewer);
    const saved = mock(() => {});
    const requested = mock(async () => true);
    let api!: PptxWorkerViewerApi;
    const view = render(<PptxEditor file={file} fonts={[]} readOnly experimentalWorkerOpen
      onReady={(value) => { api = value; }} onSave={saved} onSaveRequest={requested}
      toolbar={<EditorToolbar mode="commands"><ToolbarCommandButton id="save" /><ToolbarCommandButton id="bold" /></EditorToolbar>} />);
    await waitFor(() => expect(api).toBeDefined());
    expect(api.commands.getState('bold').enabled).toBe(false);
    expect(api.commands.getState('slideshow').enabled).toBe(false);
    expect(api.commands.getState('exportPng').enabled).toBe(false);
    fireEvent.click(view.getByTestId('pptx-save'));
    await waitFor(() => expect(saved).toHaveBeenCalledWith(new Uint8Array([8, 9])));
    expect(requested).toHaveBeenCalledTimes(1);
    await act(async () => { await api.commands.execute('zoom', { scale: 1.5 }); });
    await waitFor(() => expect(painted.at(-1)?.scale).toBe(1.5));
  });

  it('bounds cached frames while retaining the active frame', async () => {
    const { viewer } = session(40);
    const changed = mock(() => {});
    const run = new ViewerSession(viewer, 1, changed, () => {}, () => {});
    run.start();
    await waitFor(() => expect(run.frame(0)).toBeDefined());
    run.didPaint(run.frame(0)!);
    for (let index = 1; index < 40; index += 1) run.visibility(index, true);
    await waitFor(() => expect(viewer.call.frame).toHaveBeenCalledTimes(40));
    expect(Array.from({ length: 40 }, (_, index) => run.frame(index)).filter(Boolean)).toHaveLength(25);
    expect(run.frame(0)).toBeDefined();
    run.dispose();
  });

  it('skips undecoded TIFF without invoking the wasm decoder', async () => {
    const images = frameImages();
    const image = frame(0);
    image.media = new Map([['tiff', new Uint8Array([0x49, 0x49, 0x2a, 0, 0, 0, 0, 0])]]);
    const decode = spyOn(pptx, 'decodePresentationImage');
    restorers.push(() => decode.mockRestore());
    expect(await images.resolve(image)('tiff')).toBeNull();
    expect(decode).not.toHaveBeenCalled();
    images.dispose();
  });

  it('uses the transport seam with a real presentation host', async () => {
    const wasm = await readFile(new URL('../../../pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url));
    const bytes = await readFile(new URL('../../../pptx/src/render/fixtures/tiff-image.pptx', import.meta.url));
    await pptx.initWasm(wasm);
    const pair = createInProcessPair();
    createPresentationSessionHost(pair.host, { initWasm: async () => {} });
    const opener = spyOn(presentationSessionOpener, 'open').mockImplementation((bytes, options = {}) =>
      createPresentationSession(bytes, options, pair.client));
    restorers.push(() => opener.mockRestore());
    let api!: PptxWorkerViewerApi;
    render(<PptxEditor file={bytes} fonts={[]} readOnly experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await waitFor(() => expect(api).toBeDefined());
    expect((await api.readContent()).ok).toBe(true);
    expect((await api.saveAsync()).byteLength).toBeGreaterThan(0);
  });

  for (const optIn of [false, true]) it(`keeps local opening when worker mode is inactive (${optIn})`, async () => {
    const opener = open(session().viewer);
    const init = spyOn(pptx, 'initWasm').mockResolvedValue(undefined);
    const failure = new Error('Local open reached');
    const local = spyOn(pptx, 'openPresentation').mockImplementation(() => { throw failure; });
    restorers.push(() => init.mockRestore(), () => local.mockRestore());
    const view = render(optIn ? <PptxEditor file={file} fonts={[]} experimentalWorkerOpen readOnly={false} /> :
      <PptxEditor file={file} fonts={[]} />);
    await waitFor(() => expect(view.getByText(failure.message)).toBeDefined());
    expect(local).toHaveBeenCalledTimes(1);
    expect(opener).not.toHaveBeenCalled();
  });
});
