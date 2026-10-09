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
import { isMacPlatform } from '../commands/descriptors';
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
let painted: { canvas: HTMLCanvasElement; list: SlideDisplayList; scale: number; dpr: number }[];

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
  visibility(indices.map((index) => ({ target: observed[index], isIntersecting }) as unknown as IntersectionObserverEntry),
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
  const paint = spyOn(pptx, 'paintSlide').mockImplementation(async (ctx, list, dpr = 1, scale = 1) => {
    painted.push({ canvas: ctx.canvas, list, scale, dpr });
  });
  restorers.push(() => paint.mockRestore());
});
afterEach(() => {
  cleanup();
  for (const restore of restorers.reverse()) restore();
  restorers.length = 0;
});
afterAll(async () => {
  await new Promise((done) => setTimeout(done));
  if (ownsDom) await GlobalRegistrator.unregister();
});

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

  it('resolves batched round-trip navigation and continues visible thumbnails', async () => {
    const { viewer, call } = session();
    open(viewer);
    let api!: PptxWorkerViewerApi;
    const view = render(<PptxEditor file={file} fonts={[]} readOnly experimentalWorkerOpen
      onReady={(value) => { api = value; }} />);
    await waitFor(() => expect(api).toBeDefined());
    let navigation!: Promise<boolean>;
    let completed = false;
    act(() => {
      api.goToSlide(2);
      navigation = api.goToSlideAsync(1).then((value) => { completed = true; return value; });
    });
    await waitFor(() => expect(completed).toBe(true), { timeout: 1000 });
    expect(await navigation).toBe(true);
    expect(view.container.querySelector('[aria-current="page"]')?.getAttribute('data-slide-index')).toBe('0');
    await waitFor(() => expect(observed).toHaveLength(4));
    await act(async () => visible([2, 3]));
    await waitFor(() => expect(call.frame.mock.calls.map(([index]) => index)).toEqual([0, 1, 2, 3]));
    await waitFor(() => expect(view.container.querySelectorAll('aside canvas')).toHaveLength(2));
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

  for (const first of ['open', 'fonts'] as const) it(`starts fonts before open resolves and waits for both (${first} first)`, async () => {
    const originalFace = globalThis.FontFace;
    const originalFonts = Object.getOwnPropertyDescriptor(document, 'fonts');
    const add = mock(() => {});
    const remove = mock(() => true);
    const loaded = deferred<void>();
    const load = mock(() => loaded.promise);
    const faces: { family: string; source: ArrayBuffer; descriptors: FontFaceDescriptors }[] = [];
    globalThis.FontFace = class {
      constructor(family: string, source: ArrayBuffer, descriptors: FontFaceDescriptors) {
        faces.push({ family, source, descriptors });
      }
      async load() { await load(); return this; }
    } as unknown as typeof FontFace;
    Object.defineProperty(document, 'fonts', { configurable: true, value: { add, delete: remove } });
    restorers.push(() => {
      globalThis.FontFace = originalFace;
      if (originalFonts) Object.defineProperty(document, 'fonts', originalFonts);
      else Reflect.deleteProperty(document, 'fonts');
    });
    const { viewer, call } = session();
    const opening = deferred<PresentationSession>();
    const opener = open(viewer).mockImplementation(() => opening.promise);
    const fonts = () => [{ family: 'Viewer Font', bold: true, italic: true, bytes: new Uint8Array([3, 4]) }];
    const view = render(<PptxEditor file={file} fonts={fonts()} readOnly experimentalWorkerOpen />);
    await waitFor(() => expect(load).toHaveBeenCalledTimes(1));
    expect(opener).toHaveBeenCalledTimes(1);
    expect(call.frame).not.toHaveBeenCalled();
    expect(add).not.toHaveBeenCalled();
    await act(async () => { if (first === 'open') opening.resolve(viewer); else loaded.resolve(); });
    expect(call.frame).not.toHaveBeenCalled();
    expect(painted).toHaveLength(0);
    await act(async () => { if (first === 'open') loaded.resolve(); else opening.resolve(viewer); });
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

  it('returns no-session results when the file is replaced during async API operations', async () => {
    const old = session();
    const next = session();
    const opener = open(old.viewer);
    const ready: PptxWorkerViewerApi[] = [];
    const onReady = (api: PptxWorkerViewerApi) => { ready.push(api); };
    const view = render(<PptxEditor file={file} fonts={[]} readOnly experimentalWorkerOpen onReady={onReady} />);
    await waitFor(() => expect(ready).toHaveLength(1));
    await waitFor(() => expect((view.getByTestId('pptx-notes-textarea') as HTMLTextAreaElement).value).toBe('Notes 0'));
    const read = deferred<pptx.PptxReadResult>();
    const version = deferred<string>();
    const found = deferred<pptx.PptxFindResult>();
    const saved = deferred<Uint8Array>();
    old.call.readContent.mockImplementation(() => read.promise);
    old.call.version.mockImplementation(() => version.promise);
    old.call.findText.mockImplementation(() => found.promise);
    const save = spyOn(old.viewer, 'save').mockImplementation(() => saved.promise);
    restorers.push(() => save.mockRestore());
    const api = ready[0];
    const request = { expectVersion: 'v1', steps: [] };
    const pending = [api.readContent(), api.version(), api.findText({ text: 'old' }), api.saveAsync(),
      api.validateEdits(request), api.applyEdits(request)];
    opener.mockResolvedValue(next.viewer);
    view.rerender(<PptxEditor file={new Uint8Array([4])} fonts={[]} readOnly experimentalWorkerOpen onReady={onReady} />);
    await waitFor(() => expect(ready).toHaveLength(2));
    await act(async () => {
      read.resolve({ ok: true, version: 'old', slides: [], stories: [] });
      version.resolve('old');
      found.resolve({ ok: true, version: 'old', matches: [], truncated: false });
      saved.resolve(new Uint8Array([7]));
      expect(await Promise.all(pending)).toEqual([null, null, null, null, null, null]);
    });
    const calls = old.call.readContent.mock.calls.length;
    expect(await api.readContent()).toBeNull();
    expect(old.call.readContent).toHaveBeenCalledTimes(calls);
    expect(await api.goToSlideAsync(1)).toBe(false);
    expect(await ready[1].version()).toBe('v1');
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
    await waitFor(() => expect(painted[painted.length - 1]?.scale).toBe(1.5));
  });

  it('bounds cached frames while retaining the active frame', async () => {
    const { viewer } = session(40);
    const changed = mock(() => {});
    const run = new ViewerSession(viewer, 1, changed, () => {}, () => {});
    run.start();
    await waitFor(() => expect(run.frame(0)).toBeDefined());
    run.didPaint(run.frame(0)!);
    for (let index = 1; index < 40; index += 1) {
      run.visibility(index, true);
      await waitFor(() => expect(viewer.call.frame).toHaveBeenCalledTimes(index + 1));
      if (index > 3) run.visibility(index - 3, false);
    }
    expect(Array.from({ length: 40 }, (_, index) => run.frame(index)).filter(Boolean)).toHaveLength(25);
    expect(run.frame(0)).toBeDefined();
    for (let index = 36; index < 40; index += 1) expect(run.frame(index)).toBeDefined();
    run.dispose();
  });

  it('keeps every visible frame cached when visible rows exceed the cache bound', async () => {
    const { viewer } = session(40);
    const run = new ViewerSession(viewer, 1, () => {}, () => {}, () => {});
    run.start();
    await waitFor(() => expect(run.frame(0)).toBeDefined());
    run.didPaint(run.frame(0)!);
    for (let index = 1; index < 30; index += 1) run.visibility(index, true);
    await waitFor(() => expect(viewer.call.frame).toHaveBeenCalledTimes(30));
    for (let index = 0; index < 30; index += 1) expect(run.frame(index)).toBeDefined();
    run.dispose();
  });

  it('trims cached frames when visible rows are hidden while retaining the active frame', async () => {
    const { viewer } = session(31);
    const run = new ViewerSession(viewer, 1, () => {}, () => {}, () => {});
    run.start();
    await waitFor(() => expect(run.frame(0)).toBeDefined());
    run.didPaint(run.frame(0)!);
    for (let index = 0; index < 30; index += 1) run.visibility(index, true);
    await waitFor(() => expect(run.frame(29)).toBeDefined());
    expect(Array.from({ length: 31 }, (_, index) => run.frame(index) !== undefined).filter(Boolean)).toHaveLength(30);
    for (let index = 1; index < 30; index += 1) run.visibility(index, false);
    expect(Array.from({ length: 31 }, (_, index) => run.frame(index) !== undefined).filter(Boolean).length).toBeLessThanOrEqual(25);
    expect(run.frame(0)).toBeDefined();
    run.dispose();
  });

  it('saves from the platform shortcut without throwing', async () => {
    const { viewer } = session();
    open(viewer);
    const saved = mock(() => {});
    const requested = mock(async () => true);
    let api!: PptxWorkerViewerApi;
    const view = render(<PptxEditor file={file} fonts={[]} readOnly experimentalWorkerOpen
      onReady={(value) => { api = value; }} onSave={saved} onSaveRequest={requested} />);
    await waitFor(() => expect(api).toBeDefined());
    const errors = mock((event: ErrorEvent) => { event.preventDefault(); });
    window.addEventListener('error', errors);
    restorers.push(() => window.removeEventListener('error', errors));
    const stage = view.container.querySelector('[tabindex="0"]')!;
    expect(() => fireEvent.keyDown(stage, { key: 's', ...(isMacPlatform() ? { metaKey: true } : { ctrlKey: true }) })).not.toThrow();
    await waitFor(() => expect(saved).toHaveBeenCalledWith(new Uint8Array([8, 9])));
    expect(requested).toHaveBeenCalledTimes(1);
    expect(viewer.save).toHaveBeenCalledTimes(1);
    expect(errors).not.toHaveBeenCalled();
  });

  it('repaints active slides and thumbnails when DPR changes and removes listeners', async () => {
    open(session().viewer);
    const originalDpr = Object.getOwnPropertyDescriptor(window, 'devicePixelRatio');
    const originalMedia = Object.getOwnPropertyDescriptor(window, 'matchMedia');
    const setDpr = (value: number) => Object.defineProperty(window, 'devicePixelRatio', { configurable: true, value });
    const makeMedia = (media: string) => {
      let change!: () => void;
      return {
        media, change: () => change(),
        addEventListener: mock((_type: string, listener: () => void) => { change = listener; }),
        removeEventListener: mock(() => {}),
      };
    };
    const queries: ReturnType<typeof makeMedia>[] = [];
    setDpr(1);
    Object.defineProperty(window, 'matchMedia', { configurable: true, value: (query: string) => {
      const media = makeMedia(query);
      queries.push(media);
      return media as unknown as MediaQueryList;
    } });
    restorers.push(() => {
      if (originalDpr) Object.defineProperty(window, 'devicePixelRatio', originalDpr);
      else Reflect.deleteProperty(window, 'devicePixelRatio');
      if (originalMedia) Object.defineProperty(window, 'matchMedia', originalMedia);
      else Reflect.deleteProperty(window, 'matchMedia');
    });
    const view = render(<PptxEditor file={file} fonts={[]} readOnly experimentalWorkerOpen />);
    await waitFor(() => expect(observed).toHaveLength(4));
    await act(async () => visible([1]));
    await waitFor(() => expect(painted).toHaveLength(2));
    expect(queries[0].media).toBe('(resolution: 1dppx)');
    for (const ratio of [2, 1]) {
      const previous = queries[queries.length - 1];
      painted = [];
      await act(async () => { setDpr(ratio); previous.change(); });
      await waitFor(() => expect(painted).toHaveLength(2));
      expect(painted.every((paint) => paint.dpr === ratio)).toBe(true);
      expect(painted.some((paint) => paint.canvas.closest('aside'))).toBe(true);
      expect(painted.some((paint) => !paint.canvas.closest('aside'))).toBe(true);
      expect(previous.removeEventListener).toHaveBeenCalledTimes(1);
      expect(queries[queries.length - 1].media).toBe(`(resolution: ${ratio}dppx)`);
    }
    view.unmount();
    expect(queries[queries.length - 1].removeEventListener).toHaveBeenCalledTimes(1);
  });

  for (const extent of [1, 20, 40]) it(`keeps fit painting positive in a ${extent}px viewport`, async () => {
    open(session().viewer);
    const originalWidth = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
    const originalHeight = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientHeight');
    Object.defineProperty(HTMLElement.prototype, 'clientWidth', { configurable: true, get: () => extent });
    Object.defineProperty(HTMLElement.prototype, 'clientHeight', { configurable: true, get: () => extent });
    restorers.push(() => {
      if (originalWidth) Object.defineProperty(HTMLElement.prototype, 'clientWidth', originalWidth);
      else Reflect.deleteProperty(HTMLElement.prototype, 'clientWidth');
      if (originalHeight) Object.defineProperty(HTMLElement.prototype, 'clientHeight', originalHeight);
      else Reflect.deleteProperty(HTMLElement.prototype, 'clientHeight');
    });
    render(<PptxEditor file={file} fonts={[]} readOnly experimentalWorkerOpen />);
    await waitFor(() => expect(painted).toHaveLength(1));
    expect(Number.isFinite(painted[0].scale)).toBe(true);
    expect(painted[0].scale).toBeGreaterThan(0);
    expect(painted[0].canvas.width).toBeGreaterThanOrEqual(1);
    expect(painted[0].canvas.height).toBeGreaterThanOrEqual(1);
  });

  it('evicts the least recently used bitmap and closes it after every paint releases it', async () => {
    const image = frame(0);
    image.media = new Map(Array.from({ length: 27 }, (_, index) => [String(index), new Uint8Array([index])] as const));
    const bitmaps = Array.from({ length: 27 }, () => ({ close: mock(() => {}) }));
    const decode = spyOn(pptx, 'decodePresentationImage').mockImplementation(async (bytes) =>
      bitmaps[bytes[0]] as unknown as ImageBitmap);
    restorers.push(() => decode.mockRestore());
    const images = frameImages();
    const first = images.resolve(image);
    const second = images.resolve(image);
    await first('0');
    await second('0');
    for (let index = 1; index < 25; index += 1) {
      const paint = images.resolve(image);
      await paint(String(index));
      paint.release();
    }
    await second('1');
    const next = images.resolve(image);
    await next('25');
    next.release();
    expect(bitmaps[0].close).not.toHaveBeenCalled();
    expect(bitmaps[1].close).not.toHaveBeenCalled();
    first.release();
    expect(bitmaps[0].close).not.toHaveBeenCalled();
    second.release();
    expect(bitmaps[0].close).toHaveBeenCalledTimes(1);
    const last = images.resolve(image);
    await last('26');
    expect(bitmaps[2].close).toHaveBeenCalledTimes(1);
    images.dispose();
    expect(bitmaps[26].close).not.toHaveBeenCalled();
    last.release();
    last.release();
    images.dispose();
    expect(decode).toHaveBeenCalledTimes(27);
    for (const bitmap of bitmaps) expect(bitmap.close).toHaveBeenCalledTimes(1);
  });

  it('closes a pending bitmap once decoding and its disposed paint finish', async () => {
    const decoded = deferred<ImageBitmap>();
    const bitmap = { close: mock(() => {}) } as unknown as ImageBitmap;
    const decode = spyOn(pptx, 'decodePresentationImage').mockImplementation(() => decoded.promise);
    restorers.push(() => decode.mockRestore());
    const images = frameImages();
    const image = frame(0);
    image.media = new Map([['image', new Uint8Array([1])]]);
    const paint = images.resolve(image);
    const pending = paint('image');
    images.dispose();
    decoded.resolve(bitmap);
    await pending;
    expect(bitmap.close).not.toHaveBeenCalled();
    paint.release();
    paint.release();
    expect(bitmap.close).toHaveBeenCalledTimes(1);
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
    expect((await api.readContent())!.ok).toBe(true);
    expect((await api.saveAsync())!.byteLength).toBeGreaterThan(0);
  });

  it('keeps local opening when worker mode is inactive', async () => {
    const opener = open(session().viewer);
    const init = spyOn(pptx, 'initWasm').mockResolvedValue(undefined);
    const failure = new Error('Local open reached');
    const local = spyOn(pptx, 'openPresentation').mockImplementation(() => { throw failure; });
    restorers.push(() => init.mockRestore(), () => local.mockRestore());
    const view = render(<PptxEditor file={file} fonts={[]} />);
    await waitFor(() => expect(view.getByText(failure.message)).toBeDefined());
    expect(local).toHaveBeenCalledTimes(1);
    expect(opener).not.toHaveBeenCalled();
  });
});
