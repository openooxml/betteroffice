import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, describe, expect, it, mock, spyOn } from 'bun:test';
import * as xlsx from '@betteroffice/xlsx';
import type {
  DisplayList, Viewport, WorkbookFrame, WorkbookSession, WorkbookSheetView,
} from '@betteroffice/xlsx';
import { XlsxEditor } from '../XlsxEditor';
import type { XlsxEditorProps, XlsxWorkerViewerApi } from '../XlsxEditor';
import { EditorToolbar } from '../components/EditorToolbar';
import { ToolbarCommandButton } from '../components/toolbar/ToolbarCommand';
import { workbookSessionOpener } from './useSessionWorkbook';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const file = new Uint8Array([1, 2, 3]);
const viewport = { x: 0, y: 0, width: 800, height: 600 };
const restorers: (() => void)[] = [];
let painted: { canvas: HTMLCanvasElement; list: DisplayList; scale: number }[];
let animationFrames: Map<number, FrameRequestCallback>;
let animationId = 0;
let resized: ResizeObserverCallback;

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

async function tick() {
  await act(async () => {
    const callbacks = [...animationFrames.values()];
    animationFrames.clear();
    for (const callback of callbacks) callback(0);
  });
}

function frame(sheet: number, viewport: Viewport): WorkbookFrame {
  return {
    sheet, viewport, version: 'v1', epoch: 1, sequence: 0, mergedRanges: [],
    displayList: {
      width: viewport.width, height: viewport.height,
      commands: [{ op: 'text', text: `Sheet ${sheet}`, x: 8, y: 18, fontSize: 11,
        color: '#000000', clip: { x: 0, y: 0, w: 96, h: 24 } }],
      grid: { startRow: 0, startCol: 0, rowOffsets: [0, 24, 48, 72], colOffsets: [0, 96, 192, 288] },
    },
  };
}

function sheetView(sheet: number): WorkbookSheetView {
  return {
    sheet, version: 'v1', contentWidth: 4000, contentHeight: 6000,
    frozenRows: 0, frozenCols: 0, frozenWidth: 0, frozenHeight: 0,
    initialScrollX: 0, initialScrollY: 0,
  };
}

function session() {
  const sheets = [{ id: 'sheet:0', index: 0, name: 'First' }, { id: 'sheet:1', index: 1, name: 'Second' }];
  let failure: Parameters<WorkbookSession['onFailure']>[0] | undefined;
  const call = {
    frame: mock(async (viewport: Viewport, options?: { sheet?: number }) => frame(options?.sheet ?? 0, viewport)),
    sheetView: mock(async (sheet: number) => sheetView(sheet)),
    cellGeometry: mock(async (sheet: number, row: number, col: number) => ({
      sheet, version: 'v1', rect: { x: col * 96, y: row * 24, w: 96, h: 24 },
      scrollPosition: { x: col * 96, y: row * 24 },
    })),
    cellInputs: mock(async (sheet: number, range: string) => ({
      sheet, version: 'v1', cells: [[{ a1: range, input: '=SUM(B1:C1)', isFormula: true }]],
    })),
    version: mock(async () => 'v1'),
    readCells: mock(async (_request: xlsx.XlsxReadRequest): Promise<xlsx.XlsxReadResult> => ({
      ok: true, version: 'v1', sheets: [], ranges: [],
      calculation: { cycleCells: [], limitedCells: [], truncated: false },
    })),
    findText: mock(async (_request: xlsx.XlsxFindRequest): Promise<xlsx.XlsxFindResult> => ({
      ok: true, version: 'v1', matches: [], truncated: false,
    })),
    validateEdits: mock(async (): Promise<xlsx.XlsxValidationResult> => { throw new Error('Unexpected mutation'); }),
    applyEdits: mock(async (): Promise<xlsx.XlsxEditResult> => { throw new Error('Unexpected mutation'); }),
    sheets: async () => sheets, calculationStatus: async () => ({ limitedCells: [] }),
    save: async () => new Uint8Array([8, 9]).buffer,
  };
  const viewer: WorkbookSession = {
    state: { format: 'xlsx', stage: 'ready', version: 0, dirty: false, sheets, activeSheet: 0 },
    call, save: mock(async () => new Uint8Array([8, 9])), on: () => () => {},
    onFailure: (listener) => { failure = listener; return () => { failure = undefined; }; },
    failure: undefined, dispose: mock(async () => {}),
  };
  return { viewer, call, fail: () => failure?.(Object.assign(new Error('Worker stopped'), { code: 'crash' as const })) };
}

function open(viewer: WorkbookSession) {
  const spy = spyOn(workbookSessionOpener, 'open').mockResolvedValue(viewer);
  restorers.push(() => spy.mockRestore());
  return spy;
}

async function opened() {
  await waitFor(() => expect(animationFrames.size).toBeGreaterThan(0));
  await tick();
}

function clampScrolling(scroll: HTMLElement) {
  for (const [property, dimension, size] of [
    ['scrollLeft', 'width', 'clientWidth'], ['scrollTop', 'height', 'clientHeight'],
  ] as const) {
    let value = 0;
    Object.defineProperty(scroll, property, {
      configurable: true,
      get: () => value,
      set: (next: number) => {
        const spacer = scroll.firstElementChild as HTMLElement;
        value = Math.max(0, Math.min(next, Math.max(0, parseFloat(spacer.style[dimension]) - scroll[size])));
      },
    });
  }
}

beforeEach(() => {
  painted = [];
  animationFrames = new Map();
  const originalRaf = globalThis.requestAnimationFrame;
  const originalCancel = globalThis.cancelAnimationFrame;
  const originalResize = globalThis.ResizeObserver;
  globalThis.requestAnimationFrame = (callback) => {
    const id = ++animationId;
    animationFrames.set(id, callback);
    return id;
  };
  globalThis.cancelAnimationFrame = (id) => { animationFrames.delete(id); };
  globalThis.ResizeObserver = class {
    constructor(callback: ResizeObserverCallback) { resized = callback; }
    observe() {}
    disconnect() {}
  } as unknown as typeof ResizeObserver;
  restorers.push(() => {
    globalThis.requestAnimationFrame = originalRaf;
    globalThis.cancelAnimationFrame = originalCancel;
    globalThis.ResizeObserver = originalResize;
  });
  for (const [name, value] of [['clientWidth', 800], ['clientHeight', 600]] as const) {
    const original = Object.getOwnPropertyDescriptor(HTMLElement.prototype, name);
    Object.defineProperty(HTMLElement.prototype, name, { configurable: true, get: () => value });
    restorers.push(() => {
      if (original) Object.defineProperty(HTMLElement.prototype, name, original);
      else Reflect.deleteProperty(HTMLElement.prototype, name);
    });
  }
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (this: HTMLCanvasElement) {
    return { canvas: this } as unknown as CanvasRenderingContext2D;
  } as unknown as HTMLCanvasElement['getContext']);
  const paint = spyOn(xlsx, 'paintDisplayList').mockImplementation((ctx, list, scale) => {
    painted.push({ canvas: ctx.canvas, list, scale });
  });
  const local = spyOn(xlsx, 'openWorkbook').mockImplementation(() => { throw new Error('Local open forbidden'); });
  const init = spyOn(xlsx, 'initWasm').mockRejectedValue(new Error('Local wasm forbidden'));
  restorers.push(() => context.mockRestore(), () => paint.mockRestore(), () => local.mockRestore(), () => init.mockRestore());
});
afterEach(() => {
  cleanup();
  for (const restore of restorers.reverse()) restore();
  restorers.length = 0;
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

describe('workbook session viewer', () => {
  it('opens and publishes onReady once after the first paint without local wasm', async () => {
    const { viewer, call } = session();
    const opener = open(viewer);
    const ready = mock((_api: XlsxWorkerViewerApi) => {});
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={ready} />);
    await waitFor(() => expect(call.sheetView).toHaveBeenCalledWith(0));
    expect(ready).not.toHaveBeenCalled();
    await opened();
    expect(painted).toHaveLength(1);
    expect(ready).toHaveBeenCalledTimes(1);
    expect(opener).toHaveBeenCalledWith(file, expect.anything());
    expect(call.frame).toHaveBeenCalledWith(viewport, { sheet: 0 });
    expect(view.getByRole('grid').textContent).toContain('Sheet 0');
    fireEvent.scroll(view.getByTestId('xlsx-scroll'));
    await tick();
    expect(ready).toHaveBeenCalledTimes(1);
    expect(xlsx.openWorkbook).not.toHaveBeenCalled();
    expect(xlsx.initWasm).not.toHaveBeenCalled();
  });

  it('ignores legacy clientId props without forwarding collaboration options or reopening', async () => {
    const { viewer } = session();
    const opener = open(viewer).mockImplementation(async (_bytes, options) => {
      if (options?.clientId !== undefined) throw new Error('clientId requires collaborative mode');
      return viewer;
    });
    const ready = mock((_api: XlsxWorkerViewerApi) => {});
    const errors = mock((_error: Error) => {});
    const legacy = { clientId: 42, collaborative: true };
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen {...legacy}
      onReady={ready} onError={errors} />);
    await opened();
    expect(opener.mock.calls[0][1]).not.toHaveProperty('clientId');
    expect(opener.mock.calls[0][1]).not.toHaveProperty('collaborative');
    expect(ready).toHaveBeenCalledTimes(1);
    expect(errors).not.toHaveBeenCalled();
    view.rerender(<XlsxEditor file={file} readOnly experimentalWorkerOpen {...{ ...legacy, clientId: 43 }}
      onReady={ready} onError={errors} />);
    await tick();
    expect(opener).toHaveBeenCalledTimes(1);
  });

  it('keeps the synchronous sentinels and forwards async reads and saving', async () => {
    const { viewer, call } = session();
    open(viewer);
    let api!: XlsxWorkerViewerApi;
    render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    expect(api.handle).toBeNull();
    expect(api.save()).toBeNull();
    expect(api.selectCells(0, xlsx.selectionAt({ row: 0, col: 0 }))).toBe(false);
    expect(await api.saveAsync()).toEqual(new Uint8Array([8, 9]));
    expect(await api.version()).toBe('v1');
    await api.readCells({ ranges: [] });
    await api.findText({ text: 'word' });
    expect(call.readCells).toHaveBeenCalledWith({ ranges: [] });
    expect(call.findText).toHaveBeenCalledWith({ text: 'word' });
    for (const method of [api.validateEdits, api.applyEdits]) {
      expect(await method({ expectVersion: 'v1', steps: [] })).toEqual({
        ok: false, version: 'v1', failure: { code: 'read-only', message: 'The editor is read-only' },
      });
    }
    expect(call.validateEdits).not.toHaveBeenCalled();
    expect(call.applyEdits).not.toHaveBeenCalled();
  });

  it('disposes on file replacement and unmount and clears ready', async () => {
    const old = session();
    const next = session();
    const opener = open(old.viewer);
    const disposeReady = mock(() => {});
    const ready = mock((_api: XlsxWorkerViewerApi) => disposeReady);
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={ready} />);
    await opened();
    const nextView = deferred<WorkbookSheetView>();
    next.call.sheetView.mockImplementationOnce(() => nextView.promise);
    opener.mockResolvedValue(next.viewer);
    view.rerender(<XlsxEditor file={new Uint8Array([4])} readOnly experimentalWorkerOpen onReady={ready} />);
    await waitFor(() => expect(next.call.sheetView).toHaveBeenCalled());
    await act(async () => nextView.resolve(sheetView(0)));
    await tick();
    expect(old.viewer.dispose).toHaveBeenCalledTimes(1);
    expect(disposeReady).toHaveBeenCalledTimes(1);
    expect(ready).toHaveBeenCalledTimes(2);
    view.unmount();
    expect(next.viewer.dispose).toHaveBeenCalledTimes(1);
    expect(disposeReady).toHaveBeenCalledTimes(2);
    expect(animationFrames.size).toBe(0);
  });

  for (const close of ['replace', 'unmount'] as const) {
    for (const settle of ['resolve', 'reject'] as const) {
      it(`cancels an in-flight open on ${close} and ignores a late ${settle}`, async () => {
        const old = session();
        const next = session();
        const opening = deferred<WorkbookSession>();
        const stopped = mock(() => {});
        let signal!: AbortSignal;
        const opener = open(old.viewer).mockImplementationOnce((_bytes, options) => {
          signal = options!.signal!;
          signal.addEventListener('abort', stopped, { once: true });
          return opening.promise;
        });
        const ready = mock((_api: XlsxWorkerViewerApi) => {});
        const errors = mock((_error: Error) => {});
        const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={ready} onError={errors} />);
        expect(signal.aborted).toBe(false);
        if (close === 'unmount') view.unmount();
        else {
          opener.mockResolvedValue(next.viewer);
          view.rerender(<XlsxEditor file={new Uint8Array([4])} readOnly experimentalWorkerOpen onReady={ready} onError={errors} />);
          await opened();
        }
        expect(signal.aborted).toBe(true);
        expect(stopped).toHaveBeenCalledTimes(1);
        expect(old.viewer.dispose).not.toHaveBeenCalled();
        await act(async () => {
          if (settle === 'resolve') opening.resolve(old.viewer);
          else opening.reject(new Error('Obsolete open failed'));
        });
        expect(old.viewer.dispose).toHaveBeenCalledTimes(settle === 'resolve' ? 1 : 0);
        expect(old.call.sheetView).not.toHaveBeenCalled();
        expect(old.call.frame).not.toHaveBeenCalled();
        expect(ready).toHaveBeenCalledTimes(close === 'replace' ? 1 : 0);
        expect(errors).not.toHaveBeenCalled();
        expect(painted).toHaveLength(close === 'replace' ? 1 : 0);
        expect(view.queryByRole('alert')).toBeNull();
      });
    }
  }

  it('terminates an opening worker and removes its transport listeners before it is ready', async () => {
    const openSession = workbookSessionOpener.open;
    const worker = Object.assign(new EventTarget(), {
      postMessage: mock((_message: unknown) => {}), terminate: mock(() => {}),
    });
    const removed = spyOn(worker, 'removeEventListener');
    restorers.push(() => removed.mockRestore());
    let opening!: Promise<WorkbookSession>;
    open(session().viewer).mockImplementationOnce((bytes, options) => {
      opening = openSession(bytes, { ...options, worker: () => worker as unknown as Worker });
      return opening;
    });
    const ready = mock((_api: XlsxWorkerViewerApi) => {});
    const errors = mock((_error: Error) => {});
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={ready} onError={errors} />);
    expect(worker.postMessage).toHaveBeenCalledTimes(1);
    view.unmount();
    expect(worker.terminate).toHaveBeenCalledTimes(1);
    expect(removed.mock.calls.map(([type]) => type)).toEqual(['message', 'error', 'messageerror']);
    await act(async () => { await opening.catch(() => {}); });
    expect(ready).not.toHaveBeenCalled();
    expect(errors).not.toHaveBeenCalled();
    expect(painted).toHaveLength(0);
  });

  it('disposes an opening that completes after replacement and drops a late frame', async () => {
    const old = session();
    const next = session();
    const opening = deferred<WorkbookSession>();
    const opener = open(old.viewer).mockImplementationOnce(() => opening.promise);
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen />);
    opener.mockResolvedValue(next.viewer);
    view.rerender(<XlsxEditor file={new Uint8Array([4])} readOnly experimentalWorkerOpen />);
    await opened();
    await act(async () => opening.resolve(old.viewer));
    expect(old.viewer.dispose).toHaveBeenCalledTimes(1);
    expect(old.call.frame).not.toHaveBeenCalled();
    const late = deferred<WorkbookFrame>();
    next.call.frame.mockImplementation(() => late.promise);
    fireEvent.scroll(view.getByTestId('xlsx-scroll'));
    await tick();
    view.unmount();
    await act(async () => late.resolve(frame(0, viewport)));
    expect(painted).toHaveLength(1);
  });

  for (const close of ['replace', 'unmount'] as const) it(`resolves late reads to null after ${close}`, async () => {
    const old = session();
    const next = session();
    const opener = open(old.viewer);
    let api!: XlsxWorkerViewerApi;
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    const read = deferred<xlsx.XlsxReadResult>();
    const found = deferred<xlsx.XlsxFindResult>();
    const version = deferred<string>();
    const saved = deferred<Uint8Array>();
    old.call.readCells.mockImplementation(() => read.promise);
    old.call.findText.mockImplementation(() => found.promise);
    old.call.version.mockImplementation(() => version.promise);
    const save = spyOn(old.viewer, 'save').mockImplementation(() => saved.promise);
    restorers.push(() => save.mockRestore());
    const pending = [api.readCells({ ranges: [] }), api.findText({ text: 'old' }), api.version(), api.saveAsync(),
      api.validateEdits({ expectVersion: 'v1', steps: [] }), api.applyEdits({ expectVersion: 'v1', steps: [] })];
    if (close === 'unmount') view.unmount();
    else {
      opener.mockResolvedValue(next.viewer);
      view.rerender(<XlsxEditor file={new Uint8Array([4])} readOnly experimentalWorkerOpen />);
    }
    await act(async () => {
      read.resolve({ ok: false, version: 'old', failure: { code: 'missing-target', message: 'old' } });
      found.resolve({ ok: true, version: 'old', matches: [], truncated: false });
      version.resolve('old');
      saved.resolve(new Uint8Array([7]));
      expect(await Promise.all(pending)).toEqual([null, null, null, null, null, null]);
    });
    const calls = old.call.readCells.mock.calls.length;
    expect(await api.readCells({ ranges: [] })).toBeNull();
    expect(old.call.readCells).toHaveBeenCalledTimes(calls);
    expect(await api.selectCellsAsync(0, xlsx.selectionAt({ row: 0, col: 0 }))).toBe(false);
  });

  it('coalesces frames, keeps one in flight and paints only the latest viewport', async () => {
    const { viewer, call } = session();
    const first = deferred<WorkbookFrame>();
    call.frame.mockImplementationOnce(() => first.promise);
    open(viewer);
    const ready = mock((_api: XlsxWorkerViewerApi) => {});
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={ready} />);
    await opened();
    const scroll = view.getByTestId('xlsx-scroll');
    for (const x of [100, 200, 300]) { scroll.scrollLeft = x; fireEvent.scroll(scroll); }
    expect(animationFrames.size).toBe(1);
    await tick();
    expect(call.frame).toHaveBeenCalledTimes(1);
    scroll.scrollLeft = 400;
    fireEvent.scroll(scroll);
    await tick();
    expect(call.frame).toHaveBeenCalledTimes(1);
    await act(async () => first.resolve(frame(0, viewport)));
    expect(call.frame).toHaveBeenCalledTimes(2);
    expect(call.frame.mock.calls[1][0].x).toBe(400);
    expect(painted).toHaveLength(1);
    expect(ready).toHaveBeenCalledTimes(1);
  });

  it('cancels superseded requests and retries without surfacing an error', async () => {
    const { viewer, call } = session();
    const cancelled = new Error('replaced');
    cancelled.name = 'SessionSuperseded';
    call.frame.mockRejectedValueOnce(cancelled);
    open(viewer);
    const errors = mock(() => {});
    render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onError={errors} />);
    await opened();
    expect(painted).toHaveLength(0);
    await tick();
    expect(painted).toHaveLength(1);
    expect(errors).not.toHaveBeenCalled();
  });

  it('drops a response after zoom changes even before the next animation frame', async () => {
    const { viewer, call } = session();
    open(viewer);
    let api!: XlsxWorkerViewerApi;
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    const stale = deferred<WorkbookFrame>();
    call.frame.mockImplementationOnce(() => stale.promise);
    fireEvent.scroll(view.getByTestId('xlsx-scroll'));
    await tick();
    await act(async () => {
      await api.commands.execute('zoom', { scale: 2 });
      stale.resolve(frame(0, viewport));
    });
    expect(painted).toHaveLength(1);
    await tick();
    expect(painted).toHaveLength(2);
    expect(call.frame.mock.calls[2][0]).toEqual({ x: 0, y: 0, width: 400, height: 300 });
    expect(painted[1].scale).toBe((window.devicePixelRatio || 1) * 2);
  });

  it('drops a replaced session frame before the next session paints', async () => {
    const old = session();
    const next = session();
    const stale = deferred<WorkbookFrame>();
    old.call.frame.mockImplementationOnce(() => stale.promise);
    const opener = open(old.viewer);
    const ready = mock((_api: XlsxWorkerViewerApi) => {});
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={ready} />);
    await opened();
    opener.mockResolvedValue(next.viewer);
    view.rerender(<XlsxEditor file={new Uint8Array([4])} readOnly experimentalWorkerOpen onReady={ready} />);
    await waitFor(() => expect(next.call.sheetView).toHaveBeenCalled());
    await act(async () => stale.resolve(frame(0, viewport)));
    expect(painted).toHaveLength(0);
    expect(ready).not.toHaveBeenCalled();
    await tick();
    expect(painted).toHaveLength(1);
    expect(ready).toHaveBeenCalledTimes(1);
    expect(old.viewer.dispose).toHaveBeenCalledTimes(1);
  });

  it('switches local sheets and drops old-sheet frames without changing engine state', async () => {
    const { viewer, call } = session();
    const old = deferred<WorkbookFrame>();
    call.frame.mockImplementationOnce(() => old.promise);
    open(viewer);
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen />);
    await opened();
    const second = deferred<WorkbookSheetView>();
    call.sheetView.mockImplementationOnce(() => second.promise);
    fireEvent.click(view.getByRole('tab', { name: 'Second' }));
    expect(call.sheetView).toHaveBeenCalledWith(1);
    await act(async () => second.resolve(sheetView(1)));
    await tick();
    await act(async () => old.resolve(frame(0, viewport)));
    expect(call.frame.mock.calls.map(([, options]) => options?.sheet)).toEqual([0, 1]);
    expect(painted).toHaveLength(1);
    expect(view.getByRole('grid').textContent).toContain('Sheet 1');
    expect(viewer.state.activeSheet).toBe(0);
    expect(Object.keys(call)).not.toContain('setActiveSheet');
  });

  it('reveals async selection and resolves only after the target paint', async () => {
    const { viewer, call } = session();
    open(viewer);
    let api!: XlsxWorkerViewerApi;
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    const target = deferred<WorkbookFrame>();
    call.frame.mockImplementationOnce(() => target.promise);
    let completed = false;
    let selected!: Promise<boolean>;
    const selection = { anchor: { row: 1, col: 1 }, focus: { row: 2, col: 2 } };
    await act(async () => { selected = api.selectCellsAsync(1, selection).then((value) => { completed = true; return value; }); });
    await tick();
    const targetViewport = call.frame.mock.calls[1][0];
    expect(targetViewport).toEqual({ ...viewport, x: 192, y: 48 });
    expect(completed).toBe(false);
    expect((view.getByTestId('xlsx-scroll') as HTMLElement).scrollLeft).toBe(192);
    await act(async () => target.resolve(frame(1, targetViewport)));
    expect(await selected).toBe(true);
    expect(view.getByTestId('xlsx-selection')).toBeDefined();
    expect(viewer.state.activeSheet).toBe(0);
    expect(await api.selectCellsAsync(-1, selection)).toBe(false);
    expect(await api.selectCellsAsync(0, { ...selection, focus: { row: -1, col: 0 } })).toBe(false);
  });

  it('extends an A1-only sheet to reveal Y91 before completing async selection', async () => {
    const { viewer, call } = session();
    call.sheetView.mockResolvedValue({ ...sheetView(0), contentWidth: 96, contentHeight: 24 });
    open(viewer);
    let api!: XlsxWorkerViewerApi;
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    const scroll = view.getByTestId('xlsx-scroll');
    clampScrolling(scroll);
    const target = deferred<WorkbookFrame>();
    call.frame.mockImplementationOnce(() => target.promise);
    let completed = false;
    let selected!: Promise<boolean>;
    await act(async () => {
      selected = api.selectCellsAsync(0, xlsx.selectionAt({ row: 90, col: 24 }))
        .then((value) => { completed = true; return value; });
    });
    const spacer = scroll.firstElementChild as HTMLElement;
    expect(spacer.style.width).toBe('2400px');
    expect(spacer.style.height).toBe('2184px');
    expect(scroll.scrollLeft).toBe(1600);
    expect(scroll.scrollTop).toBe(1584);
    await tick();
    expect(completed).toBe(false);
    const requested = call.frame.mock.calls[1][0];
    expect(requested).toEqual({ ...viewport, x: 1600, y: 1584 });
    await act(async () => target.resolve({
      ...frame(0, requested),
      displayList: { ...frame(0, requested).displayList,
        grid: { startRow: 90, startCol: 24, rowOffsets: [576, 600], colOffsets: [704, 800] } },
    }));
    expect(await selected).toBe(true);
    const outline = view.getByTestId('xlsx-selection') as HTMLElement;
    expect(outline.style.left).toBe('704px');
    expect(outline.style.top).toBe('576px');
  });

  it('accepts painted frozen-pane geometry after fractional zoom rounds the reveal scroll', async () => {
    const { viewer, call } = session();
    call.sheetView.mockResolvedValue({ ...sheetView(0),
      frozenRows: 1, frozenCols: 1, frozenWidth: 96, frozenHeight: 24 });
    call.cellGeometry.mockImplementation(async (sheet, row, col) => ({
      sheet, version: 'v1', rect: { x: col * 96, y: row * 24, w: 96, h: 24 },
      scrollPosition: { x: Math.max(0, col * 96 - 96), y: Math.max(0, row * 24 - 24) },
    }));
    open(viewer);
    let api!: XlsxWorkerViewerApi;
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    await act(async () => { await api.commands.execute('zoom', { scale: 1.3 }); });
    await tick();
    const scroll = view.getByTestId('xlsx-scroll');
    for (const property of ['scrollLeft', 'scrollTop']) {
      let value = 0;
      Object.defineProperty(scroll, property, {
        configurable: true, get: () => value, set: (next: number) => { value = Math.round(next); },
      });
    }
    call.frame.mockImplementation(async (viewport, options) => ({
      ...frame(options?.sheet ?? 0, viewport),
      displayList: { ...frame(options?.sheet ?? 0, viewport).displayList,
        grid: { startRow: 0, startCol: 0, colIndices: [0, 2, 3],
          rowOffsets: [0, 24, 48 - viewport.y, 72 - viewport.y, 96 - viewport.y],
          colOffsets: [0, 96, 288 - viewport.x, 384 - viewport.x] } },
    }));
    let selected!: Promise<boolean>;
    await act(async () => { selected = api.selectCellsAsync(0, xlsx.selectionAt({ row: 2, col: 2 })); });
    expect(scroll.scrollLeft).toBe(125);
    await tick();
    expect(await selected).toBe(true);
    expect(parseFloat((view.getByTestId('xlsx-selection') as HTMLElement).style.left)).toBeCloseTo(96 * 1.3);
  });

  it('accepts a painted target with a positive partial viewport intersection', async () => {
    const { viewer, call } = session();
    call.cellGeometry.mockResolvedValue({ sheet: 0, version: 'v1',
      rect: { x: 0, y: 0, w: 900, h: 700 }, scrollPosition: { x: 0, y: 0 } });
    call.frame.mockImplementation(async (viewport, options) => ({
      ...frame(options?.sheet ?? 0, viewport),
      displayList: { ...frame(options?.sheet ?? 0, viewport).displayList,
        grid: { startRow: 0, startCol: 0, rowOffsets: [0, 700], colOffsets: [0, 900] } },
    }));
    open(viewer);
    let api!: XlsxWorkerViewerApi;
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    let selected!: Promise<boolean>;
    await act(async () => { selected = api.selectCellsAsync(0, xlsx.selectionAt({ row: 0, col: 0 })); });
    await tick();
    expect(await selected).toBe(true);
    expect(view.getByTestId('xlsx-selection')).toBeDefined();
  });

  it('selects hidden B1 inside a visible A1:B1 merge after painting', async () => {
    const { viewer, call } = session();
    call.sheetView.mockResolvedValue({ ...sheetView(0), contentWidth: 96, contentHeight: 24 });
    call.cellGeometry.mockResolvedValue({ sheet: 0, version: 'v1',
      rect: { x: 96, y: 0, w: 0, h: 24 }, scrollPosition: { x: 96, y: 0 } });
    call.frame.mockImplementation(async (viewport, options) => ({
      ...frame(options?.sheet ?? 0, viewport),
      mergedRanges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }],
      displayList: { ...frame(options?.sheet ?? 0, viewport).displayList,
        grid: { startRow: 0, startCol: 0, rowOffsets: [0, 24], colOffsets: [0, 96, 96] } },
    }));
    open(viewer);
    let api!: XlsxWorkerViewerApi;
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    const scroll = view.getByTestId('xlsx-scroll');
    clampScrolling(scroll);
    let selected!: Promise<boolean>;
    await act(async () => { selected = api.selectCellsAsync(0, xlsx.selectionAt({ row: 0, col: 1 })); });
    await tick();
    expect(await selected).toBe(true);
    const outline = view.getByTestId('xlsx-selection') as HTMLElement;
    expect(outline.style.left).toBe('0px');
    expect(outline.style.width).toBe('96px');
  });

  it('returns false when a matching paint does not visibly contain the target geometry', async () => {
    const { viewer, call } = session();
    call.sheetView.mockResolvedValue({ ...sheetView(0), contentWidth: 96, contentHeight: 24 });
    open(viewer);
    let api!: XlsxWorkerViewerApi;
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    const scroll = view.getByTestId('xlsx-scroll');
    for (const property of ['scrollLeft', 'scrollTop']) {
      Object.defineProperty(scroll, property, { configurable: true, get: () => 0, set: () => {} });
    }
    let selected!: Promise<boolean>;
    await act(async () => { selected = api.selectCellsAsync(0, xlsx.selectionAt({ row: 90, col: 24 })); });
    await tick();
    expect(call.frame.mock.calls[1][0]).toEqual(viewport);
    expect(painted).toHaveLength(2);
    expect(await selected).toBe(false);
    expect(view.queryByTestId('xlsx-selection')).toBeNull();
  });

  it('cancels an async selection replaced by a newer navigation', async () => {
    const { viewer, call } = session();
    open(viewer);
    let api!: XlsxWorkerViewerApi;
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    const geometry = deferred<Awaited<ReturnType<WorkbookSession['call']['cellGeometry']>>>();
    call.cellGeometry.mockImplementation(() => geometry.promise);
    const selected = api.selectCellsAsync(1, xlsx.selectionAt({ row: 2, col: 2 }));
    fireEvent.click(view.getByRole('tab', { name: 'First' }));
    await act(async () => geometry.resolve({ sheet: 1, version: 'v1', rect: { x: 192, y: 48, w: 96, h: 24 },
      scrollPosition: { x: 192, y: 48 } }));
    expect(await selected).toBe(false);
  });

  it('refuses stale selection completion when the revealed viewport is scrolled away', async () => {
    const { viewer, call } = session();
    open(viewer);
    let api!: XlsxWorkerViewerApi;
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    const target = deferred<WorkbookFrame>();
    call.frame.mockImplementationOnce(() => target.promise);
    let selected!: Promise<boolean>;
    await act(async () => { selected = api.selectCellsAsync(0, xlsx.selectionAt({ row: 2, col: 2 })); });
    await tick();
    const requested = call.frame.mock.calls[1][0];
    const scroll = view.getByTestId('xlsx-scroll');
    scroll.scrollLeft = 2000;
    fireEvent.scroll(scroll);
    await tick();
    await act(async () => target.resolve(frame(0, requested)));
    expect(await selected).toBe(false);
    expect(painted).toHaveLength(2);
  });

  it('matches canvas size, scroll extent and DPR times zoom paint coordinates', async () => {
    const { viewer, call } = session();
    call.sheetView.mockResolvedValue({ ...sheetView(0), initialScrollX: 120, initialScrollY: 48,
      frozenRows: 1, frozenCols: 1, frozenWidth: 96, frozenHeight: 24 });
    open(viewer);
    const original = Object.getOwnPropertyDescriptor(window, 'devicePixelRatio');
    Object.defineProperty(window, 'devicePixelRatio', { configurable: true, value: 2 });
    restorers.push(() => {
      if (original) Object.defineProperty(window, 'devicePixelRatio', original);
      else Reflect.deleteProperty(window, 'devicePixelRatio');
    });
    let api!: XlsxWorkerViewerApi;
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    await act(async () => { await api.commands.execute('zoom', { scale: 1.5 }); });
    await tick();
    const last = painted[painted.length - 1];
    expect(last.canvas.width).toBe(1600);
    expect(last.canvas.height).toBe(1200);
    expect(last.canvas.style.width).toBe('800px');
    expect(last.canvas.style.height).toBe('600px');
    expect(last.scale).toBe(3);
    expect(call.frame.mock.calls[call.frame.mock.calls.length - 1][0]).toEqual({
      x: 80, y: 32, width: 800 / 1.5, height: 400,
    });
    const spacer = view.getByTestId('xlsx-scroll').firstElementChild as HTMLElement;
    expect(spacer.style.width).toBe('6000px');
    expect(spacer.style.height).toBe('9000px');
    act(() => resized([], {} as ResizeObserver));
    expect(animationFrames.size).toBe(1);
  });

  it('uses painted merge geometry for pointer and keyboard selections and accessibility', async () => {
    const { viewer, call } = session();
    call.frame.mockImplementation(async (viewport, options) => ({
      ...frame(options?.sheet ?? 0, viewport),
      mergedRanges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }],
    }));
    open(viewer);
    let api!: XlsxWorkerViewerApi;
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    expect((view.getByTestId('xlsx-selection') as HTMLElement).style.width).toBe('192px');
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.mouseDown(scroll, { clientX: 200, clientY: 30 });
    fireEvent.mouseUp(window);
    expect((view.getByTestId('xlsx-selection') as HTMLElement).style.left).toBe('192px');
    fireEvent.keyDown(scroll, { key: 'ArrowDown', shiftKey: true });
    expect((view.getByTestId('xlsx-selection') as HTMLElement).style.height).toBe('48px');
    expect(view.getAllByRole('gridcell').filter((cell) => cell.getAttribute('aria-selected') === 'true')).toHaveLength(2);
    act(() => api.clearSelection());
    expect(view.queryByTestId('xlsx-selection')).toBeNull();
    act(() => api.focus());
    expect(document.activeElement).toBe(scroll);
  });

  it('keeps old painted geometry while a scrolled frame is pending', async () => {
    const { viewer, call } = session();
    open(viewer);
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen />);
    await opened();
    const next = deferred<WorkbookFrame>();
    call.frame.mockImplementationOnce(() => next.promise);
    const scroll = view.getByTestId('xlsx-scroll');
    scroll.scrollLeft = 200;
    fireEvent.scroll(scroll);
    await tick();
    fireEvent.mouseDown(scroll, { clientX: 110, clientY: 30 });
    fireEvent.mouseUp(window);
    expect((view.getByTestId('xlsx-selection') as HTMLElement).style.left).toBe('96px');
    expect(view.getByRole('grid').textContent).toContain('Sheet 0');
    await act(async () => next.resolve(frame(0, { ...viewport, x: 200 })));
    expect(painted).toHaveLength(1);
  });

  it('binds custom chrome, formula reads, save and read-only command gates', async () => {
    const { viewer, call } = session();
    open(viewer);
    const saved = mock((_bytes: Uint8Array) => {});
    let api!: XlsxWorkerViewerApi;
    const toolbar = <EditorToolbar mode="commands"><ToolbarCommandButton id="save" /><EditorToolbar.FormulaBar /></EditorToolbar>;
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen toolbar={toolbar}
      onSave={saved} onReady={(value) => { api = value; }} />);
    await opened();
    await waitFor(() => expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).value).toBe('=SUM(B1:C1)'));
    expect(call.cellInputs).toHaveBeenCalledWith(0, 'A1');
    expect(api.commands.getState('bold').enabled).toBe(false);
    expect(api.commands.getState('exportPng').enabled).toBe(false);
    expect(api.commands.getState('proposalsPanel').enabled).toBe(false);
    fireEvent.click(view.getByTestId('xlsx-save'));
    await waitFor(() => expect(saved).toHaveBeenCalledWith(new Uint8Array([8, 9])));
    view.rerender(<XlsxEditor file={file} readOnly experimentalWorkerOpen toolbar={toolbar} showToolbar={false} />);
    expect(view.queryByTestId('xlsx-toolbar')).toBeNull();
  });

  it('opens external links safely and reveals internal links through cell geometry', async () => {
    const { viewer, call } = session();
    call.frame.mockImplementation(async (viewport, options) => {
      const result = frame(options?.sheet ?? 0, viewport);
      result.displayList.hyperlinks = [
        { top: 0, bottom: 0, left: 0, right: 0, externalTarget: 'https://example.com/report', tooltip: 'Report' },
        { top: 0, bottom: 0, left: 1, right: 1, location: 'Second!C3' },
      ];
      return result;
    });
    open(viewer);
    const external = spyOn(window, 'open').mockReturnValue(null);
    restorers.push(() => external.mockRestore());
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen />);
    await opened();
    const scroll = view.getByTestId('xlsx-scroll');
    for (const x of [20, 110]) {
      fireEvent.mouseDown(scroll, { clientX: x, clientY: 12 });
      fireEvent.mouseUp(window);
      fireEvent.click(scroll, { clientX: x, clientY: 12 });
    }
    expect(external).toHaveBeenCalledWith('https://example.com/report', '_blank', 'noopener,noreferrer');
    await waitFor(() => expect(call.cellGeometry).toHaveBeenCalledWith(1, 2, 2));
    await tick();
    expect(view.getByRole('tab', { name: 'Second' }).getAttribute('aria-selected')).toBe('true');
    expect(viewer.state.activeSheet).toBe(0);
  });

  it('surfaces opening and session failures without any main-thread fallback', async () => {
    const { viewer, fail } = session();
    const opener = open(viewer);
    const errors = mock((_error: Error) => {});
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen onError={errors} />);
    await opened();
    act(() => fail());
    expect(view.getByRole('alert').textContent).toContain('Worker stopped');
    const failure = new Error('Worker unavailable');
    opener.mockRejectedValue(failure);
    view.rerender(<XlsxEditor file={new Uint8Array([4])} readOnly experimentalWorkerOpen onError={errors} />);
    await waitFor(() => expect(view.getByRole('alert').textContent).toContain(failure.message));
    expect(errors).toHaveBeenCalledWith(failure);
    expect(xlsx.openWorkbook).not.toHaveBeenCalled();
    expect(xlsx.initWasm).not.toHaveBeenCalled();
  });

  it('reports unsupported collaboration without opening a session', async () => {
    const opener = open(session().viewer);
    const collaboration = { collaboration: { clientId: 1 } };
    const view = render(<XlsxEditor file={file} readOnly experimentalWorkerOpen {...collaboration} />);
    await waitFor(() => expect(view.getByRole('alert').textContent).toContain('Collaboration is unavailable'));
    expect(opener).not.toHaveBeenCalled();
  });

  for (const props of [
    { readOnly: false, experimentalWorkerOpen: true },
    { readOnly: true, experimentalWorkerOpen: false },
    { readOnly: true },
  ]) it(`keeps local dispatch without both flags (${JSON.stringify(props)})`, async () => {
    const opener = open(session().viewer);
    render(<XlsxEditor {...props as XlsxEditorProps} />);
    expect(opener).not.toHaveBeenCalled();
    expect(painted.length).toBeGreaterThanOrEqual(1);
  });
});
