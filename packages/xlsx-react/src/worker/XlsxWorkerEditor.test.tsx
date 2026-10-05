import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, describe, expect, it, mock, spyOn } from 'bun:test';
import * as xlsx from '@betteroffice/xlsx';
import type { DisplayList, EditResult, WorkbookEditPeer, WorkbookHandle, WorkbookSession, XlsxEditRequest } from '@betteroffice/xlsx';
import { SessionFailure } from '../../../../shared/office-session';
import { XlsxEditor } from '../XlsxEditor';
import { EditorToolbar } from '../components/EditorToolbar';
import { defineXlsxPlugin } from '../plugins/defineXlsxPlugin';
import type { XlsxPluginContext } from '../plugins/types';
import { XlsxWorkerEditorCollaborationError, type XlsxWorkerEditorApi } from './createWorkerEditorApi';
import { editableWorkbookSessionBackend } from './useEditableSessionWorkbook';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
const file = new Uint8Array([1, 2, 3]);
let animationFrames: Map<number, FrameRequestCallback>;
let nextAnimation = 0;
let painted: DisplayList[];
let observers: { disconnect: ReturnType<typeof mock> }[];

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, failed) => { resolve = done; reject = failed; });
  return { promise, resolve, reject };
}

async function tick() {
  await act(async () => {
    const callbacks = [...animationFrames.values()];
    animationFrames.clear();
    for (const callback of callbacks) callback(0);
  });
}

async function advance(count = 8) {
  for (let index = 0; index < count; index++) await tick();
}

function harness() {
  const log: string[] = [];
  const cells = new Map<string, string>([['0:0:0', 'initial']]);
  const charts: xlsx.ChartRegion[] = [];
  const updates = new Set<() => void>();
  let active = 0;
  let sequence = 0;
  let version = 1;
  let failure: SessionFailure | undefined;
  let failureListener: Parameters<WorkbookSession['onFailure']>[0] | undefined;
  const info = () => ({ sheetIds: ['sheet:0', 'sheet:1'], sheetNames: ['First', 'Second'], activeSheet: active,
    contentWidth: 4000, contentHeight: 6000, frozenRows: 0, frozenCols: 0, initialScrollX: 0, initialScrollY: 0 });
  const cell = (sheet: number, row: number, col: number) => ({ a1: `${String.fromCharCode(65 + col)}${row + 1}`,
    input: cells.get(`${sheet}:${row}:${col}`) ?? '', isFormula: (cells.get(`${sheet}:${row}:${col}`) ?? '').startsWith('=') });
  const display = (source: string): DisplayList => ({ width: 800, height: 600,
    commands: Array.from(cells.entries()).filter(([key]) => key.startsWith(`${active}:`)).map(([key, text]) => {
      const [, row, col] = key.split(':').map(Number);
      return { op: 'text' as const, text: `${source}:${text}`, x: col * 96 + 8, y: row * 24 + 18,
        fontSize: 11, color: '#000000', clip: { x: col * 96, y: row * 24, w: 96, h: 24 } };
    }),
    grid: { startRow: 0, startCol: 0, rowOffsets: [0, 24, 48, 72, 96], colOffsets: [0, 96, 192, 288] },
    charts,
  });
  const result = (): EditResult => ({ applied: true, sheetInfo: info() });
  const mutate = (sheet: number, row: number, col: number, input: string) => {
    log.push(`edit:${sheet}:${row}:${col}:${input}`);
    cells.set(`${sheet}:${row}:${col}`, input);
    version += 1;
    sequence += 1;
    for (const update of updates) update();
    void Promise.resolve().then(() => { log.push(`replay:${input}`); });
    return result();
  };
  const peerMethods = {
    sheetInfo: info, sheetCount: () => 2, version: () => `v${version}`, cell,
    rangeCells: (sheet: number, range: string) => {
      const match = /^([A-Z])(\d+)/.exec(range)!;
      return [[cell(sheet, Number(match[2]) - 1, match[1].charCodeAt(0) - 65)]];
    },
    cellPosition: (_sheet: number, row: number, col: number) => ({ x: col * 96, y: row * 24 }),
    displayList: mock(() => { log.push('display:peer'); return display('peer'); }),
    mergedRanges: () => [], listProposals: () => [],
    historyState: () => ({ canUndo: true, canRedo: true }),
    selectionFormatting: () => null,
    onUpdate: (listener: () => void) => { updates.add(listener); return () => { updates.delete(listener); }; },
    readCells: () => ({ ok: true, version: `v${version}`, sheets: [], ranges: [],
      calculation: { cycleCells: [], limitedCells: [], truncated: false } }),
    dispose: mock(() => { log.push('dispose:peer'); }),
    editCell: mock(() => { throw new Error('Direct peer mutation forbidden'); }),
    applyEdits: mock(() => { throw new Error('Direct peer batch forbidden'); }),
    setActiveSheet: mock(() => { throw new Error('Direct peer navigation forbidden'); }),
    save: mock(() => { throw new Error('Full peer save forbidden'); }),
  };
  const peer = peerMethods as unknown as WorkbookHandle;
  const editMethods = {
    get state() { return failure ? 'failed' as const : 'ready' as const; },
    get error() { return failure; },
    get sentSequence() { return sequence; },
    editCell: mock(mutate),
    editCells: mock((sheet: number, edits: xlsx.CellInputEdit[]) => {
      log.push('batch');
      for (const edit of edits) mutate(sheet, edit.row, edit.col, edit.input);
      return result();
    }),
    applyEdits: mock((_request: XlsxEditRequest) => {
      log.push('plugin:batch');
      mutate(0, 0, 0, 'plugin');
      return { ok: true, applied: true, baseVersion: 'v1', version: `v${version}`, source: 'host', receipts: [],
        changedSheets: [], calculation: { changed: [], cycleCells: [], limitedCells: [], truncated: false } };
    }),
    setActiveSheet: mock((sheet: number) => { active = sheet; sequence += 1; log.push(`sheet:${sheet}`); }),
    moveChart: mock((_sheet: number, id: string, dx: number, dy: number) => {
      sequence += 1; log.push(`chart:${id}:${dx}:${dy}`); return result();
    }),
    undo: mock(() => { sequence += 1; log.push('undo'); return result(); }),
    flush: mock(async () => {}), save: mock(async () => new Uint8Array([8, 9]).buffer),
    recoverySave: mock(() => ({ bytes: new Uint8Array([7]).buffer, recovery: true as const })),
    dispose: mock(() => { log.push('dispose:facade'); }),
  };
  const edits = editMethods as unknown as WorkbookEditPeer;
  const sessionMethods = {
    frame: mock(async (viewport: xlsx.Viewport, options?: { sheet?: number }): Promise<xlsx.WorkbookFrame> => {
      log.push('frame:worker');
      return { sheet: options?.sheet ?? active, viewport, version: `v${version}`, epoch: 1,
        sequence, mergedRanges: [], displayList: display('worker') };
    }),
    sheetView: mock(async (sheet: number) => ({ ...info(), sheet, version: `v${version}`, frozenWidth: 0, frozenHeight: 0 })),
    cellInputs: mock(async (sheet: number, range: string) => ({ sheet, version: `v${version}`,
      cells: peerMethods.rangeCells(sheet, range) })),
  };
  const session: WorkbookSession = {
    state: { format: 'xlsx', stage: 'ready', version: 0, dirty: false, activeSheet: 0,
      sheets: [{ id: 'sheet:0', index: 0, name: 'First' }, { id: 'sheet:1', index: 1, name: 'Second' }] },
    call: sessionMethods as unknown as WorkbookSession['call'], save: mock(async () => new Uint8Array([0])),
    on: () => () => {},
    onFailure(listener) { failureListener = listener; return () => { log.push('off:failure'); failureListener = undefined; }; },
    get failure() { return failure; },
    dispose: mock(async () => { log.push('dispose:session'); }),
  };
  const open = spyOn(editableWorkbookSessionBackend, 'open').mockResolvedValue(session);
  const hydrate = spyOn(editableWorkbookSessionBackend, 'hydrate').mockResolvedValue(peer);
  const attach = spyOn(editableWorkbookSessionBackend, 'attach').mockReturnValue(edits);
  restorers.push(() => open.mockRestore(), () => hydrate.mockRestore(), () => attach.mockRestore());
  return { log, cells, charts, peer, peerMethods, edits, editMethods, session, sessionMethods, open, hydrate,
    fail(error = new SessionFailure('crash', 'Worker stopped')) { failure = error; failureListener?.(error); return error; } };
}

async function opened() {
  await waitFor(() => expect(animationFrames.size).toBeGreaterThan(0));
  await tick();
}

beforeEach(() => {
  animationFrames = new Map(); painted = []; observers = [];
  const originalRaf = globalThis.requestAnimationFrame;
  const originalCancel = globalThis.cancelAnimationFrame;
  const originalResize = globalThis.ResizeObserver;
  globalThis.requestAnimationFrame = (callback) => {
    const id = ++nextAnimation; animationFrames.set(id, callback); return id;
  };
  globalThis.cancelAnimationFrame = (id) => { animationFrames.delete(id); };
  globalThis.ResizeObserver = class {
    disconnect = mock(() => {});
    constructor() { observers.push(this); }
    observe() {}
  } as unknown as typeof ResizeObserver;
  restorers.push(() => {
    globalThis.requestAnimationFrame = originalRaf; globalThis.cancelAnimationFrame = originalCancel;
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
  const paint = spyOn(xlsx, 'paintDisplayList').mockImplementation((_context, list) => { painted.push(list); });
  const local = spyOn(xlsx, 'openWorkbook').mockImplementation(() => { throw new Error('Local open forbidden'); });
  const init = spyOn(xlsx, 'initWasm').mockRejectedValue(new Error('Local wasm forbidden'));
  const proposals = spyOn(xlsx, 'isProposalsAvailable').mockReturnValue(false);
  restorers.push(() => context.mockRestore(), () => paint.mockRestore(), () => local.mockRestore(),
    () => init.mockRestore(), () => proposals.mockRestore());
});
afterEach(() => {
  cleanup();
  for (const restore of restorers.reverse()) restore();
  restorers.length = 0;
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

describe('workbook worker editor', () => {
  it('opens retained worker editor sessions with omitted and explicit false readOnly', async () => {
    for (const props of [{ experimentalWorkerOpen: true as const }, { readOnly: false as const, experimentalWorkerOpen: true as const }]) {
      const host = harness();
      const ready = mock((_api: XlsxWorkerEditorApi) => {});
      const view = render(<XlsxEditor file={file} {...props} onReady={ready} showToolbar={false} />);
      await opened();
      expect(host.open).toHaveBeenCalledWith(file, { signal: expect.any(AbortSignal), retainPeerHydration: true });
      expect(painted.length).toBeGreaterThan(0);
      expect(ready).toHaveBeenCalledTimes(1);
      expect(host.hydrate).toHaveBeenCalledTimes(1);
      expect(xlsx.openWorkbook).not.toHaveBeenCalled();
      expect(xlsx.initWasm).not.toHaveBeenCalled();
      view.unmount();
      for (const restore of restorers.splice(-3).reverse()) restore();
    }
  });

  it('shows DOM text across two frames before sync mutation, peer publication and the worker frame', async () => {
    const host = harness();
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    host.log.length = 0; painted.length = 0;
    host.editMethods.editCell.mockImplementation((sheet, row, col, input) => {
      expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('typed');
      expect(painted).toHaveLength(0);
      host.log.push('sync'); host.cells.set(`${sheet}:${row}:${col}`, input);
      return { applied: true, sheetInfo: host.peer.sheetInfo() };
    });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 't' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'typed' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('typed');
    await act(async () => {});
    await tick();
    expect(host.editMethods.editCell).not.toHaveBeenCalled();
    await tick();
    expect(host.log.indexOf('sync')).toBeLessThan(host.log.indexOf('display:peer'));
    expect(painted[0].commands.some((command) => command.op === 'text' && command.text === 'peer:typed')).toBe(true);
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    expect(host.log).not.toContain('frame:worker');
    await tick();
    expect(host.log.indexOf('display:peer')).toBeLessThan(host.log.indexOf('frame:worker'));
    expect(painted[painted.length - 1].commands.some((command) => command.op === 'text' && command.text === 'worker:typed')).toBe(true);
  });

  it('keeps typing, caret and accepted commits while eager hydration waits', async () => {
    const host = harness();
    const pending = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(pending.promise);
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'a' });
    const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    fireEvent.change(input, { target: { value: 'before hydration' } });
    input.setSelectionRange(3, 3);
    await act(async () => pending.resolve(host.peer));
    expect(document.activeElement).toBe(input);
    expect(input.value).toBe('before hydration');
    expect(input.selectionStart).toBe(3);
    fireEvent.keyDown(input, { key: 'Tab' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'b' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await advance();
    expect(host.editMethods.editCell.mock.calls.map((args) => args.slice(0, 4))).toEqual([
      [0, 0, 0, 'before hydration'], [0, 0, 1, 'b'],
    ]);
  });

  it('replays formula bar and gesture-captured clipboard edits in order', async () => {
    const host = harness();
    const reading = deferred<string>();
    const clipboard = spyOn(navigator.clipboard, 'readText').mockReturnValue(reading.promise);
    restorers.push(() => clipboard.mockRestore());
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen
      toolbar={<EditorToolbar mode="commands"><EditorToolbar.FormulaBar /></EditorToolbar>} />);
    await opened();
    fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value: '=12' } });
    fireEvent.keyDown(view.getByTestId('xlsx-formula-input'), { key: 'Enter' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
    expect(clipboard).toHaveBeenCalledTimes(1);
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'z' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'last' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Tab' });
    await act(async () => reading.resolve('pasted'));
    await advance(12);
    expect(host.log.filter((entry) => entry.startsWith('replay:'))).toEqual(['replay:=12', 'replay:pasted', 'replay:last']);
    expect(host.peerMethods.editCell).not.toHaveBeenCalled();
    expect(host.peerMethods.save).not.toHaveBeenCalled();
  });

  it('preserves later typing and caret when a pre-hydration cell prefill arrives', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    const prefill = deferred<xlsx.WorkbookCellInputs>();
    host.hydrate.mockReturnValue(hydration.promise);
    host.sessionMethods.cellInputs.mockImplementation(() => prefill.promise);
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    fireEvent.change(input, { target: { value: 'later typing' } });
    input.setSelectionRange(2, 2);
    await act(async () => prefill.resolve({ sheet: 0, version: 'v1',
      cells: [[{ a1: 'A1', input: 'old prefill', isFormula: false }]] }));
    await act(async () => hydration.resolve(host.peer));
    expect(input.value).toBe('later typing');
    expect(input.selectionStart).toBe(2);
    expect(document.activeElement).toBe(input);
  });

  it('keeps a later live draft open when an earlier accepted commit publishes', async () => {
    const host = harness();
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'a' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Tab' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'b' });
    const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    fireEvent.change(input, { target: { value: 'still typing' } });
    input.setSelectionRange(4, 4);
    await advance();
    expect(host.editMethods.editCell).toHaveBeenCalledTimes(1);
    expect(view.getByTestId('xlsx-cell-editor')).toBe(input);
    expect(input.value).toBe('still typing');
    expect(input.selectionStart).toBe(4);
    expect(document.activeElement).toBe(input);
  });

  it('rejects stale worker frames without erasing pending text or peer pixels', async () => {
    const host = harness();
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    const stale = await host.sessionMethods.frame({ x: 0, y: 0, width: 800, height: 600 }, { sheet: 0 });
    const pending = deferred<xlsx.WorkbookFrame>();
    host.sessionMethods.frame.mockImplementationOnce(() => pending.promise);
    fireEvent.scroll(view.getByTestId('xlsx-scroll'));
    await tick();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'new' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await act(async () => {});
    await tick();
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('new');
    await tick();
    expect(painted[painted.length - 1].commands.some((command) => command.op === 'text' && command.text === 'peer:new')).toBe(true);
    const count = painted.length;
    await act(async () => pending.resolve(stale));
    expect(painted).toHaveLength(count);
    await advance();
    expect(painted[painted.length - 1].commands.some((command) => command.op === 'text' && command.text === 'worker:new')).toBe(true);
  });

  it('orders sheet navigation after accepted edits and mutates only the facade', async () => {
    const host = harness();
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 's' });
    fireEvent.click(view.getByRole('tab', { name: 'Second' }));
    await advance();
    expect(host.log.indexOf('edit:0:0:0:s')).toBeLessThan(host.log.indexOf('sheet:1'));
    expect(host.editMethods.setActiveSheet).toHaveBeenCalledWith(1);
    expect(host.peerMethods.setActiveSheet).not.toHaveBeenCalled();
    expect(view.getByRole('tab', { name: 'Second' }).getAttribute('aria-selected')).toBe('true');
  });

  it('commits a chart drag through the ordered facade after its preview', async () => {
    const host = harness();
    host.charts.push({ id: 'chart:1', label: 'Chart', movable: true,
      rect: { x: 100, y: 100, w: 200, h: 100 }, clip: { x: 100, y: 100, w: 200, h: 100 } });
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { clientX: 120, clientY: 120, button: 0 });
    fireEvent.mouseMove(view.getByTestId('xlsx-scroll'), { clientX: 150, clientY: 140, buttons: 1 });
    expect(view.getByTestId('xlsx-chart-selection').style.transform).toBe('translate(30px, 20px)');
    fireEvent.mouseUp(window, { clientX: 150, clientY: 140, button: 0 });
    await act(async () => {});
    await tick();
    expect(host.editMethods.moveChart).not.toHaveBeenCalled();
    await tick();
    expect(host.editMethods.moveChart).toHaveBeenCalledWith(0, 'chart:1', 30, 20);
  });

  it('saves pending formula input through the facade and keeps PNG unavailable', async () => {
    const host = harness();
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value: '=24' } });
    host.editMethods.save.mockImplementation(async () => {
      expect(host.cells.get('0:0:0')).toBe('=24');
      expect(painted[painted.length - 1].commands.some((command) => command.op === 'text' && command.text === 'peer:=24')).toBe(true);
      return new Uint8Array([8, 9]).buffer;
    });
    let saving!: Promise<Uint8Array | null>;
    act(() => { saving = api.save(); });
    await advance();
    expect(await saving).toEqual(new Uint8Array([8, 9]));
    expect(host.editMethods.save).toHaveBeenCalledTimes(1);
    expect(host.peerMethods.save).not.toHaveBeenCalled();
    expect(api.commands.getState('exportPng')).toMatchObject({ enabled: false, disabledReason: { code: 'png-unavailable' } });
  });

  it('prints after the accepted edited viewport has been painted', async () => {
    const host = harness();
    const print = spyOn(window, 'print').mockImplementation(() => {
      expect(painted[painted.length - 1].commands.some((command) => command.op === 'text' && command.text.endsWith(':print'))).toBe(true);
    });
    restorers.push(() => print.mockRestore());
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'p' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'print' } });
    let printing!: ReturnType<XlsxWorkerEditorApi['commands']['execute']>;
    act(() => { printing = api.commands.execute('print', null); });
    expect(print).not.toHaveBeenCalled();
    await advance();
    expect(await printing).toMatchObject({ ok: true });
    expect(print).toHaveBeenCalledTimes(1);
  });

  it('surfaces typed failure once and stops accepting cell, formula and clipboard input', async () => {
    const host = harness();
    const errors = mock((_error: Error) => {});
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen onError={errors}
      onReady={(value) => { api = value; }} />);
    await opened();
    const failure = new SessionFailure('crash', 'Worker stopped');
    act(() => host.fail(failure));
    expect(view.getByRole('alert').textContent).toContain('Worker stopped');
    expect(api.failure).toBe(failure);
    await expect(api.editCellAsync(0, 0, 0, 'host')).rejects.toBe(failure);
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x' });
    fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value: 'formula' } });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
    await advance();
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    expect(host.editMethods.editCell).not.toHaveBeenCalled();
    expect(host.editMethods.editCells).not.toHaveBeenCalled();
    expect(errors).toHaveBeenCalledTimes(1);
  });

  it('disposes readiness cleanup, input, paints, observers, facade, peer and session on unmount', async () => {
    const host = harness();
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
      onReady={() => () => { host.log.push('dispose:ready'); }} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'q' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await act(async () => {});
    view.unmount();
    await advance();
    expect(animationFrames.size).toBe(0);
    expect(observers.every((observer) => observer.disconnect.mock.calls.length === 1)).toBe(true);
    expect(host.log.filter((entry) => entry.startsWith('dispose:'))).toEqual([
      'dispose:ready', 'dispose:facade', 'dispose:peer', 'dispose:session',
    ]);
    expect(host.editMethods.editCell).not.toHaveBeenCalled();
  });

  it('activates plugins after the real peer and routes granted batches through the facade', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    let context!: XlsxPluginContext<null>;
    const initialize = mock((value: XlsxPluginContext<null>) => { context = value; });
    const plugin = defineXlsxPlugin({ id: 'review', createState: () => null, initialize });
    render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} plugins={[plugin]}
      pluginGrants={{ review: { document: 'write', editBatches: true } }} />);
    await opened();
    expect(initialize).not.toHaveBeenCalled();
    await act(async () => hydration.resolve(host.peer));
    await waitFor(() => expect(initialize).toHaveBeenCalledTimes(1));
    let applied: unknown;
    await act(async () => {
      void context.run(async (current) => {
        applied = await current.edits!.applyEdits({ expectVersion: 'v1', source: 'host', steps: [] });
      });
    });
    await advance();
    expect(applied).toMatchObject({ ok: true, applied: true });
    expect(host.editMethods.applyEdits).toHaveBeenCalledTimes(1);
    expect(host.peerMethods.applyEdits).not.toHaveBeenCalled();
  });

  it('recovers retained unapplied UI and host edits exactly once without worker writes', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    const recovery = spyOn(xlsx, 'createWorkbookRecoveryMutators').mockImplementation(() => ({
      ...host.edits, editCell: (sheet, row, col, input) => {
        host.log.push(`recover:${input}`);
        host.cells.set(`${sheet}:${row}:${col}`, input);
        return { applied: true, sheetInfo: host.peer.sheetInfo() };
      },
    }));
    restorers.push(() => recovery.mockRestore());
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
      onReady={(value) => { api = value; }} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'u' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    let hostEdit!: Promise<EditResult | null>;
    act(() => { hostEdit = api.editCellAsync(0, 0, 1, 'host'); void hostEdit.catch(() => {}); });
    act(() => host.fail());
    await expect(hostEdit).rejects.toBeInstanceOf(SessionFailure);
    await act(async () => hydration.resolve(host.peer));
    let recovered!: { bytes: Uint8Array; recovery: true };
    await act(async () => { recovered = await api.recoverySave(); });
    expect(recovered).toEqual({ bytes: new Uint8Array([7]), recovery: true });
    expect(host.log.filter((entry) => entry.startsWith('recover:'))).toEqual(['recover:u', 'recover:host']);
    expect(host.editMethods.editCell).not.toHaveBeenCalled();
    expect(host.editMethods.save).not.toHaveBeenCalled();
    await expect(api.editCellAsync(0, 0, 0, 'later')).rejects.toBeInstanceOf(SessionFailure);
  });

  it('fails recovery with a typed error when retained operations cannot be materialized', async () => {
    const host = harness();
    let api!: XlsxWorkerEditorApi;
    render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
    await opened();
    act(() => host.fail());
    const recovery = spyOn(xlsx, 'createWorkbookRecoveryMutators').mockImplementation(() => {
      throw new xlsx.WorkbookEditPeerFailedError(new Error('Retained operation unavailable'));
    });
    restorers.push(() => recovery.mockRestore());
    await expect(api.recoverySave()).rejects.toBeInstanceOf(xlsx.WorkbookEditPeerFailedError);
    expect(host.editMethods.recoverySave).not.toHaveBeenCalled();
  });

  it('includes the retained live draft in recovery and continues refusing normal input', async () => {
    const host = harness();
    const recovery = spyOn(xlsx, 'createWorkbookRecoveryMutators').mockImplementation(() => ({
      ...host.edits, editCell: (sheet, row, col, input) => {
        host.log.push(`recover:${input}`);
        host.cells.set(`${sheet}:${row}:${col}`, input);
        return { applied: true, sheetInfo: host.peer.sheetInfo() };
      },
    }));
    restorers.push(() => recovery.mockRestore());
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
      onReady={(value) => { api = value; }} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'l' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'live draft' } });
    act(() => host.fail());
    await act(async () => { await api.recoverySave(); });
    expect(host.log.filter((entry) => entry.startsWith('recover:'))).toEqual(['recover:live draft']);
    expect(host.editMethods.editCell).not.toHaveBeenCalled();
    expect(api.hydrated).toBe(false);
    await expect(api.editCellAsync(0, 0, 0, 'later')).rejects.toBeInstanceOf(SessionFailure);
  });

  it('rejects collaboration with a typed error before opening a worker session', async () => {
    const host = harness();
    const errors = mock((_error: Error) => {});
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen onError={errors}
      {...{ collaboration: { clientId: 1 } }} />);
    await waitFor(() => expect(view.getByRole('alert').textContent).toContain('Collaboration is unavailable'));
    expect(errors.mock.calls[0][0]).toBeInstanceOf(XlsxWorkerEditorCollaborationError);
    expect(host.open).not.toHaveBeenCalled();
  });
});
