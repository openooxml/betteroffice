import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, describe, expect, it, mock, spyOn } from 'bun:test';
import * as xlsx from '@betteroffice/xlsx';
import type { DisplayList, EditResult, WorkbookEditPeer, WorkbookHandle, WorkbookSession, XlsxEditRequest } from '@betteroffice/xlsx';
import { SessionFailure } from '../../../../shared/office-session';
import { workbookEditPeerInternals } from '../../../xlsx/src/session/editPeerInternals';
import { workbookSessionInternals, type WorkbookReplayEnvelope, type WorkbookReplayReply } from '../../../xlsx/src/session/replay';
import { XlsxEditor } from '../XlsxEditor';
import { EditorToolbar } from '../components/EditorToolbar';
import { defineXlsxPlugin } from '../plugins/defineXlsxPlugin';
import type { XlsxPluginContext } from '../plugins/types';
import { workbookSessionOpener } from '../viewer/useSessionWorkbook';
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

function harness(realFacade = false) {
  const log: string[] = [];
  const cells = new Map<string, string>([['0:0:0', 'initial']]);
  const workerCells = new Map(cells);
  const charts: xlsx.ChartRegion[] = [];
  const updates = new Set<() => void>();
  let active = 0;
  let sequence = 0;
  let version = 1;
  let failure: SessionFailure | undefined;
  const failureListeners = new Set<Parameters<WorkbookSession['onFailure']>[0]>();
  const info = () => ({ sheetIds: ['sheet:0', 'sheet:1'], sheetNames: ['First', 'Second'], activeSheet: active,
    contentWidth: 4000, contentHeight: 6000, frozenRows: 0, frozenCols: 0, initialScrollX: 0, initialScrollY: 0 });
  const cell = (sheet: number, row: number, col: number) => ({ a1: `${String.fromCharCode(65 + col)}${row + 1}`,
    input: cells.get(`${sheet}:${row}:${col}`) ?? '', isFormula: (cells.get(`${sheet}:${row}:${col}`) ?? '').startsWith('=') });
  const display = (source: string, values = cells): DisplayList => ({ width: 800, height: 600,
    commands: Array.from(values.entries()).filter(([key]) => key.startsWith(`${active}:`)).map(([key, text]) => {
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
    if (sheet < 0 || sheet > 1 || row < 0 || row > 1048575 || col < 0 || col > 16383) throw new RangeError('Invalid cell target');
    if (input.length > 32767) throw new Error('Cell text exceeds Excel length limit');
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
    readCells: mock((_request: xlsx.XlsxReadRequest): xlsx.XlsxReadResult => ({
      ok: true, version: `v${version}`, sheets: [], ranges: [],
      calculation: { cycleCells: [], limitedCells: [], truncated: false },
    })),
    dispose: mock(() => { log.push('dispose:peer'); }),
    setCalculationContext: mock(() => {}),
    editCell: mock((sheet: number, row: number, col: number, input: string) => {
      if (realFacade) return mutate(sheet, row, col, input);
      throw new Error('Direct peer mutation forbidden');
    }),
    applyEdits: mock((request: XlsxEditRequest): xlsx.XlsxEditResult => {
      if (!realFacade) throw new Error('Direct peer batch forbidden');
      if (request.expectVersion !== `v${version}`) return { ok: false, version: `v${version}`,
        failure: { code: 'stale-version', message: 'Version changed' } };
      log.push('plugin:batch');
      for (const step of request.steps) {
        if (step.op !== 'setCellInputs') continue;
        const start = step.target.range.kind === 'rowCol' ? step.target.range.start : { row: 0, col: 0 };
        step.inputs.forEach((row, dr) => row.forEach((input, dc) => mutate(0, start.row + dr, start.col + dc, input)));
      }
      return { ok: true, applied: true, baseVersion: request.expectVersion, version: `v${version}`, source: 'host', receipts: [],
        changedSheets: [], calculation: { changed: [], cycleCells: [], limitedCells: [], truncated: false } };
    }),
    editCells: mock((sheet: number, edits: xlsx.CellInputEdit[]) => {
      if (!realFacade) throw new Error('Direct peer bulk mutation forbidden');
      log.push('batch');
      for (const edit of edits) mutate(sheet, edit.row, edit.col, edit.input);
      return result();
    }),
    patchRangeStyle: mock(() => { log.push('format'); return result(); }),
    setActiveSheet: mock(() => { throw new Error('Direct peer navigation forbidden'); }),
    moveChart: mock((_sheet: number, id: string, dx: number, dy: number) => {
      if (!realFacade) throw new Error('Direct peer chart mutation forbidden');
      const chart = charts.find((entry) => entry.id === id)!;
      chart.rect = { ...chart.rect, x: chart.rect.x + dx, y: chart.rect.y + dy };
      log.push(`recover:chart:${id}:${dx}:${dy}`);
      version += 1;
      return result();
    }),
    save: mock(() => {
      if (realFacade) return new TextEncoder().encode(JSON.stringify([...cells]));
      throw new Error('Full peer save forbidden');
    }),
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
    applyEdits: mock((_request: XlsxEditRequest): xlsx.XlsxEditResult => {
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
    onFailure(listener) {
      failureListeners.add(listener);
      return () => { log.push('off:failure'); failureListeners.delete(listener); };
    },
    get failure() { return failure; },
    dispose: mock(async () => { log.push('dispose:session'); }),
  };
  const open = spyOn(editableWorkbookSessionBackend, 'open').mockResolvedValue(session);
  const hydrate = spyOn(editableWorkbookSessionBackend, 'hydrate').mockResolvedValue(peer);
  const replay = mock(async (envelope: WorkbookReplayEnvelope): Promise<WorkbookReplayReply> => ({
    sequence: envelope.sequence, revision: version, version, result: result(),
  }));
  const preview = mock(async (viewport: xlsx.Viewport, sheet: number, ops: readonly import('../../../xlsx/src/session/replay').WorkbookReplayOp[]) => {
    const values = new Map(cells);
    for (const op of ops) if (op.method === 'editCell') {
      if (op.args[3].startsWith('=') && new TextEncoder().encode(op.args[3].slice(1)).byteLength > 32768) {
        throw new Error('Formula exceeds the length limit');
      }
      values.set(`${op.args[0]}:${op.args[1]}:${op.args[2]}`, op.args[3]);
    }
    return { sheet, viewport, version: `v${version}`, epoch: 0, sequence,
      mergedRanges: [], displayList: display('preview', values) };
  });
  workbookSessionInternals.set(session, { replay, editPeerAttached: false, preview,
    cellInput: async (sheet, row, col) => workerCells.get(`${sheet}:${row}:${col}`) ?? '' });
  if (!realFacade) workbookEditPeerInternals.set(edits, {
    fail: () => {},
    applyQueuedOp: (op) => {
      if (op.method === 'applyEdits') return editMethods.applyEdits(op.args[0]);
      if (op.method === 'editCell') return editMethods.editCell(...op.args);
      throw new Error('Unexpected mock operation');
    },
    applyRecoveryOp: (op) => {
      if (op.method === 'editCell') return xlsx.createWorkbookRecoveryMutators(edits).editCell(...op.args);
      throw new Error('Unexpected mock recovery operation');
    },
  });
  let attached: WorkbookEditPeer | null = null;
  const attach = spyOn(editableWorkbookSessionBackend, 'attach').mockImplementation((options) => {
    attached = realFacade ? xlsx.createWorkbookEditPeer(options) : edits;
    return attached;
  });
  restorers.push(() => open.mockRestore(), () => hydrate.mockRestore(), () => attach.mockRestore());
  return { log, cells, charts, peer, peerMethods, edits, editMethods, session, sessionMethods, open, hydrate, replay, preview,
    get attached() { return attached; },
    fail(error = new SessionFailure('crash', 'Worker stopped')) {
      failure = error;
      for (const listener of [...failureListeners]) listener(error);
      return error;
    } };
}

function promisedClipboard() {
  const originalItem = Object.getOwnPropertyDescriptor(globalThis, 'ClipboardItem');
  const originalWrite = Object.getOwnPropertyDescriptor(navigator.clipboard, 'write');
  class Item {
    constructor(readonly data: Record<string, Promise<Blob>>) {}
  }
  const payloads: Promise<Blob>[] = [];
  const writes: Promise<void>[] = [];
  const copied: string[] = [];
  const write = mock((items: Item[]) => {
    const payload = items[0].data['text/plain'];
    payloads.push(payload);
    const writing = payload.then(async (blob) => { copied.push(await blob.text()); });
    void writing.catch(() => {});
    writes.push(writing);
    return writing;
  });
  Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: Item });
  Object.defineProperty(navigator.clipboard, 'write', { configurable: true, value: write });
  restorers.push(() => {
    if (originalItem) Object.defineProperty(globalThis, 'ClipboardItem', originalItem);
    else Reflect.deleteProperty(globalThis, 'ClipboardItem');
    if (originalWrite) Object.defineProperty(navigator.clipboard, 'write', originalWrite);
    else Reflect.deleteProperty(navigator.clipboard, 'write');
  });
  return { write, payloads, writes, copied };
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

  it('shows the preview across two frames before mutation and retains it until worker adoption', async () => {
    const host = harness();
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    host.log.length = 0; painted.length = 0;
    const worker = deferred<xlsx.WorkbookFrame>();
    host.sessionMethods.frame.mockImplementation(async () => {
      host.log.push('frame:worker');
      return worker.promise;
    });
    const write = host.editMethods.editCell.getMockImplementation()!;
    host.editMethods.editCell.mockImplementation((sheet, row, col, input) => {
      expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('typed');
      expect(painted.some((list) => list.commands.some((command) =>
        command.op === 'text' && command.text === 'preview:typed'))).toBe(true);
      host.log.push('sync');
      return write(sheet, row, col, input);
    });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 't' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'typed' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('typed');
    await act(async () => {});
    await tick();
    expect(host.editMethods.editCell).not.toHaveBeenCalled();
    await tick();
    expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'typed']]);
    expect(host.peerMethods.displayList).not.toHaveBeenCalled();
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('typed');
    expect(host.log).not.toContain('frame:worker');
    await tick();
    expect(host.log.indexOf('sync')).toBeLessThan(host.log.indexOf('frame:worker'));
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('typed');
    const [viewport, options] = host.sessionMethods.frame.mock.calls.at(-1)!;
    await act(async () => worker.resolve({ sheet: options?.sheet ?? 0, viewport, version: 'v2', epoch: 2, sequence: 1,
      displayList: { width: 800, height: 600, commands: [{ op: 'text', text: 'worker:typed', x: 8, y: 18, fontSize: 11, color: '#000000' }] } }));
    await advance();
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    expect(host.peerMethods.displayList).not.toHaveBeenCalled();
    expect(painted.every((list) => list.commands.every((command) =>
      command.op !== 'text' || !command.text.startsWith('peer:')))).toBe(true);
    expect(painted[painted.length - 1].commands.some((command) => command.op === 'text' && command.text === 'worker:typed')).toBe(true);
  });

  it('paints an accepted commit preview while eager hydration is held before mutation', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
      onReady={(value) => { api = value; }} />);
    await opened();
    const worker = deferred<xlsx.WorkbookFrame>();
    host.sessionMethods.frame.mockReturnValue(worker.promise);
    expect(host.hydrate).toHaveBeenCalledTimes(1);
    expect(api.hydrated).toBe(false);
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 't' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'typed before hydration' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('typed before hydration');
    await advance(3);
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('typed before hydration');
    expect(api.hydrated).toBe(false);
    expect(host.editMethods.editCell).not.toHaveBeenCalled();
    expect(host.cells.get('0:0:0')).toBe('initial');
    expect(host.log).not.toContain('display:peer');
    await act(async () => hydration.resolve(host.peer));
    expect(api.hydrated).toBe(true);
    await advance();
    expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'typed before hydration']]);
    expect(host.peerMethods.displayList).not.toHaveBeenCalled();
    expect(host.cells.get('0:0:0')).toBe('typed before hydration');
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('typed before hydration');
    await advance();
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('typed before hydration');
    const [viewport, options] = host.sessionMethods.frame.mock.calls.at(-1)!;
    await act(async () => worker.resolve({ sheet: options?.sheet ?? 0, viewport, version: 'v2', epoch: 2, sequence: 1,
      displayList: { width: 800, height: 600, commands: [{ op: 'text', text: 'worker:typed before hydration', x: 8, y: 18, fontSize: 11, color: '#000000' }] } }));
    await advance();
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    expect(host.peerMethods.displayList).not.toHaveBeenCalled();
    expect(painted.every((list) => list.commands.every((command) =>
      command.op !== 'text' || !command.text.startsWith('peer:')))).toBe(true);
    expect(painted[painted.length - 1].commands.some((command) =>
      command.op === 'text' && command.text === 'worker:typed before hydration')).toBe(true);
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

  it('preserves an untouched cell when Enter precedes hydration and prefill', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    const prefill = deferred<xlsx.WorkbookCellInputs>();
    host.hydrate.mockReturnValue(hydration.promise);
    host.sessionMethods.cellInputs.mockReturnValue(prefill.promise);
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('');
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(host.cells.get('0:0:0')).toBe('initial');
    expect(host.editMethods.editCell).not.toHaveBeenCalled();
    await act(async () => prefill.resolve({ sheet: 0, version: 'v1',
      cells: [[{ a1: 'A1', input: 'initial', isFormula: false }]] }));
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    expect(host.cells.get('0:0:0')).toBe('initial');

    view.unmount();
    for (const restore of restorers.splice(-3).reverse()) restore();
    const reopenedHost = harness();
    const nextHydration = deferred<WorkbookHandle>();
    const nextPrefill = deferred<xlsx.WorkbookCellInputs>();
    reopenedHost.hydrate.mockReturnValue(nextHydration.promise);
    reopenedHost.sessionMethods.cellInputs.mockReturnValue(nextPrefill.promise);
    const reopened = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.keyDown(reopened.getByTestId('xlsx-scroll'), { key: 'n' });
    fireEvent.change(reopened.getByTestId('xlsx-cell-editor'), { target: { value: 'new' } });
    fireEvent.keyDown(reopened.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    fireEvent.keyDown(reopened.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
    fireEvent.keyDown(reopened.getByTestId('xlsx-scroll'), { key: 'F2' });
    await act(async () => nextPrefill.resolve({ sheet: 0, version: 'v1',
      cells: [[{ a1: 'A1', input: 'initial', isFormula: false }]] }));
    expect((reopened.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('initial');
    fireEvent.keyDown(reopened.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    expect(reopenedHost.cells.get('0:0:0')).toBe('initial');
    expect(reopenedHost.editMethods.editCell).not.toHaveBeenCalled();
    await act(async () => nextHydration.resolve(reopenedHost.peer));
    await advance();
    expect(reopenedHost.cells.get('0:0:0')).toBe('new');
    expect(reopenedHost.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'new']]);
    expect(reopened.queryByTestId('xlsx-cell-editor')).toBeNull();
  });

  it('retains later typing across queued host navigation and saves its original target', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
      onReady={(value) => { api = value; }} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'a' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Tab' });
    let navigation!: Promise<boolean>;
    act(() => { navigation = api.selectCellsAsync(1, xlsx.selectionAt({ row: 0, col: 0 })); });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'b' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'later typing' } });
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(await navigation).toBe(true);
    expect(host.cells.get('0:0:0')).toBe('a');
    expect(host.cells.has('0:0:1')).toBe(false);
    let saving!: Promise<Uint8Array | null>;
    act(() => { saving = api.save(); });
    await advance();
    expect(await saving).toEqual(new Uint8Array([8, 9]));
    expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'a'], [0, 0, 1, 'later typing']]);
    expect(host.log.indexOf('sheet:1')).toBeLessThan(host.log.indexOf('edit:0:0:1:later typing'));
  });

  it('continues cell input at its original target after queued navigation paints another sheet', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
      onReady={(value) => { api = value; }} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'a' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Tab' });
    let navigation!: Promise<boolean>;
    act(() => { navigation = api.selectCellsAsync(1, xlsx.selectionAt({ row: 2, col: 2 })); });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'b' });
    const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    fireEvent.change(input, { target: { value: 'before navigation' } });
    input.setSelectionRange(3, 3);
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(await navigation).toBe(true);
    expect(view.getByRole('tab', { name: 'Second' }).getAttribute('aria-selected')).toBe('true');
    expect(view.getByTestId('xlsx-cell-editor')).toBe(input);
    expect(input.value).toBe('before navigation');
    expect(input.selectionStart).toBe(3);
    expect(document.activeElement).toBe(input);
    fireEvent.change(input, { target: { value: 'continued cell input' } });
    fireEvent.keyDown(input, { key: 'Enter' });
    await advance();
    expect(host.cells.get('0:0:1')).toBe('continued cell input');
    expect(host.cells.has('1:2:2')).toBe(false);
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'g' });
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('g');
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await advance();
    expect(host.editMethods.editCell.mock.calls).toEqual([
      [0, 0, 0, 'a'], [0, 0, 1, 'continued cell input'], [1, 1, 1, 'g'],
    ]);
  });

  it('continues formula input at its original target after queued navigation executes', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen
      toolbar={<EditorToolbar mode="commands"><EditorToolbar.FormulaBar /></EditorToolbar>}
      onReady={(value) => { api = value; }} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'a' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Tab' });
    let navigation!: Promise<boolean>;
    act(() => { navigation = api.selectCellsAsync(1, xlsx.selectionAt({ row: 2, col: 2 })); });
    const input = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
    fireEvent.change(input, { target: { value: '=1' } });
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(await navigation).toBe(true);
    expect(input.value).toBe('=1');
    fireEvent.change(input, { target: { value: '=12' } });
    fireEvent.change(input, { target: { value: '=123' } });
    fireEvent.keyDown(input, { key: 'Enter' });
    await advance();
    expect(host.cells.get('0:0:1')).toBe('=123');
    expect(host.cells.has('1:2:2')).toBe(false);
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'g' });
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('g');
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await advance();
    expect(host.editMethods.editCell.mock.calls).toEqual([
      [0, 0, 0, 'a'], [0, 0, 1, '=123'], [1, 1, 1, 'g'],
    ]);
  });

  it('enqueues a continued cell draft at its original target before replacing it with formula input', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen
      toolbar={<EditorToolbar mode="commands"><EditorToolbar.FormulaBar /></EditorToolbar>}
      onReady={(value) => { api = value; }} />);
    await opened();
    let navigation!: Promise<boolean>;
    act(() => { navigation = api.selectCellsAsync(1, xlsx.selectionAt({ row: 2, col: 2 })); });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'c' });
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(await navigation).toBe(true);
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'continued' } });
    fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value: '=4' } });
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    fireEvent.keyDown(view.getByTestId('xlsx-formula-input'), { key: 'Enter' });
    await advance();
    expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'continued'], [1, 2, 2, '=4']]);
  });

  it('keeps an untouched stale prefill open while its preceding accepted commit drains', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    const prefill = deferred<xlsx.WorkbookCellInputs>();
    host.hydrate.mockReturnValue(hydration.promise);
    host.sessionMethods.cellInputs.mockReturnValue(prefill.promise);
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'new' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    await act(async () => prefill.resolve({ sheet: 0, version: 'v1',
      cells: [[{ a1: 'A1', input: 'initial', isFormula: false }]] }));
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('initial');
    expect(host.cells.get('0:0:0')).toBe('initial');
    expect(host.editMethods.editCell).not.toHaveBeenCalled();
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('initial');
    expect(host.cells.get('0:0:0')).toBe('new');
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await advance();
    expect(host.cells.get('0:0:0')).toBe('new');
    expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'new']]);
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
  });

  it('preserves an accepted commit when untouched hydrated prefill precedes its preview', async () => {
    const host = harness();
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'new' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await act(async () => {});
    await tick();
    expect(host.editMethods.editCell).not.toHaveBeenCalled();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('initial');
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await advance();
    expect(host.cells.get('0:0:0')).toBe('new');
    expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'new']]);
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
  });

  it('admits cut in the gesture and copies the preceding accepted value before clearing it', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    const originalItem = Object.getOwnPropertyDescriptor(globalThis, 'ClipboardItem');
    class Item {
      constructor(readonly data: Record<string, Promise<Blob>>) {}
    }
    Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: Item });
    const originalWrite = Object.getOwnPropertyDescriptor(navigator.clipboard, 'write');
    let copied: string | null = null;
    const write = mock(async (items: Item[]) => {
      copied = await (await items[0].data['text/plain']).text();
      host.log.push(`clipboard:${copied}`);
    });
    Object.defineProperty(navigator.clipboard, 'write', { configurable: true, value: write });
    restorers.push(() => {
      if (originalItem) Object.defineProperty(globalThis, 'ClipboardItem', originalItem);
      else Reflect.deleteProperty(globalThis, 'ClipboardItem');
      if (originalWrite) Object.defineProperty(navigator.clipboard, 'write', originalWrite);
      else Reflect.deleteProperty(navigator.clipboard, 'write');
    });
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'accepted' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true });
    expect(write).toHaveBeenCalledTimes(1);
    expect(copied).toBeNull();
    expect(host.cells.get('0:0:0')).toBe('initial');
    await act(async () => hydration.resolve(host.peer));
    await advance(12);
    expect<unknown>(copied).toBe('accepted');
    expect(host.cells.get('0:0:0')).toBe('');
    expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'accepted']]);
    expect(host.editMethods.editCells.mock.calls).toEqual([[0, [{ row: 0, col: 0, input: '' }]]]);
    expect(host.log.indexOf('edit:0:0:0:accepted')).toBeLessThan(host.log.indexOf('clipboard:accepted'));
    expect(host.log.indexOf('clipboard:accepted')).toBeLessThan(host.log.indexOf('edit:0:0:0:'));
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

  it('keeps a newer entry open when an earlier commit writes the same cell and value', async () => {
    const host = harness();
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'a' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'a' });
    const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    await advance();
    expect(view.getByTestId('xlsx-cell-editor')).toBe(input);
    expect(input.value).toBe('a');
    fireEvent.change(input, { target: { value: 'later same cell' } });
    fireEvent.keyDown(input, { key: 'Enter' });
    await advance();
    expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'a'], [0, 0, 0, 'later same cell']]);
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
  });

  it('drains an accepted cut payload and mutation before the replaced document is disposed', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    const clipboard = promisedClipboard();
    const errors = mock((_error: Error) => {});
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} onError={errors} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true });
    expect(clipboard.write).toHaveBeenCalledTimes(1);
    expect(clipboard.copied).toEqual([]);
    host.open.mockReturnValueOnce(new Promise<WorkbookSession>(() => {}));
    view.rerender(<XlsxEditor file={new Uint8Array([4])} experimentalWorkerOpen showToolbar={false} onError={errors} />);
    expect(host.session.dispose).not.toHaveBeenCalled();
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(await (await clipboard.payloads[0]).text()).toBe('initial');
    await clipboard.writes[0];
    expect(clipboard.copied).toEqual(['initial']);
    expect(host.editMethods.editCells.mock.calls).toEqual([[0, [{ row: 0, col: 0, input: '' }]]]);
    await waitFor(() => expect(host.session.dispose).toHaveBeenCalledTimes(1));
    expect(host.log.indexOf('batch')).toBeLessThan(host.log.indexOf('dispose:facade'));
    expect(host.editMethods.editCells).toHaveBeenCalledTimes(1);
    expect(host.cells.get('0:0:0')).toBe('');
    expect(errors).not.toHaveBeenCalled();
  });

  it('rejects the pending cut payload and clipboard write when admission is refused', async () => {
    const host = harness();
    host.charts.push({ id: 'chart:1', label: 'Chart', movable: true,
      rect: { x: 100, y: 100, w: 200, h: 100 }, clip: { x: 100, y: 100, w: 200, h: 100 } });
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    const clipboard = promisedClipboard();
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
      onReady={(value) => { api = value; }} />);
    await opened();
    let navigation!: Promise<boolean>;
    act(() => { navigation = api.selectCellsAsync(1, xlsx.selectionAt({ row: 0, col: 0 })); });
    fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { clientX: 120, clientY: 120, button: 0 });
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(await navigation).toBe(true);
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true });
    expect(clipboard.write).toHaveBeenCalledTimes(1);
    await expect(clipboard.payloads[0]).rejects.toMatchObject({ code: 'gesture-active' });
    await expect(clipboard.writes[0]).rejects.toMatchObject({ code: 'gesture-active' });
    expect(host.editMethods.editCells).not.toHaveBeenCalled();
    expect(api.failure).toBeNull();
    fireEvent.mouseUp(window, { clientX: 120, clientY: 120, button: 0 });
  });

  it('refuses cut before admission without promised clipboard writes and keeps editing available', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    const originalItem = Object.getOwnPropertyDescriptor(globalThis, 'ClipboardItem');
    Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: undefined });
    const write = spyOn(navigator.clipboard, 'writeText').mockRejectedValue(new Error('Gesture expired'));
    restorers.push(() => {
      if (originalItem) Object.defineProperty(globalThis, 'ClipboardItem', originalItem);
      else Reflect.deleteProperty(globalThis, 'ClipboardItem');
      write.mockRestore();
    });
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
      onReady={(value) => { api = value; }} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'new' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true });
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(write).not.toHaveBeenCalled();
    expect(host.editMethods.editCells).not.toHaveBeenCalled();
    expect(host.cells.get('0:0:0')).toBe('new');
    expect(api.failure).toBeNull();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 's' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'still editable' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await advance();
    expect(host.cells.get('0:0:0')).toBe('still editable');
  });

  it('refuses cut when ClipboardItem rejects promised payloads without failing the editor', async () => {
    const host = harness();
    const originalItem = Object.getOwnPropertyDescriptor(globalThis, 'ClipboardItem');
    class Item {
      constructor() { throw new TypeError('Promised clipboard data is unavailable'); }
    }
    Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: Item });
    restorers.push(() => {
      if (originalItem) Object.defineProperty(globalThis, 'ClipboardItem', originalItem);
      else Reflect.deleteProperty(globalThis, 'ClipboardItem');
    });
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
      onReady={(value) => { api = value; }} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true });
    await advance();
    expect(host.editMethods.editCells).not.toHaveBeenCalled();
    expect(host.cells.get('0:0:0')).toBe('initial');
    expect(api.failure).toBeNull();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 's' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await advance();
    expect(host.cells.get('0:0:0')).toBe('s');
  });

  for (const held of [true, false]) {
    it(`copies the preceding accepted value while ${held ? 'hydration' : 'its preview'} waits`, async () => {
      const host = harness();
      const hydration = deferred<WorkbookHandle>();
      if (held) host.hydrate.mockReturnValue(hydration.promise);
      const clipboard = promisedClipboard();
      const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
      await opened();
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
      fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'new' } });
      fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'c', ctrlKey: true });
      expect(clipboard.write).toHaveBeenCalledTimes(1);
      expect(clipboard.copied).toEqual([]);
      expect(host.cells.get('0:0:0')).toBe('initial');
      if (held) await act(async () => hydration.resolve(host.peer));
      await advance();
      expect(clipboard.copied).toEqual(['new']);
      expect(host.cells.get('0:0:0')).toBe('new');
      expect(host.editMethods.editCells).not.toHaveBeenCalled();
    });
  }

  it('copies through the FIFO without promised clipboard writes', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    const originalItem = Object.getOwnPropertyDescriptor(globalThis, 'ClipboardItem');
    Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: undefined });
    const write = spyOn(navigator.clipboard, 'writeText').mockResolvedValue(undefined);
    restorers.push(() => {
      if (originalItem) Object.defineProperty(globalThis, 'ClipboardItem', originalItem);
      else Reflect.deleteProperty(globalThis, 'ClipboardItem');
      write.mockRestore();
    });
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'new' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'c', ctrlKey: true });
    expect(write).not.toHaveBeenCalled();
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(write).toHaveBeenCalledTimes(1);
    expect(write).toHaveBeenCalledWith('new');
    expect(host.cells.get('0:0:0')).toBe('new');
    expect(host.editMethods.editCells).not.toHaveBeenCalled();
  });

  it('rejects stale worker frames without erasing the preview before matching worker adoption', async () => {
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
    expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'new']]);
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('new');
    expect(painted[painted.length - 1].commands.some((command) => command.op === 'text' && command.text === 'preview:new')).toBe(true);
    const count = painted.length;
    await act(async () => pending.resolve(stale));
    expect(painted).toHaveLength(count);
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('new');
    await advance();
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    expect(host.peerMethods.displayList).not.toHaveBeenCalled();
    expect(painted.every((list) => list.commands.every((command) =>
      command.op !== 'text' || !command.text.startsWith('peer:')))).toBe(true);
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

  it('keeps the origin sheet for a chart drag released after queued navigation', async () => {
    const host = harness();
    host.charts.push({ id: 'chart:1', label: 'Chart', movable: true,
      rect: { x: 100, y: 100, w: 200, h: 100 }, clip: { x: 100, y: 100, w: 200, h: 100 } });
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
      onReady={(value) => { api = value; }} />);
    await opened();
    let navigation!: Promise<boolean>;
    act(() => { navigation = api.selectCellsAsync(1, xlsx.selectionAt({ row: 0, col: 0 })); });
    fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { clientX: 120, clientY: 120, button: 0 });
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(await navigation).toBe(true);
    expect(view.getByRole('tab', { name: 'Second' }).getAttribute('aria-selected')).toBe('true');
    await act(async () => { expect(await api.commands.execute('zoom', { scale: 2 })).toMatchObject({ ok: true }); });
    await advance();
    fireEvent.mouseUp(window, { clientX: 150, clientY: 140, button: 0 });
    await advance();
    expect(host.editMethods.moveChart.mock.calls).toEqual([[0, 'chart:1', 30, 20]]);
    expect(host.log.indexOf('sheet:1')).toBeLessThan(host.log.indexOf('chart:chart:1:30:20'));
    expect(api.failure).toBeNull();
  });

  it('recovers an accepted chart nudge when the worker fails before its timer fires', async () => {
    const host = harness(true);
    host.charts.push({ id: 'chart:1', label: 'Chart', movable: true,
      rect: { x: 100, y: 100, w: 200, h: 100 }, clip: { x: 100, y: 100, w: 200, h: 100 } });
    host.peerMethods.save.mockImplementation(() => new TextEncoder().encode(JSON.stringify(host.charts[0].rect)));
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
      onReady={(value) => { api = value; }} />);
    await opened();
    fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { clientX: 120, clientY: 120, button: 0 });
    fireEvent.mouseUp(window, { clientX: 120, clientY: 120, button: 0 });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowDown', shiftKey: true });
    expect(view.getByTestId('xlsx-chart-selection').style.transform).toBe('translate(1px, 10px)');
    expect(host.peerMethods.moveChart).not.toHaveBeenCalled();
    act(() => host.fail());
    let saved!: { bytes: Uint8Array; recovery: true };
    await act(async () => { saved = await api.recoverySave(); });
    expect(saved.recovery).toBe(true);
    expect(JSON.parse(new TextDecoder().decode(saved.bytes))).toEqual({ x: 101, y: 110, w: 200, h: 100 });
    expect(host.peerMethods.moveChart.mock.calls).toEqual([[0, 'chart:1', 1, 10]]);
    expect(host.replay).not.toHaveBeenCalled();
    expect(host.session.save).not.toHaveBeenCalled();
    expect(host.peerMethods.save).toHaveBeenCalledTimes(1);
    await advance();
    expect(host.peerMethods.moveChart).toHaveBeenCalledTimes(1);
  });

  it('saves pending formula input through the facade and keeps PNG unavailable', async () => {
    const host = harness();
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    const worker = deferred<xlsx.WorkbookFrame>();
    host.sessionMethods.frame.mockReturnValue(worker.promise);
    fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value: '=24' } });
    host.editMethods.save.mockImplementation(async () => {
      expect(host.cells.get('0:0:0')).toBe('=24');
      expect(host.peerMethods.displayList).not.toHaveBeenCalled();
      expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('=24');
      return new Uint8Array([8, 9]).buffer;
    });
    let saving!: Promise<Uint8Array | null>;
    act(() => { saving = api.save(); });
    await advance();
    expect(await saving).toEqual(new Uint8Array([8, 9]));
    expect(host.editMethods.save).toHaveBeenCalledTimes(1);
    expect(host.peerMethods.save).not.toHaveBeenCalled();
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('=24');
    const [viewport, options] = host.sessionMethods.frame.mock.calls.at(-1)!;
    await act(async () => worker.resolve({ sheet: options?.sheet ?? 0, viewport, version: 'v2', epoch: 2, sequence: 1,
      displayList: { width: 800, height: 600, commands: [{ op: 'text', text: 'worker:=24', x: 8, y: 18, fontSize: 11, color: '#000000' }] } }));
    await advance();
    expect(host.peerMethods.displayList).not.toHaveBeenCalled();
    expect(painted[painted.length - 1].commands.some((command) => command.op === 'text' && command.text === 'worker:=24')).toBe(true);
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    expect(api.commands.getState('exportPng')).toMatchObject({ enabled: false, disabledReason: { code: 'png-unavailable' } });
  });

  it('prints after the accepted edited viewport has been painted', async () => {
    const host = harness();
    const original = Object.getOwnPropertyDescriptor(window, 'print');
    const print = mock(() => {
      expect(painted[painted.length - 1].commands.some((command) => command.op === 'text' && command.text === 'worker:print')).toBe(true);
      expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
      expect(host.peerMethods.displayList).not.toHaveBeenCalled();
    });
    Object.defineProperty(window, 'print', { configurable: true, value: print });
    restorers.push(() => {
      if (original) Object.defineProperty(window, 'print', original);
      else Reflect.deleteProperty(window, 'print');
    });
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
    expect(host.editMethods.editCell).toHaveBeenCalledWith(0, 0, 0, 'print');
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
    expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'q']]);
    expect(host.log.indexOf('edit:0:0:0:q')).toBeLessThan(host.log.indexOf('dispose:facade'));
    expect(host.editMethods.flush).toHaveBeenCalledTimes(1);
  });

  it('reports throwing readiness cleanup while replacing the document and disposing its resources', async () => {
    const host = harness();
    const error = new Error('Cleanup failed');
    const errors = mock((_error: Error) => {});
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
      onError={errors} onReady={() => () => { throw error; }} />);
    await opened();
    view.rerender(<XlsxEditor file={new Uint8Array([4])} experimentalWorkerOpen showToolbar={false}
      onError={errors} />);
    await opened();
    expect(errors).toHaveBeenCalledTimes(1);
    expect(errors).toHaveBeenCalledWith(error);
    expect(host.editMethods.dispose).toHaveBeenCalledTimes(1);
    expect(host.peerMethods.dispose).toHaveBeenCalledTimes(1);
    expect(host.session.dispose).toHaveBeenCalledTimes(1);
    expect(view.queryByRole('alert')).toBeNull();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'r' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await advance();
    expect(host.cells.get('0:0:0')).toBe('r');
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

  it('recovers applied and retained edits through the real facade after local paint failure', async () => {
    const host = harness(true);
    const failure = new Error('Peer publication failed');
    const errors = mock((_error: Error) => {});
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
      onError={errors} onReady={(value) => { api = value; }} />);
    await opened();
    host.sessionMethods.frame.mockRejectedValueOnce(failure);
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'a' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    let retained!: Promise<EditResult | null>;
    act(() => { retained = api.editCellAsync(0, 0, 1, 'retained'); void retained.catch(() => {}); });
    await advance();
    await expect(retained).rejects.toBe(failure);
    expect(api.failure).toBe(failure);
    expect(host.session.failure).toBeUndefined();
    expect(host.attached!.state).toBe('failed');
    expect(errors).toHaveBeenCalledTimes(1);
    await expect(api.save()).rejects.toBe(failure);
    const replays = host.replay.mock.calls.length;
    let saved!: { bytes: Uint8Array; recovery: true };
    await act(async () => { saved = await api.recoverySave(); });
    expect(saved.recovery).toBe(true);
    expect(JSON.parse(new TextDecoder().decode(saved.bytes))).toEqual([
      ['0:0:0', 'a'], ['0:0:1', 'retained'],
    ]);
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'a'], [0, 0, 1, 'retained']]);
    expect(host.replay).toHaveBeenCalledTimes(replays);
    expect(host.session.save).not.toHaveBeenCalled();
    expect(host.peerMethods.save).toHaveBeenCalledTimes(1);
    await expect(api.editCellAsync(0, 0, 0, 'later')).rejects.toBe(failure);
    await expect(api.recoverySave()).rejects.toThrow('already saved');
  });

  it('recovers accepted host edits when local readiness fails before facade attachment', async () => {
    const host = harness(true);
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    const failure = new Error('Readiness failed');
    let api!: XlsxWorkerEditorApi;
    let editing!: Promise<EditResult | null>;
    render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => {
      api = value;
      editing = api.editCellAsync(0, 0, 1, 'accepted');
      void editing.catch(() => {});
      throw failure;
    }} />);
    await opened();
    await expect(editing).rejects.toBe(failure);
    expect(host.attached).toBeNull();
    await act(async () => hydration.resolve(host.peer));
    expect(host.session.failure).toBeUndefined();
    expect(host.attached!.state).toBe('failed');
    let saved!: { bytes: Uint8Array; recovery: true };
    await act(async () => { saved = await api.recoverySave(); });
    expect(JSON.parse(new TextDecoder().decode(saved.bytes))).toEqual([
      ['0:0:0', 'initial'], ['0:0:1', 'accepted'],
    ]);
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 1, 'accepted']]);
    expect(host.replay).not.toHaveBeenCalled();
    expect(host.session.save).not.toHaveBeenCalled();
    await expect(api.save()).rejects.toBe(failure);
  });

  it('discards refused reads while recovering retained host edits through the real facade', async () => {
    const host = harness(true);
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    host.peerMethods.readCells.mockReturnValue({
      ok: false, version: 'v1', failure: { code: 'missing-target', message: 'Sheet is missing' },
    });
    let api!: XlsxWorkerEditorApi;
    render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
      onReady={(value) => { api = value; }} />);
    await opened();
    let reading!: ReturnType<XlsxWorkerEditorApi['readCells']>;
    let editing!: Promise<EditResult | null>;
    act(() => {
      reading = api.readCells({ ranges: [{ sheetId: 'missing', range: { kind: 'a1', a1: 'A1' } }] });
      editing = api.editCellAsync(0, 0, 1, 'retained');
      void reading.catch(() => {});
      void editing.catch(() => {});
    });
    let failure!: SessionFailure;
    act(() => { failure = host.fail(); });
    await expect(reading).rejects.toBe(failure);
    await expect(editing).rejects.toBe(failure);
    await act(async () => hydration.resolve(host.peer));
    let saved!: { bytes: Uint8Array; recovery: true };
    await act(async () => { saved = await api.recoverySave(); });
    expect(JSON.parse(new TextDecoder().decode(saved.bytes))).toEqual([
      ['0:0:0', 'initial'], ['0:0:1', 'retained'],
    ]);
    expect(host.peerMethods.readCells).not.toHaveBeenCalled();
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 1, 'retained']]);
    expect(host.replay).not.toHaveBeenCalled();
    expect(host.session.save).not.toHaveBeenCalled();
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


function reviewEdit(view: ReturnType<typeof render>, value: string) {
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
  fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value } });
  fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
}

function reviewBatch(value: string, version = 'v1'): XlsxEditRequest {
  return { expectVersion: version, steps: [{ op: 'setCellInputs',
    target: { sheetId: 'sheet:0', range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 0, col: 0 } } },
    inputs: [[value]], expect: { cells: [[{ displayText: 'initial' }]] } }] };
}

function replaceClipboard(name: 'write' | 'writeText' | 'readText', value: unknown) {
  const original = Object.getOwnPropertyDescriptor(navigator.clipboard, name);
  Object.defineProperty(navigator.clipboard, name, { configurable: true, value });
  restorers.push(() => {
    if (original) Object.defineProperty(navigator.clipboard, name, original);
    else Reflect.deleteProperty(navigator.clipboard, name);
  });
}

for (const wait of ['hydration', 'preview'] as const) {
  for (const transition of ['replacement', 'unmount', 'readOnly'] as const) {
    it(`drains accepted edits before replacing or disposing their document (${wait}, ${transition})`, async () => {
      const host = harness(true);
      const hydration = deferred<WorkbookHandle>();
      if (wait === 'hydration') host.hydrate.mockReturnValue(hydration.promise);
      let api!: XlsxWorkerEditorApi;
      const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
        onReady={(value) => { api = value; }} />);
      await opened();
      reviewEdit(view, 'retiring edit');
      await act(async () => {});
      const acknowledged = deferred<WorkbookReplayReply>();
      host.replay.mockReturnValueOnce(acknowledged.promise);
      if (transition === 'unmount') view.unmount();
      else if (transition === 'readOnly') {
        const viewerOpen = spyOn(workbookSessionOpener, 'open').mockReturnValue(new Promise<WorkbookSession>(() => {}));
        restorers.push(() => viewerOpen.mockRestore());
        view.rerender(<XlsxEditor file={file} experimentalWorkerOpen readOnly showToolbar={false} />);
      }
      else {
        host.open.mockReturnValueOnce(new Promise<WorkbookSession>(() => {}));
        view.rerender(<XlsxEditor file={new Uint8Array([4])} experimentalWorkerOpen showToolbar={false} />);
      }
      if (wait === 'hydration') await act(async () => hydration.resolve(host.peer));
      await advance();
      expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'retiring edit']]);
      expect(host.session.dispose).not.toHaveBeenCalled();
      expect(host.peerMethods.dispose).not.toHaveBeenCalled();
      expect(host.replay).toHaveBeenCalledTimes(1);
      await act(async () => acknowledged.resolve({ sequence: 1, revision: 1, version: 1, result: undefined }));
      await waitFor(() => expect(host.session.dispose).toHaveBeenCalledTimes(1));
      expect(host.attached!.acknowledgedSequence).toBe(1);
      expect(host.peerMethods.editCell).toHaveBeenCalledTimes(1);
      expect(api.hydrated).toBe(false);
    });
  }
}

it('drains a buffered chart nudge before its timer expires on disposal', async () => {
  const host = harness();
  host.charts.push({ id: 'chart', rect: { x: 192, y: 96, w: 192, h: 96 }, movable: true } as xlsx.ChartRegion);
  const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
  await opened();
  fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { clientX: 210, clientY: 110, button: 0 });
  fireEvent.mouseUp(window, { clientX: 210, clientY: 110, button: 0 });
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
  view.unmount();
  await advance();
  expect(host.editMethods.moveChart).toHaveBeenCalledWith(0, 'chart', 1, 0);
  expect(host.log.indexOf('chart:chart:1:0')).toBeLessThan(host.log.indexOf('dispose:facade'));
});

it('cancels the original pending cut before replaying it during recovery', async () => {
  const host = harness(true);
  const clipboard = promisedClipboard();
  const completion = deferred<void>();
  clipboard.write.mockImplementation((items) => {
    clipboard.payloads.push(items[0].data['text/plain']);
    return completion.promise;
  });
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
  await opened();
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true });
  await advance();
  await clipboard.payloads[0];
  let newer!: Promise<EditResult | null>;
  act(() => { newer = api.editCellAsync(0, 0, 1, 'newer'); void newer.catch(() => {}); host.fail(); });
  await expect(newer).rejects.toBeInstanceOf(SessionFailure);
  let recovered!: Promise<{ bytes: Uint8Array; recovery: true }>;
  act(() => { recovered = api.recoverySave(); });
  await act(async () => completion.resolve());
  const saved = await recovered;
  expect(host.peerMethods.editCells).toHaveBeenCalledTimes(1);
  expect(host.peerMethods.setCalculationContext).toHaveBeenCalledTimes(2);
  expect(JSON.parse(new TextDecoder().decode(saved.bytes))).toEqual([['0:0:0', ''], ['0:0:1', 'newer']]);
});

for (const mode of ['unsupported-cut', 'construct-cut', 'throw-cut', 'deny-cut', 'deny-copy'] as const) {
  it(`shows clipboard refusals and does not mark denied cuts as applied (${mode})`, async () => {
    const host = harness();
    const clipboard = promisedClipboard();
    const denied = new DOMException('Clipboard permission denied', 'NotAllowedError');
    if (mode === 'unsupported-cut' || mode === 'deny-copy') replaceClipboard('write', undefined);
    if (mode === 'construct-cut') {
      const original = Object.getOwnPropertyDescriptor(globalThis, 'ClipboardItem')!;
      Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: class { constructor() { throw denied; } } });
      restorers.push(() => Object.defineProperty(globalThis, 'ClipboardItem', original));
    }
    if (mode === 'throw-cut') clipboard.write.mockImplementation(() => { throw denied; });
    if (mode === 'deny-cut') clipboard.write.mockRejectedValue(denied);
    if (mode === 'deny-copy') replaceClipboard('writeText', mock(async () => { throw denied; }));
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: mode === 'deny-copy' ? 'c' : 'x', ctrlKey: true });
    await advance();
    expect(view.getByTestId('xlsx-input-refusal').textContent).toMatch(/clipboard|permission/i);
    expect(view.queryByTestId('xlsx-error')).toBeNull();
    expect(host.cells.get('0:0:0')).toBe('initial');
    expect(host.editMethods.editCells).not.toHaveBeenCalled();
    expect(host.replay).not.toHaveBeenCalled();
    expect(api.failure).toBeNull();
    reviewEdit(view, 'still editable');
    await advance();
    expect(host.cells.get('0:0:0')).toBe('still editable');
  });
}

it('renders target-changed and gesture-active command refusals', async () => {
  const host = harness();
  host.charts.push({ id: 'chart', rect: { x: 192, y: 96, w: 192, h: 96 }, movable: true } as xlsx.ChartRegion);
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor file={file} experimentalWorkerOpen onReady={(value) => { api = value; }} />);
  await opened();
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'b', ctrlKey: true });
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
  await advance();
  expect(view.getByTestId('xlsx-input-refusal').textContent).toMatch(/selection changed/i);
  expect(host.peerMethods.patchRangeStyle).not.toHaveBeenCalled();
  fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { clientX: 210, clientY: 110, button: 0 });
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 's', ctrlKey: true });
  await advance();
  expect(view.getByTestId('xlsx-input-refusal').textContent).toMatch(/chart drag/i);
  expect(host.editMethods.save).not.toHaveBeenCalled();
  expect(api.failure).toBeNull();
  fireEvent.mouseUp(window, { clientX: 210, clientY: 110, button: 0 });
  fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { clientX: 8, clientY: 8, button: 0 });
  reviewEdit(view, 'after refusal');
  await advance();
  expect(host.cells.get('0:0:0')).toBe('after refusal');
});

for (const route of ['cut', 'bold'] as const) {
  for (const barrier of ['flush', 'save'] as const) {
    it(`rejects covering flush and save when an admitted mutation was refused (${route}, ${barrier})`, async () => {
      const host = harness();
      const clipboard = promisedClipboard();
      clipboard.write.mockRejectedValue(new DOMException('Clipboard denied', 'NotAllowedError'));
      let api!: XlsxWorkerEditorApi;
      const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
      await opened();
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: route === 'cut' ? 'x' : 'b', ctrlKey: true });
      if (route === 'bold') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
      let covering!: Promise<unknown>;
      act(() => { covering = api[barrier](); void covering.catch(() => {}); });
      await advance();
      await expect(covering).rejects.toBeInstanceOf(Error);
      expect(view.getByTestId('xlsx-input-refusal')).toBeTruthy();
      expect(host.editMethods.editCells).not.toHaveBeenCalled();
      expect(host.peerMethods.patchRangeStyle).not.toHaveBeenCalled();
      expect(host.editMethods.save).not.toHaveBeenCalled();
      expect(api.failure).toBeNull();
    });
  }
}

it('recovers an applied plugin batch after post-commit publication failure', async () => {
  const host = harness(true);
  let context!: XlsxPluginContext<null>;
  let api!: XlsxWorkerEditorApi;
  const plugin = defineXlsxPlugin({ id: 'recover-batch', createState: () => null, initialize(value) { context = value; } });
  render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} plugins={[plugin]}
    pluginGrants={{ 'recover-batch': { document: 'write', editBatches: true } }} onReady={(value) => { api = value; }} />);
  await opened();
  await waitFor(() => expect(context).toBeTruthy());
  const publication = spyOn(host.peerMethods, 'listProposals').mockImplementationOnce(() => { throw new Error('Publication failed'); });
  restorers.push(() => publication.mockRestore());
  let pending!: Promise<unknown>;
  act(() => {
    void context.run(async (current) => { await current.edits!.applyEdits(reviewBatch('plugin')); });
    pending = api.editCellAsync(0, 0, 1, 'later host');
    void pending.catch(() => {});
  });
  await advance();
  await expect(pending).rejects.toThrow('Publication failed');
  const saved = await api.recoverySave();
  expect(JSON.parse(new TextDecoder().decode(saved.bytes))).toEqual([['0:0:0', 'plugin'], ['0:0:1', 'later host']]);
  expect(host.peerMethods.applyEdits).toHaveBeenCalledTimes(1);
  expect(host.peerMethods.applyEdits.mock.calls[0][0].expectVersion).toBe('v1');
  expect(host.peerMethods.save).toHaveBeenCalledTimes(1);
});

for (const pending of ['read', 'write'] as const) {
  it(`recovers accepted edits with pending plugin operations after plugin shutdown (${pending})`, async () => {
    const host = harness(true);
    let context!: XlsxPluginContext<null>;
    let api!: XlsxWorkerEditorApi;
    let stopped = false;
    const plugin = defineXlsxPlugin({ id: 'pending-plugin', createState: () => null,
      initialize(value) { context = value; value.onCleanup(() => { stopped = true; }); } });
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} plugins={[plugin]}
      pluginGrants={{ 'pending-plugin': { document: 'write', editBatches: true } }} onReady={(value) => { api = value; }} />);
    await opened();
    await waitFor(() => expect(context).toBeTruthy());
    reviewEdit(view, 'accepted');
    act(() => { void context.run(async (current) => {
      if (pending === 'read') await current.read.readCells({ ranges: [] });
      else await current.edits!.applyEdits({ ...reviewBatch('plugin', 'v2'), steps: reviewBatch('plugin').steps.map(({ expect: _expect, ...step }) => step) });
    }); });
    act(() => host.fail());
    await waitFor(() => expect(stopped).toBe(true));
    const saved = await api.recoverySave();
    expect(JSON.parse(new TextDecoder().decode(saved.bytes))).toEqual([['0:0:0', pending === 'read' ? 'accepted' : 'plugin']]);
    expect(host.peerMethods.editCell).toHaveBeenCalledTimes(1);
    expect(host.peerMethods.applyEdits).toHaveBeenCalledTimes(pending === 'write' ? 1 : 0);
    expect(host.peerMethods.readCells).not.toHaveBeenCalled();
    expect(host.peerMethods.save).toHaveBeenCalledTimes(1);
  });
}

it('keeps refused nonmutating commands out of mutation recovery', async () => {
  const host = harness(true);
  const hydration = deferred<WorkbookHandle>();
  host.hydrate.mockReturnValue(hydration.promise);
  const print = spyOn(window, 'print').mockImplementation(() => {});
  restorers.push(() => print.mockRestore());
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
  await opened();
  let requests!: Promise<unknown>[];
  act(() => {
    requests = [api.commands.execute('print', null), api.commands.execute('exportPng', null), api.commands.execute('bold', null)];
    void api.editCellAsync(0, 0, 1, 'valid').catch(() => {});
    api.clearSelection();
    host.fail();
  });
  await Promise.all(requests);
  await act(async () => hydration.resolve(host.peer));
  const saved = await api.recoverySave();
  expect(JSON.parse(new TextDecoder().decode(saved.bytes))).toContainEqual(['0:0:1', 'valid']);
  expect(print).not.toHaveBeenCalled();
  expect(view.getByTestId('xlsx-input-refusal')).toBeTruthy();
  expect(host.peerMethods.save).toHaveBeenCalledTimes(1);
});

it('keeps the workbook editable and saveable after clipboard permission denial', async () => {
  const host = harness(true);
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
  await opened();
  reviewEdit(view, 'valid');
  await advance();
  replaceClipboard('readText', mock(async () => { throw new DOMException('Paste denied', 'NotAllowedError'); }));
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
  await advance();
  expect(view.getByTestId('xlsx-input-refusal').textContent).toContain('Paste denied');
  expect(view.queryByTestId('xlsx-error')).toBeNull();
  expect(api.failure).toBeNull();
  expect(host.cells.get('0:0:0')).toBe('valid');
  const save = api.save();
  await advance();
  expect(await save).toEqual(new Uint8Array([0]));
  expect(host.session.save).toHaveBeenCalledTimes(1);
  expect(host.peerMethods.save).not.toHaveBeenCalled();
});

for (const invalid of ['text', 'formula', 'sheet', 'cell'] as const) {
  it(`preserves valid edits after rejected cell input (${invalid})`, async () => {
    const host = harness(true);
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
    await opened();
    reviewEdit(view, 'valid');
    await advance();
    const rejected = api.editCellAsync(invalid === 'sheet' ? 99 : 0, invalid === 'cell' ? -1 : 0, 1,
      invalid === 'formula' ? `=${'1+'.repeat(17000)}1` : invalid === 'text' ? 'x'.repeat(32768) : 'bad target');
    void rejected.catch(() => {});
    await advance();
    await expect(rejected).rejects.toBeInstanceOf(Error);
    expect(view.queryByTestId('xlsx-error')).toBeNull();
    expect(view.getByTestId('xlsx-input-refusal')).toBeTruthy();
    expect(api.failure).toBeNull();
    expect(host.cells.get('0:0:0')).toBe('valid');
    expect(host.peerMethods.editCell).toHaveBeenCalledTimes(1);
    const corrected = api.editCellAsync(0, 0, 1, 'corrected');
    await advance();
    await corrected;
    act(() => host.fail());
    const saved = await api.recoverySave();
    expect(JSON.parse(new TextDecoder().decode(saved.bytes))).toEqual([['0:0:0', 'valid'], ['0:0:1', 'corrected']]);
  });
}

for (const source of ['cell', 'formula'] as const) {
  it(`recovers accepted IME text when worker failure ends composition (${source})`, async () => {
    const host = harness(true);
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    if (source === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    const input = source === 'cell' ? view.getByTestId('xlsx-cell-editor') : view.getByTestId('xlsx-formula-input');
    fireEvent.compositionStart(input);
    fireEvent.change(input, { target: { value: '日本語' } });
    let dependent!: Promise<EditResult | null>;
    act(() => { dependent = api.editCellAsync(0, 0, 1, 'dependent'); void dependent.catch(() => {}); host.fail(); });
    await expect(dependent).rejects.toBeInstanceOf(SessionFailure);
    const saved = await api.recoverySave();
    expect(JSON.parse(new TextDecoder().decode(saved.bytes))).toEqual([['0:0:0', '日本語'], ['0:0:1', 'dependent']]);
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, '日本語'], [0, 0, 1, 'dependent']]);
  });
}

it('restores an accepted preview after a later host batch is refused', async () => {
  const host = harness(true);
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
    onReady={(value) => { api = value; }} />);
  await opened();
  const worker = deferred<xlsx.WorkbookFrame>();
  host.sessionMethods.frame.mockReturnValue(worker.promise);
  reviewEdit(view, 'accepted preview');
  await advance();
  const rejected = api.applyEdits(reviewBatch('refused preview', 'stale'));
  await advance();
  expect(await rejected).toMatchObject({ ok: false });
  expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('accepted preview');
  expect(host.cells.get('0:0:0')).toBe('accepted preview');
  expect(host.attached!.sentSequence).toBe(1);
  expect(api.failure).toBeNull();
  const [viewport] = host.sessionMethods.frame.mock.calls.at(-1)!;
  await act(async () => worker.resolve({ sheet: 0, viewport, version: 'v2', epoch: 1, sequence: 1,
    displayList: { width: 800, height: 600, commands: [{ op: 'text', text: 'worker:accepted preview',
      x: 8, y: 18, fontSize: 11, color: '#000000' }] } }));
  await advance();
  expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
});

it('recovers on the first attempt with a pending failed flush', async () => {
  const host = harness(true);
  const hydration = deferred<WorkbookHandle>();
  host.hydrate.mockReturnValue(hydration.promise);
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
  await opened();
  reviewEdit(view, 'accepted');
  const flush = api.flush();
  void flush.catch(() => {});
  act(() => host.fail());
  await expect(flush).rejects.toBeInstanceOf(SessionFailure);
  await act(async () => hydration.resolve(host.peer));
  const saved = await api.recoverySave();
  expect(JSON.parse(new TextDecoder().decode(saved.bytes))).toEqual([['0:0:0', 'accepted']]);
  expect(host.peerMethods.editCell).toHaveBeenCalledTimes(1);
  expect(host.peerMethods.save).toHaveBeenCalledTimes(1);
  expect(host.session.save).not.toHaveBeenCalled();
});

for (const route of ['paste', 'delete', 'host', 'bulk', 'formula-save', 'formula-navigation'] as const) {
  it(`paints every accepted pre-hydration cell edit before hydration resolves (${route})`, async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    replaceClipboard('readText', mock(async () => 'pasted'));
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    let accepted: Promise<unknown> | undefined;
    if (route === 'paste') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
    else if (route === 'delete') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'Delete' });
    else if (route === 'host') act(() => { accepted = api.editCellAsync(0, 0, 0, 'host'); });
    else if (route === 'bulk') act(() => { accepted = api.applyEdits(reviewBatch('plugin')); });
    else {
      fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value: '=24' } });
      act(() => { accepted = route === 'formula-save' ? api.save() : api.selectCellsAsync(1, xlsx.selectionAt({ row: 0, col: 0 })); });
    }
    await advance();
    const expected = route === 'delete' ? '' : route === 'paste' ? 'pasted' : route === 'host' ? 'host' : route === 'bulk' ? 'plugin' : '=24';
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe(expected);
    expect(host.preview).toHaveBeenCalled();
    expect(painted.some((list) => list.commands.some((cmd) => cmd.op === 'text' && cmd.text === `preview:${expected}`))).toBe(true);
    expect(api.hydrated).toBe(false);
    expect(host.editMethods.editCell).not.toHaveBeenCalled();
    expect(host.editMethods.editCells).not.toHaveBeenCalled();
    await act(async () => hydration.resolve(host.peer));
    await advance();
    await accepted;
    expect(host.cells.get('0:0:0')).toBe(expected);
  });
}

for (const preceding of [false, true]) {
  it(`previews the resolved untouched F2 value before hydration (${preceding ? 'accepted input' : 'snapshot'})`, async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    const prefill = deferred<xlsx.WorkbookCellInputs>();
    host.hydrate.mockReturnValue(hydration.promise);
    host.sessionMethods.cellInputs.mockReturnValue(prefill.promise);
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    if (preceding) {
      reviewEdit(view, 'new');
      await advance();
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
    }
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await advance();
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe(preceding ? 'new' : 'initial');
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(host.cells.get('0:0:0')).toBe(preceding ? 'new' : 'initial');
    await act(async () => prefill.resolve({ sheet: 0, version: 'v1', cells: [[{ a1: 'A1', input: 'initial', isFormula: false }]] }));
    expect(host.cells.get('0:0:0')).toBe(preceding ? 'new' : 'initial');
  });
}

it('resolves an untouched F2 preview before preparing a later host edit', async () => {
  const host = harness();
  const hydration = deferred<WorkbookHandle>();
  const snapshot = deferred<string>();
  host.hydrate.mockReturnValue(hydration.promise);
  workbookSessionInternals.get(host.session)!.cellInput = () => snapshot.promise;
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false}
    onReady={(value) => { api = value; }} />);
  await opened();
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
  fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
  const later = api.editCellAsync(0, 0, 0, 'later');
  await act(async () => snapshot.resolve('initial'));
  await advance();
  expect(host.preview.mock.calls[0][2]).toMatchObject([{ method: 'editCell', args: [0, 0, 0, 'initial'] }]);
  expect(host.preview.mock.calls.at(-1)![2]).toMatchObject([{ method: 'editCell', args: [0, 0, 0, 'later'] }]);
  expect(host.editMethods.editCell).not.toHaveBeenCalled();
  await act(async () => hydration.resolve(host.peer));
  await advance();
  await later;
  expect(host.cells.get('0:0:0')).toBe('later');
  expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'later']]);
});

it('keeps the preview until worker adoption and never paints a peer frame', async () => {
  const host = harness();
  const view = render(<XlsxEditor file={file} experimentalWorkerOpen showToolbar={false} />);
  await opened();
  host.peerMethods.displayList.mockImplementation(() => { throw new Error('Peer grid paint forbidden'); });
  const worker = deferred<xlsx.WorkbookFrame>();
  host.sessionMethods.frame.mockReturnValue(worker.promise);
  reviewEdit(view, 'worker owned');
  await advance();
  expect(host.cells.get('0:0:0')).toBe('worker owned');
  expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('worker owned');
  expect(host.peerMethods.displayList).not.toHaveBeenCalled();
  expect(painted.every((list) => list.commands.every((cmd) => cmd.op !== 'text' || !cmd.text.startsWith('peer:')))).toBe(true);
  const [viewport, options] = host.sessionMethods.frame.mock.calls.at(-1)!;
  await act(async () => worker.resolve({ sheet: options?.sheet ?? 0, viewport, version: 'v2', epoch: 2, sequence: 1,
    displayList: { width: 800, height: 600, commands: [{ op: 'text', text: 'worker:worker owned', x: 8, y: 18, fontSize: 11, color: '#000000' }] } }));
  await advance();
  expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
  expect(host.peerMethods.displayList).not.toHaveBeenCalled();
  expect(painted.at(-1)!.commands).toContainEqual({ op: 'text', text: 'worker:worker owned', x: 8, y: 18, fontSize: 11, color: '#000000' });
});

for (const route of ['save', 'saveAsync', 'command', 'shortcut', 'toolbar'] as const) {
  it(`routes every normal XLSX save through the worker and saves the peer only during recovery (${route})`, async () => {
    const host = harness(true);
    let api!: XlsxWorkerEditorApi;
    const delivery = mock((_bytes: Uint8Array) => {});
    const view = render(<XlsxEditor file={file} experimentalWorkerOpen onReady={(value) => { api = value; }} onSave={delivery} />);
    await opened();
    const ack = deferred<WorkbookReplayReply>();
    host.replay.mockReturnValueOnce(ack.promise);
    reviewEdit(view, 'saved');
    let saving: Promise<unknown> | undefined;
    if (route === 'save' || route === 'saveAsync') act(() => { saving = api[route](); });
    else if (route === 'command') act(() => { saving = api.commands.execute('save', null); });
    else if (route === 'shortcut') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 's', ctrlKey: true });
    else fireEvent.click(view.getByRole('button', { name: /^save/i }));
    await advance();
    expect(host.session.save).not.toHaveBeenCalled();
    expect(host.peerMethods.save).not.toHaveBeenCalled();
    await act(async () => ack.resolve({ sequence: 1, revision: 1, version: 1, result: undefined }));
    await advance();
    await saving;
    expect(host.session.save).toHaveBeenCalledTimes(1);
    expect(host.peerMethods.save).not.toHaveBeenCalled();
    if (route !== 'save' && route !== 'saveAsync') expect(delivery).toHaveBeenCalledWith(new Uint8Array([0]));
    act(() => host.fail());
    const recovered = await api.recoverySave();
    expect(recovered.recovery).toBe(true);
    expect(host.peerMethods.save).toHaveBeenCalledTimes(1);
    expect(host.session.save).toHaveBeenCalledTimes(1);
    expect(JSON.parse(new TextDecoder().decode(recovered.bytes))).toEqual([['0:0:0', 'saved']]);
  });
}
