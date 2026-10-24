import { GlobalRegistrator } from '@happy-dom/global-registrator';
import type { AwarenessPeer, CollaborationAwareness, CollaborationReplica } from '@betteroffice/xlsx/collaboration';
import { createAsyncCollaborationReplica } from '@betteroffice/xlsx/collaboration';
import { Profiler, StrictMode } from 'react';
import { flushSync } from 'react-dom';
import { afterAll, afterEach, beforeEach, describe, expect, it, mock, spyOn } from 'bun:test';
import * as xlsx from '@betteroffice/xlsx';
import * as xlsxWorker from '@betteroffice/xlsx/worker';
import type { DisplayList, EditResult, WorkbookHandle, WorkbookSession, XlsxEditRequest } from '@betteroffice/xlsx';
import type { WorkbookEditPeer } from '@betteroffice/xlsx/worker';
import { SessionFailure } from '../../../../shared/office-session';
import { workbookEditPeerInternals } from '../../../xlsx/src/session/editPeerInternals';
import { applyWorkbookCommand, createWorkbookCommandState } from '../../../xlsx/src/session/commandHandlers';
import { workbookSessionInternals, type WorkbookReplayEnvelope, type WorkbookReplayReply } from '../../../xlsx/src/session/replay';
import { XlsxEditor, type XlsxWorkerViewerApi } from '../XlsxEditor';
import type { XlsxAsyncEditorApi, XlsxAsyncEditorProps } from '../index';
import { EditorToolbar } from '../components/EditorToolbar';
import { ToolbarCommandButton } from '../components/toolbar/ToolbarCommand';
import { XlsxCommandProvider } from '../commands/XlsxCommandProvider';
import { PLAIN_FORMATTING, testSessionCommands } from '../commands/testing';
import { useXlsxCommands } from '../commands/hooks';
import { XlsxCommandAdmissionError, xlsxCommandController } from '../commands/createXlsxCommandStore';

import * as peerlessInput from '../commands/peerlessWorkerInputCoordinator';
import type { PeerlessWorkerCommand } from '../commands/peerlessWorkerInputCoordinator';
import type { XlsxCommandStore } from '../commands/types';
import { defineXlsxPlugin } from '../plugins/defineXlsxPlugin';
import type { XlsxPluginContext, XlsxPluginSelection } from '../plugins/types';
import { workbookSessionOpener } from '../viewer/useSessionWorkbook';
import { XlsxWorkerEditorCollaborationError, type XlsxWorkerEditorApi } from './createWorkerEditorApi';
import { XlsxWorkerEditor, xlsxPeerlessEditorLoader } from './XlsxWorkerEditor';
import { editableWorkbookSessionBackend } from './useEditableSessionWorkbook';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, getConfig, render, waitFor, within } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
const file = new Uint8Array([1, 2, 3]);
let animationFrames: Map<number, FrameRequestCallback>;
let frameRequests: (() => void)[];
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

function harness(realFacade = false, normalizeInput = false) {
  const normalize = (input: string) => !normalizeInput ? input : input === '001' ? '1' :
    input === 'true' ? 'TRUE' : input === 'false' ? 'FALSE' : input.replace(/^'/, '');
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
    cells.set(`${sheet}:${row}:${col}`, normalize(input));
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
    cellRect: (_sheet: number, row: number, col: number) => ({ x: col * 96, y: row * 24, w: 96, h: 24 }),
    displayList: mock(() => { log.push('display:peer'); return display('peer'); }),
    mergedRanges: () => [], listProposals: () => [],
    selectionData: (sheet: number) => {
      const used = [...cells.entries()].filter(([key, input]) => input && key.startsWith(`${sheet}:`))
        .map(([key]) => key.split(':').map(Number));
      return { usedRange: used.length === 0 ? null : {
        top: Math.min(...used.map(([, row]) => row!)), left: Math.min(...used.map(([, , col]) => col!)),
        bottom: Math.max(...used.map(([, row]) => row!)), right: Math.max(...used.map(([, , col]) => col!)),
      } };
    },
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
    get acknowledgedSequence() { return sequence; },
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
    selectionData: mock<WorkbookSession['call']['selectionData']>(async (sheet) => peerMethods.selectionData(sheet)),
    cellGeometry: mock(async (sheet: number, row: number, col: number): Promise<xlsx.WorkbookCellGeometry> => ({
      sheet, version: `v${version}`, rect: { x: col * 96, y: row * 24, w: 96, h: 24 },
      scrollPosition: { x: col * 96, y: row * 24 },
    })),
  };
  const sessionSave = mock<WorkbookSession['save']>(async () => new Uint8Array([0]));
  const session: WorkbookSession = {
    ...testSessionCommands(),
    state: { format: 'xlsx', stage: 'ready', version: 0, dirty: false, activeSheet: 0,
      sheets: [{ id: 'sheet:0', index: 0, name: 'First' }, { id: 'sheet:1', index: 1, name: 'Second' }] },
    call: sessionMethods as unknown as WorkbookSession['call'], save: sessionSave,
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
    for (const op of ops) {
      if (op.method === 'editCells') {
        for (const edit of op.args[1]) values.set(`${op.args[0]}:${edit.row}:${edit.col}`, normalize(edit.input));
      }
      if (op.method === 'applyEdits') {
        for (const step of op.args[0].steps) {
          if (step.op !== 'setCellInputs') continue;
          const sheet = Number(step.target.sheetId.split(':')[1]);
          const start = step.target.range.kind === 'rowCol' ? step.target.range.start : { row: 0, col: 0 };
          step.inputs.forEach((row, dr) => row.forEach((input, dc) =>
            values.set(`${sheet}:${start.row + dr}:${start.col + dc}`, normalize(input))));
        }
      }
      if (op.method !== 'editCell') continue;
      if (op.args[3].startsWith('=') && new TextEncoder().encode(op.args[3].slice(1)).byteLength > 32768) {
        throw new Error('Formula exceeds the length limit');
      }
      values.set(`${op.args[0]}:${op.args[1]}:${op.args[2]}`, normalize(op.args[3]));
    }
    return { sheet, viewport, version: `v${version}`, epoch: 0, sequence,
      mergedRanges: [], displayList: display('preview', values) };
  });
  workbookSessionInternals.set(session, { replay, editPeerAttached: false, preview,
    cellInput: async (sheet, row, col) => workerCells.get(`${sheet}:${row}:${col}`) ?? '' });
  if (!realFacade) workbookEditPeerInternals.set(edits, {
    fail: () => {},
    whenAcknowledged: async () => {},
    applyQueuedOp: (op) => {
      if (op.method === 'applyEdits') return editMethods.applyEdits(op.args[0]);
      if (op.method === 'editCell') return editMethods.editCell(...op.args);
      if (op.method === 'editCells') return editMethods.editCells(...op.args);
      throw new Error('Unexpected mock operation');
    },
    applyRecoveryOp: (op) => {
      if (op.method === 'editCell') return xlsxWorker.createWorkbookRecoveryMutators(edits).editCell(...op.args);
      if (op.method === 'editCells') return peerMethods.editCells(...op.args);
      throw new Error('Unexpected mock recovery operation');
    },
  });
  let attached: WorkbookEditPeer | null = null;
  const attach = spyOn(editableWorkbookSessionBackend, 'attach').mockImplementation((options) => {
    attached = realFacade ? xlsxWorker.createWorkbookEditPeer(options) : edits;
    return attached;
  });
  restorers.push(() => open.mockRestore(), () => hydrate.mockRestore(), () => attach.mockRestore());
  return { log, cells, charts, peer, peerMethods, edits, editMethods, session, sessionSave, sessionMethods, open, hydrate, replay, preview,
    get attached() { return attached; },
    fail(error = new SessionFailure('crash', 'Worker stopped')) {
      failure = error;
      for (const listener of [...failureListeners]) listener(error);
      return error;
    } };
}

function promisedClipboard(completion?: Promise<void>) {
  const originalItem = Object.getOwnPropertyDescriptor(globalThis, 'ClipboardItem');
  const originalWrite = Object.getOwnPropertyDescriptor(navigator.clipboard, 'write');
  class Item {
    constructor(readonly data: Record<string, Promise<Blob>>) {}
  }
  const payloads: Promise<Blob>[] = [];
  const writes: Promise<void>[] = [];
  const copied: string[] = [];
  const itemsWritten: Item[] = [];
  const write = mock((items: Item[]) => {
    itemsWritten.push(items[0]);
    const payload = items[0].data['text/plain'];
    payloads.push(payload);
    const writing = payload.then(async (blob) => {
      copied.push(await blob.text());
      if (completion) await completion;
    });
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
  return { write, payloads, writes, copied, itemsWritten };
}

function markedClipboard(completion?: Promise<void>) {
  const clipboard = promisedClipboard(completion);
  const latest = () => clipboard.itemsWritten[clipboard.itemsWritten.length - 1];
  replaceClipboard('readText', mock(async () => (await latest().data['text/plain']).text()));
  replaceClipboard('read', mock(async () => {
    const item = latest();
    return item ? [{ types: ['text/plain', 'text/html'],
      getType: async (type: string) => item.data[type],
    } as unknown as ClipboardItem] : [];
  }));
  return clipboard;
}

function textClipboard() {
  const original = Object.getOwnPropertyDescriptor(globalThis, 'ClipboardItem');
  Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: undefined });
  restorers.push(() => {
    if (original) Object.defineProperty(globalThis, 'ClipboardItem', original);
    else Reflect.deleteProperty(globalThis, 'ClipboardItem');
  });
  const writing = deferred<void>();
  const writeText = mock((_text: string) => writing.promise);
  const write = mock(() => Promise.resolve());
  replaceClipboard('writeText', writeText);
  replaceClipboard('write', write);
  return { writing, writeText, write };
}

async function opened() {
  if (animationFrames.size === 0) {
    await getConfig().asyncWrapper(() => new Promise<void>((resolve) => frameRequests.push(resolve)));
  }
  await tick();
}

beforeEach(() => {
  const originalCanvas = Object.getOwnPropertyDescriptor(globalThis, 'OffscreenCanvas');
  Object.defineProperty(globalThis, 'OffscreenCanvas', { configurable: true, writable: true, value: class {} });
  restorers.push(() => {
    if (originalCanvas) Object.defineProperty(globalThis, 'OffscreenCanvas', originalCanvas);
    else Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
  });
  const originalWorker = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
  Object.defineProperty(globalThis, 'Worker', { configurable: true, writable: true, value: class {} });
  restorers.push(() => {
    if (originalWorker) Object.defineProperty(globalThis, 'Worker', originalWorker);
    else Reflect.deleteProperty(globalThis, 'Worker');
  });
  animationFrames = new Map(); frameRequests = []; painted = []; observers = [];
  const originalRaf = globalThis.requestAnimationFrame;
  const originalCancel = globalThis.cancelAnimationFrame;
  const originalResize = globalThis.ResizeObserver;
  globalThis.requestAnimationFrame = (callback) => {
    const id = ++nextAnimation; animationFrames.set(id, callback);
    for (const notify of frameRequests.splice(0)) notify();
    return id;
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
afterEach(async () => {
  cleanup();
  await advance();
  for (const restore of restorers.reverse()) restore();
  restorers.length = 0;
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

function peerlessHarness(toolbarCommands = false, workerCommands = false, defaultToolbar = false) {
  const listeners = new Set<(ack: xlsxWorker.WorkbookAck) => void>();
  const failures = new Set<Parameters<WorkbookSession['onFailure']>[0]>();
  const sent: PeerlessWorkerCommand[] = [];
  const commandWaiters = new Map<number, ReturnType<typeof deferred<PeerlessWorkerCommand>>>();
  const commandAt = (index: number) => {
    if (sent[index]) return Promise.resolve(sent[index]);
    let waiter = commandWaiters.get(index);
    if (!waiter) { waiter = deferred<PeerlessWorkerCommand>(); commandWaiters.set(index, waiter); }
    return waiter.promise;
  };
  const replies = new Map<number, ReturnType<typeof deferred<xlsxWorker.CommandStatus>>>();
  const presentations: { seq: number; resolve(): void }[] = [];
  const cells = new Map<string, string>([['0:0:0', 'initial']]);
  const charts: xlsx.ChartRegion[] = [];
  const hidden = { rows: [] as [number, number][], cols: [] as [number, number][] };
  let throughSeq = 0;
  const confirmedBasis = { epoch: 1, version: 'v1' };
  let presentedThroughSeq = 0;
  let selection: xlsxWorker.WorkbookCommandSelection | undefined;
  let editState: Partial<xlsxWorker.WorkbookEditState> = {
    formatting: { ...PLAIN_FORMATTING }, formattingAvailable: true, selectionEditable: true,
    canUndo: true, canRedo: true, readOnly: false, mergedRanges: [], activeSheet: 0,
    sheets: [{ id: 'sheet:0', index: 0, name: 'First' }, { id: 'sheet:1', index: 1, name: 'Second' }],
  };
  const display = (viewport?: xlsx.Viewport): DisplayList => ({ width: 800, height: 600, charts,
    grid: { startRow: Math.floor((viewport?.y ?? 0) / 24), startCol: 0,
      hiddenRowSpans: hidden.rows, hiddenColSpans: hidden.cols,
      rowIndices: Array.from({ length: 8 }, (_, index) => Math.floor((viewport?.y ?? 0) / 24) + index)
        .filter((row) => !hidden.rows.some(([start, end]) => row >= start && row <= end)).slice(0, 4),
      colIndices: Array.from({ length: 6 }, (_, col) => col)
        .filter((col) => !hidden.cols.some(([start, end]) => col >= start && col <= end)).slice(0, 3),
      rowOffsets: [0, 24, 48, 72, 96], colOffsets: [0, 96, 192, 288] },
    commands: [...cells].map(([key, text]) => {
      const [, row, col] = key.split(':').map(Number);
      return { op: 'text' as const, text, x: col * 96 + 8, y: row * 24 + 18, fontSize: 11, color: '#000000' };
    }),
  });
  const sheetView = mock(async (sheet: number): Promise<xlsx.WorkbookSheetView> => ({
    sheet, version: 'v1', contentWidth: 4000, contentHeight: 6000, frozenRows: 0, frozenCols: 0,
    frozenWidth: 0, frozenHeight: 0, initialScrollX: 0, initialScrollY: 0,
  }));
  const selectionData = mock<WorkbookSession['call']['selectionData']>(async () => ({ edge: null, region: null }));
  const cellInputs = mock(async (sheet: number, a1: string) => {
    const [, column, row] = /^([A-Z]+)(\d+)/.exec(a1)!;
    const col = [...column].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1;
    return { sheet, version: `v${version}`,
      cells: [[{ a1, input: cells.get(`${sheet}:${Number(row) - 1}:${col}`) ?? '', isFormula: false }]],
    };
  });
  const cellGeometry = mock(async (sheet: number, row: number, col: number): Promise<xlsx.WorkbookCellGeometry> => ({
    sheet, version: 'v1', rect: { x: col * 96, y: row * 24, w: 96, h: 24 },
    scrollPosition: { x: col * 96, y: row * 24 },
  }));
  const frame = mock(async (viewport: xlsx.Viewport, options?: { sheet?: number }) => ({
    displayList: display(viewport), viewport, sheet: options?.sheet ?? 0, sequence: 100 + throughSeq,
    version: `v${version}`, epoch: 1,
  }));
  const session: WorkbookSession = {
    ...testSessionCommands(),
    confirmedBasis,
    state: { format: 'xlsx', stage: 'ready', version: 0, dirty: false, activeSheet: 0, sheets: editState.sheets! },
    get editState() { return editState; }, get selection() { return selection; },
    failure: undefined,
    get throughSeq() { return throughSeq; }, get presentedThroughSeq() { return presentedThroughSeq; },
    call: { sheetView, cellInputs, cellGeometry, selectionData, frame } as unknown as WorkbookSession['call'],
    submit: mock(async (command: xlsxWorker.WorkbookCommand) => {
      sent.push(command);
      commandWaiters.get(sent.length - 1)?.resolve(command);
      const reply = deferred<xlsxWorker.CommandStatus>();
      replies.set(command.seq, reply);
      return reply.promise;
    }),
    onAck: (listener) => { listeners.add(listener); return () => { listeners.delete(listener); }; },
    on: () => () => {}, onFailure: (listener) => { failures.add(listener); return () => { failures.delete(listener); }; },
    save: mock(async () => new Uint8Array([8, 9])), dispose: mock(async () => {}),
    markPresented: mock((paint) => {
      presentedThroughSeq = Math.max(presentedThroughSeq, paint.coversThroughSeq);
      for (const pending of presentations.splice(0)) {
        if (pending.seq <= presentedThroughSeq) pending.resolve();
        else presentations.push(pending);
      }
      return true;
    }),
    whenPresented: mock(async (seq = throughSeq) => {
      if (seq > presentedThroughSeq) await new Promise<void>((resolve) => presentations.push({ seq, resolve }));
    }),
  };
  const open = spyOn(editableWorkbookSessionBackend, 'open').mockResolvedValue(session);
  const hydrate = spyOn(editableWorkbookSessionBackend, 'hydrate').mockImplementation(() => { throw new Error('Peer forbidden'); });
  const attach = spyOn(editableWorkbookSessionBackend, 'attach').mockImplementation(() => { throw new Error('Edit peer forbidden'); });
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
    save() {}, restore() {}, setTransform() {}, beginPath() {}, rect() {}, clip() {},
  } as unknown as CanvasRenderingContext2D);
  restorers.push(() => open.mockRestore(), () => hydrate.mockRestore(), () => attach.mockRestore(), () => context.mockRestore());
  const ready = mock(() => {});
  const renderCount = mock(() => {});
  let commands!: XlsxCommandStore;
  function Chrome() {
    commands = useXlsxCommands();
    return <EditorToolbar mode="commands">
      {toolbarCommands && <EditorToolbar.Toolbar />}
      <EditorToolbar.FormulaBar />
    </EditorToolbar>;
  }
  const mount = async (onSave?: (bytes: Uint8Array) => void, readOnly = false, hostProps: Partial<Omit<XlsxAsyncEditorProps, 'file'>> = {}) => {
    editState = { ...editState, readOnly };
    const view = render(<Profiler id="peerless" onRender={renderCount}>
      <XlsxWorkerEditor backend="peerless" file={file} readOnly={readOnly} toolbar={defaultToolbar ? undefined : <Chrome />} onReady={ready} onSave={onSave} {...hostProps} />
    </Profiler>);
    await opened();
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    return view;
  };
  const push = (ack: xlsxWorker.WorkbookAck) => {
    throughSeq = Math.max(throughSeq, ack.throughSeq);
    confirmedBasis.version = ack.version;
    editState = { ...editState, ...ack.editState };
    if (ack.selection) selection = ack.selection;
    for (const listener of listeners) listener(ack);
    for (const status of ack.statuses) replies.get(status.seq)?.resolve(status);
  };
  let version = 1;
  const workbook = {
    sheetCount: () => 2,
    sheetInfo: () => ({ sheetIds: ['sheet:0', 'sheet:1'], sheetNames: ['First', 'Second'], activeSheet: 0 }),
    version: () => `v${version}`,
    historyState: () => ({ canUndo: version > 1, canRedo: false, undoDepth: version - 1, redoDepth: 0 }),
    validateEdits: () => ({ ok: true, version: `v${version}` }),
    readCells: () => ({ ok: true, sheets: [{ editable: true }, { editable: true }] }),
    selectionFormatting: () => PLAIN_FORMATTING,
    mergedRanges: () => [], calculationStatus: () => ({ limitedCells: [] }),
    editCell: (sheet: number, row: number, col: number, input: string) => {
      cells.set(`${sheet}:${row}:${col}`, input);
      version++;
      return { applied: true };
    },
    editCells: (sheet: number, edits: xlsx.CellInputEdit[]) => {
      for (const cell of edits) cells.set(`${sheet}:${cell.row}:${cell.col}`, cell.input);
      version++;
      return { applied: true };
    },
  } as unknown as WorkbookHandle;
  const commandState = createWorkbookCommandState(session.commandIdentity, session.confirmedBasis,
    { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }, ['sheet:0', 'sheet:1']);
  const ack = (command: PeerlessWorkerCommand, options: {
    status?: xlsxWorker.CommandStatus['status']; paint?: boolean; result?: unknown; selection?: xlsxWorker.WorkbookCommandSelection;
  } = {}) => {
    const { frame: _wireFrame, ...applied } = workerCommands ? applyWorkbookCommand(workbook, commandState, command) : {};
    const reading = command.kind === 'select' || command.kind === 'applyOps' && 'toolbar' in command.payload &&
      (command.payload.toolbar.id === 'paintFormat' || command.payload.toolbar.id === 'refreshProposals');
    const status = { seq: command.seq, status: options.status ?? 'applied', committed: !reading,
      ...(options.status === 'refused' ? { reason: 'invalid-step' } : {}),
      result: options.result ?? (command.kind === 'applyUpdate' ? { remoteUpdate: true, proposals: [] } : undefined) };
    if (status.status === 'applied' && command.kind === 'editCell') {
      const [sheet, row, col, text] = command.payload.args;
      cells.set(`${sheet}:${row}:${col}`, text);
      if (!workerCommands && command.seq > throughSeq) version++;
    }
    if (status.status === 'applied' && command.kind === 'applyUpdate') {
      cells.set('0:0:0', new TextDecoder().decode(command.payload.update));
      version++;
    }
    if (!workerCommands && status.status === 'applied' && command.seq > throughSeq &&
      (command.kind === 'undo' || command.kind === 'redo')) version++;
    const request = command.frame;
    push({ throughSeq: command.seq, version: `v${version}`, statuses: [status], ...applied,
      ...(options.selection ? { selection: options.selection } : {}),
      ...(options.paint && request ? { frame: {
        ...session.commandIdentity, coversThroughSeq: command.seq,
        requestGeneration: request.requestGeneration ?? request.generation,
        zoom: request.zoom, dpr: request.dpr, sheet: request.sheet, viewport: request.viewport,
        displayList: display(request.viewport), sequence: 100 + command.seq, epoch: 1, version: `v${version}`,
      } } : {}),
    });
  };
  const replace = (view: ReturnType<typeof render>, replacement = new Uint8Array([4])) => view.rerender(
    <Profiler id="peerless" onRender={renderCount}>
      <XlsxWorkerEditor backend="peerless" file={replacement} toolbar={<Chrome />} onReady={ready} />
    </Profiler>);
  const asyncEditor = (props: Partial<XlsxAsyncEditorProps> = {}) => (
    <XlsxEditor file={file} toolbar={<Chrome />} {...props} />
  );
  const setReadOnly = (view: ReturnType<typeof render>, readOnly: boolean) => view.rerender(
    <Profiler id="peerless" onRender={renderCount}>
      <XlsxWorkerEditor backend="peerless" file={file} readOnly={readOnly} toolbar={<Chrome />} onReady={ready} />
    </Profiler>);
  const fail = (error: SessionFailure) => { for (const listener of failures) listener(error); };
  return { mount, replace, asyncEditor, setReadOnly, fail, ready, open, sent, commandAt, ack, push, session, frame, sheetView, cellInputs, cellGeometry, selectionData, hydrate, attach, cells, charts, hidden,
    renderCount, commands: () => commands };
}

for (const [key, reverse, address] of [
  ['ArrowRight', { key: 'ArrowLeft' }, 'C1'],
  ['Tab', { key: 'Tab', shiftKey: true }, 'C1'],
  ['ArrowDown', { key: 'ArrowUp' }, 'A3'],
  ['Enter', { key: 'Enter', shiftKey: true }, 'A3'],
] as const) {
  it(`navigates past hidden tracks with ${key} and its reverse in the default peerless editor`, async () => {
    const host = peerlessHarness();
    host.hidden.rows.push([1, 1]); host.hidden.cols.push([1, 1]);
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    await act(async () => { fireEvent.keyDown(scroll, { key }); });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe(address);
    const forward = host.sent[host.sent.length - 1]!;
    await act(async () => host.ack(forward));
    await act(async () => { fireEvent.keyDown(scroll, reverse); });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
    expect(host.sent[host.sent.length - 1]).toMatchObject({ kind: 'select', payload: { selection: {
      activeCell: { row: 0, col: 0 },
    } } });
    expect(host.hydrate).not.toHaveBeenCalled();
  });
}

for (const [key, address] of [['Tab', 'C1'], ['Enter', 'A3']] as const) {
  it(`skips hidden tracks after committing cell input with ${key}`, async () => {
    const host = peerlessHarness();
    host.hidden.rows.push([1, 1]); host.hidden.cols.push([1, 1]);
    const view = await host.mount();
    await act(async () => { fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x' }); });
    await act(async () => { fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key }); });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe(address);
    expect(host.sent.some((command) => command.kind === 'editCell')).toBe(true);
    expect(host.hydrate).not.toHaveBeenCalled();
  });
}

it('extends past hidden tracks while keeping the peerless selection anchor', async () => {
  const host = peerlessHarness();
  host.hidden.rows.push([1, 1]); host.hidden.cols.push([1, 1]);
  const view = await host.mount();
  const scroll = view.getByTestId('xlsx-scroll');
  await act(async () => { fireEvent.keyDown(scroll, { key: 'ArrowRight', shiftKey: true }); });
  await act(async () => { fireEvent.keyDown(scroll, { key: 'ArrowDown', shiftKey: true }); });
  expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
  expect(host.sent[host.sent.length - 1]).toMatchObject({ kind: 'select', payload: { selection: {
    range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 2, col: 2 } },
    activeCell: { row: 0, col: 0 },
  } } });
});

it('pages over offscreen hidden rows in the peerless editor', async () => {
  const host = peerlessHarness();
  host.hidden.rows.push([10, 49]);
  const view = await host.mount();
  const scroll = view.getByTestId('xlsx-scroll');
  await act(async () => { fireEvent.keyDown(scroll, { key: 'PageDown' }); });
  expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A66');
  expect(host.cellGeometry).toHaveBeenCalledWith(0, 65, 0);
});

it('keeps peerless selection when the remaining rows and columns are hidden', async () => {
  const host = peerlessHarness();
  host.hidden.rows.push([1, 1_048_575]); host.hidden.cols.push([1, 16_383]);
  const view = await host.mount();
  const scroll = view.getByTestId('xlsx-scroll');
  for (const key of ['ArrowRight', 'ArrowDown', 'Tab', 'Enter', 'PageDown']) {
    await act(async () => { fireEvent.keyDown(scroll, { key }); });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
  }
  await act(async () => { fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true }); });
  expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
  expect(host.hydrate).not.toHaveBeenCalled();
});

for (const modifier of [{ ctrlKey: true }, { metaKey: true }]) {
  it(`selects the current region then the full extent with ${Object.keys(modifier)[0]} without moving focus`, async () => {
    const host = peerlessHarness();
    host.selectionData.mockResolvedValue({ region: { top: 0, left: 0, bottom: 2, right: 2 } });
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.mouseDown(scroll, { clientX: 110, clientY: 30 });
    fireEvent.mouseUp(window);
    await act(async () => { for (const command of host.sent) host.ack(command); });
    await act(async () => { fireEvent.keyDown(scroll, { key: 'a', ...modifier }); });
    expect(host.selectionData).toHaveBeenLastCalledWith(0, 1, 1, undefined);
    const first = host.sent[host.sent.length - 1]!;
    expect(first).toMatchObject({ kind: 'select', payload: { selection: {
      range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 2, col: 2 } },
      activeCell: { row: 1, col: 1 },
    } } });
    await act(async () => host.ack(first, { selection: {
      sheetId: 'sheet:0', range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 2, col: 2 } },
      activeCell: { row: 1, col: 1 },
    } }));
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('B2');
    expect((view.getByTestId('xlsx-selection') as HTMLElement).style.width).toBe('288px');
    expect((view.getByTestId('xlsx-selection') as HTMLElement).style.height).toBe('72px');
    await act(async () => { fireEvent.keyDown(scroll, { key: 'A', ...modifier }); });
    expect(host.sent[host.sent.length - 1]).toMatchObject({ kind: 'select', payload: { selection: {
      range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 1_048_575, col: 16_383 } },
      activeCell: { row: 1, col: 1 },
    } } });
    expect(host.hydrate).not.toHaveBeenCalled();
  });
}

it('reveals an offscreen data edge and extends from the original anchor', async () => {
  const host = peerlessHarness();
  host.selectionData.mockResolvedValue({ edge: { row: 179, col: 0 } });
  const view = await host.mount();
  const scroll = view.getByTestId('xlsx-scroll');
  await act(async () => { fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true, shiftKey: true }); });
  expect(host.selectionData).toHaveBeenCalledWith(0, 0, 0, 'down');
  expect(host.sent[0]).toMatchObject({ kind: 'select', payload: { selection: {
    range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 179, col: 0 } },
    activeCell: { row: 0, col: 0 },
  } } });
  expect(host.cellGeometry).toHaveBeenCalledWith(0, 179, 0);
  expect(scroll.scrollTop).toBe(180 * 24 - 600 + 2);
});

it('keeps a revealed distant cell in reach after an edit refreshes the sheet bounds', async () => {
  const host = peerlessHarness();
  host.selectionData.mockResolvedValue({ edge: { row: 9_999, col: 0 } });
  const view = await host.mount();
  const scroll = view.getByTestId('xlsx-scroll');
  await act(async () => { fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true }); });
  const extent = scroll.firstElementChild as HTMLElement;
  expect(extent.style.height).toBe('240002px');
  const reads = host.sheetView.mock.calls.length;
  await act(async () => { fireEvent.keyDown(scroll, { key: 'Delete' }); });
  expect(host.sent[host.sent.length - 1]).toMatchObject({ kind: 'applyOps' });
  await act(async () => { for (const command of [...host.sent]) host.ack(command); });
  expect(host.sheetView.mock.calls.length).toBeGreaterThan(reads);
  expect(extent.style.height).toBe('240002px');
  expect(extent.style.width).toBe('4000px');
});

for (const modifier of [{ ctrlKey: true }, { metaKey: true }]) {
  it(`reveals the used-range corner with ${Object.keys(modifier)[0]}+End and retains it at 150% zoom`, async () => {
    const host = peerlessHarness();
    host.selectionData.mockResolvedValue({ usedRange: { top: 0, left: 0, bottom: 9_999, right: 701 } });
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    await act(async () => { fireEvent.keyDown(scroll, { key: 'End', ...modifier }); });
    expect(host.selectionData).toHaveBeenCalledWith(0, 0, 0, 'usedRange');
    expect(host.sent[0]).toMatchObject({ kind: 'select', payload: { selection: {
      range: { kind: 'rowCol', start: { row: 9_999, col: 701 }, end: { row: 9_999, col: 701 } },
      activeCell: { row: 9_999, col: 701 },
    } } });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('ZZ10000');
    expect(host.cellGeometry).toHaveBeenCalledWith(0, 9_999, 701);
    expect(scroll.scrollLeft).toBe(702 * 96 - 800 + 2);
    expect(scroll.scrollTop).toBe(10_000 * 24 - 600 + 2);
    const extent = scroll.firstElementChild as HTMLElement;
    expect(extent.style.width).toBe('67394px');
    expect(extent.style.height).toBe('240002px');
    await act(async () => { await host.commands().execute('zoom', { scale: 1.5 }); });
    expect(scroll.scrollLeft).toBeCloseTo(702 * 96 * 1.5 - 800 + 2, 6);
    expect(scroll.scrollTop).toBeCloseTo(10_000 * 24 * 1.5 - 600 + 2, 6);
    expect(extent.style.width).toBe('101091px');
    expect(extent.style.height).toBe('360003px');
    await act(async () => { fireEvent.keyDown(scroll, { key: 'Home', ...modifier }); });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
    expect(scroll.scrollLeft).toBe(0);
    expect(scroll.scrollTop).toBe(0);
    expect(host.hydrate).not.toHaveBeenCalled();
  });
}

it('moves plain End to the last data cell of the row and Shift+End extends to it', async () => {
  const host = peerlessHarness();
  host.selectionData.mockResolvedValue({ edge: { row: 0, col: 7 } });
  const view = await host.mount();
  const scroll = view.getByTestId('xlsx-scroll');
  await act(async () => { fireEvent.keyDown(scroll, { key: 'End', shiftKey: true }); });
  expect(host.selectionData).toHaveBeenCalledWith(0, 0, 0, 'rowEnd');
  expect(host.sent[0]).toMatchObject({ kind: 'select', payload: { selection: {
    range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 0, col: 7 } },
    activeCell: { row: 0, col: 0 },
  } } });
  host.selectionData.mockResolvedValue({ edge: null });
  await act(async () => { fireEvent.keyDown(scroll, { key: 'End' }); });
  expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
});

for (const change of ['sheet', 'replacement']) {
  it(`drops an obsolete data edge after ${change}`, async () => {
    const host = peerlessHarness();
    const answer = deferred<xlsx.SelectionData>();
    host.selectionData.mockReturnValueOnce(answer.promise);
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.keyDown(scroll, { key: 'ArrowDown', metaKey: true });
    if (change === 'sheet') await act(async () => fireEvent.click(view.getByRole('tab', { name: 'Second' })));
    else await act(async () => host.replace(view));
    const commands = host.sent.length;
    const geometryReads = host.cellGeometry.mock.calls.length;
    await act(async () => answer.resolve({ edge: { row: 179, col: 0 } }));
    expect(host.sent).toHaveLength(commands);
    expect(host.cellGeometry).toHaveBeenCalledTimes(geometryReads);
  });
}

it('clears the cell a pending peerless data jump lands on, then applies later arrows in order', async () => {
  const host = peerlessHarness();
  const answer = deferred<xlsx.SelectionData>();
  host.selectionData.mockReturnValueOnce(answer.promise);
  const view = await host.mount();
  const scroll = view.getByTestId('xlsx-scroll');
  const before = host.sent.length;
  fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true });
  fireEvent.keyDown(scroll, { key: 'Delete' });
  fireEvent.keyDown(scroll, { key: 'ArrowRight' });
  await act(async () => {});
  expect(host.sent).toHaveLength(before);
  await act(async () => answer.resolve({ edge: { row: 3, col: 0 } }));
  const sent = host.sent.slice(before);
  expect(sent.map((command) => command.kind)).toEqual(['select', 'applyOps', 'select']);
  expect(sent[1]).toMatchObject({ payload: { toolbar: { id: 'clear', target: { range: {
    start: { row: 3, col: 0 }, end: { row: 3, col: 0 },
  } } } } });
  expect(sent[2]).toMatchObject({ payload: { selection: { activeCell: { row: 3, col: 1 } } } });
});

it('types keys queued behind a pending peerless data jump into the cell they open', async () => {
  const host = peerlessHarness();
  const answer = deferred<xlsx.SelectionData>();
  host.selectionData.mockReturnValueOnce(answer.promise);
  const view = await host.mount();
  const scroll = view.getByTestId('xlsx-scroll');
  fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true });
  for (const key of ['Delete', 'a', 'b', 'c', 'Enter']) fireEvent.keyDown(scroll, { key });
  await act(async () => answer.resolve({ edge: { row: 3, col: 0 } }));
  expect(host.sent.filter((command) => command.kind === 'editCell')[0]).toMatchObject({ payload: { args: [0, 3, 0, 'abc'] } });
});

it('resumes a data jump queued after a typed cell commits behind a pending peerless jump', async () => {
  const host = peerlessHarness();
  const answer = deferred<xlsx.SelectionData>();
  host.selectionData.mockImplementation(async (_sheet, _row, _col, direction) => direction === 'down'
    ? answer.promise : { edge: direction === 'right' ? { row: 4, col: 5 } : null, merge: null });
  const view = await host.mount();
  const scroll = view.getByTestId('xlsx-scroll');
  fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true });
  fireEvent.keyDown(scroll, { key: 'Delete' });
  fireEvent.keyDown(scroll, { key: 'a' });
  fireEvent.keyDown(scroll, { key: 'Enter' });
  fireEvent.keyDown(scroll, { key: 'ArrowRight', ctrlKey: true });
  fireEvent.keyDown(scroll, { key: 'Delete' });
  await act(async () => answer.resolve({ edge: { row: 3, col: 0 } }));
  await act(async () => {});
  expect(host.sent.filter((command) => command.kind === 'editCell')[0]).toMatchObject({ payload: { args: [0, 3, 0, 'a'] } });
  expect(host.sent.filter((command) => command.kind === 'applyOps').slice(-1)[0]).toMatchObject({ payload: { toolbar: {
    id: 'clear', target: { range: { start: { row: 4, col: 5 }, end: { row: 4, col: 5 } } },
  } } });
});

it('cancels a typed cell with Escape queued behind a pending peerless data jump', async () => {
  const host = peerlessHarness();
  const answer = deferred<xlsx.SelectionData>();
  host.selectionData.mockReturnValueOnce(answer.promise);
  const view = await host.mount();
  const scroll = view.getByTestId('xlsx-scroll');
  fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true });
  for (const key of ['Delete', 'a', 'Escape', 'Enter']) fireEvent.keyDown(scroll, { key });
  await act(async () => answer.resolve({ edge: { row: 3, col: 0 } }));
  expect(host.sent.some((command) => command.kind === 'editCell')).toBe(false);
});

it('starts clipboard access in the key press behind a pending peerless data jump', async () => {
  const host = peerlessHarness();
  const answer = deferred<xlsx.SelectionData>();
  host.selectionData.mockReturnValueOnce(answer.promise);
  const clipboard = promisedClipboard();
  const readText = mock(async () => 'pasted');
  replaceClipboard('readText', readText);
  const view = await host.mount();
  const scroll = view.getByTestId('xlsx-scroll');
  fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true });
  fireEvent.keyDown(scroll, { key: 'c', ctrlKey: true });
  fireEvent.keyDown(scroll, { key: 'v', ctrlKey: true });
  expect(clipboard.write).toHaveBeenCalledTimes(1);
  expect(readText).toHaveBeenCalledTimes(1);
  expect(host.sent.some((command) => command.kind === 'select')).toBe(false);
  await act(async () => answer.resolve({ edge: { row: 3, col: 0 } }));
});

it('chains rapid peerless data jumps from the latest accepted target', async () => {
  const host = peerlessHarness();
  const first = deferred<xlsx.SelectionData>();
  host.selectionData.mockReturnValueOnce(first.promise).mockResolvedValueOnce({ edge: { row: 0, col: 9 } });
  const view = await host.mount();
  const scroll = view.getByTestId('xlsx-scroll');
  fireEvent.keyDown(scroll, { key: 'ArrowRight', ctrlKey: true });
  fireEvent.keyDown(scroll, { key: 'ArrowRight', ctrlKey: true });
  expect(host.selectionData).toHaveBeenCalledTimes(1);
  await act(async () => first.resolve({ edge: { row: 0, col: 3 } }));
  expect(host.selectionData.mock.calls[1]).toEqual([0, 0, 3, 'right']);
  expect(host.sent).toHaveLength(2);
  expect(host.sent[1]).toMatchObject({ kind: 'select', payload: { selection: { activeCell: { row: 0, col: 9 } } } });
});

it('keeps a pending jump when a region selection is acknowledged at a corner', async () => {
  const host = peerlessHarness();
  const edge = deferred<xlsx.SelectionData>();
  host.selectionData.mockResolvedValueOnce({ region: { top: 0, left: 0, bottom: 2, right: 2 } })
    .mockReturnValueOnce(edge.promise);
  const view = await host.mount();
  const scroll = view.getByTestId('xlsx-scroll');
  await act(async () => { fireEvent.keyDown(scroll, { key: 'a', ctrlKey: true }); });
  fireEvent.keyDown(scroll, { key: 'ArrowRight', ctrlKey: true });
  await act(async () => host.ack(host.sent[0]!, { selection: {
    sheetId: 'sheet:0', range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 2, col: 2 } },
    activeCell: { row: 0, col: 0 },
  } }));
  await act(async () => edge.resolve({ edge: { row: 0, col: 2 } }));
  expect(host.sent).toHaveLength(2);
  expect(host.sent[1]).toMatchObject({ kind: 'select', payload: { selection: { activeCell: { row: 0, col: 2 } } } });
});

it('re-queries peerless data after a remote version change', async () => {
  const host = peerlessHarness();
  const first = deferred<xlsx.SelectionData>();
  const fresh = deferred<xlsx.SelectionData>();
  host.selectionData.mockReturnValueOnce(first.promise).mockReturnValueOnce(fresh.promise);
  const view = await host.mount();
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight', ctrlKey: true });
  host.session.confirmedBasis.version = 'v2';
  await act(async () => first.resolve({ edge: { row: 0, col: 3 } }));
  expect(host.selectionData).toHaveBeenCalledTimes(2);
  expect(host.sent).toHaveLength(0);
  await act(async () => fresh.resolve({ edge: { row: 0, col: 7 } }));
  expect(host.sent).toHaveLength(1);
  expect(host.sent[0]).toMatchObject({ kind: 'select', payload: { selection: { activeCell: { row: 0, col: 7 } } } });
});

it('falls back to the peerless directional limit after a query failure', async () => {
  const host = peerlessHarness();
  host.selectionData.mockRejectedValueOnce(new Error('Query failed'));
  const view = await host.mount();
  await act(async () => { fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight', ctrlKey: true }); });
  expect(host.sent[0]).toMatchObject({ kind: 'select', payload: { selection: { activeCell: { row: 0, col: 16_383 } } } });
  expect(view.queryByRole('alert')).toBeNull();
});

it('default readiness opens one worker session without a main-thread engine or edit peer', async () => {
  const host = peerlessHarness();
  const engine = spyOn(xlsx, 'openWorkbook');
  const ready = mock((_api: XlsxAsyncEditorApi) => {});
  try {
    render(host.asyncEditor({ onReady: ready }));
    await opened();
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    expect(host.open).toHaveBeenCalledTimes(1);
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
    expect(engine).not.toHaveBeenCalled();
    const api = ready.mock.calls[0]![0];
    expect(api.hydrated).toBe(true);
    expect('handle' in api).toBe(false);
    expect(api.save).toBeFunction();
  } finally { engine.mockRestore(); }
});

it('saving while F2 prefill is pending waits for the seed and preserves the untouched cell', async () => {
  const host = peerlessHarness();
  const ready = mock((_api: XlsxAsyncEditorApi) => {});
  const view = render(host.asyncEditor({ onReady: ready }));
  await opened();
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const prefill = deferred<Awaited<ReturnType<WorkbookSession['call']['cellInputs']>>>();
  const requested = host.cellInputs.mock.calls.length;
  host.cellInputs.mockImplementationOnce(() => prefill.promise);
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
  await waitFor(() => expect(host.cellInputs.mock.calls.length).toBeGreaterThan(requested));
  expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('initial');
  let saving!: Promise<unknown>;
  await act(async () => {
    saving = ready.mock.calls[0]![0].commands.execute('save', null);
    await Promise.resolve();
  });
  await act(async () => {});
  expect(host.session.save).not.toHaveBeenCalled();
  expect(view.getByTestId('xlsx-cell-editor')).toBeTruthy();
  await act(async () => prefill.resolve({ sheet: 0, version: 'v1',
    cells: [[{ a1: 'A1', input: 'initial', isFormula: false }]] }));
  const edits = host.sent.filter((command) => command.kind === 'editCell');
  await act(async () => {
    for (const command of host.sent) host.ack(command);
    await saving;
  });
  expect(edits).toEqual([]);
  expect(host.cells.get('0:0:0')).toBe('initial');
  expect(host.session.save).toHaveBeenCalledTimes(1);
  expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('initial');
});

for (const key of ['z', 'y', 'b', 'i']) {
  it(`routes Mod+${key} to the text input during an uncached pending prefill`, async () => {
    const host = peerlessHarness();
    const pending = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
    host.cellInputs.mockReturnValue(pending.promise);
    const view = await host.mount();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    const input = view.getByTestId('xlsx-cell-editor');
    expect(document.activeElement).toBe(input);
    fireEvent.keyDown(input, { key, ctrlKey: true });
    expect(host.sent).toEqual([]);
    expect(view.getByTestId('xlsx-cell-editor')).toBe(input);
    await act(async () =>
      pending.resolve({
        sheet: 0,
        version: 'v1',
        cells: [[{ a1: 'A1', input: 'initial', isFormula: false }]],
      }),
    );
  });
}

for (const reopen of [false, true]) {
  it(`updates B1 formula input when F2 interrupts its read${reopen ? ' and reopens after Escape' : ''}`, async () => {
    const host = peerlessHarness();
    host.cells.set('0:0:1', 'second');
    const view = await host.mount();
    const formula = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
    expect(formula.value).toBe('initial');
    const focused = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
    const prefill = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
    host.cellInputs.mockReturnValueOnce(focused.promise);
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
    await waitFor(() => expect(host.cellInputs).toHaveBeenLastCalledWith(0, 'B1'));
    await act(async () => {
      for (const command of host.sent) host.ack(command);
    });
    host.cellInputs.mockReturnValueOnce(prefill.promise);
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    await act(async () =>
      focused.resolve({
        sheet: 0,
        version: 'v1',
        cells: [[{ a1: 'B1', input: 'second', isFormula: false }]],
      }),
    );
    expect(formula.value).toBe('second');
    if (reopen) {
      await act(async () =>
        fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), {
          key: 'Escape',
        }),
      );
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
      expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('second');
      expect(formula.value).toBe('second');
    }
    await act(async () =>
      prefill.resolve({
        sheet: 0,
        version: 'v1',
        cells: [[{ a1: 'B1', input: 'second', isFormula: false }]],
      }),
    );
  });
}

it('updates the focused formula input across a proposal refresh without a paint', async () => {
  let coordinator!: ReturnType<typeof peerlessInput.createPeerlessWorkerInputCoordinator>;
  const original = peerlessInput.createPeerlessWorkerInputCoordinator;
  const creating = spyOn(peerlessInput, 'createPeerlessWorkerInputCoordinator').mockImplementation((hooks) => {
    coordinator = original(hooks);
    return coordinator;
  });
  restorers.push(() => creating.mockRestore());
  const host = peerlessHarness();
  const view = await host.mount();
  const focused = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
  host.cellInputs.mockReturnValueOnce(focused.promise);
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
  await waitFor(() => expect(host.cellInputs).toHaveBeenLastCalledWith(0, 'B1'));
  let ticket!: ReturnType<typeof coordinator.command>;
  act(() => {
    ticket = coordinator.command(
      {
        kind: 'applyOps',
        payload: { toolbar: { id: 'refreshProposals', args: null } },
      },
      { after: null, frame: undefined },
    );
  });
  await act(async () =>
    focused.resolve({
      sheet: 0,
      version: 'v1',
      cells: [[{ a1: 'B1', input: 'second', isFormula: false }]],
    }),
  );
  expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).value).toBe('second');
  await act(async () => {
    for (const command of host.sent) host.ack(command);
    await ticket.result;
  });
});

for (const cached of [true, false]) describe(`peerless F2 prefill keyboard replay (${cached ? 'cached' : 'uncached'})`, () => {
  const value = 'Q3 Sales Dashboard';
  const preview = cached ? value : '';
  const key = (input: HTMLInputElement, name: string, options: { ctrlKey?: boolean; metaKey?: boolean; altKey?: boolean; shiftKey?: boolean } = {}) => {
    const allowed = fireEvent.keyDown(input, { key: name, ...options });
    if (allowed && (options.ctrlKey || options.metaKey) && !options.altKey && name.toLowerCase() === 'a') input.select();
  };
  const type = (input: HTMLInputElement, text: string) => {
    act(() => {
      const allowed = input.dispatchEvent(new InputEvent('beforeinput', {
        bubbles: true, cancelable: true, inputType: 'insertText', data: text,
      }));
      if (allowed) {
        const previous = input.value;
        input.setRangeText(text, input.selectionStart ?? 0, input.selectionEnd ?? 0, 'end');
        const value = input.value;
        const caret = input.selectionStart ?? value.length;
        input.value = previous;
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, value);
        input.setSelectionRange(caret, caret);
        input.dispatchEvent(new Event('input', { bubbles: true }));
      }
    });
  };
  const end = (input: HTMLInputElement, data: string) => {
    const event = new Event('compositionend', { bubbles: true });
    Object.defineProperty(event, 'data', { value: data });
    fireEvent(input, event);
  };
  const setup = async (onSave: (bytes: Uint8Array) => void = () => {}) => {
    const host = peerlessHarness();
    host.cells.set('0:0:0', value);
    const focused = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
    const read = host.cellInputs.getMockImplementation()!;
    if (!cached) host.cellInputs.mockReturnValue(focused.promise);
    const view = await host.mount(onSave);
    host.cellInputs.mockImplementation(read);
    const prefill = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
    const requested = host.cellInputs.mock.calls.length;
    host.cellInputs.mockReturnValueOnce(prefill.promise);
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    await waitFor(() => expect(host.cellInputs.mock.calls.length).toBeGreaterThan(requested));
    const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    const resolve = () => act(async () => {
      const result = { sheet: 0, version: 'v1', cells: [[{ a1: 'A1', input: value, isFormula: false }]] };
      focused.resolve(result);
      prefill.resolve(result);
    });
    return { host, view, input, resolve };
  };

  it('selects all before prefill and replaces after prefill', async () => {
    const { input, resolve } = await setup();
    key(input, 'a', { ctrlKey: true });
    await resolve();
    type(input, 'probe');
    expect(input.value).toBe('probe');
  });

  it('selects all and replaces before prefill', async () => {
    const { input, resolve } = await setup();
    key(input, 'a', { ctrlKey: true });
    type(input, 'probe');
    await resolve();
    expect(input.value).toBe('probe');
  });

  it('appends at the end after prefill', async () => {
    const { input, resolve } = await setup();
    await resolve();
    type(input, 'X');
    expect(input.value).toBe(`${value}X`);
    expect(input.selectionStart).toBe(value.length + 1);
  });

  it('appends at the end before prefill', async () => {
    const { input, resolve } = await setup();
    type(input, 'X');
    await resolve();
    expect(input.value).toBe(`${value}X`);
    expect(input.selectionStart).toBe(value.length + 1);
  });

  it('replays native selection against the visible value before prefill', async () => {
    const { input, resolve } = await setup();
    input.select();
    key(input, 'Delete');
    type(input, 'X');
    await resolve();
    expect(input.value).toBe(cached ? 'X' : `${value}X`);
    expect(input.selectionStart).toBe(cached ? 1 : value.length + 1);
  });

  it('closes Enter immediately and commits once after replay with the selection moved down', async () => {
    const { host, view, input, resolve } = await setup();
    key(input, 'a', { ctrlKey: true });
    type(input, 'probe');
    key(input, 'Enter');
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    expect(host.sent.filter((command) => command.kind === 'editCell')).toEqual([]);
    await resolve();
    const edits = host.sent.filter((command) => command.kind === 'editCell');
    expect(edits).toHaveLength(1);
    expect(edits[0]).toMatchObject({ payload: { args: [0, 0, 0, 'probe'], selection: {
      range: { kind: 'rowCol', start: { row: 1, col: 0 }, end: { row: 1, col: 0 } },
    } } });
  });

  for (const [platform, modifier] of [['Linux', { ctrlKey: true }], ['MacIntel', { metaKey: true }]] as const) {
    it(`replaces a held seed with select all, Delete and typed keys on ${platform}`, async () => {
      const previous = Object.getOwnPropertyDescriptor(navigator, 'platform');
      Object.defineProperty(navigator, 'platform', { configurable: true, value: platform });
      restorers.push(() => { if (previous) Object.defineProperty(navigator, 'platform', previous); else Reflect.deleteProperty(navigator, 'platform'); });
      const { host, view, input, resolve } = await setup();
      const text = 'saved edit';
      key(input, 'a', modifier);
      key(input, 'Delete');
      for (const character of text) { key(input, character); type(input, character); }
      expect(host.sent.filter((command) => command.kind === 'editCell')).toEqual([]);
      await resolve();
      expect(input.value).toBe(text);
      expect(input.selectionStart).toBe(text.length);
      key(input, 'Enter');
      expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
      await waitFor(() => expect(host.sent.some((command) => command.kind === 'editCell')).toBe(true));
      const edits = host.sent.filter((command) => command.kind === 'editCell');
      expect(edits).toHaveLength(1);
      expect(edits[0]).toMatchObject({ payload: { args: [0, 0, 0, text] } });
    });
  }

  it('drops held keys and closes immediately on Escape', async () => {
    const { host, view, input, resolve } = await setup();
    type(input, 'X');
    key(input, 'Escape');
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    await resolve();
    expect(host.sent.filter((command) => command.kind === 'editCell')).toEqual([]);
    expect(host.cells.get('0:0:0')).toBe(value);
  });

  it('commits an untouched prefilled cell without an edit', async () => {
    const { host, input, resolve } = await setup();
    await resolve();
    key(input, 'Enter');
    expect(host.sent.filter((command) => command.kind === 'editCell')).toEqual([]);
    expect(host.cells.get('0:0:0')).toBe(value);
  });

  it('loads F2 on the next cell when the previous edit is refused', async () => {
    const host = peerlessHarness();
    host.cells.set('0:1:0', 'below');
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.keyDown(scroll, { key: 'a' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'abc' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    fireEvent.keyDown(scroll, { key: 'F2' });
    await act(async () => host.ack(host.sent.find((command) => command.kind === 'editCell')!, { status: 'refused' }));
    expect(view.queryByTestId('xlsx-error')).toBeNull();
    expect(view.getByTestId('xlsx-refused-draft').textContent).toBe('abc');
    await waitFor(() => expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('below'));
    expect(host.cells.get('0:0:0')).toBe('initial');
  });

  it('loads F2 while an earlier refused edit is still listed', async () => {
    const host = peerlessHarness();
    host.cells.set('0:2:0', 'third');
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.keyDown(scroll, { key: 'r' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    fireEvent.keyDown(scroll, { key: 'n' });
    await act(async () => host.ack(host.sent[0]!, { status: 'refused' }));
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await act(async () => host.ack(host.sent.find((command) => command.kind === 'editCell' && command.payload.args[3] === 'n')!));
    fireEvent.keyDown(scroll, { key: 'F2' });
    await act(async () => {});
    expect(view.queryByTestId('xlsx-error')).toBeNull();
    await waitFor(() => expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('third'));
    expect(view.getByTestId('xlsx-refused-draft').textContent).toBe('r');
  });

  describe('prefill reservation regressions', () => {
    const acknowledge = (host: ReturnType<typeof peerlessHarness>, history?: string[]) => {
      const submit = host.session.submit;
      host.session.submit = (command) => {
        const result = submit(command);
        queueMicrotask(() => {
          if (command.kind === 'editCell') history?.push(host.cells.get('0:0:0')!);
          if (command.kind === 'undo' && history) host.cells.set('0:0:0', history.pop()!);
          host.ack(command, { paint: true });
        });
        return result;
      };
      restorers.push(() => { host.session.submit = submit; });
    };

    it('includes held input in saved bytes before saving', async () => {
      const delivered = mock((_bytes: Uint8Array) => {});
      const { host, input, resolve } = await setup(delivered);
      acknowledge(host);
      const save = spyOn(host.session, 'save').mockImplementation(async () => new TextEncoder().encode(host.cells.get('0:0:0')));
      restorers.push(() => save.mockRestore());
      type(input, 'X');
      let saving!: Promise<unknown>;
      await act(async () => { saving = host.commands().execute('save', null); });
      await resolve();
      await act(async () => { await saving; });
      expect(host.sent.filter((command) => command.kind === 'editCell').map((command) => command.payload.args))
        .toEqual([[0, 0, 0, `${value}X`]]);
      expect(new TextDecoder().decode(delivered.mock.calls[0]![0])).toBe(`${value}X`);
    });

    it('closes and moves immediately so the next cell accepts typing in order', async () => {
      const { host, view, input, resolve } = await setup();
      acknowledge(host);
      type(input, 'X');
      key(input, 'Enter');
      expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'Y' });
      key(view.getByTestId('xlsx-cell-editor') as HTMLInputElement, 'Enter');
      await resolve();
      await waitFor(() => expect(host.cells.get('0:1:0')).toBe('Y'));
      expect(host.sent.filter((command) => command.kind === 'editCell').map((command) => command.payload.args))
        .toEqual([[0, 0, 0, `${value}X`], [0, 1, 0, 'Y']]);
    });

    it('keeps the earlier committed input when Escape follows Enter', async () => {
      const { host, view, input, resolve } = await setup();
      acknowledge(host);
      type(input, 'X');
      key(input, 'Enter');
      fireEvent.keyDown(view.queryByTestId('xlsx-cell-editor') ?? view.getByTestId('xlsx-scroll'), { key: 'Escape' });
      await resolve();
      await waitFor(() => expect(host.cells.get('0:0:0')).toBe(`${value}X`));
      expect(host.sent.filter((command) => command.kind === 'editCell')).toHaveLength(1);
    });

    it('retains a refused reserved edit without refocusing the previous cell', async () => {
      const { host, view, input, resolve } = await setup();
      type(input, 'X');
      key(input, 'Enter');
      await resolve();
      const edit = host.sent.find((command) => command.kind === 'editCell')!;
      await act(async () => host.ack(edit, { status: 'refused' }));
      expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A2');
      expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
      expect(document.activeElement).toBe(view.getByTestId('xlsx-scroll'));
      expect(view.getByTestId('xlsx-refused-draft').textContent).toBe(`${value}X`);
      expect(host.cells.get('0:0:0')).toBe(value);
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowDown' });
      await act(async () => {});
      expect(host.sent.filter((command) => command.kind === 'editCell')).toHaveLength(1);
      expect(host.sent[host.sent.length - 1]).toMatchObject({ kind: 'select' });
    });

    it('reads the reserved edit when F2 reopens the same cell', async () => {
      const { host, view, input, resolve } = await setup();
      acknowledge(host);
      type(input, 'X');
      const scroll = view.getByTestId('xlsx-scroll');
      fireEvent.mouseDown(scroll, { clientX: 150, clientY: 40 });
      expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
      fireEvent.mouseDown(scroll, { clientX: 20, clientY: 10 });
      fireEvent.keyDown(scroll, { key: 'F2' });
      await resolve();
      const reopened = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
      await waitFor(() => expect(reopened.value).toBe(`${value}X`));
      type(reopened, 'Y');
      key(reopened, 'Enter');
      await waitFor(() => expect(host.cells.get('0:0:0')).toBe(`${value}XY`));
      expect(host.sent.filter((command) => command.kind === 'editCell').map((command) => command.payload.args[3]))
        .toEqual([`${value}X`, `${value}XY`]);
    });

    it('orders click-away input before Undo and reverts that edit', async () => {
      const { host, view, input, resolve } = await setup();
      const history: string[] = ['older'];
      acknowledge(host, history);
      type(input, 'X');
      fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { clientX: 150, clientY: 40 });
      let undo!: Promise<unknown>;
      await act(async () => { undo = host.commands().execute('undo', null); });
      await resolve();
      await act(async () => { await undo; });
      expect(host.sent.filter((command) => command.kind === 'editCell' || command.kind === 'undo').map((command) => command.kind))
        .toEqual(['editCell', 'undo']);
      expect(host.cells.get('0:0:0')).toBe(value);
    });

    it('replays held text and an IME result against the current cell input', async () => {
      const { host, input, resolve } = await setup();
      acknowledge(host);
      type(input, 'X');
      fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: '変' } });
      end(input, '変');
      await act(async () => { await Promise.resolve(); });
      fireEvent.keyUp(input, { key: 'Enter' });
      key(input, 'Enter');
      await resolve();
      await waitFor(() => expect(host.cells.get('0:0:0')).toBe(`${value}X変`));
    });

    it('waits for composition to finish if the prefill arrives during IME', async () => {
      const { host, input, resolve } = await setup();
      acknowledge(host);
      type(input, 'X');
      fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: '変' } });
      await resolve();
      expect(input.value).toBe('変');
      end(input, '変');
      await act(async () => { await Promise.resolve(); });
      fireEvent.keyUp(input, { key: 'Enter' });
      key(input, 'Enter');
      await waitFor(() => expect(host.cells.get('0:0:0')).toBe(`${value}X変`));
    });

    it('keeps one IME insert when trailing native input follows the replay', async () => {
      const { host, input, resolve } = await setup();
      acknowledge(host);
      type(input, 'X');
      fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: `${preview}仮` } });
      end(input, '変');
      await act(async () => { await Promise.resolve(); });
      await resolve();
      expect(input.value).toBe(`${value}X変`);
      fireEvent.change(input, { target: { value: `${preview}変` } });
      expect(input.value).toBe(`${value}X変`);
      type(input, '変');
      expect(input.value).toBe(`${value}X変`);
      fireEvent.keyUp(input, { key: 'Enter' });
      key(input, 'Enter');
      await waitFor(() => expect(host.cells.get('0:0:0')).toBe(`${value}X変`));
      expect(host.sent.filter((command) => command.kind === 'editCell')).toHaveLength(1);
    });

    it('captures the final loading IME value on toolbar blur and restores keyboard focus', async () => {
      const { host, view, input, resolve } = await setup();
      acknowledge(host);
      type(input, 'X');
      fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: '仮' } });
      input.addEventListener('blur', () => {
        input.value = '変';
        end(input, '変');
      }, { once: true });
      let bold!: Promise<unknown>;
      await act(async () => { bold = host.commands().execute('bold', null); });
      await resolve();
      await act(async () => { await bold; });
      expect(host.sent.filter((command) => command.kind === 'editCell').map((command) => command.payload.args))
        .toEqual([[0, 0, 0, `${value}X変`]]);
      expect(document.activeElement).toBe(view.getByTestId('xlsx-scroll'));
    });

    it('holds successive IME results once and preserves them when a later composition cancels', async () => {
      const { host, input, resolve } = await setup();
      acknowledge(host);
      type(input, 'X');
      fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: `${preview}変` } });
      end(input, '変');
      await act(async () => {});
      input.setSelectionRange(preview.length + 1, preview.length + 1);
      fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: `${preview}変仮` } });
      end(input, '数');
      await act(async () => {});
      fireEvent.change(input, { target: { value: `${preview}変数` } });
      input.setSelectionRange(preview.length + 2, preview.length + 2);
      fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: `${preview}変数仮` } });
      end(input, '');
      await act(async () => {});
      fireEvent.change(input, { target: { value: `${preview}変数` } });
      await resolve();
      expect(input.value).toBe(`${value}X変数`);
      fireEvent.keyUp(input, { key: 'Enter' });
      key(input, 'Enter');
      await waitFor(() => expect(host.cells.get('0:0:0')).toBe(`${value}X変数`));
      expect(host.sent.filter((command) => command.kind === 'editCell')).toHaveLength(1);
    });

    it('reserves a loading IME commit when blur follows compositionend in the same task', async () => {
      const { host, view, input, resolve } = await setup();
      acknowledge(host);
      type(input, 'X');
      fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: '変' } });
      end(input, '変');
      fireEvent.blur(input);
      await resolve();
      await waitFor(() => expect(host.cells.get('0:0:0')).toBe(`${value}X変`));
      expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
      expect(host.sent.filter((command) => command.kind === 'editCell')).toHaveLength(1);
    });

    it('includes a detached held commit in Save', async () => {
      const { host, view, input, resolve } = await setup();
      acknowledge(host);
      const saved: string[] = [];
      const save = spyOn(host.session, 'save').mockImplementation(async () => {
        saved.push(host.cells.get('0:0:0')!);
        return new TextEncoder().encode(saved[saved.length - 1]);
      });
      restorers.push(() => save.mockRestore());
      type(input, 'X');
      fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { clientX: 150, clientY: 40 });
      let saving!: Promise<unknown>;
      await act(async () => { saving = host.commands().execute('save', null); });
      expect(save).not.toHaveBeenCalled();
      await resolve();
      await act(async () => { await saving; });
      expect(host.cells.get('0:0:0')).toBe(`${value}X`);
      expect(host.sent.filter((command) => command.kind === 'editCell')).toHaveLength(1);
      expect(host.session.save).toHaveBeenCalledTimes(1);
      expect(saved).toEqual([`${value}X`]);
    });

    it('drains held input before the read-only fence', async () => {
      const { host, view, input, resolve } = await setup();
      acknowledge(host);
      type(input, 'X');
      host.setReadOnly(view, true);
      await resolve();
      await waitFor(() => expect(host.sent.some((command) => command.kind === 'setReadOnly')).toBe(true));
      expect(host.sent.filter((command) => command.kind === 'editCell' || command.kind === 'setReadOnly').map((command) => command.kind))
        .toEqual(['editCell', 'setReadOnly']);
      expect(host.cells.get('0:0:0')).toBe(`${value}X`);
    });

    it('uses non-mac Ctrl Shift Left to replace only the last word', async () => {
      const platform = Object.getOwnPropertyDescriptor(navigator, 'platform');
      Object.defineProperty(navigator, 'platform', { configurable: true, value: 'Linux' });
      restorers.push(() => { if (platform) Object.defineProperty(navigator, 'platform', platform); else Reflect.deleteProperty(navigator, 'platform'); });
      const { input, resolve } = await setup();
      key(input, 'ArrowLeft', { ctrlKey: true, shiftKey: true });
      type(input, 'X');
      await resolve();
      expect(input.value).toBe('Q3 Sales X');
    });

    it('uses mac Option Shift Left for words and Cmd Left for the line', async () => {
      const platform = Object.getOwnPropertyDescriptor(navigator, 'platform');
      Object.defineProperty(navigator, 'platform', { configurable: true, value: 'MacIntel' });
      restorers.push(() => { if (platform) Object.defineProperty(navigator, 'platform', platform); else Reflect.deleteProperty(navigator, 'platform'); });
      const { input, resolve } = await setup();
      key(input, 'ArrowLeft', { altKey: true, shiftKey: true });
      type(input, 'X');
      key(input, 'ArrowLeft', { metaKey: true });
      type(input, 'Y');
      await resolve();
      expect(input.value).toBe('YQ3 Sales X');
    });

    for (const scenario of [
      { name: 'Home and Delete', platform: 'Linux', keys: [{ key: 'Home' }, { key: 'Delete' }], expected: value.slice(1) },
      { name: 'Ctrl Backspace', platform: 'Linux', keys: [{ key: 'Backspace', ctrlKey: true }], expected: 'Q3 Sales ' },
      { name: 'Option Delete', platform: 'MacIntel', keys: [{ key: 'Home' }, { key: 'Delete', altKey: true }], expected: ' Sales Dashboard' },
      { name: 'Cmd Backspace', platform: 'MacIntel', keys: [{ key: 'Backspace', metaKey: true }], expected: '' },
    ]) it(`replays ${scenario.name} in the grid while cell geometry is pending`, async () => {
      const platform = Object.getOwnPropertyDescriptor(navigator, 'platform');
      Object.defineProperty(navigator, 'platform', { configurable: true, value: scenario.platform });
      restorers.push(() => { if (platform) Object.defineProperty(navigator, 'platform', platform); else Reflect.deleteProperty(navigator, 'platform'); });
      const host = peerlessHarness();
      const view = await host.mount();
      const geometry = deferred<xlsx.WorkbookCellGeometry>();
      host.cellGeometry.mockReturnValueOnce(geometry.promise);
      host.selectionData.mockResolvedValueOnce({ edge: { row: 250, col: 0 } });
      const scroll = view.getByTestId('xlsx-scroll');
      await act(async () => { fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true }); });
      act(() => host.ack(host.sent[0]!, { paint: true }));
      acknowledge(host);
      const prefill = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
      const inputs = host.cellInputs.getMockImplementation()!;
      host.cellInputs.mockImplementation((sheet, a1) => a1 === 'A251' ? prefill.promise : inputs(sheet, a1));
      fireEvent.keyDown(scroll, { key: 'F2' });
      expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
      for (const event of scenario.keys) fireEvent.keyDown(scroll, event);
      fireEvent.keyDown(scroll, { key: 'Enter' });
      await act(async () => prefill.resolve({ sheet: 0, version: 'v1', cells: [[{ a1: 'A251', input: value, isFormula: false }]] }));
      await waitFor(() => expect(host.sent.some((command) => command.kind === 'editCell')).toBe(true));
      expect(host.sent.filter((command) => command.kind === 'editCell')[0]).toMatchObject({ payload: { args: [0, 250, 0, scenario.expected] } });
      await act(async () => geometry.resolve({ sheet: 0, version: 'v1', rect: { x: 0, y: 6000, w: 96, h: 24 }, scrollPosition: { x: 0, y: 6000 } }));
    });
  });

  it('commits held input on blur without taking focus back from the host', async () => {
    const { host, input, resolve } = await setup();
    const button = document.createElement('button');
    document.body.appendChild(button);
    try {
      type(input, 'X');
      act(() => button.focus());
      expect(host.sent.filter((command) => command.kind === 'editCell')).toEqual([]);
      await resolve();
      const edits = host.sent.filter((command) => command.kind === 'editCell');
      expect(edits).toHaveLength(1);
      expect(edits[0]).toMatchObject({ payload: { args: [0, 0, 0, `${value}X`] } });
      expect(document.activeElement).toBe(button);
    } finally { button.remove(); }
  });

  it('appends an Option-composed character in the grid while cell geometry is pending', async () => {
    const platform = Object.getOwnPropertyDescriptor(navigator, 'platform');
    Object.defineProperty(navigator, 'platform', { configurable: true, value: 'MacIntel' });
    restorers.push(() => { if (platform) Object.defineProperty(navigator, 'platform', platform); else Reflect.deleteProperty(navigator, 'platform'); });
    const host = peerlessHarness();
    const view = await host.mount();
    const geometry = deferred<xlsx.WorkbookCellGeometry>();
    host.cellGeometry.mockReturnValueOnce(geometry.promise);
    host.selectionData.mockResolvedValueOnce({ edge: { row: 250, col: 0 } });
    const scroll = view.getByTestId('xlsx-scroll');
    await act(async () => { fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true }); });
    act(() => host.ack(host.sent[0]!, { paint: true }));
    fireEvent.keyDown(scroll, { key: 'a' });
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    fireEvent.keyDown(scroll, { key: '@', altKey: true });
    fireEvent.keyDown(scroll, { key: 'Enter' });
    expect(host.sent.filter((command) => command.kind === 'editCell')[0]).toMatchObject({ payload: { args: [0, 250, 0, 'a@'] } });
    await act(async () => geometry.resolve({ sheet: 0, version: 'v1', rect: { x: 0, y: 6000, w: 96, h: 24 }, scrollPosition: { x: 0, y: 6000 } }));
  });
});

for (const decision of [true, false, undefined]) {
  it(`peerless save uses the host decision (${String(decision)})`, async () => {
    const host = peerlessHarness();
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const requested = mock(async () => decision);
    const delivered = mock((_bytes: Uint8Array) => {});
    render(host.asyncEditor({ onReady: ready, onSaveRequest: requested, onSave: delivered }));
    await opened();
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    let result!: unknown;
    await act(async () => { result = await ready.mock.calls[0]![0].commands.execute('save', null); });
    expect(result).toEqual({ ok: true, status: decision === true ? 'executed' : 'noop' });
    expect(requested).toHaveBeenCalledTimes(1);
    expect(host.session.save).toHaveBeenCalledTimes(decision === true ? 1 : 0);
    expect(delivered).toHaveBeenCalledTimes(decision === true ? 1 : 0);
  });
}

describe('peerless chrome locale', () => {
  beforeEach(() => {
    const originalRect = HTMLElement.prototype.getBoundingClientRect;
    HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
      if (this.getAttribute('role') !== 'toolbar') return originalRect.call(this);
      return { width: 4000, height: 28, top: 0, left: 0, right: 4000, bottom: 28, x: 0, y: 0, toJSON() {} } as DOMRect;
    };
    restorers.push(() => { HTMLElement.prototype.getBoundingClientRect = originalRect; });
  });

  const german = { _lang: 'de', toolbar: { bold: 'Fett', formulaBarLabel: 'Formelzeile',
    formulaPlaceholder: 'Wert oder Formel', nameBoxPlaceholder: 'Zelle' },
    editor: { appLabel: 'Tabelleneditor', sheetTabsLabel: 'Tabellen', openError: 'Öffnen fehlgeschlagen',
      stoppedError: 'Der Editor hat ein Problem' },
    a11y: { gridLabel: 'Tabellenraster' },
    commands: { reasons: { nothingToUndo: 'Nichts rückgängig zu machen.', inputFailed: 'Eingabe fehlgeschlagen.',
      readOnly: 'Diese Arbeitsmappe ist schreibgeschützt.' } },
  };
  const french = { _lang: 'fr', toolbar: { bold: 'Gras', formulaBarLabel: 'Barre de formule',
    formulaPlaceholder: 'Valeur ou formule', nameBoxPlaceholder: 'Cellule' },
    editor: { appLabel: 'Éditeur de tableau', sheetTabsLabel: 'Feuilles', openError: 'Échec de chargement',
      stoppedError: "L'éditeur a rencontré un problème" },
    a11y: { gridLabel: 'Grille de tableau' },
    commands: { reasons: { nothingToUndo: 'Rien à annuler.', inputFailed: 'Échec de saisie.',
      readOnly: 'Ce classeur est en lecture seule.' } },
  };

  it('sends host Undo behind unacknowledged text with empty history', async () => {
    const host = peerlessHarness();
    host.push({ throughSeq: 0, version: 'v1', statuses: [], editState: { canUndo: false, canRedo: false } });
    const ready = deferred<XlsxAsyncEditorApi>();
    let view!: ReturnType<typeof render>;
    await act(async () => {
      view = render(<XlsxEditor file={file} onReady={(api) => { ready.resolve(api); }} />);
    });
    await tick();
    const api = await ready.promise;
    expect(api.commands.getState('undo').enabled).toBe(false);
    let submitted = deferred<PeerlessWorkerCommand>();
    const submit = host.session.submit.bind(host.session);
    host.session.submit = (command) => { const result = submit(command); submitted.resolve(command); return result; };
    act(() => fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x' }));
    let edit!: PeerlessWorkerCommand;
    await act(async () => {
      fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
      edit = await submitted.promise;
    });
    expect(edit).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'x'] } });
    expect(api.commands.getState('undo').enabled).toBe(true);
    submitted = deferred<PeerlessWorkerCommand>();
    let undo!: Promise<unknown>;
    let command!: PeerlessWorkerCommand;
    let settled = false;
    await act(async () => {
      undo = api.commands.execute('undo', null).then((result) => { settled = true; return result; });
      command = await submitted.promise;
    });
    expect(host.session.editState.canUndo).toBe(false);
    expect(host.session.throughSeq).toBe(0);
    expect(host.sent.map((entry) => entry.kind)).toEqual(['editCell', 'undo']);
    expect(command.seq).toBeGreaterThan(edit.seq);
    expect(settled).toBe(false);
    expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
    await act(async () => { host.ack(edit); host.ack(command); expect(await undo).toMatchObject({ ok: true }); });
    expect(settled).toBe(true);
    expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
  });

  for (const entry of ['public', 'standalone'] as const) {
    it(`localizes a ${entry} worker read-only refusal after a locale change`, async () => {
      const host = peerlessHarness();
      const ready = deferred<XlsxAsyncEditorApi>();
      const onReady = (api: XlsxAsyncEditorApi) => { ready.resolve(api); };
      const editor = (i18n: typeof german) => entry === 'public'
        ? <XlsxEditor file={file} i18n={i18n} onReady={onReady} />
        : <XlsxWorkerEditor backend="peerless" file={file} i18n={i18n} onReady={onReady} />;
      let view!: ReturnType<typeof render>;
      await act(async () => { view = render(editor(german)); });
      await tick();
      await ready.promise;
      const submitted = deferred<PeerlessWorkerCommand>();
      const submit = host.session.submit.bind(host.session);
      host.session.submit = (command) => { const result = submit(command); submitted.resolve(command); return result; };
      act(() => fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x' }));
      let command!: PeerlessWorkerCommand;
      await act(async () => {
        fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
        command = await submitted.promise;
      });
      expect(command).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'x'] } });
      await act(async () => host.push({ throughSeq: command.seq, version: 'v1',
        statuses: [{ seq: command.seq, status: 'refused', reason: 'read-only' }] }));
      const alert = view.getByTestId('xlsx-input-refusal');
      expect(alert.textContent).toBe(german.commands.reasons.readOnly);
      await act(async () => view.rerender(editor(french)));
      expect(view.getByTestId('xlsx-input-refusal')).toBe(alert);
      expect(alert.textContent).toBe(french.commands.reasons.readOnly);
      expect(host.open).toHaveBeenCalledTimes(1);
    });

    it(`updates ${entry} labels and refusals without reopening`, async () => {
      const host = peerlessHarness();
      const ready = mock((_api: XlsxAsyncEditorApi) => {});
      const editor = (i18n: typeof german) => entry === 'public'
        ? <XlsxEditor file={file} i18n={i18n} onReady={ready} />
        : <XlsxWorkerEditor backend="peerless" file={file} i18n={i18n} onReady={ready} />;
      let view!: ReturnType<typeof render>;
      await act(async () => { view = render(editor(german)); });
      await tick();
      expect(ready).toHaveBeenCalledTimes(1);
      const api = ready.mock.calls[0]![0];
      const scroll = view.getByTestId('xlsx-scroll');
      const formula = view.getByTestId('xlsx-formula-input');
      view.getByRole('button', { name: 'Fett' });
      view.getByRole('application', { name: 'Tabelleneditor' });
      view.getByRole('tablist', { name: 'Tabellen' });
      view.getByRole('group', { name: 'Formelzeile' });
      view.getByRole('grid', { name: /Tabellenraster/ });
      expect(formula.getAttribute('aria-label')).toBe('Wert oder Formel');
      expect(view.getByTestId('xlsx-name-box').getAttribute('aria-label')).toBe('Zelle');
      await act(async () => {
        for (const command of host.sent) host.ack(command);
        host.push({ throughSeq: host.session.throughSeq, version: 'v1', statuses: [],
          editState: { canUndo: false, canRedo: false } });
      });
      await act(async () => {
        expect(await api.commands.execute('undo', null)).toMatchObject({ ok: false,
          failure: { code: 'nothing-to-undo', message: german.commands.reasons.nothingToUndo } });
      });
      expect(view.getByTestId('xlsx-input-refusal').textContent).toBe(german.commands.reasons.nothingToUndo);
      await act(async () => view.rerender(editor(french)));
      view.getByRole('button', { name: 'Gras' });
      view.getByRole('application', { name: 'Éditeur de tableau' });
      view.getByRole('tablist', { name: 'Feuilles' });
      view.getByRole('group', { name: 'Barre de formule' });
      view.getByRole('grid', { name: /Grille de tableau/ });
      expect(view.getByTestId('xlsx-formula-input')).toBe(formula);
      expect(formula.getAttribute('aria-label')).toBe('Valeur ou formule');
      expect(view.getByTestId('xlsx-name-box').getAttribute('aria-label')).toBe('Cellule');
      expect(view.getByTestId('xlsx-input-refusal').textContent).toBe(french.commands.reasons.nothingToUndo);

      const submitted = deferred<PeerlessWorkerCommand>();
      const submit = host.session.submit.bind(host.session);
      host.session.submit = (command) => { const reply = submit(command); submitted.resolve(command); return reply; };
      let command!: PeerlessWorkerCommand;
      act(() => fireEvent.keyDown(scroll, { key: 'x' }));
      await act(async () => {
        fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
        command = await submitted.promise;
      });
      await act(async () => host.ack(command, { status: 'refused' }));
      expect(view.getByTestId('xlsx-input-refusal').textContent).toContain(french.commands.reasons.inputFailed);
      const failure = new SessionFailure('crash', 'Worker stopped');
      await act(async () => host.fail(failure));
      expect(view.getByTestId('xlsx-error').textContent).toContain(french.editor.stoppedError);
      expect(view.getByTestId('xlsx-error').getAttribute('data-failure')).toBe('stopped');
      expect(view.getByTestId('xlsx-error').querySelector('details pre')?.textContent).toContain('Worker stopped');
      expect(api.commands.getState('bold').enabled).toBe(false);
      await act(async () => view.rerender(editor(german)));
      expect(view.getByTestId('xlsx-error').textContent).toContain(german.editor.stoppedError);
      expect(view.getByTestId('xlsx-input-refusal').textContent).toContain(german.commands.reasons.inputFailed);
      expect(view.getByTestId('xlsx-scroll')).toBe(scroll);
      expect(host.open).toHaveBeenCalledTimes(1);
      expect(ready).toHaveBeenCalledTimes(1);
      expect(ready.mock.calls[0]![0]).toBe(api);
      expect(host.session.dispose).not.toHaveBeenCalled();
      expect(host.hydrate).not.toHaveBeenCalled();
      expect(host.attach).not.toHaveBeenCalled();
    });
  }

  it('disables external chrome during replacement and binds the next async API', async () => {
    const host = peerlessHarness();
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    let view!: ReturnType<typeof render>;
    await act(async () => { view = render(<XlsxEditor file={file} showToolbar={false} onReady={ready} />); });
    await tick();
    expect(ready).toHaveBeenCalledTimes(1);
    const first = ready.mock.calls[0]![0];
    const chrome = (commands: XlsxCommandStore) => <XlsxCommandProvider commands={commands}>
      <EditorToolbar mode="commands"><ToolbarCommandButton id="bold" /></EditorToolbar>
    </XlsxCommandProvider>;
    const toolbar = render(chrome(first.commands));
    const next = deferred<WorkbookSession>();
    host.open.mockImplementationOnce(() => next.promise);
    const replacement = file.slice();
    await act(async () => view.rerender(<XlsxEditor file={replacement} showToolbar={false} onReady={ready} />));
    expect(toolbar.getByRole('button', { name: 'Bold' }).getAttribute('aria-disabled')).toBe('true');
    await expect(first.commands.execute('bold', null)).rejects.toMatchObject({ code: 'document-replaced' });
    await act(async () => next.resolve(host.session));
    await tick();
    expect(ready).toHaveBeenCalledTimes(2);
    const second = ready.mock.calls[1]![0];
    toolbar.rerender(chrome(second.commands));
    expect(toolbar.getByRole('button', { name: 'Bold' }).getAttribute('aria-disabled')).toBeNull();
    const submitted = deferred<PeerlessWorkerCommand>();
    const submit = host.session.submit.bind(host.session);
    host.session.submit = (command) => { const reply = submit(command); submitted.resolve(command); return reply; };
    let result!: Promise<unknown>;
    let command!: PeerlessWorkerCommand;
    await act(async () => { result = second.commands.execute('bold', null); command = await submitted.promise; });
    await act(async () => { host.ack(command); await result; });
    expect(host.session.presentedThroughSeq).toBe(0);
    await act(async () => host.fail(new SessionFailure('crash', 'Worker stopped')));
    expect(toolbar.getByRole('button', { name: 'Bold' }).getAttribute('aria-disabled')).toBe('true');
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
  });
});
async function mountPeerlessClipboard(host: ReturnType<typeof peerlessHarness>) {
  const ready = deferred<XlsxAsyncEditorApi>();
  let view!: ReturnType<typeof render>;
  await act(async () => { view = render(host.asyncEditor({ onReady: (api) => ready.resolve(api) })); });
  await tick();
  await ready.promise;
  return view;
}

function clipboardReads(host: ReturnType<typeof peerlessHarness>) {
  return host.cellInputs.mock.calls.filter(([, a1]) => a1.includes(':'));
}

function collaborativePeerlessHarness() {
  const host = peerlessHarness();
  let applyRemote: ((update: Uint8Array) => Promise<unknown>) | undefined;
  const bindCollaboration = mock((apply: (update: Uint8Array) => Promise<unknown>,
    _applyChunks?: (updates: Uint8Array[]) => Promise<unknown>,
    _applyPipe?: Parameters<NonNullable<WorkbookSession['bindCollaboration']>>[2]) => {
    applyRemote = apply;
    return () => { if (applyRemote === apply) applyRemote = undefined; };
  });
  const collaborationStateVector = async () => Uint8Array.of(1);
  const collaborationStateAsUpdate = async () => Uint8Array.of(2);
  const replica = createAsyncCollaborationReplica({ clientId: 101,
    encodeStateVector: collaborationStateVector, encodeStateAsUpdate: collaborationStateAsUpdate,
    applyUpdate: async (update) => {
      if (!applyRemote) throw new Error('Collaboration not bound');
      return applyRemote(update.slice());
    },
    onUpdate: () => () => {},
  });
  Object.defineProperty(host.session, 'replica', { value: replica });
  host.session.bindCollaboration = bindCollaboration;
  Object.assign(host.session.call, { collaborationStateVector, collaborationStateAsUpdate });
  return { ...host, replica, bindCollaboration };
}

function replacePrint(value: (() => void) | undefined) {
  const original = Object.getOwnPropertyDescriptor(window, 'print');
  Object.defineProperty(window, 'print', { configurable: true, value });
  restorers.push(() => {
    if (original) Object.defineProperty(window, 'print', original);
    else Reflect.deleteProperty(window, 'print');
  });
}

describe('XlsxEditor async API', () => {
  for (const source of ['cell', 'formula'] as const) {
    it(`restores ${source} keyboard focus after submitting composed text once and running a toolbar command`, async () => {
      const host = peerlessHarness(true);
      const view = await host.mount();
      if (source === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
      const input = view.getByTestId(source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input') as HTMLInputElement;
      if (source === 'cell') await waitFor(() => expect(input.value).toBe('initial'));
      act(() => input.focus());
      fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: 'unfinished' } });
      const blur = spyOn(input, 'blur');
      input.addEventListener('blur', () => {
        input.value = 'final composed text';
        fireEvent.compositionEnd(input);
      }, { once: true });
      restorers.push(() => blur.mockRestore());
      const bold = view.container.querySelector<HTMLButtonElement>('button[aria-label="Bold"]')!;
      await act(async () => fireEvent.click(bold));
      expect(blur).toHaveBeenCalledTimes(1);
      await waitFor(() => expect(host.sent).toHaveLength(2));
      expect(host.sent[0]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'final composed text'] } });
      expect(host.sent[1]).toMatchObject({ kind: 'applyOps', payload: { toolbar: { id: 'bold' } } });
      await act(async () => {
        fireEvent.change(input, { target: { value: 'final composed text' } });
        fireEvent.compositionEnd(input);
        for (const command of host.sent) host.ack(command, { paint: true });
      });
      await advance();
      expect(host.sent.filter(({ kind }) => kind === 'editCell')).toHaveLength(1);
      expect(host.cells.get('0:0:0')).toBe('final composed text');
      expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).value).toBe('final composed text');
      expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
      const sink = view.getByTestId(source === 'cell' ? 'xlsx-scroll' : 'xlsx-formula-input');
      expect(document.activeElement).toBe(sink);
      if (source === 'cell') {
        fireEvent.keyDown(document.activeElement!, { key: 'F2' });
        await waitFor(() => expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('final composed text'));
      } else {
        fireEvent.change(document.activeElement!, { target: { value: 'final composed textx' } });
        expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).value).toBe('final composed textx');
        fireEvent.keyDown(document.activeElement!, { key: 'Enter' });
        await waitFor(() => expect(host.sent.filter(({ kind }) => kind === 'editCell')).toHaveLength(2));
        const edits = host.sent.filter(({ kind }) => kind === 'editCell');
        expect(edits[edits.length - 1]).toMatchObject({ payload: { args: [0, 0, 0, 'final composed textx'] } });
      }
    });

    it(`preserves host dialog focus after saving ${source} IME composition`, async () => {
      const host = peerlessHarness(false, false, true);
      const dialog = render(<input aria-label="Save dialog input" />);
      const hostInput = dialog.getByRole('textbox', { name: 'Save dialog input' });
      const saved = mock(() => hostInput.focus());
      const view = await host.mount(saved);
      if (source === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
      const input = view.getByTestId(source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input') as HTMLInputElement;
      if (source === 'cell') await waitFor(() => expect(input.value).toBe('initial'));
      act(() => input.focus());
      fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: 'unfinished' } });
      input.addEventListener('blur', () => {
        input.value = 'final composed text';
        fireEvent.compositionEnd(input);
      }, { once: true });
      await act(async () => fireEvent.click(view.getByTestId('xlsx-save')));
      await waitFor(() => expect(host.sent).toHaveLength(1));
      expect(host.sent[0]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'final composed text'] } });
      await act(async () => {
        fireEvent.change(input, { target: { value: 'final composed text' } });
        fireEvent.compositionEnd(input);
        host.ack(host.sent[0], { paint: true });
      });
      await advance();
      await waitFor(() => expect(saved).toHaveBeenCalledTimes(1));
      expect(host.session.save).toHaveBeenCalledTimes(1);
      expect(host.sent).toHaveLength(1);
      expect(host.cells.get('0:0:0')).toBe('final composed text');
      expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
      expect(document.activeElement === hostInput).toBe(true);
    });
  }

  it('names legacySyncApi when default onReplica encoding is called synchronously', async () => {
    const host = collaborativePeerlessHarness();
    const replicas = mock((_replica: CollaborationReplica | null) => {});
    render(host.asyncEditor({ collaboration: { onReplica: replicas } }));
    await opened();
    await waitFor(() => expect(replicas).toHaveBeenCalledWith(host.replica));
    const replica = replicas.mock.calls.find(([value]) => value !== null)![0]!;
    expect(() => replica.encodeStateVector()).toThrow('legacySyncApi');
    expect(() => replica.encodeStateAsUpdate()).toThrow('legacySyncApi');
    await expect(host.replica.async.encodeStateVector()).resolves.toEqual(Uint8Array.of(1));
  });

  for (const route of ['renderPng', 'renderRangePng', 'exportStructured', 'exportMarkdown', 'printDisplayList', 'exportPng'] as const) {
    for (const status of ['applied', 'refused'] as const) {
      it(`exports acknowledged composition and preserves input refusals (${route}, ${status})`, async () => {
        const host = peerlessHarness(true);
        const values: string[] = [];
        const read = () => {
          const value = host.cells.get('0:0:0')!;
          values.push(value);
          return value;
        };
        const sheet = { sheetId: 'sheet:0', index: 0, name: 'First' };
        const renderPng = mock<WorkbookSession['call']['renderPng']>(async () => new TextEncoder().encode(read()));
        const renderRangePng = mock<WorkbookSession['call']['renderRangePng']>(async () => new TextEncoder().encode(read()));
        const exportStructured = mock<WorkbookSession['call']['exportStructured']>(async () => {
          const value = read();
          return { ok: true, version: 'v1', content: {
            schemaVersion: 1, anchorScope: 'session', dateSystem: '1900', truncated: false,
            calculation: { policy: 'asStored', freshness: 'unverified' },
            included: { hiddenSheets: false, hiddenRows: false, hiddenColumns: false, definedNames: true, hiddenNames: false },
            definedNames: [], diagnostics: [], sheets: [{
              id: 'sheet:0', anchor: { kind: 'sheet', sheet }, kind: 'worksheet', visibility: 'visible',
              usedRange: 'A1', selectedRange: 'A1', truncated: false, hiddenRows: [], hiddenColumns: [],
              merges: [], tables: [], hyperlinks: [], objects: [], source: null, cells: [{
                id: 'A1', anchor: { kind: 'cell', sheet, a1: 'A1' }, value: { kind: 'text', value },
                formula: null, displayText: value, numberFormat: 'General', formulaResult: null, merge: null,
              }],
            }],
          } };
        });
        const exportMarkdown = mock<WorkbookSession['call']['exportMarkdown']>(async () => ({
          ok: true, version: 'v1', content: { markdown: read(), anchors: [], diagnostics: [], truncated: false },
        }));
        const printDisplayList = mock<WorkbookSession['call']['printDisplayList']>(async () => ({
          width: 120, height: 48,
          commands: [{ op: 'text', text: read(), x: 8, y: 18, fontSize: 11, color: '#000000' }],
        }));
        Object.assign(host.session, { pngExport: true });
        Object.assign(host.session.call, { renderPng, renderRangePng, exportStructured, exportMarkdown, printDisplayList });
        for (const name of ['createObjectURL', 'revokeObjectURL'] as const) {
          const descriptor = Object.getOwnPropertyDescriptor(URL, name);
          Object.defineProperty(URL, name, { configurable: true, value: () => 'blob:png' });
          restorers.push(() => {
            if (descriptor) Object.defineProperty(URL, name, descriptor);
            else Reflect.deleteProperty(URL, name);
          });
        }
        const click = spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
        restorers.push(() => click.mockRestore());
        const ready = mock((_api: XlsxAsyncEditorApi) => {});
        const view = render(host.asyncEditor({ onReady: ready }));
        await opened();
        await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
        const api = ready.mock.calls[0]![0];
        fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
        const input = view.getByTestId('xlsx-cell-editor');
        await waitFor(() => expect((input as HTMLInputElement).value).toBe('initial'));
        fireEvent.compositionStart(input);
        fireEvent.change(input, { target: { value: 'composed' } });
        const invoke = (): Promise<unknown> => {
          switch (route) {
            case 'renderPng': return api.renderPng({ x: 0, y: 0, width: 120, height: 48 });
            case 'renderRangePng': return api.renderRangePng({ range: 'A1' });
            case 'exportStructured': return api.exportStructured();
            case 'exportMarkdown': return api.exportMarkdown();
            case 'printDisplayList': return api.printDisplayList(0, 'A1', {
              dpi: 144, maxDigitWidth: 9, defaultRowHeightPt: 18,
              fontSizePt: 12, fontFamily: 'Arial', fontAscent: 12, fontDescent: 4,
            }, false);
            case 'exportPng': return api.commands.execute('exportPng', null);
          }
        };
        let exporting!: Promise<unknown>;
        act(() => { exporting = invoke(); void exporting.catch(() => {}); });
        expect(values).toEqual([]);
        await act(async () => fireEvent.compositionEnd(input));
        await waitFor(() => expect(host.sent.some((command) => command.kind === 'editCell')).toBe(true));
        expect(values).toEqual([]);
        await act(async () => {
          for (const command of host.sent) host.ack(command, { status });
          await exporting.catch(() => {});
        });
        if (status === 'refused') {
          if (route === 'exportPng') await expect(exporting).resolves.toMatchObject({ ok: false, failure: { code: 'input-failed' } });
          else await expect(exporting).rejects.toMatchObject({ code: 'input-failed' });
          expect(values).toEqual([]);
        } else {
          expect(values).toEqual(['composed']);
          const result = await exporting;
          if (route === 'renderPng' || route === 'renderRangePng') expect(new TextDecoder().decode(result as Uint8Array)).toBe('composed');
          else if (route === 'exportStructured') expect(result).toMatchObject({ content: { sheets: [{ cells: [{ displayText: 'composed' }] }] } });
          else if (route === 'exportMarkdown') expect(result).toMatchObject({ content: { markdown: 'composed' } });
          else if (route === 'printDisplayList') expect(result).toMatchObject({ commands: [{ text: 'composed' }] });
          else expect(result).toMatchObject({ ok: true });
        }
        expect(host.hydrate).not.toHaveBeenCalled();
        expect(host.attach).not.toHaveBeenCalled();
      });
    }
  }

  it('downloads viewport PNG from the worker and resolves the export API', async () => {
    const host = peerlessHarness(true);
    const png = Uint8Array.of(137, 80, 78, 71, 13, 10, 26, 10);
    const renderPng = mock(async (_viewport: xlsx.Viewport, _options?: { sheet?: number }) => png);
    const renderRangePng = mock(async (_range?: { range?: string; scale?: number }, _options?: { sheet?: number }) => png);
    const structured: xlsx.XlsxExportResult<xlsx.XlsxStructuredContent> = { ok: true, version: 'v1', content: {
      schemaVersion: 1, anchorScope: 'session', dateSystem: '1900', truncated: false,
      calculation: { policy: 'asStored', freshness: 'unverified' },
      included: { hiddenSheets: false, hiddenRows: false, hiddenColumns: false, definedNames: true, hiddenNames: false },
      sheets: [], definedNames: [], diagnostics: [],
    } };
    const markdown: xlsx.XlsxExportResult<xlsx.XlsxMarkdownContent> = { ok: true, version: 'v1', content: {
      markdown: 'Data', anchors: [], diagnostics: [], truncated: false,
    } };
    const exportStructured = mock(async (_options?: xlsx.XlsxExportOptions) => structured);
    const exportMarkdown = mock(async (_options?: xlsx.XlsxExportOptions, _markdown?: xlsx.XlsxMarkdownOptions) => markdown);
    const list: DisplayList = { width: 120, height: 48, commands: [] };
    const printDisplayList = mock(async (_sheet: number, _range: string, _metrics: xlsx.PrintMetrics, _gridlines: boolean) => list);
    Object.assign(host.session, { pngExport: true });
    Object.assign(host.session.call, { renderPng, renderRangePng, exportStructured, exportMarkdown, printDisplayList });
    const createUrl = mock((_blob: Blob) => 'blob:png');
    const revokeUrl = mock((_url: string) => {});
    for (const [name, value] of [['createObjectURL', createUrl], ['revokeObjectURL', revokeUrl]] as const) {
      const descriptor = Object.getOwnPropertyDescriptor(URL, name);
      Object.defineProperty(URL, name, { configurable: true, value });
      restorers.push(() => {
        if (descriptor) Object.defineProperty(URL, name, descriptor);
        else Reflect.deleteProperty(URL, name);
      });
    }
    const downloads: { name: string; href: string }[] = [];
    const click = spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (this: HTMLAnchorElement) {
      downloads.push({ name: this.download, href: this.href });
    });
    restorers.push(() => click.mockRestore());
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const view = render(host.asyncEditor({ onReady: ready, fileName: 'Report.XLSX' }));
    await opened();
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    const api = ready.mock.calls[0]![0];
    expect(api.commands.getState('exportPng').enabled).toBe(true);
    await act(async () => fireEvent.click(view.getByTestId('xlsx-export-png')));
    await waitFor(() => expect(downloads).toEqual([{ name: 'Report.png', href: 'blob:png' }]));
    expect(renderPng).toHaveBeenCalledWith({ x: 0, y: 0, width: 800, height: 600 }, { sheet: 0 });
    const blob = createUrl.mock.calls[0]![0];
    expect(blob.type).toBe('image/png');
    expect(new Uint8Array(await blob.arrayBuffer())).toEqual(png);
    expect(revokeUrl).toHaveBeenCalledWith('blob:png');
    const viewport = { x: 10, y: 20, width: 120, height: 48 };
    const options = { scope: [{ sheet: 1, range: 'A1:B2' }], maxCells: 2 };
    const metrics = { dpi: 144, maxDigitWidth: 9, defaultRowHeightPt: 18,
      fontSizePt: 12, fontFamily: 'Arial', fontAscent: 12, fontDescent: 4 };
    await expect(api.renderPng(viewport)).resolves.toEqual(png);
    await expect(api.renderRangePng({ range: 'A1:B2', scale: 2 })).resolves.toEqual(png);
    await expect(api.exportStructured(options)).resolves.toEqual(structured);
    await expect(api.exportMarkdown(options, { maxRows: 1 })).resolves.toEqual(markdown);
    await expect(api.printDisplayList(1, 'A1:B2', metrics, false)).resolves.toEqual(list);
    expect(renderPng).toHaveBeenLastCalledWith(viewport, { sheet: 0 });
    expect(renderRangePng).toHaveBeenCalledWith({ range: 'A1:B2', scale: 2 }, { sheet: 0 });
    expect(exportStructured).toHaveBeenCalledWith(options);
    expect(exportMarkdown).toHaveBeenCalledWith(options, { maxRows: 1 });
    expect(printDisplayList).toHaveBeenCalledWith(1, 'A1:B2', metrics, false);
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
    view.unmount();
    await expect(api.exportStructured()).rejects.toMatchObject({ code: 'document-replaced' });
  });

  it('uses worker PNG availability and captures export scope before composition finishes', async () => {
    const host = peerlessHarness(true);
    const exportStructured = mock(async (_options?: xlsx.XlsxExportOptions) => ({ ok: true, version: 'v1' }));
    Object.assign(host.session, { pngExport: false });
    Object.assign(host.session.call, { exportStructured });
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const view = render(host.asyncEditor({ onReady: ready }));
    await opened();
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    const api = ready.mock.calls[0]![0];
    expect(api.commands.getState('exportPng')).toMatchObject({ enabled: false,
      disabledReason: { code: 'png-unavailable' } });
    const scroll = view.getByTestId('xlsx-scroll');
    await act(async () => fireEvent.keyDown(scroll, { key: 'x' }));
    const input = view.getByTestId('xlsx-cell-editor');
    fireEvent.compositionStart(input);
    fireEvent.change(input, { target: { value: '17' } });
    const options = { scope: [{ sheet: 0, range: 'A1' }] };
    let exporting!: Promise<unknown>;
    act(() => { exporting = api.exportStructured(options); });
    options.scope[0].sheet = 1;
    expect(exportStructured).not.toHaveBeenCalled();
    await act(async () => fireEvent.compositionEnd(input));
    await waitFor(() => expect(host.sent.some((command) => command.kind === 'editCell')).toBe(true));
    await act(async () => {
      for (const command of host.sent) host.ack(command);
      await exporting;
    });
    expect(exportStructured).toHaveBeenCalledWith({ scope: [{ sheet: 0, range: 'A1' }] });
  });

  it('opens only a worker workbook and delegates toolbar command state', async () => {
    const host = peerlessHarness();
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const view = render(host.asyncEditor({ onReady: ready, plugins: [] }));
    await opened();
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    const api = ready.mock.calls[0]![0];
    expect(host.open).toHaveBeenCalledTimes(1);
    expect(host.open.mock.calls[0]![1]).toMatchObject({ includeEditState: true });
    expect(host.open.mock.calls[0]![1]?.retainPeerHydration).toBeUndefined();
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
    expect(api).not.toBe(host.session);
    expect(api.hydrated).toBe(true);
    expect(api.failure).toBeNull();
    await api.whenHydrated();
    expect(api.commands).not.toBe(host.commands());
    expect(api.commands.getDescriptor('bold')).toBe(host.commands().getDescriptor('bold'));
    expect(api.commands.getState('bold')).toBe(host.commands().getState('bold'));
    expect(api.commands.getState('bold').enabled).toBe(true);
    api.focus();
    expect(document.activeElement).toBe(view.getByTestId('xlsx-scroll'));
  });

  it('cleans up replaced and unmounted workbooks and refuses stale APIs', async () => {
    const host = peerlessHarness();
    const cleanups = [mock(() => {}), mock(() => {})];
    let index = 0;
    const ready = mock((_api: XlsxAsyncEditorApi) => cleanups[index++]);
    const view = render(host.asyncEditor({ onReady: ready }));
    await opened();
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    const first = ready.mock.calls[0]![0];
    const choices = xlsxCommandController(first.commands)!;
    const prepared = choices.prepare('fontSize');
    view.rerender(host.asyncEditor({ file: file.slice(), onReady: ready }));
    expect(cleanups[0]).toHaveBeenCalledTimes(1);
    expect(host.session.dispose).toHaveBeenCalledTimes(1);
    await expect(first.whenHydrated()).rejects.toMatchObject({ code: 'document-replaced' });
    expect(() => first.focus()).toThrow(XlsxCommandAdmissionError);
    await opened();
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(2));
    const second = ready.mock.calls[1]![0];
    expect(second).not.toBe(first);
    expect(second.hydrated).toBe(true);
    expect(() => choices.prepare('textColor')).toThrow('The workbook was replaced');
    await expect(prepared.execute({ points: 14 })).rejects.toMatchObject({ code: 'document-replaced' });
    await expect(first.commands.execute('bold', null)).rejects.toMatchObject({ code: 'document-replaced' });
    expect(host.sent).toHaveLength(0);
    expect(host.cells.get('0:0:0')).toBe('initial');
    expect(second.commands.getState('fontSize').value).toBe(PLAIN_FORMATTING.fontSize);
    const command = second.commands.execute('bold', null);
    await waitFor(() => expect(host.sent).toHaveLength(1));
    await act(async () => host.ack(host.sent[0]!, { paint: true }));
    await expect(command).resolves.toMatchObject({ ok: true });
    view.unmount();
    expect(cleanups[0]).toHaveBeenCalledTimes(1);
    expect(cleanups[1]).toHaveBeenCalledTimes(1);
    await expect(second.whenHydrated()).rejects.toMatchObject({ code: 'document-replaced' });
    await expect(second.commands.execute('bold', null)).rejects.toMatchObject({ code: 'document-replaced' });
    expect(host.sent).toHaveLength(1);
    expect(() => second.focus()).toThrow(XlsxCommandAdmissionError);
  });

  it('reports a live failure and rejects hydration after worker failure', async () => {
    const host = peerlessHarness();
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const onError = mock((_error: Error) => {});
    render(host.asyncEditor({ onReady: ready, onError }));
    await opened();
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    const api = ready.mock.calls[0]![0];
    expect(api.hydrated).toBe(true);
    expect(api.failure).toBeNull();
    await api.whenHydrated();
    expect(onError).not.toHaveBeenCalled();
    const failure = new SessionFailure('crash', 'Worker stopped');
    await act(async () => host.fail(failure));
    expect(onError).toHaveBeenCalledTimes(1);
    expect(api.failure).toBe(onError.mock.calls[0]![0]);
    expect(api.failure).toMatchObject({ cause: failure });
    expect(api.hydrated).toBe(false);
    await expect(api.whenHydrated()).rejects.toBe(api.failure);
  });

  it('changes readOnly through ordered commands without reopening', async () => {
    const host = peerlessHarness();
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const view = render(host.asyncEditor({ onReady: ready, readOnly: true, plugins: [] }));
    await opened();
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    const scroll = view.getByTestId('xlsx-scroll');
    await waitFor(() => expect(host.sent[0]?.kind).toBe('setReadOnly'));
    expect(host.sent[0]).toMatchObject({ payload: { readOnly: true } });
    await act(async () => {
      host.ack(host.sent[0]);
      host.push({ throughSeq: host.sent[0].seq, version: 'v1', statuses: [], editState: { readOnly: true } });
    });
    view.rerender(host.asyncEditor({ onReady: ready, readOnly: false, plugins: [] }));
    await waitFor(() => expect(host.sent[1]?.kind).toBe('setReadOnly'));
    expect(host.sent[1]).toMatchObject({ payload: { readOnly: false } });
    await act(async () => {
      host.ack(host.sent[1]);
      host.push({ throughSeq: host.sent[1].seq, version: 'v1', statuses: [], editState: { readOnly: false } });
    });
    await advance();
    expect(view.getByTestId('xlsx-scroll')).toBe(scroll);
    expect(ready).toHaveBeenCalledTimes(1);
    expect(host.open).toHaveBeenCalledTimes(1);
    expect(host.session.dispose).not.toHaveBeenCalled();
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
  });

  it('publishes region bounds in collaboration presence while retaining the active cell', async () => {
    const host = collaborativePeerlessHarness();
    const cursor = mock((_cursor: Parameters<CollaborationAwareness['setCursor']>[0]) => {});
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const collaboration = { clientId: 101, initialUpdate: Uint8Array.of(7),
      provider: { peers: [], setCursor: cursor, onAwareness: () => () => {} } };
    host.selectionData.mockResolvedValue({ region: { top: 0, left: 0, bottom: 2, right: 2 } });
    const view = render(host.asyncEditor({ collaboration, onReady: ready }));
    await opened();
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.mouseDown(scroll, { clientX: 110, clientY: 30 });
    fireEvent.mouseUp(window);
    await act(async () => { fireEvent.keyDown(scroll, { key: 'a', ctrlKey: true }); });
    expect(cursor).toHaveBeenLastCalledWith({ sheet: 'sheet:0', anchor: { row: 0, col: 0 }, head: { row: 2, col: 2 } });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('B2');
  });

  it('mounts collaboration, paints remote changes and clears merged presence on departure and replacement', async () => {
    const host = collaborativePeerlessHarness();
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const onError = mock((_error: Error) => {});
    const replicas = mock((_replica: CollaborationReplica | null) => {});
    const awareness = new Set<(peers: readonly AwarenessPeer[]) => void>();
    const cursor = mock(() => {});
    const peer: AwarenessPeer = { clientId: 202, clock: 1, lastSeen: Date.now(), cursorMovedAt: Date.now(),
      user: { name: 'Peer', color: '#0B57D0' },
      cursor: { sheet: 'sheet:0', anchor: { row: 0, col: 1 }, head: { row: 0, col: 1 } } };
    const collaboration = { clientId: 101, initialUpdate: Uint8Array.of(7), onReplica: replicas,
      provider: { peers: [peer], setCursor: cursor,
        onAwareness(listener: (peers: readonly AwarenessPeer[]) => void) {
          awareness.add(listener); listener([peer]); return () => { awareness.delete(listener); };
        } } };
    host.frame.mockImplementation(async (viewport, options) => ({ viewport, sheet: options?.sheet ?? 0,
      sequence: host.session.throughSeq, epoch: 2, version: 'v2',
      displayList: { width: 800, height: 600,
        grid: { startRow: 0, startCol: 0, rowOffsets: [0, 24, 48], colOffsets: [0, 96, 192, 288] },
        commands: [{ op: 'text', text: host.cells.get('0:0:0')!, x: 8, y: 18, fontSize: 11, color: '#000000' }] },
      mergedRanges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 2 } }] }));
    const view = render(host.asyncEditor({ collaboration, onReady: ready, onError }));
    await opened();
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    expect(onError).not.toHaveBeenCalled();
    expect(host.open.mock.calls[0]![1]).toMatchObject({ collaborative: true, clientId: 101, initialUpdate: Uint8Array.of(7) });
    expect(replicas).toHaveBeenCalledWith(host.replica);
    expect(host.bindCollaboration).toHaveBeenCalled();
    expect(view.getByTestId('xlsx-remote-selection-202').style.width).toBe('288px');
    expect(view.getByTestId('xlsx-presence-chip-202')).toBeTruthy();
    expect(cursor).toHaveBeenCalledWith({ sheet: 'sheet:0', anchor: { row: 0, col: 0 }, head: { row: 0, col: 0 } });
    const api = ready.mock.calls[0]![0];
    expect(await api.collaboration!.encodeStateVector()).toEqual(Uint8Array.of(1));
    let merging!: Promise<unknown>;
    await act(async () => { merging = api.collaboration!.applyUpdate(new TextEncoder().encode('remote')); });
    expect(host.sent[0]).toMatchObject({ kind: 'applyUpdate', seq: 1 });
    await act(async () => host.ack(host.sent[0]));
    await merging;
    await advance();
    expect(painted[painted.length - 1].commands.some((command) => command.op === 'text' && command.text === 'remote')).toBe(true);
    await act(async () => { for (const listener of awareness) listener([]); });
    expect(view.queryByTestId('xlsx-remote-selection-202')).toBeNull();
    expect(view.queryByTestId('xlsx-presence-chip-202')).toBeNull();
    await act(async () => { for (const listener of awareness) listener([peer]); });
    view.rerender(host.asyncEditor({ file: file.slice(), onReady: ready }));
    expect(view.queryByTestId('xlsx-remote-selection-202')).toBeNull();
    expect(view.queryByTestId('xlsx-presence-chip-202')).toBeNull();
    expect(replicas).toHaveBeenLastCalledWith(null);
    expect(cursor).toHaveBeenLastCalledWith(null);
    expect(awareness.size).toBe(0);
    await expect(api.collaboration!.encodeStateVector()).rejects.toMatchObject({ code: 'document-replaced' });
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
  });

  it('keeps the cell draft when clicking beneath a remote selection', async () => {
    const host = collaborativePeerlessHarness();
    const peer: AwarenessPeer = { clientId: 202, clock: 1, lastSeen: Date.now(), cursorMovedAt: Date.now(),
      user: { name: 'Peer', color: '#0B57D0' },
      cursor: { sheet: 'sheet:0', anchor: { row: 0, col: 0 }, head: { row: 0, col: 0 } } };
    const view = render(host.asyncEditor({ collaboration: { clientId: 101,
      provider: { peers: [peer], setCursor: () => {},
        onAwareness(listener) { listener([peer]); return () => {}; } } } }));
    await opened();
    await act(async () => fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' }));
    const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    fireEvent.change(input, { target: { value: 'draft' } });
    const remote = view.getByTestId('xlsx-remote-selection-202');
    const layer = remote.closest('[style*="pointer-events: none"]');
    expect(layer).not.toBeNull();
    expect(view.getByTestId('xlsx-remote-label-202').closest('[style*="pointer-events: none"]')).toBe(layer);
    await act(async () => {
      fireEvent.mouseDown(input, { clientX: 12, clientY: 12, button: 0 });
      fireEvent.click(input, { clientX: 12, clientY: 12, button: 0 });
    });
    expect(view.getByTestId('xlsx-cell-editor')).toBe(input);
    expect(input.value).toBe('draft');
    expect(document.activeElement).toBe(input);
    expect(host.sent).toHaveLength(0);
    expect(host.cells.get('0:0:0')).toBe('initial');
  });

  it('orders pipe commands with local input and admits them while read-only', async () => {
    const host = collaborativePeerlessHarness();
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const collaboration = { clientId: 101 };
    const view = render(host.asyncEditor({ collaboration, onReady: ready }));
    await opened();
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    const applyPipe = host.bindCollaboration.mock.calls[0]![2]!;
    expect(typeof applyPipe).toBe('function');
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'local' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    const payload = { provider: 1, ops: [{ op: 'connect' as const, seq: 1 }] };
    let merging!: Promise<unknown>;
    await act(async () => { merging = applyPipe(payload); });
    expect(host.sent.map(({ kind, seq }) => [kind, seq])).toEqual([['editCell', 1], ['collabPipe', 2]]);
    expect(host.sent[1]).toMatchObject({ payload });
    expect(host.sent[1].afterSeq).toBeUndefined();
    await act(async () => {
      host.ack(host.sent[0], { paint: true });
      host.ack(host.sent[1], { result: { remoteUpdate: true, proposals: [] }, paint: true });
      await merging;
    });
    view.rerender(host.asyncEditor({ collaboration, onReady: ready, readOnly: true }));
    await waitFor(() => expect(host.sent[2]?.kind).toBe('setReadOnly'));
    await act(async () => { merging = applyPipe({ provider: 1, ops: [{ op: 'disconnect', seq: 2 }] }); });
    expect(host.sent.map(({ kind, seq }) => [kind, seq])).toEqual([
      ['editCell', 1], ['collabPipe', 2], ['setReadOnly', 3], ['collabPipe', 4],
    ]);
    await act(async () => { host.ack(host.sent[2]); host.ack(host.sent[3]); await merging; });
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
  });

  it('coalesces pipe ack frames and keeps print waiting for the installed frame', async () => {
    const host = collaborativePeerlessHarness();
    const view = render(host.asyncEditor({ collaboration: { clientId: 101 } }));
    await opened();
    await advance();
    const input = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
    await waitFor(() => expect(input.value).toBe('initial'));
    const frames = host.frame.mock.calls.length;
    const inputs = host.cellInputs.mock.calls.length;
    const paints = painted.length;
    const canvas = view.container.querySelector('canvas')!;
    const applyPipe = host.bindCollaboration.mock.calls[0]![2]!;
    const applying: Promise<unknown>[] = [];
    await act(async () => {
      applying.push(applyPipe({ provider: 1, ops: [{ op: 'connect', seq: 1 }] }));
      applying.push(applyPipe({ provider: 1, ops: [{ op: 'disconnect', seq: 2 }] }));
    });
    expect(host.sent.map((command) => command.kind)).toEqual(['collabPipe', 'collabPipe']);
    host.cells.set('0:0:0', 'intermediate');
    await act(async () => host.ack(host.sent[0], { paint: true, result: { remoteUpdate: true } }));
    expect(painted).toHaveLength(paints);
    expect(host.session.presentedThroughSeq).toBe(0);
    expect(host.cellInputs.mock.calls.length).toBe(inputs);
    host.cells.set('0:0:0', 'latest');
    await act(async () => {
      host.ack(host.sent[1], { paint: true, result: { remoteUpdate: true } });
      await Promise.all(applying);
    });
    const print = mock(() => {});
    replacePrint(print);
    const printing = host.commands().execute('print', null);
    await act(async () => {});
    expect(print).not.toHaveBeenCalled();
    expect(painted).toHaveLength(paints);
    expect(host.session.presentedThroughSeq).toBe(0);
    expect(canvas.dataset.commandCoverage).toBe('0');
    await tick();
    expect(await printing).toMatchObject({ ok: true });
    expect(print).toHaveBeenCalledTimes(1);
    expect(painted).toHaveLength(paints + 1);
    expect(painted[paints].commands.some((entry) => entry.op === 'text' && entry.text === 'latest')).toBe(true);
    expect(host.session.presentedThroughSeq).toBe(2);
    expect(canvas.dataset.commandCoverage).toBe('2');
    expect(host.cellInputs.mock.calls.length).toBe(inputs + 1);
    expect(input.value).toBe('latest');
    await advance();
    expect(host.frame.mock.calls.length).toBe(frames);
    expect(painted).toHaveLength(paints + 1);
  });

  it('presents a pipe frame covering a pending local edit without waiting for animation', async () => {
    const host = collaborativePeerlessHarness();
    const view = render(host.asyncEditor({ collaboration: { clientId: 101 } }));
    await opened();
    await advance();
    const applyPipe = host.bindCollaboration.mock.calls[0]![2]!;
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'local' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    let applying!: Promise<unknown>;
    await act(async () => { applying = applyPipe({ provider: 1, ops: [{ op: 'connect', seq: 1 }] }); });
    expect(host.sent.map((command) => command.kind)).toEqual(['editCell', 'collabPipe']);
    const pipe = host.sent[1];
    const request = pipe.frame!;
    host.cells.set('0:0:0', 'local');
    const frame = await host.frame(request.viewport, { sheet: request.sheet });
    await act(async () => {
      host.push({ throughSeq: pipe.seq, version: 'v1',
        statuses: host.sent.map((command) => ({ seq: command.seq, status: 'applied' })),
        frame: { ...frame, ...host.session.commandIdentity, coversThroughSeq: pipe.seq,
          requestGeneration: request.requestGeneration ?? request.generation, zoom: request.zoom, dpr: request.dpr } });
      await applying;
    });
    expect(host.session.presentedThroughSeq).toBe(pipe.seq);
    expect(view.container.querySelector('canvas')!.dataset.commandCoverage).toBe(String(pipe.seq));
    expect(painted[painted.length - 1].commands.some((entry) => entry.op === 'text' && entry.text === 'local')).toBe(true);
  });

  it('presents a pipe no-op without installing a frame or fetching cell input', async () => {
    const host = collaborativePeerlessHarness();
    const view = render(host.asyncEditor({ collaboration: { clientId: 101 } }));
    await opened();
    await advance();
    await waitFor(() => expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).value).toBe('initial'));
    const frames = host.frame.mock.calls.length;
    const inputs = host.cellInputs.mock.calls.length;
    const paints = painted.length;
    const applyPipe = host.bindCollaboration.mock.calls[0]![2]!;
    let applying!: Promise<unknown>;
    await act(async () => { applying = applyPipe({ provider: 1, ops: [{ op: 'connect', seq: 1 }] }); });
    const command = await host.commandAt(0);
    let presented = false;
    const waiting = host.session.whenPresented(command.seq).then(() => { presented = true; });
    await act(async () => { host.ack(command, { status: 'noop' }); await applying; });
    await waitFor(() => expect(presented).toBe(true));
    await waiting;
    fireEvent.scroll(view.getByTestId('xlsx-scroll'));
    await advance();
    expect(view.container.querySelector('canvas')!.dataset.commandCoverage).toBe(String(command.seq));
    expect(host.frame.mock.calls.length).toBe(frames);
    expect(host.cellInputs.mock.calls.length).toBe(inputs);
    expect(painted).toHaveLength(paints);
  });

  for (const attached of [true, false]) {
    it(`uses ${attached ? 'the attached frame' : 'one fallback frame'} for a remote update`, async () => {
      const host = collaborativePeerlessHarness();
      const ready = mock((_api: XlsxAsyncEditorApi) => {});
      const view = render(host.asyncEditor({ collaboration: { clientId: 101 }, onReady: ready }));
      await opened();
      await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
      await advance();
      const frames = host.frame.mock.calls.length;
      let applying!: Promise<unknown>;
      await act(async () => { applying = ready.mock.calls[0]![0].collaboration!.applyUpdate(new TextEncoder().encode('remote')); });
      const command = host.sent[host.sent.length - 1];
      expect(command.kind).toBe('applyUpdate');
      await act(async () => { host.ack(command, { paint: attached }); await applying; });
      if (attached) {
        expect(painted[painted.length - 1].commands.some((entry) => entry.op === 'text' && entry.text === 'remote')).toBe(true);
        expect(host.session.markPresented).toHaveBeenLastCalledWith({ ...host.session.commandIdentity, coversThroughSeq: command.seq });
      }
      await advance();
      expect(host.frame.mock.calls.length).toBe(frames + (attached ? 0 : 1));
      expect(view.container.querySelector('canvas')!.dataset.commandCoverage).toBe(String(command.seq));
    });
  }

  for (const attached of [true, false]) {
    it(`uses ${attached ? 'the attached frame' : 'one fallback frame'} for a remote update from the worker replica`, async () => {
      const host = collaborativePeerlessHarness();
      const view = render(host.asyncEditor({ collaboration: { clientId: 101 } }));
      await opened();
      await advance();
      const frames = host.frame.mock.calls.length;
      const applying = host.replica.async!.applyUpdate(new TextEncoder().encode('remote'));
      const command = await host.commandAt(0);
      await act(async () => { host.ack(command, { paint: attached }); await applying; });
      if (attached) {
        expect(painted[painted.length - 1].commands.some((entry) => entry.op === 'text' && entry.text === 'remote')).toBe(true);
        expect(host.session.presentedThroughSeq).toBe(command.seq);
      }
      await advance();
      expect(host.frame.mock.calls.length).toBe(frames + (attached ? 0 : 1));
      expect(view.container.querySelector('canvas')!.dataset.commandCoverage).toBe(String(command.seq));
    });
  }

  it('sequences remote updates with local edits and read-only changes', async () => {
    const host = collaborativePeerlessHarness();
    const ready = mock((_api: XlsxAsyncEditorApi) => {});
    const collaboration = { clientId: 101 };
    const view = render(host.asyncEditor({ collaboration, onReady: ready }));
    await opened();
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    expect(host.bindCollaboration).toHaveBeenCalled();
    const api = ready.mock.calls[0]![0];
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'local' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    let applied = false;
    let merging!: Promise<unknown>;
    await act(async () => {
      merging = api.collaboration!.applyUpdate(new TextEncoder().encode('remote')).then(() => { applied = true; });
    });
    expect(host.sent.map(({ kind, seq }) => [kind, seq])).toEqual([['editCell', 1], ['applyUpdate', 2]]);
    expect(host.session.throughSeq).toBe(0);
    expect(host.cells.get('0:0:0')).toBe('initial');
    expect(applied).toBe(false);
    await act(async () => host.ack(host.sent[0], { paint: true }));
    expect(host.cells.get('0:0:0')).toBe('local');
    expect(applied).toBe(false);
    await act(async () => host.ack(host.sent[1], { paint: true }));
    await merging;
    expect(applied).toBe(true);
    expect(host.cells.get('0:0:0')).toBe('remote');
    view.rerender(host.asyncEditor({ collaboration, onReady: ready, readOnly: true }));
    await waitFor(() => expect(host.sent[2]?.kind).toBe('setReadOnly'));
    await act(async () => {
      merging = api.collaboration!.applyUpdate(new TextEncoder().encode('read-only remote'));
    });
    expect(host.sent.map(({ kind, seq }) => [kind, seq])).toEqual([
      ['editCell', 1], ['applyUpdate', 2], ['setReadOnly', 3], ['applyUpdate', 4],
    ]);
    await act(async () => {
      host.ack(host.sent[2]);
      host.push({ throughSeq: 3, version: 'v2', statuses: [], editState: { readOnly: true } });
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    });
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    expect(host.session.editState.readOnly).toBe(true);
    expect(host.cells.get('0:0:0')).toBe('remote');
    await act(async () => host.ack(host.sent[3], { paint: true }));
    await merging;
    await advance();
    expect(host.sent).toHaveLength(4);
    expect(host.session.throughSeq).toBe(4);
    expect(painted[painted.length - 1].commands.some((command) =>
      command.op === 'text' && command.text === 'read-only remote')).toBe(true);
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
  });
});

describe('worker editor open and empty states', () => {
  it('shows a styled, localized open failure with the technical details and reports it once', async () => {
    const open = spyOn(editableWorkbookSessionBackend, 'open').mockRejectedValue(new Error('Unsupported zip entry'));
    restorers.push(() => open.mockRestore());
    const onError = mock((_error: Error) => {});
    const view = render(<XlsxWorkerEditor backend="peerless" file={file} onError={onError}
      i18n={{ editor: { openError: 'Öffnen fehlgeschlagen' } }} />);
    await waitFor(() => expect(view.queryByTestId('xlsx-error')).not.toBeNull());
    const failure = view.getByTestId('xlsx-error');
    expect(failure.getAttribute('role')).toBe('alert');
    expect(failure.getAttribute('data-failure')).toBe('open');
    expect(failure.textContent).toContain('Öffnen fehlgeschlagen');
    expect(failure.textContent).toContain('The file may be damaged');
    expect(failure.querySelector('summary')?.textContent).toBe('Show details');
    expect(failure.querySelector('details pre')?.textContent).toContain('Unsupported zip entry');
    expect(view.getByTestId('xlsx-workspace').contains(failure)).toBe(true);
    expect(within(failure).queryByRole('button')).toBeNull();
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError.mock.calls[0][0].message).toContain('Unsupported zip entry');
  });

  it('asks for a workbook when no file is given', async () => {
    const view = render(<XlsxWorkerEditor backend="peerless" />);
    await waitFor(() => expect(view.queryByTestId('xlsx-empty')).not.toBeNull());
    expect(view.getByTestId('xlsx-empty').getAttribute('role')).toBe('status');
    expect(view.getByTestId('xlsx-empty').textContent).toBe('Open an XLSX file');
    expect(view.queryByTestId('xlsx-error')).toBeNull();
  });
});

describe('peerless editor chunk failures', () => {
  beforeEach(() => {
    const errors = spyOn(console, 'error').mockImplementation(() => {});
    restorers.push(() => errors.mockRestore());
  });

  async function rejectChunk(reject: (reason: unknown) => void, failure: Error) {
    try { await act(async () => { reject(failure); }); }
    catch (error) { if (error !== failure) throw error; }
  }

  it('keeps host siblings mounted and shows the editor error after a rejected import', async () => {
    const chunk = deferred<Awaited<ReturnType<typeof xlsxPeerlessEditorLoader.load>>>();
    const load = spyOn(xlsxPeerlessEditorLoader, 'load').mockReturnValue(chunk.promise);
    restorers.push(() => load.mockRestore());
    const view = render(<><span>Host sibling</span><XlsxWorkerEditor backend="peerless" file={file} /></>);
    await rejectChunk(chunk.reject, new Error('Chunk fetch failed'));

    expect(view.queryByText('Host sibling')).not.toBeNull();
    expect(view.getByTestId('xlsx-error').getAttribute('role')).toBe('alert');
    expect(view.getByTestId('xlsx-error').textContent).toContain("The editor couldn't load");
    expect(view.getByTestId('xlsx-error').textContent).toContain('Failed to load');
  });

  it('reports a typed chunk load error once even when the host callback throws', async () => {
    const chunk = deferred<Awaited<ReturnType<typeof xlsxPeerlessEditorLoader.load>>>();
    const load = spyOn(xlsxPeerlessEditorLoader, 'load').mockReturnValue(chunk.promise);
    restorers.push(() => load.mockRestore());
    const onError = mock((_error: Error) => { throw new Error('Host callback failed'); });
    const view = render(<StrictMode><span>Host sibling</span>
      <XlsxWorkerEditor backend="peerless" file={file} onError={onError} />
    </StrictMode>);
    const failure = new Error('Chunk fetch failed');
    await rejectChunk(chunk.reject, failure);
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError.mock.calls[0][0]).toBeInstanceOf(Error);
    expect(onError.mock.calls[0][0]).toMatchObject({
      name: 'XlsxWorkerEditorChunkLoadError', code: 'chunk-load-failed', cause: failure,
    });
    view.rerender(<StrictMode><span>Host sibling</span>
      <XlsxWorkerEditor backend="peerless" file={new Uint8Array([4])} onError={onError} />
    </StrictMode>);
    await advance();
    expect(onError).toHaveBeenCalledTimes(1);
    expect(view.queryByText('Host sibling')).not.toBeNull();
  });

  it('retries a transient import failure on a later mount without opening a main-thread peer', async () => {
    const host = peerlessHarness();
    const chunk = deferred<Awaited<ReturnType<typeof xlsxPeerlessEditorLoader.load>>>();
    const originalLoad = xlsxPeerlessEditorLoader.load;
    const load = spyOn(xlsxPeerlessEditorLoader, 'load').mockImplementation(originalLoad)
      .mockReturnValueOnce(chunk.promise);
    restorers.push(() => load.mockRestore());
    const onError = mock((_error: Error) => {});
    const failed = render(<XlsxWorkerEditor backend="peerless" file={file} onError={onError} />);
    const failure = new Error('Transient chunk fetch failure');
    await rejectChunk(chunk.reject, failure);
    failed.unmount();

    try {
      render(<XlsxWorkerEditor backend="peerless" file={file} onError={onError} onReady={host.ready} />);
      await opened();
    } catch (error) { if (error !== failure) throw error; }

    expect(load).toHaveBeenCalledTimes(2);
    await waitFor(() => expect(host.ready).toHaveBeenCalledTimes(1));
    expect(onError).toHaveBeenCalledTimes(1);
    expect(host.open).toHaveBeenCalledTimes(1);
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
  });
});

describe('peerless workbook worker editor', () => {
  it('bounds visible Enter commits with unchanged formatting and geometry', async () => {
    const formattingRenders = mock(() => {});
    const formulaRenders = mock(() => {});
    const Toolbar = EditorToolbar.Toolbar;
    const FormulaBar = EditorToolbar.FormulaBar;
    const toolbar = spyOn(EditorToolbar, 'Toolbar').mockImplementation((props) =>
      <Profiler id="formatting" onRender={formattingRenders}><Toolbar {...props} /></Profiler>);
    const formula = spyOn(EditorToolbar, 'FormulaBar').mockImplementation((props) =>
      <Profiler id="formula" onRender={formulaRenders}><FormulaBar {...props} /></Profiler>);
    const accessibility = spyOn(xlsx, 'buildA11yGrid');
    restorers.push(() => toolbar.mockRestore(), () => formula.mockRestore(), () => accessibility.mockRestore());
    const host = peerlessHarness(false, false, true);
    for (let row = 1; row <= 3; row++) host.cells.set(`0:${row}:0`, `next ${row}`);
    const view = await host.mount();
    await advance();
    expect(formattingRenders).toHaveBeenCalled();
    expect(formulaRenders).toHaveBeenCalled();
    for (let index = 0; index < 3; index++) {
      await act(async () => fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x' }));
      await act(async () => fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: `edit ${index}` } }));
      const geometry = deferred<xlsx.WorkbookSheetView>();
      host.sheetView.mockReturnValueOnce(geometry.promise);
      const inputs: { a1: string; reply: ReturnType<typeof deferred<Awaited<ReturnType<typeof host.cellInputs>>>> }[] = [];
      host.cellInputs.mockImplementation((_sheet, a1) => {
        const reply = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
        inputs.push({ a1, reply });
        return reply.promise;
      });
      formattingRenders.mockClear(); formulaRenders.mockClear(); host.renderCount.mockClear(); accessibility.mockClear();
      await act(async () => fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' }));
      expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe(`A${index + 2}`);
      expect(accessibility).toHaveBeenCalledTimes(1);
      expect(formattingRenders).not.toHaveBeenCalled();
      const command = host.sent[index]!;
      expect(command.kind).toBe('editCell');
      await act(async () => host.ack(command));
      const request = command.frame!;
      const frame = await host.frame(request.viewport, { sheet: request.sheet });
      await act(async () => host.push({ throughSeq: command.seq, version: frame.version, statuses: [], frame: {
        ...frame, ...host.session.commandIdentity, coversThroughSeq: command.seq,
        requestGeneration: request.requestGeneration ?? request.generation, zoom: request.zoom, dpr: request.dpr,
      } }));
      await tick();
      expect(accessibility).toHaveBeenCalledTimes(2);
      const beforeGeometry = host.renderCount.mock.calls.length;
      await act(async () => geometry.resolve({ sheet: 0, version: `v${index + 2}`, contentWidth: 4000, contentHeight: 6000,
        frozenRows: 0, frozenCols: 0, frozenWidth: 0, frozenHeight: 0, initialScrollX: 0, initialScrollY: 0 }));
      expect(host.renderCount.mock.calls.length).toBe(beforeGeometry);
      await act(async () => {
        for (const { a1, reply } of inputs) reply.resolve({ sheet: 0, version: frame.version,
          cells: [[{ a1, input: `next ${index + 1}`, isFormula: false }]] });
      });
      expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).value).toBe(`next ${index + 1}`);
      expect(accessibility).toHaveBeenCalledTimes(2);
      expect(formattingRenders).not.toHaveBeenCalled();
      expect(formulaRenders.mock.calls.length).toBeLessThanOrEqual(2);
      expect(host.renderCount.mock.calls.length).toBeLessThanOrEqual(3);
      const beforeAck = host.renderCount.mock.calls.length;
      await act(async () => host.push({ throughSeq: command.seq, version: frame.version,
        statuses: [{ seq: command.seq, status: 'noop', committed: false }], editState: { canUndo: true, canRedo: true } }));
      await advance();
      expect(host.renderCount.mock.calls.length).toBe(beforeAck);
      expect(accessibility).toHaveBeenCalledTimes(2);
    }
  });

  for (const source of ['formula', 'cell'] as const) it(`formula input after a ${source} Enter waits for the next cell text and keeps a whole selection when it lands`, async () => {
    const host = peerlessHarness();
    host.cells.set('0:0:0', 'first');
    host.cells.set('0:1:0', 'second');
    const view = await host.mount();
    const input = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
    await waitFor(() => expect(input.value).toBe('first'));
    const next = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
    host.cellInputs.mockImplementation(() => next.promise);
    if (source === 'formula') {
      fireEvent.focus(input);
      fireEvent.change(input, { target: { value: 'edited' } });
      await act(async () => fireEvent.keyDown(input, { key: 'Enter' }));
    } else {
      await act(async () => fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'e' }));
      const cell = view.getByTestId('xlsx-cell-editor');
      fireEvent.change(cell, { target: { value: 'edited' } });
      await act(async () => fireEvent.keyDown(cell, { key: 'Enter' }));
    }
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A2');
    expect(host.sent[0]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'edited'] } });
    await act(async () => host.ack(host.sent[0]!, { paint: true }));
    await advance();
    expect(input.value).toBe('');
    input.focus();
    input.select();
    await act(async () => next.resolve({ sheet: 0, version: 'v1', cells: [[{ a1: 'A2', input: 'second', isFormula: false }]] }));
    expect(input.value).toBe('second');
    expect([input.selectionStart, input.selectionEnd]).toEqual([0, 'second'.length]);
  });

  it('sealing formula input preserves its text and protects a newer draft from covering paint', async () => {
    const host = peerlessHarness();
    host.cells.set('0:0:0', 'A');
    const view = await host.mount();
    const input = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
    await waitFor(() => expect(input.value).toBe('A'));
    fireEvent.focus(input);
    fireEvent.change(input, { target: { value: `${input.value}B` } });
    const bold = host.commands().execute('bold', null);
    await act(async () => {});
    expect(host.sent.map((command) => command.kind)).toEqual(['editCell', 'applyOps']);
    expect(input.value).toBe('AB');
    fireEvent.change(input, { target: { value: `${input.value}C` } });
    await act(async () => host.ack(host.sent[0], { paint: true }));
    expect(input.value).toBe('ABC');
    await act(async () => host.ack(host.sent[1]));
    await bold;
    fireEvent.keyDown(input, { key: 'Enter' });
    expect(host.sent[2]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'ABC'] } });
    await act(async () => host.ack(host.sent[2], { paint: true }));
    await advance();
    expect(host.cells.get('0:0:0')).toBe('ABC');
  });

  it('submitted formula text remains visible until covering paint without resubmission', async () => {
    const host = peerlessHarness();
    host.cells.set('0:0:0', 'A');
    const view = await host.mount();
    const input = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
    await waitFor(() => expect(input.value).toBe('A'));
    fireEvent.change(input, { target: { value: 'AB' } });
    const bold = host.commands().execute('bold', null);
    await act(async () => {});
    await act(async () => { host.ack(host.sent[0]); host.ack(host.sent[1]); });
    await bold;
    expect(input.value).toBe('AB');
    const italic = host.commands().execute('italic', null);
    await act(async () => {});
    expect(host.sent.map((command) => command.kind)).toEqual(['editCell', 'applyOps', 'applyOps']);
    expect(input.value).toBe('AB');
    await act(async () => host.ack(host.sent[2], { paint: true }));
    await italic;
    await advance();
    expect(input.value).toBe('AB');
  });

  it('keeps focused formula reads across cancelled cell input and scrolling', async () => {
    const host = peerlessHarness();
    host.cells.set('0:0:1', 'second');
    const view = await host.mount();
    const input = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
    const name = view.getByTestId('xlsx-name-box') as HTMLInputElement;
    const scroll = view.getByTestId('xlsx-scroll');
    const canvas = view.container.querySelector('canvas')!;
    await waitFor(() => expect(input.value).toBe('initial'));
    await advance();
    const reply = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
    const result = { sheet: 0, version: 'v1', cells: [[{ a1: 'B1', input: 'second', isFormula: false }]] };
    host.cellInputs.mockReturnValue(reply.promise);
    await act(async () => fireEvent.keyDown(scroll, { key: 'ArrowRight' }));
    const selection: xlsxWorker.WorkbookCommandSelection = { sheetId: 'sheet:0', activeCell: { row: 0, col: 1 },
      range: { kind: 'rowCol', start: { row: 0, col: 1 }, end: { row: 0, col: 1 } } };
    expect(host.sent).toHaveLength(1);
    expect(host.sent[0]).toMatchObject({ kind: 'select', payload: { selection } });
    await act(async () => host.ack(host.sent[0], { paint: true, selection }));
    await advance();
    expect(name.value).toBe('B1');
    expect(host.cellInputs).toHaveBeenLastCalledWith(0, 'B1');
    fireEvent.keyDown(scroll, { key: 'x' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Escape' });
    await act(async () => reply.resolve(result));
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    expect(input.value).toBe('second');
    host.cellInputs.mockResolvedValue({ ...result, version: 'v2' });
    const nextFrame = host.frame.getMockImplementation()!;
    host.frame.mockImplementationOnce(async (...args) => ({ ...await nextFrame(...args), version: 'v2' }));
    const reads = host.cellInputs.mock.calls.length;
    const frames = host.frame.mock.calls.length;
    const coverage = canvas.dataset.commandCoverage;
    expect(coverage).toBe('1');
    scroll.scrollTop = 24;
    fireEvent.scroll(scroll);
    await advance();
    expect(host.frame.mock.calls.length).toBeGreaterThan(frames);
    expect(canvas.dataset.commandCoverage).toBe(coverage);
    expect(host.cellInputs.mock.calls.length).toBeGreaterThan(reads);
    expect(host.cellInputs).toHaveBeenLastCalledWith(0, 'B1');
    expect(name.value).toBe('B1');
    expect(input.value).toBe('second');
    expect(host.sent).toHaveLength(1);
  });

  it('refetches focused cell input on selection-range, scroll, painted-version and selected-cell changes', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    await advance();
    const input = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
    await waitFor(() => expect(input.value).toBe('initial'));
    const reads = host.cellInputs.mock.calls.length;
    await act(async () => host.setReadOnly(view, false));
    await advance();
    expect(host.cellInputs.mock.calls.length).toBe(reads);
    const scroll = view.getByTestId('xlsx-scroll');
    await act(async () => fireEvent.keyDown(scroll, { key: 'ArrowRight', shiftKey: true }));
    expect(host.cellInputs.mock.calls.length).toBe(reads + 1);
    expect(host.cellInputs).toHaveBeenLastCalledWith(0, 'A1');
    await act(async () => host.ack(host.sent[0], { paint: true }));
    expect(host.cellInputs.mock.calls.length).toBe(reads + 2);
    expect(input.value).toBe('initial');
    scroll.scrollTop = 24;
    fireEvent.scroll(scroll);
    await advance();
    expect(host.cellInputs.mock.calls.length).toBe(reads + 3);
    expect(host.cellInputs).toHaveBeenLastCalledWith(0, 'A1');
    const beforeVersionChange = host.cellInputs.mock.calls.length;
    const nextFrame = host.frame.getMockImplementation()!;
    host.frame.mockImplementationOnce(async (...args) => ({ ...await nextFrame(...args), version: 'v2' }));
    host.cells.set('0:0:0', 'changed');
    scroll.scrollTop = 48;
    fireEvent.scroll(scroll);
    await advance();
    expect(host.cellInputs.mock.calls.length).toBe(beforeVersionChange + 1);
    expect(host.cellInputs).toHaveBeenLastCalledWith(0, 'A1');
    expect(input.value).toBe('changed');
    await act(async () => fireEvent.keyDown(scroll, { key: 'ArrowRight' }));
    expect(host.cellInputs.mock.calls.length).toBe(beforeVersionChange + 2);
    expect(host.cellInputs).toHaveBeenLastCalledWith(0, 'B1');
  });

  for (const route of ['blur', 'composition'] as const) {
    it(`disables Redo after a same-cell formula ${route} commit`, async () => {
      const host = peerlessHarness(true);
      const selection: xlsxWorker.WorkbookCommandSelection = { sheetId: 'sheet:0', activeCell: { row: 0, col: 0 },
        range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 0, col: 0 } } };
      host.push({ throughSeq: 0, version: 'v1', statuses: [], selection, editState: { canUndo: true, canRedo: false } });
      const view = await host.mount();
      const input = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
      const redo = view.getByTestId('xlsx-redo') as HTMLButtonElement;
      await waitFor(() => expect(input.value).toBe('initial'));
      await advance();
      expect(redo.getAttribute('aria-disabled')).toBe('true');
      expect(host.commands().getState('redo').enabled).toBe(false);
      fireEvent.focus(input);
      if (route === 'composition') fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: 'edited' } });
      if (route === 'composition') fireEvent.compositionEnd(input);
      else fireEvent.blur(input);
      await act(async () => {});
      expect(host.sent).toHaveLength(1);
      expect(host.sent[0]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'edited'], selection } });
      expect(host.commands().getState('redo').enabled).toBe(true);
      expect(redo.getAttribute('aria-disabled')).toBeNull();
      await act(async () => host.ack(host.sent[0], { paint: true, selection }));
      await advance();
      expect(host.session.editState).toMatchObject({ canUndo: true, canRedo: false });
      expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
      expect(host.cells.get('0:0:0')).toBe('edited');
      expect(host.commands().getState('redo')).toMatchObject({ enabled: false, disabledReason: { code: 'nothing-to-redo' } });
      expect(redo.getAttribute('aria-disabled')).toBe('true');
    });
  }

  it('Redo reaches the worker behind an unacknowledged Undo', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'a' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await act(async () => {
      host.ack(host.sent[0], { paint: true });
      host.push({ throughSeq: 1, version: 'v1', statuses: [], editState: { canUndo: true, canRedo: false } });
    });
    expect(host.commands().getState('redo').enabled).toBe(false);
    const undoing = host.commands().execute('undo', null);
    const redoing = host.commands().execute('redo', null);
    await act(async () => {});
    expect(host.sent.map((command) => command.kind)).toEqual(['editCell', 'undo', 'redo']);
    await act(async () => { host.ack(host.sent[1]); host.ack(host.sent[2], { paint: true }); });
    expect(await undoing).toMatchObject({ ok: true });
    expect(await redoing).toMatchObject({ ok: true });
  });

  for (const print of [undefined, mock(() => { throw new Error('Host print failed'); })]) {
    it(`${print ? 'throwing' : 'missing'} window.print leaves the worker session editable`, async () => {
      replacePrint(print);
      const host = peerlessHarness();
      const view = await host.mount();
      expect(await host.commands().execute('print', null)).toMatchObject({ ok: false, failure: { code: 'command-failed' } });
      expect(view.queryByTestId('xlsx-error')).toBeNull();
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x' });
      fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
      expect(host.sent[0]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'x'] } });
      await act(async () => host.ack(host.sent[0], { paint: true }));
      expect(host.cells.get('0:0:0')).toBe('x');
    });
  }

  it('a throwing Save callback leaves the worker session editable', async () => {
    const host = peerlessHarness();
    const view = await host.mount(() => { throw new Error('Host save failed'); });
    expect(await host.commands().execute('save', null)).toMatchObject({ ok: false, failure: { code: 'command-failed' } });
    expect(view.queryByTestId('xlsx-error')).toBeNull();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    expect(host.sent[0]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'x'] } });
    await act(async () => host.ack(host.sent[0], { paint: true }));
    expect(host.cells.get('0:0:0')).toBe('x');
  });

  it('formula Enter navigates after IME acknowledgement paint and conversion keyup', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const input = view.getByTestId('xlsx-formula-input');
    fireEvent.focus(input);
    fireEvent.compositionStart(input);
    fireEvent.change(input, { target: { value: '確定' } });
    fireEvent.keyDown(input, { key: 'Enter', isComposing: true });
    fireEvent.compositionEnd(input);
    await act(async () => {});
    await act(async () => host.ack(host.sent[0], { paint: true }));
    expect(host.session.presentedThroughSeq).toBe(host.sent[0].seq);
    fireEvent.keyUp(input, { key: 'Enter' });
    fireEvent.keyDown(input, { key: 'Enter' });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A2');
    expect(host.sent.map((command) => command.kind)).toEqual(['editCell', 'select']);
    expect(host.sent[1]).toMatchObject({ payload: { selection: { activeCell: { row: 1, col: 0 } } } });
  });

  it('mouse format painting applies when a held selection finishes after acknowledgement and paint', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const capturing = host.commands().execute('paintFormat', null);
    await act(async () => {});
    await act(async () => host.ack(host.sent[0], { result: { formatToken: 'format:1' } }));
    await capturing;
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.mouseDown(scroll, { clientX: 120, clientY: 12, button: 0 });
    expect(host.sent[1].kind).toBe('select');
    await act(async () => host.ack(host.sent[1], { paint: true }));
    expect(host.sent.filter((command) => command.kind === 'applyFormat')).toHaveLength(0);
    fireEvent.mouseUp(window, { clientX: 120, clientY: 12, button: 0 });
    await act(async () => {});
    expect(host.sent[2]).toMatchObject({ kind: 'applyFormat', payload: {
      formatToken: 'format:1', target: { sheetId: 'sheet:0', range: {
        kind: 'rowCol', start: { row: 0, col: 1 }, end: { row: 0, col: 1 },
      } },
    } });
    await act(async () => host.ack(host.sent[2], { paint: true }));
    expect(host.commands().getState('paintFormat').active).toBe(false);
  });

  it('toolbar formatting preserves the reverse selection anchor through the next edit', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.keyDown(scroll, { key: 'ArrowDown' });
    fireEvent.keyDown(scroll, { key: 'ArrowDown' });
    fireEvent.keyDown(scroll, { key: 'ArrowUp', shiftKey: true });
    fireEvent.keyDown(scroll, { key: 'ArrowUp', shiftKey: true });
    const selection: xlsxWorker.WorkbookCommandSelection = { sheetId: 'sheet:0', activeCell: { row: 2, col: 0 },
      range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 2, col: 0 } } };
    await act(async () => host.ack(host.sent[3], { paint: true, selection }));
    const formatting = host.commands().execute('bold', null);
    await act(async () => {});
    const bold = host.sent[4];
    expect(bold.payload).toMatchObject({ selection });
    await act(async () => host.ack(bold, { paint: true,
      selection: 'selection' in bold.payload ? bold.payload.selection : undefined }));
    await formatting;
    fireEvent.keyDown(scroll, { key: 'x' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    expect(host.sent[5]).toMatchObject({ kind: 'editCell', payload: { args: [0, 2, 0, 'x'] } });
  });

  it('a delayed toolbar reply preserves newer clipboard-blocked navigation', async () => {
    const host = peerlessHarness();
    const first = deferred<string>();
    const second = deferred<string>();
    replaceClipboard('readText', mock().mockReturnValueOnce(first.promise).mockReturnValueOnce(second.promise));
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.keyDown(scroll, { key: 'v', ctrlKey: true });
    const formatting = host.commands().execute('bold', null);
    fireEvent.keyDown(scroll, { key: 'v', ctrlKey: true });
    fireEvent.keyDown(scroll, { key: 'ArrowRight' });
    await act(async () => first.resolve('first'));
    expect(host.sent.map((command) => command.kind)).toEqual(['editCells', 'applyOps']);
    await act(async () => {
      host.ack(host.sent[0]);
      host.ack(host.sent[1], { paint: true, selection: { sheetId: 'sheet:0',
        range: { kind: 'a1', a1: 'A1' } } });
    });
    await formatting;
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('B1');
    await act(async () => second.resolve('second'));
    await act(async () => {
      host.ack(host.sent[2]);
      host.ack(host.sent[3], { paint: true, selection: { sheetId: 'sheet:0',
        range: { kind: 'a1', a1: 'B1' } } });
    });
    fireEvent.keyDown(scroll, { key: 'x' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    expect(host.sent[4]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 1, 'x'] } });
  });

  for (const range of [false, true]) {
    it(`sheet tabs retain independent ${range ? 'reverse ranges' : 'active cells'} before and after acknowledgement`, async () => {
      const host = peerlessHarness();
      const view = await host.mount();
      const scroll = view.getByTestId('xlsx-scroll');
      const name = view.getByTestId('xlsx-name-box') as HTMLInputElement;
      const acknowledgeTab = async () => {
        const command = host.sent[host.sent.length - 1];
        if (command.kind !== 'setActiveSheet') throw new Error('Expected a sheet tab command');
        expect(command.payload.selection).toBeDefined();
        await act(async () => host.ack(command, { selection: command.payload.selection }));
      };
      await act(async () => {
        for (const key of ['ArrowRight', 'ArrowRight', 'ArrowDown', 'ArrowDown']) fireEvent.keyDown(scroll, { key });
      });
      if (range) {
        fireEvent.keyDown(scroll, { key: 'ArrowUp', shiftKey: true });
        fireEvent.keyDown(scroll, { key: 'ArrowLeft', shiftKey: true });
      }
      const first = host.sent[host.sent.length - 1].payload;
      await act(async () => { fireEvent.click(view.getByRole('tab', { name: 'First' })); });
      expect(name.value).toBe('C3');
      await acknowledgeTab();
      expect(name.value).toBe('C3');
      await act(async () => { fireEvent.click(view.getByRole('tab', { name: 'Second' })); });
      expect(name.value).toBe('A1');
      await acknowledgeTab();
      expect(name.value).toBe('A1');
      await act(async () => {
        for (const key of ['ArrowRight', 'ArrowDown', 'ArrowDown', 'ArrowDown']) fireEvent.keyDown(scroll, { key });
      });
      const second = host.sent[host.sent.length - 1].payload;
      await act(async () => { fireEvent.click(view.getByRole('tab', { name: 'First' })); });
      expect(name.value).toBe('C3');
      expect(host.sent[host.sent.length - 1]).toMatchObject({ kind: 'setActiveSheet', payload: first });
      await acknowledgeTab();
      expect(name.value).toBe('C3');
      await act(async () => { fireEvent.click(view.getByRole('tab', { name: 'Second' })); });
      expect(name.value).toBe('B4');
      expect(host.sent[host.sent.length - 1]).toMatchObject({ kind: 'setActiveSheet', payload: second });
      await acknowledgeTab();
      expect(name.value).toBe('B4');
    });
  }

  for (const first of [{ x: 0, y: 0 }, { x: 384, y: 240 }]) {
    it(`sheet tabs restore independent scroll positions from ${first.x},${first.y} without revealing the selection`, async () => {
      const host = peerlessHarness();
      const initialView = await host.sheetView(0);
      host.sheetView.mockImplementation(async (sheet) => ({ ...initialView, sheet,
        initialScrollX: 96, initialScrollY: 24,
      }));
      host.push({ throughSeq: 0, version: 'v1', statuses: [], selection: {
        sheetId: 'sheet:0', range: { kind: 'a1', a1: 'C3' },
      } });
      const view = await host.mount();
      const scroll = view.getByTestId('xlsx-scroll');
      const name = view.getByTestId('xlsx-name-box') as HTMLInputElement;
      const tab = async (label: string) => {
        await act(async () => fireEvent.click(view.getByRole('tab', { name: label })));
        const command = host.sent[host.sent.length - 1];
        if (command.kind !== 'setActiveSheet') throw new Error('Expected a sheet tab command');
        await act(async () => host.ack(command, { selection: command.payload.selection }));
      };
      expect(scroll.scrollLeft).toBe(96);
      expect(scroll.scrollTop).toBe(24);
      scroll.scrollLeft = first.x; scroll.scrollTop = first.y;
      await tab('Second');
      expect(scroll.scrollLeft).toBe(96);
      expect(scroll.scrollTop).toBe(24);
      scroll.scrollLeft = 192; scroll.scrollTop = 72;
      await tab('First');
      expect(name.value).toBe('C3');
      expect(scroll.scrollLeft).toBe(first.x);
      expect(scroll.scrollTop).toBe(first.y);
      await tab('Second');
      expect(scroll.scrollLeft).toBe(192);
      expect(scroll.scrollTop).toBe(72);
      expect(host.cellGeometry).not.toHaveBeenCalled();
    });
  }

  it('cross-sheet hyperlinks reveal their target even when the initial sheet view arrives later', async () => {
    const host = peerlessHarness();
    const frame = await host.frame({ x: 0, y: 0, width: 800, height: 600 }, { sheet: 0 });
    host.frame.mockResolvedValueOnce({ ...frame, displayList: { ...frame.displayList,
      hyperlinks: [{ top: 0, bottom: 0, left: 0, right: 0, location: 'Second!C3' }],
    } });
    const view = await host.mount();
    await advance();
    const scroll = view.getByTestId('xlsx-scroll');
    const nextView = deferred<xlsx.WorkbookSheetView>();
    host.sheetView.mockReturnValueOnce(nextView.promise);
    await act(async () => {
      fireEvent.mouseDown(scroll, { clientX: 12, clientY: 12, button: 0 });
      fireEvent.mouseUp(window, { clientX: 12, clientY: 12, button: 0 });
      fireEvent.click(scroll, { clientX: 12, clientY: 12 });
    });
    expect(host.cellGeometry).toHaveBeenLastCalledWith(1, 2, 2);
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('C3');
    expect(scroll.scrollLeft).toBe(192);
    expect(scroll.scrollTop).toBe(48);
    await act(async () => nextView.resolve({ sheet: 1, version: 'v1', contentWidth: 4000, contentHeight: 6000,
      frozenRows: 0, frozenCols: 0, frozenWidth: 0, frozenHeight: 0, initialScrollX: 960, initialScrollY: 240,
    }));
    expect(scroll.scrollLeft).toBe(192);
    expect(scroll.scrollTop).toBe(48);
  });

  for (const withSelection of [false, true]) {
    it(`sheet tabs keep selections by id after a reorder acknowledgement ${withSelection ? 'with' : 'without'} a selection`, async () => {
      const host = peerlessHarness();
      const view = await host.mount();
      const scroll = view.getByTestId('xlsx-scroll');
      const name = view.getByTestId('xlsx-name-box') as HTMLInputElement;
      fireEvent.keyDown(scroll, { key: 'ArrowRight' });
      scroll.scrollLeft = 288; scroll.scrollTop = 120;
      await act(async () => { fireEvent.click(view.getByRole('tab', { name: 'Second' })); });
      await act(async () => {
        for (const key of ['ArrowRight', 'ArrowRight', 'ArrowDown', 'ArrowDown']) fireEvent.keyDown(scroll, { key });
      });
      const selection: xlsxWorker.WorkbookCommandSelection = { sheetId: 'sheet:1', range: { kind: 'a1', a1: 'C3' } };
      await act(async () => host.ack(host.sent[host.sent.length - 1], { selection }));
      scroll.scrollLeft = 192; scroll.scrollTop = 72;
      const undoing = host.commands().execute('undo', null);
      await act(async () => {});
      const undo = host.sent[host.sent.length - 1];
      expect(undo.kind).toBe('undo');
      await act(async () => host.push({ throughSeq: undo.seq, version: 'v2',
        statuses: [{ seq: undo.seq, status: 'applied', committed: true }],
        editState: { activeSheet: 0, sheets: [
          { id: 'sheet:1', index: 0, name: 'Second' }, { id: 'sheet:0', index: 1, name: 'First' },
        ] }, ...(withSelection ? { selection } : {}),
      }));
      await undoing;
      expect(view.getByRole('tab', { name: 'Second' }).getAttribute('aria-selected')).toBe('true');
      expect(name.value).toBe('C3');
      expect(scroll.scrollLeft).toBe(192);
      expect(scroll.scrollTop).toBe(72);
      await act(async () => { fireEvent.click(view.getByRole('tab', { name: 'First' })); });
      expect(name.value).toBe('B1');
      expect(scroll.scrollLeft).toBe(288);
      expect(scroll.scrollTop).toBe(120);
      await act(async () => { fireEvent.click(view.getByRole('tab', { name: 'Second' })); });
      expect(name.value).toBe('C3');
      expect(scroll.scrollLeft).toBe(192);
      expect(scroll.scrollTop).toBe(72);
    });
  }

  it('sheet tabs retain selections restored by command acknowledgements', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.keyDown(scroll, { key: 'ArrowRight' });
    const undoing = host.commands().execute('undo', null);
    await act(async () => {});
    await act(async () => host.ack(host.sent[1], { selection: {
      sheetId: 'sheet:1', range: { kind: 'a1', a1: 'B4:D5' }, activeCell: { row: 3, col: 1 },
    } }));
    await undoing;
    await act(async () => { fireEvent.click(view.getByRole('tab', { name: 'First' })); });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('B1');
    await act(async () => { fireEvent.click(view.getByRole('tab', { name: 'Second' })); });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('B4');
    expect(host.sent[host.sent.length - 1]).toMatchObject({ payload: { selection: {
      sheetId: 'sheet:1', activeCell: { row: 3, col: 1 }, range: {
        kind: 'rowCol', start: { row: 3, col: 1 }, end: { row: 4, col: 3 },
      },
    } } });
  });

  it('opens with the session selection when it is known', async () => {
    const host = peerlessHarness();
    host.push({ throughSeq: 0, version: 'v1', statuses: [], selection: {
      sheetId: 'sheet:0', range: { kind: 'a1', a1: 'C3' },
    } });
    const view = await host.mount();
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('C3');
  });

  it('clears remembered sheet selections and scroll positions when the workbook is replaced', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
    const scroll = view.getByTestId('xlsx-scroll');
    scroll.scrollLeft = 288; scroll.scrollTop = 120;
    await act(async () => { fireEvent.click(view.getByRole('tab', { name: 'Second' })); });
    await act(async () => { fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowDown' }); });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A2');
    await act(async () => {});
    scroll.scrollLeft = 192; scroll.scrollTop = 72;
    await act(async () => { fireEvent.click(view.getByRole('tab', { name: 'First' })); });
    await act(async () => host.replace(view));
    await opened();
    await waitFor(() => expect(host.ready).toHaveBeenCalledTimes(2));
    await act(async () => fireEvent.click(view.getByRole('tab', { name: 'First' })));
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
    expect(scroll.scrollLeft).toBe(0);
    expect(scroll.scrollTop).toBe(0);
    await act(async () => fireEvent.click(view.getByRole('tab', { name: 'Second' })));
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
    expect(scroll.scrollLeft).toBe(0);
    expect(scroll.scrollTop).toBe(0);
  });

  it('rapid tab changes send both active-sheet intents before acknowledgement and Save', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    await act(async () => { fireEvent.click(view.getByRole('tab', { name: 'Second' })); });
    await act(async () => { fireEvent.click(view.getByRole('tab', { name: 'First' })); });
    expect(host.sent.map((command) => command.kind)).toEqual(['setActiveSheet', 'setActiveSheet']);
    expect(host.sent[1]).toMatchObject({ payload: { args: [0] } });
    await act(async () => { host.ack(host.sent[0]); host.ack(host.sent[1]); });
    expect(await host.commands().execute('save', null)).toMatchObject({ ok: true });
    expect(host.session.save).toHaveBeenCalledTimes(1);
    expect(view.getByRole('tab', { name: 'First' }).getAttribute('aria-selected')).toBe('true');
  });

  it('paste refreshes used sheet dimensions without shrinking keyboard limits after acknowledgement', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    host.sheetView.mockResolvedValueOnce({ sheet: 0, version: 'v2', contentWidth: 8000, contentHeight: 24000,
      frozenRows: 0, frozenCols: 0, frozenWidth: 0, frozenHeight: 0, initialScrollX: 0, initialScrollY: 0 });
    replaceClipboard('readText', mock(async () => Array.from({ length: 1000 }, () => 'value').join('\n')));
    fireEvent.keyDown(scroll, { key: 'v', ctrlKey: true });
    await act(async () => {});
    await act(async () => host.ack(host.sent[0], { paint: true }));
    expect(host.sheetView).toHaveBeenLastCalledWith(0);
    expect((scroll.firstElementChild as HTMLElement).style.height).toBe('24000px');
    expect((scroll.firstElementChild as HTMLElement).style.width).toBe('8000px');
    await act(async () => { fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true }); });
    expect(host.selectionData).toHaveBeenLastCalledWith(0, 0, 0, 'down');
    expect(host.sent[1]).toMatchObject({ kind: 'select', payload: { selection: {
      activeCell: { row: 1_048_575, col: 0 },
    } } });
  });

  it('a late geometry read cannot overwrite a newer content acknowledgement', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const first = deferred<xlsx.WorkbookSheetView>();
    const second = deferred<xlsx.WorkbookSheetView>();
    host.sheetView.mockReturnValueOnce(first.promise).mockReturnValueOnce(second.promise);
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.keyDown(scroll, { key: 'a' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await act(async () => host.ack(host.sent[0], { paint: true }));
    fireEvent.keyDown(scroll, { key: 'b' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await act(async () => host.ack(host.sent[1], { paint: true }));
    const geometry = { sheet: 0, version: 'v2', contentWidth: 8000, contentHeight: 48000,
      frozenRows: 0, frozenCols: 0, frozenWidth: 0, frozenHeight: 0, initialScrollX: 0, initialScrollY: 0 };
    await act(async () => second.resolve(geometry));
    await act(async () => first.resolve({ ...geometry, version: 'v1', contentHeight: 24000 }));
    expect((scroll.firstElementChild as HTMLElement).style.height).toBe('48000px');
  });

  it('paste echoes the visible tail after an offscreen prefix through editor input', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    scroll.scrollTop = 999 * 24;
    fireEvent.scroll(scroll);
    await advance();
    replaceClipboard('readText', mock(async () => Array.from({ length: 1000 }, (_, row) => `pasted ${row + 1}`).join('\n')));
    fireEvent.keyDown(scroll, { key: 'v', ctrlKey: true });
    await act(async () => {});
    expect(host.sent[0]).toMatchObject({ kind: 'editCells' });
    expect(painted[painted.length - 1].commands.some((cmd) => cmd.op === 'text' && cmd.text === 'pasted 1000')).toBe(true);
  });

  it('large TSV paste submits through the editor without overflowing the argument stack', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    replaceClipboard('readText', mock(async () => Array.from({ length: 100_000 }, () => 'x').join('\n')));
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
    await act(async () => {});
    const command = host.sent[0];
    expect(command).toMatchObject({ kind: 'editCells', payload: { selection: {
      activeCell: { row: 0, col: 0 },
    } } });
    if (command.kind !== 'editCells') throw new Error('Expected pasted cells');
    expect(command.payload.args[1]).toHaveLength(100_000);
    expect(command.payload.args[1][99_999]).toEqual({ row: 99_999, col: 0, input: 'x' });
    expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
  });

  it('refuses TSV paste above the cell limit without changing the source', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    replaceClipboard('readText', mock(async () => Array.from({ length: 100_001 }, () => 'x').join('\n')));
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
    await act(async () => {});
    expect(view.getByTestId('xlsx-input-refusal').textContent).toContain('100000-cell limit');
    expect(host.sent).toEqual([]);
    expect([...host.cells]).toEqual([['0:0:0', 'initial']]);
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
  });

  it('refuses rebased paste of an array snapshot before submitting a mutation', async () => {
    const host = peerlessHarness();
    const clipboard = markedClipboard();
    const view = await mountPeerlessClipboard(host);
    const cell = { a1: 'A1', input: '=SEQUENCE(2)', isFormula: true, isArrayFormula: true };
    host.cellInputs.mockResolvedValueOnce({ sheet: 0, version: 'v1', cells: [[cell]] });
    const scroll = view.getByTestId('xlsx-scroll');
    await act(async () => fireEvent.keyDown(scroll, { key: 'c', ctrlKey: true }));
    await act(async () => { await clipboard.writes[0]; });
    await act(async () => fireEvent.keyDown(scroll, { key: 'v', ctrlKey: true }));
    expect(host.sent).toEqual([]);
    expect(view.getByTestId('xlsx-input-refusal').textContent).toMatch(/array or spill/i);
    expect(host.cells.get('0:0:0')).toBe('initial');
  });

  it('pastes a whole-sheet copy as the used range it copied', async () => {
    const host = peerlessHarness();
    const clipboard = markedClipboard();
    const view = await mountPeerlessClipboard(host);
    host.selectionData.mockResolvedValue({ edge: null, region: null, usedRange: { top: 0, left: 0, bottom: 0, right: 0 } });
    const scroll = view.getByTestId('xlsx-scroll');
    let acked = 0;
    const acknowledge = () => act(async () => { for (const command of host.sent.slice(acked)) host.ack(command); acked = host.sent.length; });
    await act(async () => fireEvent.keyDown(scroll, { key: 'a', ctrlKey: true }));
    await acknowledge();
    await act(async () => fireEvent.keyDown(scroll, { key: 'c', ctrlKey: true }));
    await act(async () => { await clipboard.writes[0]; });
    expect(host.cellInputs).toHaveBeenLastCalledWith(0, 'A1:A1');
    await act(async () => fireEvent.keyDown(scroll, { key: 'ArrowDown' }));
    await acknowledge();
    await act(async () => fireEvent.keyDown(scroll, { key: 'v', ctrlKey: true }));
    await act(async () => {});
    expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
    expect(host.sent[host.sent.length - 1]).toMatchObject({ kind: 'editCells', payload: {
      args: [0, [{ row: 1, col: 0, input: 'initial' }]],
    } });
  });

  it('read-only proposal panel loads proposals through a worker read command', async () => {
    const host = peerlessHarness();
    const available = spyOn(xlsx, 'isProposalsAvailable').mockReturnValue(true);
    restorers.push(() => available.mockRestore());
    const view = await host.mount(undefined, true);
    const opening = host.commands().execute('proposalsPanel', { open: true });
    await act(async () => {});
    expect(await opening).toMatchObject({ ok: true });
    expect(host.sent[0]).toMatchObject({ kind: 'applyOps', payload: { toolbar: { id: 'refreshProposals' } } });
    await act(async () => host.ack(host.sent[0], { result: { proposals: [
      { id: 'review', agentId: 'Reviewer', note: 'Pending review', cells: [] },
    ] } }));
    expect(view.getByText('Pending review')).toBeDefined();
    expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
  });

  it('mouse-confirmed cell composition permits Tab after acknowledgement and paint', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    const input = view.getByTestId('xlsx-cell-editor');
    await waitFor(() => expect((input as HTMLInputElement).value).toBe('initial'));
    fireEvent.compositionStart(input);
    fireEvent.change(input, { target: { value: '確定' } });
    fireEvent.compositionEnd(input);
    await act(async () => {});
    await act(async () => host.ack(host.sent[0], { paint: true }));
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Tab' });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('B1');
    expect(host.sent.map((command) => command.kind)).toEqual(['editCell', 'select']);
  });

  for (const [id, testId, reason] of [
    ['searchMenus', 'xlsx-search-menus', 'unsupported-command'],
    ['exportPng', 'xlsx-export-png', 'png-unavailable'],
  ] as const) {
    it(`${id} is ${reason === 'unsupported-command' ? 'hidden' : 'disabled'} and sends no worker command`, async () => {
      const host = peerlessHarness(true);
      const view = await host.mount();
      if (reason === 'unsupported-command') expect(view.queryByTestId(testId)).toBeNull();
      else {
        const button = view.getByTestId(testId) as HTMLButtonElement;
        expect(button.disabled).toBe(true);
        fireEvent.click(button);
      }
      expect(host.commands().getState(id)).toMatchObject({ enabled: false, disabledReason: { code: reason } });
      expect(await host.commands().execute(id, null)).toMatchObject({ ok: false, failure: { code: reason } });
      expect(host.sent).toEqual([]);
    });
  }

  it('puts Save and Export PNG in a File group after history, not in the formula row', async () => {
    const host = peerlessHarness(false, false, true);
    const view = await host.mount();
    const file = view.getByTestId('xlsx-save').closest('[role="group"]')!;
    expect(file.getAttribute('aria-label')).toBe('File');
    expect(file.contains(view.getByTestId('xlsx-export-png'))).toBe(true);
    const toolbar = view.getByTestId('xlsx-toolbar');
    const groups = [...toolbar.querySelectorAll('[role="group"]')].map((group) => group.getAttribute('aria-label'));
    expect(groups.indexOf('File')).toBe(groups.indexOf('History') + 1);
    expect(groups).not.toContain('Search');
    expect(file.contains(view.getByTestId('xlsx-name-box'))).toBe(false);
  });

  for (const action of ['Enter', 'Tab', 'paste', 'delete', 'clear'] as const) {
    it(`${action} paints through command reply`, async () => {
      const host = peerlessHarness(false, action === 'paste');
      const view = await host.mount();
      const frames = host.frame.mock.calls.length;
      if (action === 'Enter' || action === 'Tab') {
        fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
        fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'new value' } });
        fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: action });
        expect(host.sent).toHaveLength(1);
        expect(host.sent[0]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'new value'] } });
        expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe(action === 'Enter' ? 'A2' : 'B1');
        expect(painted[painted.length - 1]?.commands.some((cmd) => cmd.op === 'text' && cmd.text === 'new value')).toBe(true);
        expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
      } else if (action === 'paste') {
        replaceClipboard('readText', mock(async () => 'one\ttwo'));
        fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
        await act(async () => {});
        expect(host.sent[0]).toMatchObject({ kind: 'editCells' });
        expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
      } else {
        fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: action === 'delete' ? 'Delete' : 'Backspace' });
        await act(async () => {});
        expect(host.sent[0]).toMatchObject({ kind: 'applyOps', payload: { toolbar: {
          id: 'clear', target: { sheetId: 'sheet:0', range: { kind: 'rowCol' } },
        } } });
      }
      const command = host.sent[0];
      await act(async () => host.ack(command, { paint: true }));
      if (action === 'paste') {
        expect(host.cells.get('0:0:0')).toBe('one');
        expect(host.cells.get('0:0:1')).toBe('two');
        expect(painted[painted.length - 1].commands.flatMap((cmd) => cmd.op === 'text' ? [cmd.text] : [])).toEqual(['one', 'two']);
      }
      expect(host.session.markPresented).toHaveBeenLastCalledWith({ ...host.session.commandIdentity, coversThroughSeq: command.seq });
      expect(host.frame.mock.calls.length).toBe(frames);
      expect(host.hydrate).not.toHaveBeenCalled();
      expect(host.attach).not.toHaveBeenCalled();
    });
  }

  it('restores refused original target', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'r' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
    const newer = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    await act(async () => host.ack(host.sent[0], { status: 'refused' }));
    expect(newer.value).toBe('n');
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A2');
    fireEvent.keyDown(newer, { key: 'Escape' });
    fireEvent.click(view.getByTestId('xlsx-refused-draft'));
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('r');
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
    expect(painted[painted.length - 1]?.commands.some((cmd) => cmd.op === 'text' && cmd.text === 'r')).toBe(false);
  });

  it('stale proposal refusal preserves independent navigation and typed text', async () => {
    const host = peerlessHarness();
    const available = spyOn(xlsx, 'isProposalsAvailable').mockReturnValue(true);
    restorers.push(() => available.mockRestore());
    const view = await host.mount();
    await host.commands().execute('proposalsPanel', { open: true });
    const refresh = await host.commandAt(0);
    await act(async () => host.ack(refresh, {
      result: { proposals: [{ id: 'p1', agentId: 'agent', note: null, cells: [] }] },
    }));
    const accepting = host.commands().execute('proposalAccept', { proposalId: 'p1' });
    await act(async () => {});
    const proposal = host.sent[1];
    expect(proposal.kind).toBe('acceptProposal');
    const scroll = view.getByTestId('xlsx-scroll');
    await act(async () => { fireEvent.keyDown(scroll, { key: 'ArrowRight' }); });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('B1');
    await act(async () => fireEvent.keyDown(scroll, { key: 't' }));
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'kept text' } });
    await act(async () => fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' }));
    const selected = host.sent[2];
    const edited = host.sent[3];
    expect(selected).toMatchObject({ kind: 'select' });
    expect(selected.afterSeq).toBeUndefined();
    expect(edited).toMatchObject({ kind: 'editCell', afterSeq: selected.seq,
      payload: { args: [0, 0, 1, 'kept text'] } });
    await act(async () => host.push({ throughSeq: proposal.seq, version: 'v1', statuses: [{
      seq: proposal.seq, status: 'refused', committed: false, reason: 'stale-proposal', result: { cells: ['A1'] },
    }] }));
    expect(await accepting).toMatchObject({ ok: false, failure: { code: 'proposal-stale' } });
    await act(async () => {
      host.ack(selected, { selection: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B1' } } });
      host.ack(edited, { paint: true, selection: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B2' } } });
    });
    expect(host.cells.get('0:0:1')).toBe('kept text');
    expect(painted[painted.length - 1].commands.some((command) => command.op === 'text' && command.text === 'kept text')).toBe(true);
    expect(view.queryByTestId('xlsx-refused-draft')).toBeNull();
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('B2');
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
  });

  it('earlier Enter acknowledgement preserves unsent clipboard-blocked navigation', async () => {
    const host = peerlessHarness();
    const capture = deferred<string>();
    replaceClipboard('readText', mock(() => capture.promise));
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.keyDown(scroll, { key: 'a' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    const first = host.sent[0];
    fireEvent.keyDown(scroll, { key: 'v', ctrlKey: true });
    fireEvent.keyDown(scroll, { key: 'ArrowRight' });
    expect(host.sent).toHaveLength(1);
    await act(async () => host.ack(first, { selection: {
      sheetId: 'sheet:0', range: { kind: 'rowCol', start: { row: 1, col: 0 }, end: { row: 1, col: 0 } },
    } }));
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('B2');
    await act(async () => { fireEvent.keyDown(scroll, { key: 'b' }); });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await act(async () => capture.resolve('paste'));
    expect(host.sent.find((command) => command.kind === 'editCell' && command.seq !== first.seq))
      .toMatchObject({ payload: { args: [0, 1, 1, 'b'] } });
  });

  for (const kind of ['rowCol', 'a1'] as const) {
    it(`Undo restores ${kind} selection and presents its frame before Print`, async () => {
      const host = peerlessHarness();
      const view = await host.mount();
      const scroll = view.getByTestId('xlsx-scroll');
      fireEvent.keyDown(scroll, { key: 'ArrowRight' });
      await act(async () => host.ack(host.sent[0], { paint: true, selection: {
        sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B1' },
      } }));
      fireEvent.keyDown(scroll, { key: 'ArrowLeft' });
      await act(async () => host.ack(host.sent[1], { paint: true, selection: {
        sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' },
      } }));
      fireEvent.keyDown(scroll, { key: 'n' });
      fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
      await act(async () => host.ack(host.sent[2], { paint: true, selection: {
        sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A2' },
      } }));
      const undoing = host.commands().execute('undo', null);
      await act(async () => {});
      const undo = host.sent[3];
      await act(async () => host.ack(undo, { paint: true, selection: {
        sheetId: 'sheet:0', range: kind === 'a1' ? { kind, a1: 'A1' } :
          { kind, start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
      } }));
      await undoing;
      expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
      const print = mock(() => {});
      replacePrint(print);
      await host.commands().execute('print', null);
      expect(print).toHaveBeenCalledTimes(1);
      expect(host.session.presentedThroughSeq).toBe(undo.seq);
      fireEvent.keyDown(scroll, { key: 'r' });
      fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
      expect(host.sent[4]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'r'] } });
    });
  }

  it('acknowledgement preserves a reverse selection rectangle and active cell', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.keyDown(scroll, { key: 'ArrowDown' });
    fireEvent.keyDown(scroll, { key: 'ArrowDown' });
    fireEvent.keyDown(scroll, { key: 'ArrowUp', shiftKey: true });
    fireEvent.keyDown(scroll, { key: 'ArrowUp', shiftKey: true });
    const selection = { sheetId: 'sheet:0', activeCell: { row: 2, col: 0 }, range: {
      kind: 'rowCol' as const, start: { row: 0, col: 0 }, end: { row: 2, col: 0 },
    } };
    await act(async () => host.ack(host.sent[3], { paint: true, selection }));
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A3');
    expect(view.getByTestId('xlsx-selection').style.height).toBe('72px');
    fireEvent.keyDown(scroll, { key: 'Delete' });
    await act(async () => {});
    expect(host.sent[4]).toMatchObject({ payload: { toolbar: { target: selection } } });
  });

  it('Escape discards the restored refusal before independent navigation formatting and Save', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.keyDown(scroll, { key: 'r' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await act(async () => host.ack(host.sent[0], { status: 'refused' }));
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('r');
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Escape' });
    fireEvent.keyDown(scroll, { key: 'ArrowRight' });
    const formatting = host.commands().execute('bold', null);
    await act(async () => {});
    expect(host.sent.map((command) => command.kind)).toEqual(['editCell', 'select', 'applyOps']);
    await act(async () => { host.ack(host.sent[1]); host.ack(host.sent[2]); });
    expect(await formatting).toMatchObject({ ok: true });
    expect(await host.commands().execute('save', null)).toMatchObject({ ok: true });
    expect(host.session.save).toHaveBeenCalledTimes(1);
  });

  it('keyboard reveals an offscreen destination before opening and committing its native input', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    const frame = host.frame.getMockImplementation()!;
    host.frame.mockImplementation(async (viewport, options) => {
      const next = await frame(viewport, options);
      const start = Math.floor(viewport.y / 24);
      const offset = start * 24 - viewport.y;
      const count = Math.ceil((viewport.height - offset) / 24);
      next.displayList.grid = { ...next.displayList.grid!, startRow: start,
        rowIndices: Array.from({ length: count }, (_, index) => start + index),
        rowOffsets: Array.from({ length: count + 1 }, (_, index) => offset + index * 24) };
      return next;
    });
    const geometry = deferred<xlsx.WorkbookCellGeometry>();
    host.cellGeometry.mockReturnValueOnce(geometry.promise);
    host.selectionData.mockResolvedValueOnce({ edge: { row: 250, col: 0 } });
    await act(async () => { fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true }); });
    const selected = host.sent[0];
    await act(async () => host.ack(selected, { paint: true, selection: {
      sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A251' },
    } }));
    fireEvent.keyDown(scroll, { key: 'z' });
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    await act(async () => geometry.resolve({ sheet: 0, version: 'v1',
      rect: { x: 0, y: 6000, w: 96, h: 24 }, scrollPosition: { x: 0, y: 6000 } }));
    await advance();
    expect(scroll.scrollTop).toBe(6024 - 600 + 2);
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('z');
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    expect(host.sent[1]).toMatchObject({ kind: 'editCell', payload: { args: [0, 250, 0, 'z'] } });
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
  });

  it('replacement cancels retained chart nudges even with the same chart id', async () => {
    const host = peerlessHarness();
    host.charts.push({ id: 'chart', label: 'Chart', movable: true, rect: { x: 200, y: 100, w: 100, h: 100 },
      clip: { x: 200, y: 100, w: 100, h: 100 } });
    const view = await host.mount();
    const original = globalThis.setTimeout;
    let callback: (() => void) | undefined;
    const timer = spyOn(globalThis, 'setTimeout').mockImplementation(((handler: () => void, delay?: number) => {
      if (delay === 250) { callback = handler; return original(() => {}, 0); }
      return original(handler, delay);
    }) as typeof setTimeout);
    restorers.push(() => timer.mockRestore());
    fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { clientX: 220, clientY: 120, button: 0 });
    fireEvent.mouseUp(window, { clientX: 220, clientY: 120, button: 0 });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
    expect(callback).toBeDefined();
    timer.mockRestore();
    host.open.mockResolvedValueOnce({ ...host.session, commandIdentity: { session: 'replacement', incarnation: 'worker' } });
    await act(async () => host.replace(view));
    await opened();
    await waitFor(() => expect(host.ready).toHaveBeenCalledTimes(2));
    await act(async () => callback!());
    expect(host.sent.filter((command) => command.kind === 'moveChart')).toEqual([]);
    expect(view.queryByTestId('xlsx-chart-selection')).toBeNull();
  });

  it('Escape cancels pending offscreen keyboard input before its frame arrives', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    const geometry = deferred<xlsx.WorkbookCellGeometry>();
    host.cellGeometry.mockReturnValueOnce(geometry.promise);
    host.selectionData.mockResolvedValueOnce({ edge: { row: 250, col: 0 } });
    await act(async () => { fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true }); });
    fireEvent.keyDown(scroll, { key: 'z' });
    fireEvent.keyDown(scroll, { key: 'Escape' });
    await act(async () => {
      host.ack(host.sent[0]);
      geometry.resolve({ sheet: 0, version: 'v1', rect: { x: 0, y: 6000, w: 96, h: 24 },
        scrollPosition: { x: 0, y: 6000 } });
    });
    await advance();
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    expect(host.sent.map((command) => command.kind)).toEqual(['select']);
  });

  it('Enter commits offscreen keyboard input while its geometry is pending', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    const geometry = deferred<xlsx.WorkbookCellGeometry>();
    host.cellGeometry.mockReturnValueOnce(geometry.promise);
    host.selectionData.mockResolvedValueOnce({ edge: { row: 250, col: 0 } });
    await act(async () => { fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true }); });
    fireEvent.keyDown(scroll, { key: 'z' });
    fireEvent.keyDown(scroll, { key: 'Enter' });
    expect(host.sent[1]).toMatchObject({ kind: 'editCell', payload: { args: [0, 250, 0, 'z'] } });
    await act(async () => {
      host.ack(host.sent[0]); host.ack(host.sent[1]);
      geometry.resolve({ sheet: 0, version: 'v1', rect: { x: 0, y: 6000, w: 96, h: 24 },
        scrollPosition: { x: 0, y: 6000 } });
    });
    await advance();
    expect(host.cells.get('0:250:0')).toBe('z');
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
  });

  it('Undo requests current coverage when its restored sheet invalidates the attached frame', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await act(async () => host.ack(host.sent[0], { paint: true }));
    const undoing = host.commands().execute('undo', null);
    await act(async () => {});
    const undo = host.sent[1];
    await act(async () => host.ack(undo, { paint: true, selection: {
      sheetId: 'sheet:1', range: { kind: 'a1', a1: 'A1' },
    } }));
    await undoing;
    const print = mock(() => {});
    replacePrint(print);
    const printing = host.commands().execute('print', null);
    expect(print).not.toHaveBeenCalled();
    await advance();
    await printing;
    expect(host.frame.mock.calls[host.frame.mock.calls.length - 1][1]?.sheet).toBe(1);
    expect(host.session.presentedThroughSeq).toBe(undo.seq);
    expect(print).toHaveBeenCalledTimes(1);
  });

  it('refusal cancels predicted dependents and restores the first target', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'a' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Tab' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'b' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Tab' });
    expect(host.sent).toHaveLength(2);
    expect(host.sent[1].afterSeq).toBe(host.sent[0].seq);
    await act(async () => host.push({ throughSeq: 2, version: 'v1', statuses: [
      { seq: 1, status: 'refused', reason: 'invalid-step', affectedSeqs: [2] },
      { seq: 2, status: 'cancelled', reason: 'dependency-refused' },
    ] }));
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('a');
    expect(painted[painted.length - 1]?.commands.some((cmd) => cmd.op === 'text' && (cmd.text === 'a' || cmd.text === 'b'))).toBe(false);
  });

  for (const outcome of ['success', 'refusal'] as const) {
    it(`cut waits for clipboard write ${outcome}`, async () => {
      const host = peerlessHarness();
      const clipboard = promisedClipboard();
      const writing = deferred<void>();
      clipboard.write.mockImplementation((items) => {
        clipboard.payloads.push(items[0].data['text/plain']);
        return writing.promise;
      });
      const view = await host.mount();
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true });
      await act(async () => {});
      expect(host.sent).toEqual([]);
      expect(await clipboard.payloads[0].then((blob) => blob.text())).toBe('initial');
      await act(async () => {
        if (outcome === 'success') writing.resolve();
        else writing.reject(new Error('Clipboard denied'));
      });
      if (outcome === 'success') expect(host.sent[0]).toMatchObject({ kind: 'applyOps', payload: { toolbar: { id: 'clear' } } });
      else {
        expect(host.sent).toEqual([]);
        expect(host.cells.get('0:0:0')).toBe('initial');
      }
    });
  }

  for (const outcome of ['success', 'refusal'] as const) {
    it(`cut without ClipboardItem preserves its target after writeText ${outcome}`, async () => {
      const host = peerlessHarness();
      const clipboard = textClipboard();
      const view = await mountPeerlessClipboard(host);
      const scroll = view.getByTestId('xlsx-scroll');
      const reading = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
      host.cellInputs.mockReturnValueOnce(reading.promise);
      fireEvent.keyDown(scroll, { key: 'x', ctrlKey: true });
      await act(async () => {});
      expect(clipboardReads(host)).toEqual([[0, 'A1:A1']]);
      expect(clipboard.writeText).not.toHaveBeenCalled();
      expect(host.sent).toEqual([]);
      await act(async () => reading.resolve({ sheet: 0, version: 'v1', cells: [
        [{ a1: 'A1', input: 'original\tvalue', isFormula: false }],
      ] }));
      expect(clipboard.writeText).toHaveBeenCalledTimes(1);
      expect(clipboard.writeText).toHaveBeenCalledWith('"original\tvalue"');
      expect(clipboard.write).not.toHaveBeenCalled();
      await act(async () => { fireEvent.keyDown(scroll, { key: 'ArrowRight' }); });
      expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('B1');
      await act(async () => fireEvent.keyDown(scroll, { key: 'n' }));
      fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
      await act(async () => {});
      expect(host.sent).toEqual([]);
      await act(async () => {
        if (outcome === 'success') clipboard.writing.resolve();
        else clipboard.writing.reject(new DOMException('Clipboard denied', 'NotAllowedError'));
      });
      const clears = host.sent.filter((command) => command.kind === 'applyOps');
      const edit = host.sent.find((command) => command.kind === 'editCell')!;
      if (outcome === 'success') {
        expect(clears).toHaveLength(1);
        const target = { sheetId: 'sheet:0', range: {
          kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 0, col: 0 },
        } };
        expect(clears[0]).toMatchObject({ payload: { toolbar: { id: 'clear', target }, selection: target } });
        expect(host.sent[0]).toBe(clears[0]);
        expect(host.sent.map((command) => command.kind)).toEqual(['applyOps', 'select', 'editCell', 'select']);
        const select = host.sent.find((command) => command.kind === 'select')!;
        expect(select.seq).toBeGreaterThan(clears[0].seq);
        expect(edit.seq).toBeGreaterThan(select.seq);
      } else {
        expect(clears).toEqual([]);
        expect(host.sent.map((command) => command.kind)).toEqual(['editCell', 'select']);
        expect(host.cells.get('0:0:0')).toBe('initial');
        expect(view.getByTestId('xlsx-input-refusal').textContent).toMatch(/clipboard denied/i);
        expect(view.queryByTestId('xlsx-error')).toBeNull();
      }
      expect(host.sent[host.sent.length - 1]).toMatchObject({ kind: 'select', payload: { selection: { activeCell: { row: 1, col: 1 } } } });
      expect(edit).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 1, 'n'] } });
      await act(async () => { for (const command of host.sent.slice()) host.ack(command); });
      expect(host.cells.get('0:0:1')).toBe('n');
      await act(async () => fireEvent.keyDown(scroll, { key: 'r' }));
      fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
      await act(async () => {});
      const next = host.sent.filter((command) => command.kind === 'editCell').slice(-1)[0]!;
      expect(next).toMatchObject({ kind: 'editCell', payload: { args: [0, 1, 1, 'r'] } });
      await act(async () => host.ack(next));
      expect(host.cells.get('0:1:1')).toBe('r');
      expect(clipboardReads(host)).toEqual([[0, 'A1:A1']]);
      expect(host.sent.filter((command) => command.kind === 'applyOps')).toHaveLength(outcome === 'success' ? 1 : 0);
      expect(host.hydrate).not.toHaveBeenCalled();
      expect(host.attach).not.toHaveBeenCalled();
    });
  }

  it('cut behind a pending paste leaves newer input after the clear', async () => {
    const host = peerlessHarness(false, true);
    const clipboard = textClipboard();
    const pasting = deferred<string>();
    replaceClipboard('readText', mock(() => pasting.promise));
    const view = await mountPeerlessClipboard(host);
    const scroll = view.getByTestId('xlsx-scroll');
    const formula = view.getByTestId('xlsx-formula-input');
    fireEvent.keyDown(scroll, { key: 'v', ctrlKey: true });
    fireEvent.change(formula, { target: { value: 'before' } });
    fireEvent.keyDown(scroll, { key: 'x', ctrlKey: true });
    fireEvent.change(formula, { target: { value: 'after' } });
    expect(host.sent).toEqual([]);
    expect(clipboardReads(host)).toEqual([]);
    await act(async () => pasting.resolve('pasted'));
    expect(host.sent.map((command) => command.kind)).toEqual(['editCells', 'editCell']);
    expect(clipboard.writeText).not.toHaveBeenCalled();
    await act(async () => host.ack(host.sent[0]));
    expect(host.sent[1]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'before'] } });
    await act(async () => host.ack(host.sent[1]));
    expect(host.sent.map((command) => command.kind)).toEqual(['editCells', 'editCell']);
    expect(clipboard.writeText).toHaveBeenCalledWith('before');
    expect((formula as HTMLInputElement).value).toBe('after');
    fireEvent.keyDown(formula, { key: 'Enter' });
    expect(host.sent).toHaveLength(2);
    await act(async () => clipboard.writing.resolve());
    expect(host.sent.map((command) => command.kind)).toEqual(['editCells', 'editCell', 'applyOps', 'editCell', 'select']);
    expect(host.sent[4]).toMatchObject({ kind: 'select', payload: { selection: { activeCell: { row: 1, col: 0 } } } });
    const target = { sheetId: 'sheet:0', range: {
      kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 0, col: 0 },
    } };
    expect(host.sent[2]).toMatchObject({ payload: { toolbar: { id: 'clear', target }, selection: target } });
    expect(host.sent[3]).toMatchObject({ payload: { args: [0, 0, 0, 'after'] } });
    expect(host.sent[3].seq).toBeGreaterThan(host.sent[2].seq);
    await act(async () => host.ack(host.sent[2]));
    expect(host.cells.get('0:0:0')).toBe('');
    await act(async () => host.ack(host.sent[3]));
    expect(host.cells.get('0:0:0')).toBe('after');
    expect(clipboardReads(host)).toEqual([[0, 'A1:A1']]);
    expect(clipboard.writeText).toHaveBeenCalledTimes(1);
    expect(clipboard.write).not.toHaveBeenCalled();
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
  });

  for (const mode of ['ClipboardItem', 'writeText'] as const) {
    for (const outcome of ['success', 'refusal'] as const) {
      it(`${mode} copy releases later edits and preserves captured inputs on ${outcome}`, async () => {
        const host = peerlessHarness();
        const completion = deferred<void>();
        const clipboard = mode === 'ClipboardItem' ? promisedClipboard(completion.promise) : textClipboard();
        const writing = 'writing' in clipboard ? clipboard.writing : completion;
        const view = await mountPeerlessClipboard(host);
        const scroll = view.getByTestId('xlsx-scroll');
        const reading = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
        host.cells.set('0:0:0', 'line\nbreak');
        let captured!: Awaited<ReturnType<typeof host.cellInputs>>;
        host.cellInputs.mockImplementationOnce((sheet, a1) => {
          captured = { sheet, version: 'v1', cells: [[{ a1, input: host.cells.get('0:0:0')!, isFormula: false }]] };
          return reading.promise;
        });
        await act(async () => fireEvent.keyDown(scroll, { key: 'c', ctrlKey: true }));
        expect(clipboardReads(host)).toEqual([[0, 'A1:A1']]);
        if ('writeText' in clipboard) {
          expect(clipboard.writeText).not.toHaveBeenCalled();
          expect(clipboard.write).not.toHaveBeenCalled();
        } else {
          expect(clipboard.write).toHaveBeenCalledTimes(1);
          expect(clipboard.copied).toEqual([]);
        }
        const formula = view.getByTestId('xlsx-formula-input');
        fireEvent.change(formula, { target: { value: 'later' } });
        await act(async () => fireEvent.keyDown(formula, { key: 'Enter' }));
        expect(host.sent).toHaveLength(1);
        const edit = host.sent.find((command) => command.kind === 'editCell')!;
        expect(edit).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'later'] } });
        expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).not.toBe('A1');
        await act(async () => host.ack(edit));
        expect(host.cells.get('0:0:0')).toBe('later');
        if ('writeText' in clipboard) expect(clipboard.writeText).not.toHaveBeenCalled();
        else expect(clipboard.copied).toEqual([]);
        await act(async () => reading.resolve(captured));
        if ('writeText' in clipboard) {
          expect(clipboard.writeText).toHaveBeenCalledTimes(1);
          expect(clipboard.writeText).toHaveBeenCalledWith('"line\nbreak"');
          expect(clipboard.write).not.toHaveBeenCalled();
        } else expect(await clipboard.payloads[0].then((blob) => blob.text())).toBe('"line\nbreak"');
        await act(async () => {
          if (outcome === 'success') writing.resolve();
          else writing.reject(new DOMException('Clipboard denied', 'NotAllowedError'));
        });
        if ('writes' in clipboard) {
          if (outcome === 'success') await clipboard.writes[0];
          else await expect(clipboard.writes[0]).rejects.toMatchObject({ name: 'NotAllowedError' });
          expect(clipboard.copied).toEqual(['"line\nbreak"']);
        }
        if (outcome === 'refusal') expect(view.getByTestId('xlsx-input-refusal').textContent).toMatch(/clipboard denied/i);
        expect(view.queryByTestId('xlsx-error')).toBeNull();
        expect(host.sent).toHaveLength(1);
        fireEvent.change(formula, { target: { value: 'next' } });
        await act(async () => fireEvent.keyDown(formula, { key: 'Enter' }));
        expect(host.sent.map((command) => command.kind)).toEqual(['editCell', 'editCell']);
        const next = host.sent.find((command) => command.kind === 'editCell' && command.seq > edit.seq)!;
        expect(next).toMatchObject({ payload: { args: [0, 1, 0, 'next'] } });
        expect(next.seq).toBeGreaterThan(edit.seq);
        await act(async () => host.ack(next));
        expect(host.cells.get('0:1:0')).toBe('next');
        expect(clipboardReads(host)).toEqual([[0, 'A1:A1']]);
        expect(host.hydrate).not.toHaveBeenCalled();
        expect(host.attach).not.toHaveBeenCalled();
      });
    }
  }

  it('cut refuses missing clipboard write access before reading or sealing input', async () => {
    const host = peerlessHarness();
    const clipboard = textClipboard();
    replaceClipboard('writeText', undefined);
    const view = await mountPeerlessClipboard(host);
    const formula = view.getByTestId('xlsx-formula-input');
    fireEvent.change(formula, { target: { value: 'draft' } });
    await act(async () => fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true }));
    expect(view.getByTestId('xlsx-input-refusal').textContent).toMatch(/clipboard write access/i);
    expect(view.queryByTestId('xlsx-error')).toBeNull();
    expect(clipboardReads(host)).toEqual([]);
    expect(clipboard.write).not.toHaveBeenCalled();
    expect(host.sent).toEqual([]);
    expect(host.cells.get('0:0:0')).toBe('initial');
    expect((formula as HTMLInputElement).value).toBe('draft');
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
  });

  for (const ending of ['replacement', 'failure'] as const) {
    it(`cut without ClipboardItem never writes after ${ending} at read fulfillment`, async () => {
      const host = peerlessHarness();
      const clipboard = textClipboard();
      const fulfilled = deferred<void>();
      const create = peerlessInput.createPeerlessWorkerInputCoordinator;
      let onRead: (() => void) | undefined;
      const observed = spyOn(peerlessInput, 'createPeerlessWorkerInputCoordinator').mockImplementation((hooks) => {
        const input = create(hooks);
        const runAfterPendingInput = input.runAfterPendingInput;
        input.runAfterPendingInput = (operation, fence) => {
          const notify = onRead;
          onRead = undefined;
          const result = runAfterPendingInput(operation, fence);
          if (notify) void result.then(() => { notify(); fulfilled.resolve(); }).catch(fulfilled.reject);
          return result;
        };
        return input;
      });
      restorers.push(() => observed.mockRestore());
      const view = await mountPeerlessClipboard(host);
      const reading = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
      host.cellInputs.mockReturnValueOnce(reading.promise);
      onRead = () => {
        expect(clipboardReads(host)).toEqual([[0, 'A1:A1']]);
        expect(clipboard.writeText).not.toHaveBeenCalled();
        if (ending === 'replacement') flushSync(() => view.rerender(host.asyncEditor({ file: file.slice() })));
        else host.fail(new SessionFailure('crash', 'Worker stopped'));
      };
      await act(async () => fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true }));
      expect(clipboardReads(host)).toEqual([[0, 'A1:A1']]);
      await act(async () => {
        reading.resolve({ sheet: 0, version: 'v1', cells: [
          [{ a1: 'A1', input: 'initial', isFormula: false }],
        ] });
        await fulfilled.promise;
      });
      expect(clipboard.writeText).not.toHaveBeenCalled();
      expect(clipboard.write).not.toHaveBeenCalled();
      expect(host.sent).toEqual([]);
      expect(host.cells.get('0:0:0')).toBe('initial');
      expect(clipboardReads(host)).toEqual([[0, 'A1:A1']]);
      expect(host.hydrate).not.toHaveBeenCalled();
      expect(host.attach).not.toHaveBeenCalled();
    });

    it(`cut without ClipboardItem never writes after ${ending} during its read`, async () => {
      const host = peerlessHarness();
      const clipboard = textClipboard();
      const view = await mountPeerlessClipboard(host);
      const reading = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
      host.cellInputs.mockReturnValueOnce(reading.promise);
      await act(async () => fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true }));
      expect(clipboardReads(host)).toEqual([[0, 'A1:A1']]);
      expect(clipboard.writeText).not.toHaveBeenCalled();
      await act(async () => {
        if (ending === 'replacement') view.rerender(host.asyncEditor({ file: file.slice() }));
        else host.fail(new SessionFailure('crash', 'Worker stopped'));
      });
      await act(async () => reading.resolve({ sheet: 0, version: 'v1', cells: [
        [{ a1: 'A1', input: 'initial', isFormula: false }],
      ] }));
      expect(clipboard.writeText).not.toHaveBeenCalled();
      expect(clipboard.write).not.toHaveBeenCalled();
      expect(host.sent).toEqual([]);
      expect(host.cells.get('0:0:0')).toBe('initial');
      expect(clipboardReads(host)).toEqual([[0, 'A1:A1']]);
      expect(host.hydrate).not.toHaveBeenCalled();
      expect(host.attach).not.toHaveBeenCalled();
    });

    it(`cut without ClipboardItem never clears after ${ending}`, async () => {
      const host = peerlessHarness();
      const clipboard = textClipboard();
      const view = await mountPeerlessClipboard(host);
      await act(async () => fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true }));
      expect(clipboard.writeText).toHaveBeenCalledWith('initial');
      expect(host.sent).toEqual([]);
      await act(async () => {
        if (ending === 'replacement') view.rerender(host.asyncEditor({ file: file.slice() }));
        else host.fail(new SessionFailure('crash', 'Worker stopped'));
      });
      await act(async () => clipboard.writing.resolve());
      expect(host.sent).toEqual([]);
      expect(host.cells.get('0:0:0')).toBe('initial');
      expect(host.hydrate).not.toHaveBeenCalled();
      expect(host.attach).not.toHaveBeenCalled();
    });
  }

  it('clear sends one compact operation for the entire selected range', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    await act(async () => { fireEvent.keyDown(scroll, { key: 'a', ctrlKey: true }); });
    fireEvent.keyDown(scroll, { key: 'Delete' });
    await act(async () => {});
    const clear = host.sent[host.sent.length - 1]!;
    expect(clear).toMatchObject({ kind: 'applyOps', payload: { toolbar: { id: 'clear', target: {
      range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 1_048_575, col: 16_383 } },
    } } } });
    expect(JSON.stringify(clear.payload).length).toBeLessThan(500);
  });

  it('prefill never overwrites newer draft', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const prefill = deferred<Awaited<ReturnType<typeof host.cellInputs>>>();
    const requested = host.cellInputs.mock.calls.length;
    host.cellInputs.mockReturnValueOnce(prefill.promise);
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    await waitFor(() => expect(host.cellInputs.mock.calls.length).toBeGreaterThan(requested));
    fireEvent.change(input, { target: { value: 'newer' } });
    await act(async () => prefill.resolve({ sheet: 0, version: 'v1', cells: [[{ a1: 'A1', input: 'stale', isFormula: false }]] }));
    expect(input.value).toBe('newer');
  });

  it('bare ack causes no editor render', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await act(async () => {});
    expect(host.sent).toHaveLength(1);
    const count = host.renderCount.mock.calls.length;
    const frames = host.frame.mock.calls.length;
    await act(async () => host.push({ throughSeq: 1, version: 'v1', statuses: [{ seq: 1, status: 'noop', committed: false }],
      editState: { canUndo: true } }));
    expect(host.renderCount.mock.calls.length).toBe(count);
    expect(host.frame.mock.calls.length).toBe(frames);
  });

  for (const source of ['cell', 'formula'] as const) {
    it(`${source} IME commits once`, async () => {
      const host = peerlessHarness();
      const view = await host.mount();
      if (source === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
      const input = view.getByTestId(source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input') as HTMLInputElement;
      if (source === 'cell') await waitFor(() => expect(input.value).toBe('initial'));
      fireEvent.focus(input);
      fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: '変換' } });
      Object.defineProperty(view.getByTestId('xlsx-scroll'), 'scrollLeft', { configurable: true, value: 96, writable: true });
      fireEvent.scroll(view.getByTestId('xlsx-scroll'));
      await advance();
      expect(input.value).toBe('変換');
      expect(host.sent).toEqual([]);
      fireEvent.compositionEnd(input);
      fireEvent.change(input, { target: { value: '変換' } });
      fireEvent.keyDown(input, { key: 'Enter' });
      fireEvent.blur(input);
      await act(async () => {});
      expect(host.sent.filter((command) => command.kind === 'editCell')).toHaveLength(1);
      expect(host.sent[0]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, '変換'] } });
      expect(host.sent.map((command) => command.kind)).toEqual(['editCell']);
    });
  }

  it('empty composition cancels', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    const input = view.getByTestId('xlsx-cell-editor');
    fireEvent.compositionStart(input);
    fireEvent.change(input, { target: { value: '' } });
    fireEvent.compositionEnd(input);
    fireEvent.blur(input);
    await act(async () => {});
    expect(host.sent).toEqual([]);
  });

  it('restores the formula input after an empty composition', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const input = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
    await waitFor(() => expect(input.value).toBe('initial'));
    await advance();
    fireEvent.focus(input);
    fireEvent.compositionStart(input);
    fireEvent.change(input, { target: { value: '' } });
    expect(input.value).toBe('');
    fireEvent.compositionEnd(input);
    await act(async () => {});
    expect(host.sent).toEqual([]);
    expect(host.cells.get('0:0:0')).toBe('initial');
    expect(input.value).toBe('initial');
  });

  it('composition fences toolbar with one captured-target commit', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const input = view.getByTestId('xlsx-formula-input');
    fireEvent.compositionStart(input);
    fireEvent.change(input, { target: { value: '入力' } });
    const bold = host.commands().execute('bold', null);
    expect(host.sent).toEqual([]);
    fireEvent.compositionEnd(input);
    fireEvent.change(input, { target: { value: '入力済み' } });
    await act(async () => {});
    expect(host.sent.map((command) => command.kind)).toEqual(['editCell', 'applyOps']);
    expect(host.sent[0]).toMatchObject({ payload: { args: [0, 0, 0, '入力済み'] } });
    expect(host.sent[1].afterSeq).toBe(host.sent[0].seq);
    await act(async () => host.push({ throughSeq: 2, version: 'v1', statuses: [
      { seq: 1, status: 'applied' }, { seq: 2, status: 'applied' },
    ] }));
    await bold;
  });

  it('conversion Enter does not navigate', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    const input = view.getByTestId('xlsx-cell-editor');
    await waitFor(() => expect((input as HTMLInputElement).value).toBe('initial'));
    fireEvent.compositionStart(input);
    fireEvent.change(input, { target: { value: '確定' } });
    fireEvent.keyDown(input, { key: 'Enter', isComposing: true });
    fireEvent.compositionEnd(input);
    fireEvent.keyDown(input, { key: 'Enter' });
    await act(async () => {});
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
    expect(host.sent).toHaveLength(1);
  });

  it('toggles armed formatting locally before capturing a fresh token', async () => {
    const host = peerlessHarness();
    await host.mount();
    const first = host.commands().execute('paintFormat', null);
    await act(async () => {});
    const capture = host.sent[0]!;
    await act(async () => host.ack(capture, { result: { formatToken: 'format:1' } }));
    expect(await first).toMatchObject({ ok: true });
    expect(host.commands().getState('paintFormat').active).toBe(true);
    await act(async () => {
      expect(await host.commands().execute('paintFormat', null)).toMatchObject({ ok: true });
    });
    expect(host.commands().getState('paintFormat').active).toBe(false);
    expect(host.sent).toHaveLength(1);
    const second = host.commands().execute('paintFormat', null);
    await act(async () => {});
    expect(host.sent).toHaveLength(2);
    await act(async () => host.ack(host.sent[1]!, { result: { formatToken: 'format:2' } }));
    expect(await second).toMatchObject({ ok: true });
    expect(host.commands().getState('paintFormat').active).toBe(true);
  });

  for (const [armed, input] of [
    [false, 'xlsx-scroll'], [true, 'xlsx-scroll'],
    [false, 'xlsx-formula-input'], [true, 'xlsx-formula-input'],
  ] as const) {
    it(`Escape cancels ${armed ? 'armed' : 'in-flight'} formatting from ${input}`, async () => {
      const host = peerlessHarness();
      const view = await host.mount();
      const capturing = host.commands().execute('paintFormat', null);
      await act(async () => {});
      const capture = host.sent[0]!;
      if (armed) {
        await act(async () => host.ack(capture, { result: { formatToken: 'format:1' } }));
        expect(await capturing).toMatchObject({ ok: true });
        expect(host.commands().getState('paintFormat').active).toBe(true);
      }
      fireEvent.keyDown(view.getByTestId(input), { key: 'Escape' });
      await act(async () => {});
      if (!armed) {
        await act(async () => host.ack(capture, { result: { formatToken: 'format:1' } }));
        expect(await capturing).toMatchObject({ ok: true });
      }
      expect(host.commands().getState('paintFormat').active).toBe(false);
      expect(host.sent).toHaveLength(1);
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
      await act(async () => {});
      expect(host.sent.some((command) => command.kind === 'applyFormat')).toBe(false);
    });
  }

  it('captures and applies fresh format tokens to consecutive targets', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    for (const col of [1, 2]) {
      const capturing = host.commands().execute('paintFormat', null);
      await act(async () => {});
      const capture = host.sent[host.sent.length - 1]!;
      expect(capture).toMatchObject({ kind: 'applyOps', payload: { toolbar: { id: 'paintFormat' } } });
      await act(async () => host.ack(capture, { result: { formatToken: `format:${col}` } }));
      expect(await capturing).toMatchObject({ ok: true });
      expect(host.commands().getState('paintFormat').active).toBe(true);
      fireEvent.keyDown(scroll, { key: 'ArrowRight' });
      await act(async () => {});
      const selected = host.sent[host.sent.length - 2]!;
      const applied = host.sent[host.sent.length - 1]!;
      expect(selected.kind).toBe('select');
      expect(applied).toMatchObject({ kind: 'applyFormat', payload: {
        formatToken: `format:${col}`, target: { sheetId: 'sheet:0', range: {
          kind: 'rowCol', start: { row: 0, col }, end: { row: 0, col },
        } },
      } });
      expect(host.commands().getState('paintFormat').active).toBe(true);
      await act(async () => { host.ack(selected); host.ack(applied); });
      expect(host.commands().getState('paintFormat').active).toBe(false);
    }
    expect(host.sent.filter((command) => command.kind === 'applyFormat')).toHaveLength(2);
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
  });

  it('keeps the format token armed after a refused apply', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const capturing = host.commands().execute('paintFormat', null);
    await act(async () => {});
    await act(async () => host.ack(host.sent[0]!, { result: { formatToken: 'format:1' } }));
    await capturing;
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
    await act(async () => {});
    const selected = host.sent[host.sent.length - 2]!;
    const applied = host.sent[host.sent.length - 1]!;
    await act(async () => { host.ack(selected); host.ack(applied, { status: 'refused' }); });
    expect(host.commands().getState('paintFormat').active).toBe(true);
    expect(host.sent.filter((command) => command.kind === 'applyFormat')).toHaveLength(1);
  });

  it('format chart and proposal routes use no peer', async () => {
    const host = peerlessHarness();
    const available = spyOn(xlsx, 'isProposalsAvailable').mockReturnValue(true);
    restorers.push(() => available.mockRestore());
    host.charts.push({ id: 'chart', label: 'Chart', movable: true, rect: { x: 200, y: 100, w: 100, h: 100 }, clip: { x: 200, y: 100, w: 100, h: 100 } });
    const view = await host.mount();
    const execute = host.commands().execute('paintFormat', null);
    await act(async () => {});
    const capture = host.sent.find((command) => command.kind === 'applyOps' && 'toolbar' in command.payload && command.payload.toolbar.id === 'paintFormat')!;
    await act(async () => host.ack(capture, { result: { formatToken: 'format:1' } }));
    await execute;
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
    await act(async () => {});
    expect(host.sent.find((command) => command.kind === 'applyFormat')).toMatchObject({ payload: { formatToken: 'format:1' } });
    fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { clientX: 220, clientY: 120, button: 0 });
    fireEvent.mouseUp(window, { clientX: 230, clientY: 125, button: 0 });
    await act(async () => {});
    expect(host.sent.find((command) => command.kind === 'moveChart')).toMatchObject({ payload: { args: [0, 'chart', 10, 5] } });
    await act(async () => host.push({ throughSeq: host.sent[host.sent.length - 1]!.seq, version: 'v1', statuses: host.sent.map((command) => ({ seq: command.seq, status: 'applied' })) }));
    const refresh = host.commands().execute('proposalsPanel', { open: true });
    await act(async () => {});
    await refresh;
    const request = host.sent.find((command) => command.kind === 'applyOps' && 'toolbar' in command.payload && command.payload.toolbar.id === 'refreshProposals')!;
    await act(async () => host.ack(request, { result: { proposals: [{ id: 'p1', agentId: 'agent', note: null, cells: [] }] } }));
    const accept = host.commands().execute('proposalAccept', { proposalId: 'p1' });
    await act(async () => {});
    const proposal = host.sent.find((command) => command.kind === 'acceptProposal')!;
    expect(proposal).toBeDefined();
    await act(async () => host.ack(proposal));
    await accept;
    const reject = host.commands().execute('proposalReject', { proposalId: 'p1' });
    await act(async () => {});
    const rejected = host.sent.find((command) => command.kind === 'rejectProposal')!;
    expect(rejected).toBeDefined();
    await act(async () => host.ack(rejected));
    await reject;
    for (const id of ['undo', 'redo'] as const) {
      const history = host.commands().execute(id, null);
      await act(async () => {});
      const command = host.sent.find((command) => command.kind === id)!;
      expect(command).toBeDefined();
      await act(async () => {
        host.push({ throughSeq: command.seq, version: 'v1', statuses: [{ seq: command.seq, status: 'applied' }],
          selection: { sheetId: 'sheet:0', range: { kind: 'rowCol', start: { row: 2, col: 0 }, end: { row: 2, col: 0 } } } });
      });
      await history;
      expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A3');
    }
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
  });

  it('print waits for presentation', async () => {
    const host = peerlessHarness();
    const view = await host.mount();
    const print = mock(() => {});
    replacePrint(print);
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    const printing = host.commands().execute('print', null);
    await act(async () => host.ack(host.sent[0]));
    expect(print).not.toHaveBeenCalled();
    await act(async () => host.ack(host.sent[0], { paint: true }));
    await printing;
    expect(print).toHaveBeenCalledTimes(1);
  });

  it('save fences the accepted draft and uses the worker bytes', async () => {
    const host = peerlessHarness();
    const delivered = mock((_bytes: Uint8Array) => {});
    const view = await host.mount(delivered);
    fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value: 'saved draft' } });
    const saving = host.commands().execute('save', null);
    await act(async () => {});
    expect(host.session.save).not.toHaveBeenCalled();
    expect(host.sent[0]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'saved draft'] } });
    await act(async () => host.ack(host.sent[0]));
    await saving;
    expect(delivered).toHaveBeenCalledWith(new Uint8Array([8, 9]));
  });
});

describe('workbook worker editor', () => {
  it('keeps legacy worker moves local before hydration and reveals offscreen pages by the nearest amount', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    const cellGeometry = mock(async (sheet: number, row: number, col: number): Promise<xlsx.WorkbookCellGeometry> => ({
      sheet, version: 'v1', rect: { x: col * 96, y: row * 24, w: 96, h: 24 },
      scrollPosition: { x: col * 96, y: row * 24 },
    }));
    Object.assign(host.sessionMethods, { cellGeometry });
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
      onReady={(ready) => { api = ready; }} />);
    await opened();
    const scroll = view.getByTestId('xlsx-scroll');
    for (const key of ['ArrowRight', 'ArrowDown', 'Tab', 'Enter']) {
      await act(async () => { fireEvent.keyDown(scroll, { key }); });
      expect(scroll.scrollLeft).toBe(0);
      expect(scroll.scrollTop).toBe(0);
    }
    expect(cellGeometry).not.toHaveBeenCalled();
    await act(async () => { await api.commands.execute('zoom', { scale: 2 }); });
    expect(scroll.scrollLeft).toBe(0);
    expect(scroll.scrollTop).toBe(0);
    expect(cellGeometry).not.toHaveBeenCalled();
    await act(async () => { fireEvent.keyDown(scroll, { key: 'PageDown' }); });
    expect(cellGeometry).toHaveBeenCalledTimes(1);
    expect(cellGeometry).toHaveBeenCalledWith(0, 14, 2);
    expect(scroll.scrollLeft).toBe(0);
    expect(scroll.scrollTop).toBe(120);
    await act(async () => hydration.resolve(host.peer));
  });

  for (const decision of [true, false, undefined] as const) {
    it(`coalesces host save requests returning ${String(decision)} after flushing worker input`, async () => {
      const host = harness();
      let api!: XlsxWorkerEditorApi;
      const request = deferred<boolean | void>();
      const onSaveRequest = mock(async () => {
        await api.flushPendingInput();
        return request.promise;
      });
      const saved = mock((_bytes: Uint8Array) => {});
      const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen
        onReady={(ready) => { api = ready; }} onSaveRequest={onSaveRequest} onSave={saved} />);
      await opened();
      await advance();
      fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value: 'host draft' } });
      let first!: ReturnType<XlsxWorkerEditorApi['commands']['execute']>;
      let second!: typeof first;
      act(() => {
        first = api.commands.execute('save', null);
        fireEvent.keyDown(view.getByTestId('xlsx-formula-input'), { key: 's', ctrlKey: true });
        second = api.commands.execute('save', null);
      });
      await advance();
      expect(onSaveRequest).toHaveBeenCalledTimes(1);
      expect(host.cells.get('0:0:0')).toBe('host draft');
      expect(saved).not.toHaveBeenCalled();
      await act(async () => { request.resolve(decision); await Promise.all([first, second]); });
      expect(host.editMethods.flush).toHaveBeenCalled();
      expect(host.editMethods.save).toHaveBeenCalledTimes(decision === true ? 1 : 0);
      expect(saved).toHaveBeenCalledTimes(decision === true ? 1 : 0);
    });
  }

  it('queries painted worker cells and refuses stale APIs after replacement', async () => {
    harness();
    const apis: XlsxWorkerEditorApi[] = [];
    const onReady = (api: XlsxWorkerEditorApi) => { apis.push(api); };
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen onReady={onReady} />);
    await opened();
    await advance();
    const canvas = view.container.querySelector('canvas')!;
    canvas.getBoundingClientRect = () => new DOMRect(40, 60, 1200, 900);
    const focused = document.activeElement;
    expect(apis[0].getPositionAtPoint(40 + 144 * 1.5, 60 + 36 * 1.5)).toEqual({ sheet: 0, row: 1, col: 1 });
    expect(apis[0].getPositionAtPoint(39, 70)).toBeNull();
    expect(document.activeElement).toBe(focused);
    await act(async () => { view.rerender(<XlsxEditor legacySyncApi file={file.slice()} experimentalWorkerOpen onReady={onReady} />); });
    expect(apis[0].getPositionAtPoint(40 + 144 * 1.5, 60 + 36 * 1.5)).toBeNull();
    await expect(apis[0].flushPendingInput()).rejects.toThrow('replaced');
  });

  it('clears a Ctrl+A sheet selection through its used cells', async () => {
    const host = harness(true);
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    const scroll = view.getByTestId('xlsx-scroll');
    await act(async () => { fireEvent.keyDown(scroll, { key: 'a', ctrlKey: true }); });
    await act(async () => { fireEvent.keyDown(scroll, { key: 'Delete' }); });
    await advance();
    expect(host.peerMethods.editCells).toHaveBeenCalledWith(0, [{ row: 0, col: 0, input: '' }]);
  });

  for (const hydrated of [false, true]) for (const key of ['Delete', 'x'] as const) {
    it(`${key === 'Delete' ? 'clears' : 'cuts'} cells added by a pending paste after Ctrl+A (${hydrated ? 'hydrated' : 'before hydration'})`, async () => {
      const host = harness();
      const hydration = deferred<WorkbookHandle>();
      if (!hydrated) host.hydrate.mockReturnValue(hydration.promise);
      const paste = deferred<string>();
      replaceClipboard('readText', mock(() => paste.promise));
      if (key === 'x') promisedClipboard();
      const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
      await opened();
      const scroll = view.getByTestId('xlsx-scroll');
      fireEvent.keyDown(scroll, { key: 'v', ctrlKey: true });
      await act(async () => { fireEvent.keyDown(scroll, { key: 'a', ctrlKey: true }); });
      fireEvent.keyDown(scroll, { key, ctrlKey: key === 'x' });
      await advance(3);
      expect(host.editMethods.editCells).not.toHaveBeenCalled();
      expect(host.cells.has('0:0:1')).toBe(false);
      await act(async () => { paste.resolve('first\toutside'); });
      if (!hydrated) {
        await advance(16);
        expect(host.preview.mock.calls[host.preview.mock.calls.length - 1]![2]).toMatchObject([
          { method: 'editCells', args: [0, [{ row: 0, col: 0, input: 'first' }, { row: 0, col: 1, input: 'outside' }]] },
        ]);
        expect(host.editMethods.editCells).not.toHaveBeenCalled();
        await act(async () => hydration.resolve(host.peer));
      }
      await advance(16);
      expect(host.editMethods.editCells.mock.calls).toEqual([
        [0, [{ row: 0, col: 0, input: 'first' }, { row: 0, col: 1, input: 'outside' }]],
        [0, [{ row: 0, col: 0, input: '' }, { row: 0, col: 1, input: '' }]],
      ]);
      expect(host.cells.get('0:0:0')).toBe('');
      expect(host.cells.get('0:0:1')).toBe('');
    });
  }

  it('preserves replacement input after a deferred Ctrl+A cut and an unchanged commit', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    const writing = deferred<void>();
    const clipboard = promisedClipboard(writing.promise);
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    const scroll = view.getByTestId('xlsx-scroll');
    await act(async () => { fireEvent.keyDown(scroll, { key: 'a', ctrlKey: true }); });
    fireEvent.keyDown(scroll, { key: 'x', ctrlKey: true });
    reviewEdit(view, 'replacement');
    await advance();
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('replacement');
    expect(host.editMethods.editCells).not.toHaveBeenCalled();
    expect(host.editMethods.editCell).not.toHaveBeenCalled();
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(clipboard.copied).toEqual(['initial']);
    fireEvent.keyDown(scroll, { key: 'ArrowUp' });
    fireEvent.keyDown(scroll, { key: 'F2' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await advance();
    expect(host.cells.get('0:0:0')).toBe('initial');
    expect(host.editMethods.editCells).not.toHaveBeenCalled();
    expect(host.editMethods.editCell).not.toHaveBeenCalled();
    await act(async () => writing.resolve());
    await advance(16);
    expect(host.editMethods.editCells.mock.calls).toEqual([
      [0, [{ row: 0, col: 0, input: '' }]],
    ]);
    expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'replacement']]);
    expect(host.cells.get('0:0:0')).toBe('replacement');
  });

  it('does not dispatch an empty clear for a Ctrl+A sheet selection', async () => {
    const host = harness();
    host.cells.clear();
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    const scroll = view.getByTestId('xlsx-scroll');
    await act(async () => { fireEvent.keyDown(scroll, { key: 'a', ctrlKey: true }); });
    fireEvent.keyDown(scroll, { key: 'Delete' });
    await advance();
    expect(host.editMethods.editCells).not.toHaveBeenCalled();
    expect(host.replay).not.toHaveBeenCalled();
  });

  it('shows visible ready commit text while the replacement frame is held without blocking input', async () => {
    const host = harness(true);
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    const adoption = deferred<void>();
    const frame = host.sessionMethods.frame.getMockImplementation()!;
    host.sessionMethods.frame.mockImplementation(async (...args) => {
      await adoption.promise;
      return frame(...args);
    });
    try {
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
      fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'new value' } });
      fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
      await act(async () => {});
      expect(host.peerMethods.editCell).toHaveBeenCalledWith(0, 0, 0, 'new value');
      expect(host.preview).not.toHaveBeenCalled();
      await advance();
      const preview = view.getByTestId('xlsx-commit-preview');
      expect(Number.parseFloat(preview.style.width)).toBeGreaterThan(0);
      expect(Number.parseFloat(preview.style.height)).toBeGreaterThan(0);
      expect(preview.tagName).toBe('DIV');
      expect(preview.textContent).toBe('new value');
      expect(getComputedStyle(preview).display).not.toBe('none');
      expect(getComputedStyle(preview).visibility).not.toBe('hidden');
      expect(preview.style.background).toBeTruthy();
      expect(preview.style.pointerEvents).toBe('none');
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x' });
      expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('x');
      fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Escape' });
      expect(view.getByTestId('xlsx-commit-preview')).toBe(preview);
      await act(async () => adoption.resolve());
      await advance();
      expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
      expect(view.container.querySelector('[data-paint-source="worker"]')?.getAttribute('data-worker-sequence')).toBe('1');
    } finally {
      await act(async () => adoption.resolve());
      await advance();
    }
  });

  it('opens retained worker editor sessions with omitted and explicit false readOnly', async () => {
    for (const props of [{ experimentalWorkerOpen: true as const }, { readOnly: false as const, experimentalWorkerOpen: true as const }]) {
      const host = harness();
      const ready = mock((_api: XlsxWorkerEditorApi) => {});
      const view = render(<XlsxEditor legacySyncApi file={file} {...props} onReady={ready} showToolbar={false} />);
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

  it('keeps edited bytes, readiness and live edit permissions across readOnly toggles', async () => {
    const host = harness();
    const viewerOpen = spyOn(workbookSessionOpener, 'open').mockResolvedValue(host.session);
    restorers.push(() => viewerOpen.mockRestore());
    host.editMethods.save.mockImplementation(async () => new TextEncoder().encode(JSON.stringify([...host.cells])).buffer);
    let api!: XlsxWorkerEditorApi;
    const ready = mock((value: XlsxWorkerEditorApi | XlsxWorkerViewerApi) => {
      if ('whenHydrated' in value) api = value;
    });
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={ready} />);
    await opened();
    await advance();
    const originalApi = api;
    let editing!: ReturnType<XlsxWorkerEditorApi['editCellAsync']>;
    act(() => { editing = api.editCellAsync(0, 0, 0, 'first edit'); });
    await advance();
    expect(await editing).toMatchObject({ applied: true });
    view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly showToolbar={false} onReady={ready} />);
    await advance();
    expect(viewerOpen).not.toHaveBeenCalled();
    expect(host.open).toHaveBeenCalledTimes(1);
    expect(host.session.dispose).not.toHaveBeenCalled();
    expect(ready).toHaveBeenCalledTimes(1);
    expect(api).toBe(originalApi);
    expect(api.hydrated).toBe(true);
    expect(await api.cellAsync(0, 0, 0)).toMatchObject({ input: 'first edit' });
    expect(painted[painted.length - 1].commands).toContainEqual(expect.objectContaining({ op: 'text', text: 'worker:first edit' }));
    expect(() => api.editCell(0, 0, 0, 'blocked')).toThrow('read-only');
    expect(await api.validateEdits(reviewBatch('blocked', 'v2'))).toMatchObject({ ok: false, failure: { code: 'read-only' } });
    expect(api.commands.getState('bold')).toMatchObject({ enabled: false, disabledReason: { code: 'read-only' } });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x' });
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    expect(JSON.parse(new TextDecoder().decode((await api.saveAsync())!))).toEqual([['0:0:0', 'first edit']]);
    view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly={false} showToolbar={false} onReady={ready} />);
    expect(api.commands.getState('bold').enabled).toBe(true);
    act(() => { editing = api.editCellAsync(0, 0, 1, 'second edit'); });
    await advance();
    expect(await editing).toMatchObject({ applied: true });
    expect(await api.cellAsync(0, 0, 0)).toMatchObject({ input: 'first edit' });
    expect(await api.cellAsync(0, 0, 1)).toMatchObject({ input: 'second edit' });
    expect(JSON.parse(new TextDecoder().decode((await api.saveAsync())!))).toEqual([
      ['0:0:0', 'first edit'], ['0:0:1', 'second edit'],
    ]);
    expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'first edit'], [0, 0, 1, 'second edit']]);
    expect(host.open).toHaveBeenCalledTimes(1);
    expect(host.hydrate).toHaveBeenCalledTimes(1);
    expect(ready).toHaveBeenCalledTimes(1);
    expect(api).toBe(originalApi);
  });

  it('reports no error or refusal for idle readOnly flips before hydration', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    const errors = mock(() => {});
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onError={errors} />);
    await opened();
    expect(host.attached).toBeNull();
    view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly showToolbar={false} onError={errors} />);
    await advance(3);
    expect(errors).not.toHaveBeenCalled();
    expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
    view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly={false} showToolbar={false} onError={errors} />);
    await advance(3);
    expect(errors).not.toHaveBeenCalled();
    expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(errors).not.toHaveBeenCalled();
    expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
    expect(host.editMethods.editCell).not.toHaveBeenCalled();
    expect(host.editMethods.flush).not.toHaveBeenCalled();
  });

  it('commits a cell draft awaiting hydration across a readOnly flip', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    host.editMethods.save.mockImplementation(async () => new TextEncoder().encode(JSON.stringify([...host.cells])).buffer);
    let api!: XlsxWorkerEditorApi;
    const onReady = (value: XlsxWorkerEditorApi | XlsxWorkerViewerApi) => { if ('whenHydrated' in value) api = value; };
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={onReady} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 't' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'pending text' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await advance(3);
    view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly showToolbar={false} onReady={onReady} />);
    let saved!: Uint8Array | null;
    act(() => { void api.saveAsync().then((bytes) => { saved = bytes; }); });
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
    expect(await api.cellAsync(0, 0, 0)).toMatchObject({ input: 'pending text' });
    expect(painted.some((frame) => frame.commands.some((command) => command.op === 'text' && command.text === 'worker:pending text'))).toBe(true);
    expect(JSON.parse(new TextDecoder().decode(saved!))).toEqual([['0:0:0', 'pending text']]);
    await expect(api.editCellAsync(0, 0, 0, 'later')).rejects.toThrow('read-only');
    view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={onReady} />);
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
  });

  it('applies a granted plugin batch queued across a readOnly flip', async () => {
    const host = harness();
    host.editMethods.save.mockImplementation(async () => new TextEncoder().encode(JSON.stringify([...host.cells])).buffer);
    let api!: XlsxWorkerEditorApi;
    const onReady = (value: XlsxWorkerEditorApi | XlsxWorkerViewerApi) => { if ('whenHydrated' in value) api = value; };
    let context!: XlsxPluginContext<null>;
    const plugin = defineXlsxPlugin({ id: 'review', createState: () => null, initialize(value) { context = value; } });
    const plugins = [plugin];
    const pluginGrants = { review: { document: 'write' as const, editBatches: true as const } };
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} plugins={plugins} pluginGrants={pluginGrants} onReady={onReady} />);
    await opened();
    await waitFor(() => expect(context).toBeTruthy());
    const acknowledgment = deferred<void>();
    workbookEditPeerInternals.get(host.edits)!.whenAcknowledged = () => acknowledgment.promise;
    let applied: unknown;
    await act(async () => {
      void context.run(async (current) => { applied = await current.edits!.applyEdits(reviewBatch('plugin')); });
    });
    expect(host.editMethods.applyEdits).not.toHaveBeenCalled();
    view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly showToolbar={false} plugins={plugins} pluginGrants={pluginGrants} onReady={onReady} />);
    await act(async () => acknowledgment.resolve());
    await advance();
    expect(applied).toMatchObject({ ok: true });
    expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
    expect(await api.cellAsync(0, 0, 0)).toMatchObject({ input: 'plugin' });
    expect(painted.some((frame) => frame.commands.some((command) => command.op === 'text' && command.text === 'worker:plugin'))).toBe(true);
    expect(JSON.parse(new TextDecoder().decode((await api.saveAsync())!))).toEqual([['0:0:0', 'plugin']]);
    await act(async () => {
      await context.run(async (current) => {
        expect(await current.edits!.applyEdits(reviewBatch('later'))).toMatchObject({ ok: false, failure: { code: 'read-only' } });
      });
    });
    expect(host.editMethods.applyEdits).toHaveBeenCalledTimes(1);
  });

  for (const kind of ['paste', 'cut'] as const) {
    it(`applies ${kind} awaiting clipboard completion across a readOnly flip`, async () => {
      const host = harness();
      host.editMethods.save.mockImplementation(async () => new TextEncoder().encode(JSON.stringify([...host.cells])).buffer);
      let api!: XlsxWorkerEditorApi;
      const onReady = (value: XlsxWorkerEditorApi | XlsxWorkerViewerApi) => { if ('whenHydrated' in value) api = value; };
      const completion = deferred<void>();
      const reading = deferred<string>();
      if (kind === 'paste') {
        const read = spyOn(navigator.clipboard, 'readText').mockReturnValue(reading.promise);
        restorers.push(() => read.mockRestore());
      } else promisedClipboard().write.mockReturnValue(completion.promise);
      const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={onReady} />);
      await opened();
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: kind === 'paste' ? 'v' : 'x', ctrlKey: true });
      await advance(3);
      view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly showToolbar={false} onReady={onReady} />);
      let saved!: Uint8Array | null;
      act(() => { void api.saveAsync().then((bytes) => { saved = bytes; }); });
      await act(async () => { reading.resolve('pasted'); completion.resolve(); });
      await advance();
      const input = kind === 'paste' ? 'pasted' : '';
      expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
      expect(await api.cellAsync(0, 0, 0)).toMatchObject({ input });
      expect(painted.some((frame) => frame.commands.some((command) => command.op === 'text' && command.text === `worker:${input}`))).toBe(true);
      expect(JSON.parse(new TextDecoder().decode(saved!))).toEqual([['0:0:0', input]]);
      await expect(api.editCellAsync(0, 0, 0, 'later')).rejects.toThrow('read-only');
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: kind === 'paste' ? 'v' : 'x', ctrlKey: true });
      await advance();
      expect(host.editMethods.editCells).toHaveBeenCalledTimes(1);
    });
  }

  for (const source of ['cell', 'formula'] as const) {
    it(`keeps a newer same-cell ${source} draft open when a sealed composition save resumes after unlocking`, async () => {
      const host = harness();
      host.editMethods.save.mockImplementation(async () => new TextEncoder().encode(JSON.stringify([...host.cells])).buffer);
      const toolbar = <EditorToolbar mode="commands"><EditorToolbar.FormulaBar /></EditorToolbar>;
      let api!: XlsxWorkerEditorApi;
      const onReady = (value: XlsxWorkerEditorApi | XlsxWorkerViewerApi) => { if ('whenHydrated' in value) api = value; };
      const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen toolbar={toolbar} onReady={onReady} />);
      await opened();
      const testId = source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input';
      if (source === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 't' });
      const input = view.getByTestId(testId) as HTMLInputElement;
      fireEvent.compositionStart(input);
      let saving!: ReturnType<XlsxWorkerEditorApi['saveAsync']>;
      act(() => { saving = api.saveAsync(); });
      input.value = 'sealed composed text';
      view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly toolbar={toolbar} onReady={onReady} />);
      view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly={false} toolbar={toolbar} onReady={onReady} />);
      if (source === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
      const newerInput = view.getByTestId(testId) as HTMLInputElement;
      fireEvent.change(newerInput, { target: { value: 'new draft' } });
      await advance();
      expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
      expect(view.getByTestId(testId)).toBe(newerInput);
      expect(newerInput.value).toBe('new draft');
      expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'sealed composed text']]);
      expect(JSON.parse(new TextDecoder().decode((await saving)!))).toEqual([['0:0:0', 'sealed composed text']]);
      fireEvent.keyDown(newerInput, { key: 'Enter' });
      await advance();
      expect(host.editMethods.editCell.mock.calls).toEqual([
        [0, 0, 0, 'sealed composed text'], [0, 0, 0, 'new draft'],
      ]);
    });

    it(`keeps DOM-only ${source} composition and its pending save across an undefined-to-false readOnly rerender`, async () => {
      const host = harness();
      host.editMethods.save.mockImplementation(async () => new TextEncoder().encode(JSON.stringify([...host.cells])).buffer);
      let api!: XlsxWorkerEditorApi;
      const onReady = (value: XlsxWorkerEditorApi | XlsxWorkerViewerApi) => { if ('whenHydrated' in value) api = value; };
      const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen onReady={onReady} />);
      await opened();
      if (source === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 't' });
      const input = view.getByTestId(source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input') as HTMLInputElement;
      fireEvent.compositionStart(input);
      let saved!: Uint8Array | null;
      act(() => { void api.saveAsync().then((bytes) => { saved = bytes; }); });
      input.value = 'DOM-only composition';
      view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly={false} onReady={onReady} />);
      await advance(3);
      expect(input.value).toBe('DOM-only composition');
      expect(saved).toBeUndefined();
      expect(host.editMethods.editCell).not.toHaveBeenCalled();
      input.value = 'final composed text';
      fireEvent.compositionEnd(input);
      await advance();
      expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
      expect(host.cells.get('0:0:0')).toBe('final composed text');
      expect(JSON.parse(new TextDecoder().decode(saved!))).toEqual([['0:0:0', 'final composed text']]);
    });

    it(`commits an open ${source} draft before locking without restoring it on another sheet`, async () => {
      const host = harness();
      host.editMethods.save.mockImplementation(async () => new TextEncoder().encode(JSON.stringify([...host.cells])).buffer);
      const toolbar = <EditorToolbar mode="commands"><EditorToolbar.FormulaBar /></EditorToolbar>;
      let api!: XlsxWorkerEditorApi;
      const onReady = (value: XlsxWorkerEditorApi | XlsxWorkerViewerApi) => { if ('whenHydrated' in value) api = value; };
      const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen toolbar={toolbar} onReady={onReady} />);
      await opened();
      await advance();
      act(() => { expect(api.selectCells(0, xlsx.selectionAt({ row: 1, col: 1 }))).toBe(true); });
      if (source === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 't' });
      fireEvent.change(view.getByTestId(source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input'), { target: { value: 'pending text' } });
      view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly toolbar={toolbar} onReady={onReady} />);
      await advance();
      expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
      expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).value).toBe('pending text');
      expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).readOnly).toBe(true);
      expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
      expect(await api.cellAsync(0, 1, 1)).toMatchObject({ input: 'pending text' });
      expect(JSON.parse(new TextDecoder().decode((await api.saveAsync())!))).toEqual([['0:0:0', 'initial'], ['0:1:1', 'pending text']]);
      await expect(api.editCellAsync(0, 1, 1, 'later')).rejects.toThrow('read-only');
      fireEvent.click(view.getByRole('tab', { name: 'Second' }));
      await waitFor(() => expect(view.getByRole('tab', { name: 'Second' }).getAttribute('aria-selected')).toBe('true'));
      view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen toolbar={toolbar} onReady={onReady} />);
      expect(view.getByRole('tab', { name: 'Second' }).getAttribute('aria-selected')).toBe('true');
      expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
      expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).value).toBe('');
    });
  }

  for (const source of ['cell', 'formula'] as const) {
    for (const newerDraft of [false, true]) {
      it(`does not restore a failed settled ${source} draft after unlocking (${newerDraft ? 'new draft' : 'no draft'})`, async () => {
        const host = harness();
        const hydration = deferred<WorkbookHandle>();
        host.hydrate.mockReturnValue(hydration.promise);
        host.editMethods.editCell.mockImplementationOnce(() => { throw new RangeError('Invalid formula'); });
        const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen />);
        await opened();
        const testId = source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input';
        if (source === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 't' });
        fireEvent.change(view.getByTestId(testId), { target: { value: '=invalid()' } });
        view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly />);
        view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly={false} />);
        if (newerDraft) {
          if (source === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
          fireEvent.change(view.getByTestId(testId), { target: { value: 'new draft' } });
        }
        await act(async () => hydration.resolve(host.peer));
        await advance();
        expect(view.getByTestId('xlsx-input-refusal').textContent).toBe('Invalid formula');
        expect(host.cells.get('0:0:0')).toBe('initial');
        if (newerDraft || source === 'formula') {
          expect((view.getByTestId(testId) as HTMLInputElement).value).toBe(newerDraft ? 'new draft' : 'initial');
        } else expect(view.queryByTestId(testId)).toBeNull();
      });
    }
  }

  for (const route of ['cell', 'formula', 'batch'] as const) {
    it(`surfaces a rejected pending ${route} commit across a readOnly flip`, async () => {
      const host = harness();
      const hydration = deferred<WorkbookHandle>();
      host.hydrate.mockReturnValue(hydration.promise);
      host.editMethods.editCell.mockImplementationOnce(() => { throw new RangeError('Invalid formula'); });
      host.editMethods.applyEdits.mockReturnValueOnce({ ok: false, version: 'v1',
        failure: { code: 'stale-version', message: 'Version changed' } });
      const toolbar = <EditorToolbar mode="commands"><EditorToolbar.FormulaBar /></EditorToolbar>;
      let api!: XlsxWorkerEditorApi;
      const onReady = (value: XlsxWorkerEditorApi | XlsxWorkerViewerApi) => { if ('whenHydrated' in value) api = value; };
      const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen toolbar={toolbar} onReady={onReady} />);
      await opened();
      if (route === 'batch') act(() => { void api.applyEdits(reviewBatch('rejected')); });
      else {
        if (route === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 't' });
        fireEvent.change(view.getByTestId(route === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input'), { target: { value: '=invalid()' } });
      }
      view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly toolbar={toolbar} onReady={onReady} />);
      await act(async () => hydration.resolve(host.peer));
      await advance();
      expect(view.getByTestId('xlsx-input-refusal').textContent).toBe(route === 'batch' ? 'Version changed' : 'Invalid formula');
      expect(host.cells.get('0:0:0')).toBe('initial');
      expect(api.failure).toBeNull();
      view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen toolbar={toolbar} onReady={onReady} />);
      expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
      expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).value).toBe('initial');
    });
  }

  it('hides the default toolbar while readOnly and keeps supplied toolbars visible', async () => {
    harness();
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen />);
    await opened();
    expect(view.getByTestId('xlsx-toolbar')).toBeTruthy();
    view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly />);
    expect(view.queryByTestId('xlsx-toolbar')).toBeNull();
    view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen />);
    expect(view.getByTestId('xlsx-toolbar')).toBeTruthy();
    view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly toolbar={<div>Host toolbar</div>} />);
    expect(view.getByTestId('xlsx-toolbar').textContent).toBe('Host toolbar');
  });

  it('applies clearing cells queued behind hydration across a readOnly flip', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'Delete' });
    await advance(3);
    expect(view.getByTestId('xlsx-commit-preview')).toBeTruthy();
    view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly showToolbar={false} />);
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
    expect(host.editMethods.editCells).toHaveBeenCalledTimes(1);
    expect(host.cells.get('0:0:0')).toBe('');
  });

  for (const pending of ['drag', 'nudge', 'queued move'] as const) {
    it(`${pending === 'queued move' ? 'applies' : 'cancels'} a chart ${pending} across a readOnly flip`, async () => {
      const host = harness();
      host.charts.push({ id: 'chart:1', label: 'Chart', movable: true, rect: { x: 192, y: 96, w: 192, h: 96 },
        clip: { x: 192, y: 96, w: 192, h: 96 } });
      const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
      await opened();
      const acknowledgment = deferred<void>();
      if (pending === 'queued move') workbookEditPeerInternals.get(host.edits)!.whenAcknowledged = () => acknowledgment.promise;
      fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { button: 0, clientX: 200, clientY: 100 });
      if (pending === 'nudge') {
        fireEvent.mouseUp(window, { button: 0, clientX: 200, clientY: 100 });
        fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
      } else fireEvent.mouseMove(view.getByTestId('xlsx-scroll'), { buttons: 1, clientX: 230, clientY: 120 });
      if (pending === 'queued move') {
        fireEvent.mouseUp(window, { button: 0, clientX: 230, clientY: 120 });
        await act(async () => {});
      }
      view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly showToolbar={false} />);
      if (pending === 'drag') fireEvent.mouseUp(window, { button: 0, clientX: 230, clientY: 120 });
      await act(async () => acknowledgment.resolve());
      await advance();
      expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
      expect(view.getByTestId('xlsx-chart-selection').style.transform).toBe('translate(0px, 0px)');
      expect(host.editMethods.moveChart).toHaveBeenCalledTimes(pending === 'queued move' ? 1 : 0);
    });
  }

  it('applies format painting queued across a readOnly flip', async () => {
    const host = harness();
    const applyFormat = mock(() => ({ applied: true, sheetInfo: host.peer.sheetInfo() }));
    Object.assign(host.peer, { captureFormat: () => ({}) });
    Object.assign(host.edits, { applyFormat });
    let api!: XlsxWorkerEditorApi;
    const onReady = (value: XlsxWorkerEditorApi | XlsxWorkerViewerApi) => { if ('whenHydrated' in value) api = value; };
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={onReady} />);
    await opened();
    await act(async () => { expect(await api.commands.execute('paintFormat', null)).toMatchObject({ ok: true }); });
    const acknowledgment = deferred<void>();
    const acknowledge = mock(() => acknowledgment.promise);
    workbookEditPeerInternals.get(host.edits)!.whenAcknowledged = acknowledge;
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
    await act(async () => {});
    expect(acknowledge).toHaveBeenCalled();
    view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly showToolbar={false} onReady={onReady} />);
    await act(async () => acknowledgment.resolve());
    await advance();
    expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
    expect(applyFormat).toHaveBeenCalledTimes(1);
  });

  for (const source of ['cell', 'formula'] as const) {
    for (const supplied of [false, true]) {
      it(`settles DOM-only ${source} composition and its pending save (${supplied ? 'supplied' : 'default'} toolbar)`, async () => {
        const host = harness();
        host.editMethods.save.mockImplementation(async () => new TextEncoder().encode(JSON.stringify([...host.cells])).buffer);
        const toolbar = supplied ? <EditorToolbar mode="commands"><EditorToolbar.FormulaBar /></EditorToolbar> : undefined;
        let api!: XlsxWorkerEditorApi;
        const onReady = (value: XlsxWorkerEditorApi | XlsxWorkerViewerApi) => { if ('whenHydrated' in value) api = value; };
        const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen toolbar={toolbar} onReady={onReady} />);
        await opened();
        if (source === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 't' });
        const input = view.getByTestId(source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input') as HTMLInputElement;
        fireEvent.compositionStart(input);
        let saved!: Uint8Array | null;
        act(() => { void api.saveAsync().then((bytes) => { saved = bytes; }); });
        input.value = 'latest composed text';
        view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly toolbar={toolbar} onReady={onReady} />);
        await advance();
        expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
        expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
        expect(await api.cellAsync(0, 0, 0)).toMatchObject({ input: 'latest composed text' });
        expect(painted.some((frame) => frame.commands.some((command) => command.op === 'text' && command.text === 'worker:latest composed text'))).toBe(true);
        expect(JSON.parse(new TextDecoder().decode(saved!))).toEqual([['0:0:0', 'latest composed text']]);
        await expect(api.editCellAsync(0, 0, 0, 'later')).rejects.toThrow('read-only');
        view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen toolbar={toolbar} onReady={onReady} />);
        expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
        expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).value).toBe('latest composed text');
      });
    }
  }

  for (const route of ['cell', 'batch'] as const) {
    for (const submittedReadOnly of [false, true]) {
      it(`uses submission-time permissions for a hydration-pending API ${route} edit (${submittedReadOnly ? 'read-only' : 'editable'})`, async () => {
        const host = harness();
        host.editMethods.save.mockImplementation(async () => new TextEncoder().encode(JSON.stringify([...host.cells])).buffer);
        const hydration = deferred<WorkbookHandle>();
        host.hydrate.mockReturnValue(hydration.promise);
        let api!: XlsxWorkerEditorApi;
        const onReady = (value: XlsxWorkerEditorApi | XlsxWorkerViewerApi) => { if ('whenHydrated' in value) api = value; };
        const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={onReady} />);
        await opened();
        if (submittedReadOnly) view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly showToolbar={false} onReady={onReady} />);
        let outcome: unknown;
        act(() => {
          void (route === 'cell' ? api.editCellAsync(0, 0, 0, 'host') : api.applyEdits(reviewBatch('plugin')))
            .then((value) => { outcome = value; }, (error) => { outcome = error; });
        });
        view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly showToolbar={false} onReady={onReady} />);
        if (submittedReadOnly) view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={onReady} />);
        await act(async () => hydration.resolve(host.peer));
        await advance();
        if (submittedReadOnly) {
          expect(outcome).toMatchObject(route === 'cell' ? { message: 'The editor is read-only' } : { ok: false, failure: { code: 'read-only' } });
          expect(host.editMethods.editCell).not.toHaveBeenCalled();
          expect(host.editMethods.applyEdits).not.toHaveBeenCalled();
          expect(host.cells.get('0:0:0')).toBe('initial');
        } else {
          const input = route === 'cell' ? 'host' : 'plugin';
          expect(outcome).toMatchObject(route === 'cell' ? { applied: true } : { ok: true });
          expect(await api.cellAsync(0, 0, 0)).toMatchObject({ input });
          expect(painted.some((frame) => frame.commands.some((command) => command.op === 'text' && command.text === `worker:${input}`))).toBe(true);
          expect(JSON.parse(new TextDecoder().decode((await api.saveAsync())!))).toEqual([['0:0:0', input]]);
          await expect(api.editCellAsync(0, 0, 0, 'later')).rejects.toThrow('read-only');
        }
      });
    }

    it(`rejects a pending API ${route} caller but drains its edit on unmount after a readOnly flip`, async () => {
      const host = harness();
      const hydration = deferred<WorkbookHandle>();
      host.hydrate.mockReturnValue(hydration.promise);
      let api!: XlsxWorkerEditorApi;
      const onReady = (value: XlsxWorkerEditorApi | XlsxWorkerViewerApi) => { if ('whenHydrated' in value) api = value; };
      const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={onReady} />);
      await opened();
      let outcome: unknown;
      act(() => {
        void (route === 'cell' ? api.editCellAsync(0, 0, 0, 'host') : api.applyEdits(reviewBatch('host')))
          .then((value) => { outcome = value; }, (error) => { outcome = error; });
      });
      await advance(3);
      view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly showToolbar={false} onReady={onReady} />);
      view.unmount();
      await act(async () => hydration.resolve(host.peer));
      await advance();
      expect(outcome).toMatchObject({ code: 'document-replaced' });
      expect(host.editMethods.editCell).toHaveBeenCalledTimes(route === 'cell' ? 1 : 0);
      expect(host.editMethods.applyEdits).toHaveBeenCalledTimes(route === 'batch' ? 1 : 0);
      const input = route === 'cell' ? 'host' : 'plugin';
      expect(host.cells.get('0:0:0')).toBe(input);
      expect(host.log.indexOf(`edit:0:0:0:${input}`)).toBeLessThan(host.log.indexOf('dispose:facade'));
      expect(host.session.dispose).toHaveBeenCalledTimes(1);
    });
  }

  it('starts read-only files in the viewer and resets the choice on file replacement', async () => {
    const host = harness();
    const viewerOpen = spyOn(workbookSessionOpener, 'open').mockResolvedValue(host.session);
    restorers.push(() => viewerOpen.mockRestore());
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly showToolbar={false} />);
    await opened();
    await advance();
    expect(viewerOpen).toHaveBeenCalledTimes(1);
    expect(host.open).not.toHaveBeenCalled();
    expect(host.hydrate).not.toHaveBeenCalled();
    view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly={false} showToolbar={false} />);
    await opened();
    await advance();
    expect(host.open).toHaveBeenCalledTimes(1);
    expect(host.hydrate).toHaveBeenCalledTimes(1);
    view.rerender(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen readOnly showToolbar={false} />);
    await advance();
    expect(viewerOpen).toHaveBeenCalledTimes(1);
    const replacement = new Uint8Array([4, 5, 6]);
    view.rerender(<XlsxEditor legacySyncApi file={replacement} experimentalWorkerOpen readOnly showToolbar={false} />);
    await opened();
    await advance();
    expect(viewerOpen).toHaveBeenCalledTimes(2);
    expect(viewerOpen.mock.calls[1][0]).toBe(replacement);
    expect(host.open).toHaveBeenCalledTimes(1);
    expect(host.hydrate).toHaveBeenCalledTimes(1);
  });

  it('writes a ready draft without preview or a paint barrier and retains it until worker adoption', async () => {
    const host = harness();
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
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
      expect(host.preview).not.toHaveBeenCalled();
      expect(painted).toHaveLength(0);
      host.log.push('sync');
      return write(sheet, row, col, input);
    });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 't' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'typed' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('typed');
    await act(async () => {});
    expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'typed']]);
    expect(host.preview).not.toHaveBeenCalled();
    expect(host.peerMethods.displayList).not.toHaveBeenCalled();
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('typed');
    expect(host.log).not.toContain('frame:worker');
    await tick();
    expect(host.log.indexOf('sync')).toBeLessThan(host.log.indexOf('frame:worker'));
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('typed');
    const [viewport, options] = host.sessionMethods.frame.mock.calls[host.sessionMethods.frame.mock.calls.length - 1]!;
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
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
    const [viewport, options] = host.sessionMethods.frame.mock.calls[host.sessionMethods.frame.mock.calls.length - 1]!;
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
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
    const reopened = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
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

  it('preserves an accepted ready commit when untouched F2 input follows its write', async () => {
    const host = harness();
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'n' });
    fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'new' } });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await act(async () => {});
    expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'new']]);
    expect(host.preview).not.toHaveBeenCalled();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('new');
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onError={errors} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true });
    expect(clipboard.write).toHaveBeenCalledTimes(1);
    expect(clipboard.copied).toEqual([]);
    host.open.mockReturnValueOnce(new Promise<WorkbookSession>(() => {}));
    view.rerender(<XlsxEditor legacySyncApi file={new Uint8Array([4])} experimentalWorkerOpen showToolbar={false} onError={errors} />);
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
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
      const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
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
    expect(host.preview).not.toHaveBeenCalled();
    expect(painted[painted.length - 1].commands.some((command) => command.op === 'text' && command.text === 'worker:initial')).toBe(true);
    const nextFrame = host.sessionMethods.frame.getMockImplementation()!;
    host.sessionMethods.frame.mockImplementationOnce(async (...args) => {
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
      return nextFrame(...args);
    });
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 's' });
    fireEvent.click(view.getByRole('tab', { name: 'Second' }));
    await advance();
    expect(host.log.indexOf('edit:0:0:0:s')).toBeLessThan(host.log.indexOf('sheet:1'));
    expect(host.editMethods.setActiveSheet).toHaveBeenCalledWith(1);
    expect(host.peerMethods.setActiveSheet).not.toHaveBeenCalled();
    expect(view.getByRole('tab', { name: 'Second' }).getAttribute('aria-selected')).toBe('true');
  });

  it('commits a chart drag through the ordered facade without waiting for animation frames', async () => {
    const host = harness();
    host.charts.push({ id: 'chart:1', label: 'Chart', movable: true,
      rect: { x: 100, y: 100, w: 200, h: 100 }, clip: { x: 100, y: 100, w: 200, h: 100 } });
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
    await opened();
    fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { clientX: 120, clientY: 120, button: 0 });
    fireEvent.mouseMove(view.getByTestId('xlsx-scroll'), { clientX: 150, clientY: 140, buttons: 1 });
    expect(view.getByTestId('xlsx-chart-selection').style.transform).toBe('translate(30px, 20px)');
    fireEvent.mouseUp(window, { clientX: 150, clientY: 140, button: 0 });
    await act(async () => {});
    expect(host.editMethods.moveChart).toHaveBeenCalledWith(0, 'chart:1', 30, 20);
    expect(view.getByTestId('xlsx-chart-selection').style.transform).toBe('translate(0px, 0px)');
  });

  it('keeps the origin sheet for a chart drag released after queued navigation', async () => {
    const host = harness();
    host.charts.push({ id: 'chart:1', label: 'Chart', movable: true,
      rect: { x: 100, y: 100, w: 200, h: 100 }, clip: { x: 100, y: 100, w: 200, h: 100 } });
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen onReady={(value) => { api = value; }} />);
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
    const [viewport, options] = host.sessionMethods.frame.mock.calls[host.sessionMethods.frame.mock.calls.length - 1]!;
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen onError={errors}
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
      onReady={() => () => { host.log.push('dispose:ready'); }} />);
    await opened();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'q' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
      onError={errors} onReady={() => () => { throw error; }} />);
    await opened();
    view.rerender(<XlsxEditor legacySyncApi file={new Uint8Array([4])} experimentalWorkerOpen showToolbar={false}
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

  it('publishes the initial A1 selection to plugins after worker hydration', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    const snapshots: XlsxPluginSelection[] = [];
    const selections: XlsxPluginSelection[] = [];
    const plugin = defineXlsxPlugin({
      id: 'selection', createState: () => null,
      onEvent(context, event) {
        snapshots.push(context.snapshot.selection);
        if (event.type === 'selection-change') selections.push(event.selection);
      },
      overlay: ({ context }) => <span data-testid="plugin-selection">{JSON.stringify(context.snapshot.selection)}</span>,
    });
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} plugins={[plugin]} />);
    await opened();
    expect(host.hydrate).toHaveBeenCalledTimes(1);
    expect(snapshots).toEqual([]);
    expect(view.queryByTestId('plugin-selection')).toBeNull();
    expect(view.queryByTestId('xlsx-selection')).not.toBeNull();
    await act(async () => hydration.resolve(host.peer));
    await waitFor(() => expect(view.queryByTestId('plugin-selection')).not.toBeNull());
    const sheetIndex = host.session.state.activeSheet;
    const initial = { sheetId: host.session.state.sheets[sheetIndex].id, sheetIndex,
      cells: xlsx.selectionAt({ row: 0, col: 0 }), chartId: null };
    expect(snapshots[snapshots.length - 1]).toEqual(initial);
    expect(selections).toEqual([initial]);
    expect(view.getByTestId('plugin-selection').textContent).toBe(JSON.stringify(initial));
  });

  it('publishes the replacement workbook selection without the previous selection', async () => {
    const host = harness();
    let api!: XlsxWorkerEditorApi;
    const onReady = (value: XlsxWorkerEditorApi | XlsxWorkerViewerApi) => { if ('whenHydrated' in value) api = value; };
    const contexts: XlsxPluginContext<null>[] = [];
    const selections: { generation: string; selection: XlsxPluginSelection }[] = [];
    const plugin = defineXlsxPlugin({
      id: 'selection', createState: () => null,
      onEvent(context, event) {
        contexts.push(context);
        if (event.type === 'selection-change') selections.push({ generation: event.generation, selection: event.selection });
      },
      overlay: ({ context }) => {
        contexts.push(context);
        return <span data-testid="plugin-selection">{JSON.stringify(context.snapshot.selection)}</span>;
      },
    });
    const plugins = [plugin];
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} plugins={plugins} onReady={onReady} />);
    await opened();
    await waitFor(() => expect(view.queryByTestId('plugin-selection')).not.toBeNull());
    const previous = { sheetId: 'sheet:0', sheetIndex: 0, cells: xlsx.selectionAt({ row: 2, col: 1 }), chartId: null };
    act(() => { expect(api.selectCells(0, previous.cells)).toBe(true); });
    await waitFor(() => expect(view.getByTestId('plugin-selection').textContent).toBe(JSON.stringify(previous)));
    const previousContext = contexts[contexts.length - 1];
    for (const restore of restorers.splice(-3).reverse()) restore();
    const replacement = harness();
    const hydration = deferred<WorkbookHandle>();
    replacement.hydrate.mockReturnValue(hydration.promise);
    const replacementFile = new Uint8Array([4, 5, 6]);
    view.rerender(<XlsxEditor legacySyncApi file={replacementFile} experimentalWorkerOpen showToolbar={false} plugins={plugins} onReady={onReady} />);
    await opened();
    expect(previousContext.lifetimeSignal.aborted).toBe(true);
    expect(host.session.dispose).toHaveBeenCalledTimes(1);
    expect(replacement.peer).not.toBe(host.peer);
    expect(view.queryByTestId('plugin-selection')).toBeNull();
    await act(async () => hydration.resolve(replacement.peer));
    await waitFor(() => expect(view.queryByTestId('plugin-selection')).not.toBeNull());
    const sheetIndex = replacement.session.state.activeSheet;
    const initial = { sheetId: replacement.session.state.sheets[sheetIndex].id, sheetIndex,
      cells: xlsx.selectionAt({ row: 0, col: 0 }), chartId: null };
    const current = contexts.filter((context) => context.snapshot.generation !== previousContext.snapshot.generation);
    expect(current.length).toBeGreaterThan(0);
    expect(current[current.length - 1].snapshot.selection).toEqual(initial);
    for (const context of current) {
      if (context.snapshot.selection !== null) expect(context.snapshot.selection).toEqual(initial);
    }
    expect(selections.filter((event) => event.generation !== previousContext.snapshot.generation)
      .map((event) => event.selection)).toEqual([initial]);
    expect(view.getByTestId('plugin-selection').textContent).toBe(JSON.stringify(initial));
  });

  it('activates plugins after the real peer and routes granted batches through the facade', async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    let context!: XlsxPluginContext<null>;
    const initialize = mock((value: XlsxPluginContext<null>) => { context = value; });
    const plugin = defineXlsxPlugin({ id: 'review', createState: () => null, initialize });
    render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} plugins={[plugin]}
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

  it('recovers acknowledged edits through the real facade after a later paint failure', async () => {
    const host = harness(true);
    const failure = new Error('Peer publication failed');
    const errors = mock((_error: Error) => {});
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
      onError={errors} onReady={(value) => { api = value; }} />);
    await opened();
    host.sessionMethods.frame.mockRejectedValueOnce(failure);
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'a' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    let retained!: Promise<EditResult | null>;
    act(() => { retained = api.editCellAsync(0, 0, 1, 'retained'); void retained.catch(() => {}); });
    await act(async () => { expect((await retained)?.applied).toBe(true); });
    await advance();
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
    render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => {
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
    render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
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
    const recovery = spyOn(xlsxWorker, 'createWorkbookRecoveryMutators').mockImplementation(() => ({
      ...host.edits, editCell: (sheet, row, col, input) => {
        host.log.push(`recover:${input}`);
        host.cells.set(`${sheet}:${row}:${col}`, input);
        return { applied: true, sheetInfo: host.peer.sheetInfo() };
      },
    }));
    restorers.push(() => recovery.mockRestore());
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
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
    render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
    await opened();
    act(() => host.fail());
    const recovery = spyOn(xlsxWorker, 'createWorkbookRecoveryMutators').mockImplementation(() => {
      throw new xlsx.WorkbookEditPeerFailedError(new Error('Retained operation unavailable'));
    });
    restorers.push(() => recovery.mockRestore());
    await expect(api.recoverySave()).rejects.toBeInstanceOf(xlsx.WorkbookEditPeerFailedError);
    expect(host.editMethods.recoverySave).not.toHaveBeenCalled();
  });

  it('includes the retained live draft in recovery and continues refusing normal input', async () => {
    const host = harness();
    const recovery = spyOn(xlsxWorker, 'createWorkbookRecoveryMutators').mockImplementation(() => ({
      ...host.edits, editCell: (sheet, row, col, input) => {
        host.log.push(`recover:${input}`);
        host.cells.set(`${sheet}:${row}:${col}`, input);
        return { applied: true, sheetInfo: host.peer.sheetInfo() };
      },
    }));
    restorers.push(() => recovery.mockRestore());
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen onError={errors}
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

function replaceClipboard(name: 'write' | 'writeText' | 'readText' | 'read', value: unknown) {
  const original = Object.getOwnPropertyDescriptor(navigator.clipboard, name);
  Object.defineProperty(navigator.clipboard, name, { configurable: true, value });
  restorers.push(() => {
    if (original) Object.defineProperty(navigator.clipboard, name, original);
    else Reflect.deleteProperty(navigator.clipboard, name);
  });
}

for (const wait of ['hydration', 'ready'] as const) {
  for (const transition of ['replacement', 'unmount'] as const) {
    it(`drains accepted edits before replacing or disposing their document (${wait}, ${transition})`, async () => {
      const host = harness(true);
      const hydration = deferred<WorkbookHandle>();
      if (wait === 'hydration') host.hydrate.mockReturnValue(hydration.promise);
      let api!: XlsxWorkerEditorApi;
      const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
        onReady={(value) => { api = value; }} />);
      await opened();
      const acknowledged = deferred<WorkbookReplayReply>();
      host.replay.mockReturnValueOnce(acknowledged.promise);
      reviewEdit(view, 'retiring edit');
      if (transition === 'unmount') view.unmount();
      else {
        host.open.mockReturnValueOnce(new Promise<WorkbookSession>(() => {}));
        view.rerender(<XlsxEditor legacySyncApi file={new Uint8Array([4])} experimentalWorkerOpen showToolbar={false} />);
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
  host.charts.push({ id: 'chart', label: 'Chart', rect: { x: 192, y: 96, w: 192, h: 96 },
    clip: { x: 192, y: 96, w: 192, h: 96 }, movable: true });
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
  await opened();
  fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { clientX: 210, clientY: 110, button: 0 });
  expect(view.getByTestId('xlsx-chart-selection').getAttribute('data-chart-id')).toBe('chart');
  fireEvent.mouseUp(window, { clientX: 210, clientY: 110, button: 0 });
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
  view.unmount();
  await advance();
  expect(host.editMethods.moveChart).toHaveBeenCalledWith(0, 'chart', 1, 0);
  expect(host.editMethods.moveChart).toHaveBeenCalledTimes(1);
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
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
  await opened();
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true });
  await advance();
  expect(await (await clipboard.payloads[0]).text()).toBe('initial');
  let newer!: Promise<EditResult | null>;
  act(() => { newer = api.editCellAsync(0, 0, 1, 'newer'); void newer.catch(() => {}); host.fail(); });
  await expect(newer).rejects.toBeInstanceOf(SessionFailure);
  let recovered!: Promise<{ bytes: Uint8Array; recovery: true }>;
  act(() => { recovered = api.recoverySave(); });
  await act(async () => completion.resolve());
  const saved = await recovered;
  expect(host.peerMethods.editCells).toHaveBeenCalledTimes(1);
  expect(host.peerMethods.editCells.mock.calls).toEqual([[0, [{ row: 0, col: 0, input: '' }]]]);
  expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 1, 'newer']]);
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
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
  host.charts.push({ id: 'chart', label: 'Chart', rect: { x: 192, y: 96, w: 192, h: 96 },
    clip: { x: 192, y: 96, w: 192, h: 96 }, movable: true });
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen onReady={(value) => { api = value; }} />);
  await opened();
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'b', ctrlKey: true });
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
  await advance();
  expect(view.getByTestId('xlsx-input-refusal').textContent).toMatch(/selection changed/i);
  expect(host.peerMethods.patchRangeStyle).not.toHaveBeenCalled();
  fireEvent.mouseDown(view.getByTestId('xlsx-scroll'), { clientX: 210, clientY: 110, button: 0 });
  expect(view.getByTestId('xlsx-chart-selection').getAttribute('data-chart-id')).toBe('chart');
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 's', ctrlKey: true });
  await advance();
  expect(view.getByTestId('xlsx-input-refusal').textContent).toBe('Finish moving the chart first.');
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
      const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
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
  render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} plugins={[plugin]}
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} plugins={[plugin]}
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

for (const route of ['plugin', 'host'] as const) {
  it(`refuses a retained stale-version batch and recovers valid edits around it (${route})`, async () => {
    const host = harness(true);
    let context!: XlsxPluginContext<null>;
    let api!: XlsxWorkerEditorApi;
    const plugin = defineXlsxPlugin({ id: 'stale-plugin', createState: () => null, initialize(value) { context = value; } });
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} plugins={[plugin]}
      pluginGrants={{ 'stale-plugin': { document: 'write', editBatches: true } }} onReady={(value) => { api = value; }} />);
    await opened();
    await api.whenHydrated();
    await waitFor(() => expect(context).toBeTruthy());
    const request = { ...reviewBatch('refused', host.peer.version()),
      steps: reviewBatch('refused').steps.map(({ expect: _expect, ...step }) => step) };
    reviewEdit(view, 'before refusal');
    act(() => {
      if (route === 'plugin') void context.run(async (current) => { await current.edits!.applyEdits(request); });
      else void api.applyEdits(request).catch(() => {});
    });
    act(() => { void api.editCellAsync(0, 0, 1, 'after refusal').catch(() => {}); host.fail(); });
    const saved = await api.recoverySave();
    expect(saved.recovery).toBe(true);
    expect(JSON.parse(new TextDecoder().decode(saved.bytes))).toEqual([
      ['0:0:0', 'before refusal'], ['0:0:1', 'after refusal'],
    ]);
    expect(view.getByTestId('xlsx-input-refusal').textContent).toMatch(/version|refused/i);
    expect(host.peerMethods.applyEdits).toHaveBeenCalledTimes(1);
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'before refusal'], [0, 0, 1, 'after refusal']]);
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
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
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
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
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
    expect(host.peerMethods.editCell).toHaveBeenCalledTimes(invalid === 'formula' ? 2 : 1);
    expect(host.preview).not.toHaveBeenCalled();
    expect(host.replay).toHaveBeenCalledTimes(1);
    expect(host.replay.mock.calls[0][0].op).toMatchObject({ method: 'editCell', args: [0, 0, 0, 'valid'] });
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen onReady={(value) => { api = value; }} />);
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
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
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
  const [viewport] = host.sessionMethods.frame.mock.calls[host.sessionMethods.frame.mock.calls.length - 1]!;
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
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
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

for (const route of ['paste', 'cut', 'delete', 'host', 'bulk', 'formula-save', 'formula-navigation'] as const) {
  it(`paints every accepted pre-hydration cell edit before hydration resolves (${route})`, async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    replaceClipboard('readText', mock(async () => 'pasted'));
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    let accepted: Promise<unknown> | undefined;
    if (route === 'cut') { promisedClipboard(); fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true }); }
    else if (route === 'paste') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
    else if (route === 'delete') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'Delete' });
    else if (route === 'host') act(() => { accepted = api.editCellAsync(0, 0, 0, 'host'); });
    else if (route === 'bulk') act(() => { accepted = api.applyEdits(reviewBatch('plugin')); });
    else {
      fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value: '=24' } });
      act(() => { accepted = route === 'formula-save' ? api.save() : api.selectCellsAsync(1, xlsx.selectionAt({ row: 0, col: 0 })); });
    }
    await advance();
    const expected = route === 'delete' || route === 'cut' ? '' : route === 'paste' ? 'pasted' : route === 'host' ? 'host' : route === 'bulk' ? 'plugin' : '=24';
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
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
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
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
    onReady={(value) => { api = value; }} />);
  await opened();
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
  fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
  const later = api.editCellAsync(0, 0, 0, 'later');
  await act(async () => snapshot.resolve('initial'));
  await advance();
  expect(host.preview.mock.calls[0][2]).toMatchObject([{ method: 'editCell', args: [0, 0, 0, 'initial'] }]);
  expect(host.preview.mock.calls[host.preview.mock.calls.length - 1]![2]).toMatchObject([{ method: 'editCell', args: [0, 0, 0, 'later'] }]);
  expect(host.editMethods.editCell).not.toHaveBeenCalled();
  await act(async () => hydration.resolve(host.peer));
  await advance();
  await later;
  expect(host.cells.get('0:0:0')).toBe('later');
  expect(host.editMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'later']]);
});

for (const route of ['cell', 'formula', 'host'] as const) {
  it(`captures a ready operation once without preview and retires its draft after the covering frame (${route})`, async () => {
    const host = harness(true);
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    expect(api.hydrated).toBe(true);
    const operations = workbookEditPeerInternals.get(host.attached!)!;
    const applying = spyOn(operations, 'applyQueuedOp');
    const clock = spyOn(Date, 'now').mockReturnValue(1_750_000_000_000);
    const setting = spyOn(host.peer, 'setCalculationContext');
    const entropy = spyOn(globalThis.crypto, 'getRandomValues');
    const requestFrame = host.sessionMethods.frame.getMockImplementation()!;
    const adoption = deferred<void>();
    host.sessionMethods.frame.mockImplementation(async (viewport, options) => {
      await adoption.promise;
      return requestFrame(viewport, options);
    });
    const input = '=NOW()+RAND()';
    let pending: Promise<EditResult | null> | undefined;
    try {
      if (route === 'cell') reviewEdit(view, input);
      else if (route === 'formula') {
        fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value: input } });
        fireEvent.keyDown(view.getByTestId('xlsx-formula-input'), { key: 'Enter' });
      } else act(() => { pending = api.editCellAsync(0, 0, 0, input); });
      await act(async () => {});
      expect(host.preview).not.toHaveBeenCalled();
      expect(applying).toHaveBeenCalledTimes(1);
      const op = applying.mock.calls[0][0];
      if (!op.calculation) throw new Error('Missing operation calculation context');
      const calculation = structuredClone(op.calculation);
      const seed = op.calculation.randSeed;
      const captured = structuredClone(op);
      expect(op.method).toBe('editCell');
      expect(op.args).toEqual([0, 0, 0, input]);
      expect(op.calculation.nowSerial).toBe(
        (1_750_000_000_000 - new Date(1_750_000_000_000).getTimezoneOffset() * 60_000) / 86_400_000 + 25_569
      );
      expect(Number.isInteger(seed)).toBe(true);
      expect(seed).toBeGreaterThanOrEqual(0);
      expect(seed).toBeLessThanOrEqual(0xffff_ffff);
      expect(entropy).toHaveBeenCalledTimes(1);
      const words = entropy.mock.calls[0][0];
      if (!(words instanceof Uint32Array)) throw new Error('Missing u32 calculation seed');
      expect(seed).toBe(words[0]);
      expect(host.replay).toHaveBeenCalledTimes(1);
      expect(host.replay.mock.calls[0][0]).toEqual({ sequence: 1, op, calculation });
      expect(setting.mock.calls).toEqual([[calculation]]);
      expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, input]]);
      expect(view.getByTestId('xlsx-commit-preview').textContent).toBe(input);
      expect(host.peerMethods.displayList).not.toHaveBeenCalled();
      expect(host.peerMethods.save).not.toHaveBeenCalled();
      expect(host.session.save).not.toHaveBeenCalled();
      clock.mockReturnValue(1_760_000_000_000);
      await advance();
      await pending;
      expect(view.getByTestId('xlsx-commit-preview').textContent).toBe(input);
      expect(host.replay).toHaveBeenCalledTimes(1);
      await act(async () => adoption.resolve());
      await advance();
      expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
      expect(view.container.querySelector('[data-paint-source="worker"]')?.getAttribute('data-worker-sequence')).toBe('1');
      expect(host.preview).not.toHaveBeenCalled();
      expect(applying).toHaveBeenCalledTimes(1);
      expect(entropy).toHaveBeenCalledTimes(1);
      expect(setting.mock.calls).toEqual([[calculation]]);
      expect(applying.mock.calls[0][0]).toBe(op);
      expect(op).toEqual(captured);
      expect(op.calculation).toEqual(calculation);
      expect(host.replay.mock.calls[0][0].calculation).toEqual(calculation);
      expect(host.peerMethods.displayList).not.toHaveBeenCalled();
    } finally {
      try {
        await act(async () => adoption.resolve());
        await advance();
        await pending;
      } finally {
        entropy.mockRestore();
        setting.mockRestore();
        clock.mockRestore();
        applying.mockRestore();
      }
    }
  });
}

it('retires an unchanged ready draft without preview, mutation or replay', async () => {
  const host = harness(true);
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
  await opened();
  const adoption = deferred<void>();
  const frame = host.sessionMethods.frame.getMockImplementation()!;
  host.sessionMethods.frame.mockImplementation(async (viewport, options) => {
    await adoption.promise;
    return frame(viewport, options);
  });
  try {
    reviewEdit(view, 'initial');
    await act(async () => {});
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('initial');
    expect(host.preview).not.toHaveBeenCalled();
    expect(host.peerMethods.editCell).not.toHaveBeenCalled();
    expect(host.peerMethods.setCalculationContext).not.toHaveBeenCalled();
    expect(host.replay).not.toHaveBeenCalled();
    await advance();
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('initial');
    await act(async () => adoption.resolve());
    await advance();
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    expect(host.cells.get('0:0:0')).toBe('initial');
    expect(host.attached!.sentSequence).toBe(0);
    expect(host.peerMethods.displayList).not.toHaveBeenCalled();
  } finally {
    await act(async () => adoption.resolve());
    await advance();
  }
});

it('discards a refused ready draft and keeps the preceding draft until its worker frame', async () => {
  const host = harness(true);
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
    onReady={(value) => { api = value; }} />);
  await opened();
  const adoption = deferred<void>();
  const frame = host.sessionMethods.frame.getMockImplementation()!;
  host.sessionMethods.frame.mockImplementation(async (viewport, options) => {
    await adoption.promise;
    return frame(viewport, options);
  });
  try {
    reviewEdit(view, 'accepted');
    await act(async () => {});
    host.peerMethods.editCell.mockImplementationOnce(() => { throw new RangeError('Invalid cell input'); });
    reviewEdit(view, 'refused');
    await act(async () => {});
    expect(host.preview).not.toHaveBeenCalled();
    expect(host.replay).toHaveBeenCalledTimes(1);
    expect(host.replay.mock.calls[0][0].op).toMatchObject({ method: 'editCell', args: [0, 0, 0, 'accepted'] });
    expect(host.attached!.sentSequence).toBe(1);
    expect(api.failure).toBeNull();
    expect(view.getByTestId('xlsx-input-refusal').textContent).toContain('Invalid cell input');
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('refused');
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('accepted');
    expect(host.cells.get('0:0:0')).toBe('accepted');
    expect(host.cells.get('0:1:0')).toBeUndefined();
    await advance();
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('accepted');
    await act(async () => adoption.resolve());
    await advance();
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('refused');
    expect(host.peerMethods.displayList).not.toHaveBeenCalled();
  } finally {
    await act(async () => adoption.resolve());
    await advance();
  }
});

for (const route of ['host', 'clipboard'] as const) {
  it(`repairs a refused bulk predecessor before a same-cell UI refusal (${route})`, async () => {
    const host = harness(true);
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
      onReady={(value) => { api = value; }} />);
    await opened();
    const writes = deferred<void>();
    const adoption = deferred<void>();
    const frame = host.sessionMethods.frame.getMockImplementation()!;
    host.sessionMethods.frame.mockImplementation(async (...args) => {
      await adoption.promise;
      return frame(...args);
    });
    const col = route === 'host' ? 0 : 1;
    const originalRead = Object.getOwnPropertyDescriptor(navigator.clipboard, 'readText');
    const reading = mock(async () => '=batchFirstRefused()\t=hostRefused()');
    const flush = spyOn(host.attached!, 'flush');
    let blocked: Promise<void> | undefined;
    let refused: Promise<unknown> | undefined;
    try {
      await act(async () => { await api.editCellAsync(0, 0, col, 'accepted'); });
      await advance();
      expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('accepted');
      flush.mockReturnValue(writes.promise);
      act(() => { blocked = api.flush(); });
      await act(async () => {});
      expect(flush).toHaveBeenCalledTimes(1);
      if (route === 'host') {
        host.peerMethods.editCell.mockImplementationOnce(() => { throw new RangeError('Invalid host formula'); });
        act(() => { refused = api.editCellAsync(0, 0, 0, '=hostRefused()').catch((error) => error); });
      } else {
        host.peerMethods.editCells.mockImplementationOnce(() => { throw new RangeError('Invalid clipboard formula'); });
        Object.defineProperty(navigator.clipboard, 'readText', { configurable: true, value: reading });
        fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      }
      await act(async () => {});
      expect(view.getByTestId('xlsx-commit-preview').textContent)
        .toBe(route === 'host' ? '=hostRefused()' : '=batchFirstRefused()');
      if (route === 'clipboard') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
      host.peerMethods.editCell.mockImplementationOnce(() => { throw new RangeError('Invalid UI formula'); });
      reviewEdit(view, '=uiRefused()');
      await act(async () => {});
      expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('=uiRefused()');
      expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, col, 'accepted']]);
      expect(host.peerMethods.editCells).not.toHaveBeenCalled();
      await act(async () => writes.resolve());
      await advance();
      await blocked;
      if (route === 'host') expect(await refused).toBeInstanceOf(RangeError);
      else {
        expect(reading).toHaveBeenCalledTimes(1);
        expect(host.peerMethods.editCells.mock.calls).toEqual([[0, [
          { row: 0, col: 0, input: '=batchFirstRefused()' },
          { row: 0, col: 1, input: '=hostRefused()' },
        ]]]);
      }
      expect(host.peerMethods.editCell.mock.calls).toEqual(route === 'host' ? [
        [0, 0, 0, 'accepted'], [0, 0, 0, '=hostRefused()'], [0, 0, 0, '=uiRefused()'],
      ] : [[0, 0, 1, 'accepted'], [0, 0, 1, '=uiRefused()']]);
      expect(api.failure).toBeNull();
      expect(view.getByTestId('xlsx-input-refusal').textContent).toContain('Invalid UI formula');
      expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('=uiRefused()');
      expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('accepted');
      expect(host.cells.get(`0:0:${col}`)).toBe('accepted');
      if (route === 'clipboard') expect(host.cells.get('0:0:0')).toBe('initial');
      expect(host.replay).toHaveBeenCalledTimes(1);
      expect(host.replay.mock.calls[0][0].op.args).toEqual([0, 0, col, 'accepted']);
      expect(host.attached!.sentSequence).toBe(1);
      fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Escape' });
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
      expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('accepted');
      fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
      await advance();
      await act(async () => { await api.flush(); });
      expect(host.peerMethods.editCell).toHaveBeenCalledTimes(route === 'host' ? 3 : 2);
      expect(host.replay).toHaveBeenCalledTimes(1);
      expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('accepted');
      await act(async () => adoption.resolve());
      await advance();
      expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
      expect(view.container.querySelector('[data-paint-source="worker"]')?.getAttribute('data-worker-sequence')).toBe('1');
      expect(host.preview).not.toHaveBeenCalled();
      expect(host.peerMethods.displayList).not.toHaveBeenCalled();
    } finally {
      try {
        await act(async () => { writes.resolve(); adoption.resolve(); });
        await advance();
        await blocked;
        await refused;
      } finally {
        flush.mockRestore();
        host.sessionMethods.frame.mockImplementation(frame);
        if (originalRead) Object.defineProperty(navigator.clipboard, 'readText', originalRead);
        else Reflect.deleteProperty(navigator.clipboard, 'readText');
      }
    }
  });
}

it('keeps a mixed batch preview cleared after an earlier UI refusal until frame adoption', async () => {
  const host = harness(true);
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
    onReady={(value) => { api = value; }} />);
  await opened();
  const writes = deferred<void>();
  const adoption = deferred<void>();
  const frame = host.sessionMethods.frame.getMockImplementation()!;
  host.sessionMethods.frame.mockImplementation(async (...args) => {
    await adoption.promise;
    return frame(...args);
  });
  const flush = spyOn(host.attached!, 'flush');
  let blocked: Promise<void> | undefined;
  let batch: ReturnType<XlsxWorkerEditorApi['applyEdits']> | undefined;
  try {
    await act(async () => { await api.editCellAsync(0, 0, 0, 'accepted'); });
    await advance();
    expect(host.attached!.acknowledgedSequence).toBe(1);
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('accepted');
    flush.mockReturnValue(writes.promise);
    act(() => { blocked = api.flush(); });
    await act(async () => {});
    expect(flush).toHaveBeenCalledTimes(1);
    host.peerMethods.editCell.mockImplementationOnce(() => { throw new RangeError('Invalid UI formula'); });
    reviewEdit(view, '=uiRefused()');
    await act(async () => {});
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('=uiRefused()');
    const initial = reviewBatch('batch', host.peer.version());
    const request: XlsxEditRequest = { ...initial, steps: [...initial.steps.map(({ expect: _expect, ...step }) => step),
      { op: 'setFormulas', target: { sheetId: 'sheet:0',
        range: { kind: 'rowCol', start: { row: 0, col: 1 }, end: { row: 0, col: 1 } } }, formulas: [['24']] }] };
    act(() => { batch = api.applyEdits(request); });
    await act(async () => {});
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'accepted']]);
    expect(host.peerMethods.applyEdits).not.toHaveBeenCalled();
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    await act(async () => writes.resolve());
    await advance();
    await blocked;
    expect(await batch).toMatchObject({ ok: true, applied: true });
    expect(host.peerMethods.applyEdits).toHaveBeenCalledWith(request);
    expect(host.peerMethods.editCell.mock.calls).toEqual([
      [0, 0, 0, 'accepted'], [0, 0, 0, '=uiRefused()'],
    ]);
    expect(api.failure).toBeNull();
    expect(view.getByTestId('xlsx-input-refusal').textContent).toContain('Invalid UI formula');
    expect(host.cells.get('0:0:0')).toBe('batch');
    expect(host.attached!.acknowledgedSequence).toBe(2);
    expect(view.container.querySelector('[data-paint-source="worker"]')?.getAttribute('data-worker-sequence')).toBe('0');
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    await advance();
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    await act(async () => adoption.resolve());
    await advance();
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    expect(view.container.querySelector('[data-paint-source="worker"]')?.getAttribute('data-worker-sequence')).toBe('2');
    expect(painted[painted.length - 1].commands.some((command) =>
      command.op === 'text' && command.text === 'worker:batch')).toBe(true);
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Escape' });
    expect(api.cell(0, 0, 0)?.input).toBe('batch');
  } finally {
    try {
      await act(async () => { writes.resolve(); adoption.resolve(); });
      await advance();
      await Promise.all([blocked, batch]);
    } finally {
      flush.mockRestore();
      host.sessionMethods.frame.mockImplementation(frame);
    }
  }
});

it('unlinks two refused host predecessors before a same-cell UI refusal', async () => {
  const host = harness(true);
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
    onReady={(value) => { api = value; }} />);
  await opened();
  const writes = deferred<void>();
  const adoption = deferred<void>();
  const frame = host.sessionMethods.frame.getMockImplementation()!;
  host.sessionMethods.frame.mockImplementation(async (...args) => {
    await adoption.promise;
    return frame(...args);
  });
  const flush = spyOn(host.attached!, 'flush');
  let blocked: Promise<void> | undefined;
  let first: Promise<unknown> | undefined;
  let second: Promise<unknown> | undefined;
  try {
    await act(async () => { await api.editCellAsync(0, 0, 0, 'P'); });
    await advance();
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('P');
    flush.mockReturnValue(writes.promise);
    act(() => { blocked = api.flush(); });
    await act(async () => {});
    expect(flush).toHaveBeenCalledTimes(1);
    const h1 = new RangeError('Invalid H1 formula');
    const h2 = new RangeError('Invalid H2 formula');
    host.peerMethods.editCell.mockImplementationOnce(() => { throw h1; });
    host.peerMethods.editCell.mockImplementationOnce(() => { throw h2; });
    host.peerMethods.editCell.mockImplementationOnce(() => { throw new RangeError('Invalid R formula'); });
    act(() => { first = api.editCellAsync(0, 0, 0, '=H1()').catch((error) => error); });
    await act(async () => {});
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('=H1()');
    act(() => { second = api.editCellAsync(0, 0, 0, '=H2()').catch((error) => error); });
    await act(async () => {});
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('=H2()');
    reviewEdit(view, '=R()');
    await act(async () => {});
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('=R()');
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'P']]);
    await act(async () => writes.resolve());
    await advance();
    await blocked;
    expect(await first).toBe(h1);
    expect(await second).toBe(h2);
    expect(host.peerMethods.editCell.mock.calls).toEqual([
      [0, 0, 0, 'P'], [0, 0, 0, '=H1()'], [0, 0, 0, '=H2()'], [0, 0, 0, '=R()'],
    ]);
    expect(api.failure).toBeNull();
    expect(view.getByTestId('xlsx-input-refusal').textContent).toContain('Invalid R formula');
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('P');
    expect(host.cells.get('0:0:0')).toBe('P');
    expect(host.replay).toHaveBeenCalledTimes(1);
    expect(host.replay.mock.calls[0][0].op.args).toEqual([0, 0, 0, 'P']);
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Escape' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('P');
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await advance();
    await act(async () => { await api.flush(); });
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('P');
    expect(host.peerMethods.editCell).toHaveBeenCalledTimes(4);
    expect(host.replay).toHaveBeenCalledTimes(1);
    await act(async () => adoption.resolve());
    await advance();
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    expect(view.container.querySelector('[data-paint-source="worker"]')?.getAttribute('data-worker-sequence')).toBe('1');
    await act(async () => { await api.flush(); });
    expect(host.peerMethods.editCell).toHaveBeenCalledTimes(4);
    expect(host.replay).toHaveBeenCalledTimes(1);
    expect(host.preview).not.toHaveBeenCalled();
  } finally {
    try {
      await act(async () => { writes.resolve(); adoption.resolve(); });
      await advance();
      await Promise.all([blocked, first, second]);
    } finally {
      flush.mockRestore();
      host.sessionMethods.frame.mockImplementation(frame);
    }
  }
});

it('acknowledges dependent host edits with RAF suspended before a later frame failure', async () => {
  const host = harness(true);
  const failure = new Error('Host predecessor publication failed');
  const errors = mock((_error: Error) => {});
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
    onError={errors} onReady={(value) => { api = value; }} />);
  await opened();
  const publication = deferred<xlsx.WorkbookFrame>();
  let suspended: { mockRestore(): void } | undefined;
  let retained: Promise<EditResult | null> | undefined;
  try {
    await act(async () => { expect((await api.editCellAsync(0, 0, 0, 'first'))?.applied).toBe(true); });
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('first');
    const initialFrames = host.sessionMethods.frame.mock.calls.length;
    host.sessionMethods.frame.mockReturnValueOnce(publication.promise);
    await tick();
    const frames = host.sessionMethods.frame.mock.calls.length;
    expect(frames).toBe(initialFrames + 1);
    const count = painted.length;
    suspended = spyOn(globalThis, 'requestAnimationFrame').mockImplementation(() => ++nextAnimation);
    let settled = false;
    act(() => {
      retained = api.editCellAsync(0, 0, 1, 'retained');
      void retained.then(() => { settled = true; }, () => { settled = true; });
    });
    await act(async () => {});
    expect(host.sessionMethods.frame).toHaveBeenCalledTimes(frames);
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('retained');
    expect(settled).toBe(true);
    await expect(retained!).resolves.toMatchObject({ applied: true });
    await act(async () => publication.reject(failure));
    await expect(retained!).resolves.toMatchObject({ applied: true });
    expect(api.failure).toBe(failure);
    expect(host.session.failure).toBeUndefined();
    expect(host.attached!.state).toBe('failed');
    expect(errors).toHaveBeenCalledTimes(1);
    expect(suspended).toHaveBeenCalled();
    expect(painted).toHaveLength(count);
    await expect(api.flush()).rejects.toBe(failure);
    await expect(api.save()).rejects.toBe(failure);
    const replays = host.replay.mock.calls.length;
    let saved!: { bytes: Uint8Array; recovery: true };
    await act(async () => { saved = await api.recoverySave(); });
    expect(saved.recovery).toBe(true);
    expect(JSON.parse(new TextDecoder().decode(saved.bytes))).toEqual([
      ['0:0:0', 'first'], ['0:0:1', 'retained'],
    ]);
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'first'], [0, 0, 1, 'retained']]);
    expect(host.replay).toHaveBeenCalledTimes(replays);
    expect(host.session.save).not.toHaveBeenCalled();
    expect(host.peerMethods.save).toHaveBeenCalledTimes(1);
  } finally {
    await act(async () => publication.reject(failure));
    await retained?.catch(() => {});
    suspended?.mockRestore();
    await advance();
  }
});

it('rejects pending dependent edits on retirement while a predecessor acknowledgement is blocked', async () => {
  const host = harness(true);
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
    onReady={(value) => { api = value; }} />);
  await opened();
  const publication = deferred<xlsx.WorkbookFrame>();
  const acknowledged = deferred<WorkbookReplayReply>();
  void publication.promise.catch(() => {});
  let suspended: { mockRestore(): void } | undefined;
  let dependent: Promise<unknown> | undefined;
  let following: Promise<unknown> | undefined;
  try {
    await act(async () => { await api.editCellAsync(0, 0, 0, 'P'); });
    const frames = host.sessionMethods.frame.mock.calls.length;
    host.sessionMethods.frame.mockReturnValueOnce(publication.promise);
    await tick();
    expect(host.sessionMethods.frame).toHaveBeenCalledTimes(frames + 1);
    suspended = spyOn(globalThis, 'requestAnimationFrame').mockImplementation(() => ++nextAnimation);
    host.replay.mockReturnValueOnce(acknowledged.promise);
    let settled = false;
    act(() => {
      dependent = api.editCellAsync(0, 0, 1, 'H').catch((error) => error);
      following = api.editCellAsync(0, 0, 2, 'I').catch((error) => error);
      void dependent.then(() => { settled = true; });
    });
    await act(async () => {});
    expect(settled).toBe(false);
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'P'], [0, 0, 1, 'H']]);
    await act(async () => view.unmount());
    expect(await dependent).toMatchObject({ code: 'document-replaced' });
    expect(await following).toMatchObject({ code: 'document-replaced' });
    expect(host.session.dispose).not.toHaveBeenCalled();
    await act(async () => acknowledged.resolve({ sequence: 2, revision: 2, version: 2, result: undefined }));
    await waitFor(() => expect(host.session.dispose).toHaveBeenCalledTimes(1));
    expect(host.peerMethods.dispose).toHaveBeenCalledTimes(1);
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'P'], [0, 0, 1, 'H'], [0, 0, 2, 'I']]);
    expect(host.sessionMethods.frame).toHaveBeenCalledTimes(frames + 1);
    expect(api.failure).toBeNull();
  } finally {
    await act(async () => { acknowledged.resolve({ sequence: 2, revision: 2, version: 2, result: undefined }); view.unmount(); });
    await act(async () => publication.reject(new Error('Retired publication')));
    await Promise.all([dependent, following]);
    suspended?.mockRestore();
    await advance();
  }
});

it('acknowledges dependent host edits before an obsolete viewport frame rejects', async () => {
  const host = harness(true);
  const errors = mock((_error: Error) => {});
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
    onError={errors} onReady={(value) => { api = value; }} />);
  await opened();
  const publication = deferred<xlsx.WorkbookFrame>();
  void publication.promise.catch(() => {});
  let suspended: { mockRestore(): void } | undefined;
  let dependent: Promise<EditResult | null> | undefined;
  try {
    const scroll = view.getByTestId('xlsx-scroll');
    const frames = host.sessionMethods.frame.mock.calls.length;
    host.sessionMethods.frame.mockReturnValueOnce(publication.promise);
    fireEvent.scroll(scroll);
    await tick();
    expect(host.sessionMethods.frame).toHaveBeenCalledTimes(frames + 1);
    scroll.scrollLeft = 160;
    fireEvent.scroll(scroll);
    suspended = spyOn(globalThis, 'requestAnimationFrame').mockImplementation(() => ++nextAnimation);
    await act(async () => { expect((await api.editCellAsync(0, 0, 0, 'P'))?.applied).toBe(true); });
    let settled = false;
    act(() => {
      dependent = api.editCellAsync(0, 0, 1, 'H');
      void dependent.then(() => { settled = true; }, () => { settled = true; });
    });
    await act(async () => {});
    expect(settled).toBe(true);
    const count = painted.length;
    await act(async () => publication.reject(new Error('Old viewport frame failed')));
    await act(async () => { expect((await dependent)?.applied).toBe(true); });
    expect(api.failure).toBeNull();
    expect(host.attached!.state).toBe('ready');
    expect(errors).not.toHaveBeenCalled();
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'P'], [0, 0, 1, 'H']]);
    expect(host.replay.mock.calls.map(([envelope]) => envelope.op.args)).toEqual([[0, 0, 0, 'P'], [0, 0, 1, 'H']]);
    expect(host.sessionMethods.frame).toHaveBeenCalledTimes(frames + 1);
    expect(painted).toHaveLength(count);
    suspended.mockRestore();
    suspended = undefined;
    await advance();
    expect(host.sessionMethods.frame.mock.calls[frames + 1][0].x).toBe(160);
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
  } finally {
    await act(async () => publication.reject(new Error('Old viewport frame failed')));
    await dependent?.catch(() => {});
    suspended?.mockRestore();
    await advance();
  }
});

it('fails pending acknowledgements through the session after a real paint error', async () => {
  const host = harness(true);
  const failure = new Error('Current viewport failed');
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
    onReady={(value) => { api = value; }} />);
  await opened();
  const publication = deferred<xlsx.WorkbookFrame>();
  const acknowledged = deferred<WorkbookReplayReply>();
  host.sessionMethods.frame.mockReturnValueOnce(publication.promise);
  fireEvent.scroll(view.getByTestId('xlsx-scroll'));
  await tick();
  const suspended = spyOn(globalThis, 'requestAnimationFrame').mockImplementation(() => ++nextAnimation);
  host.replay.mockReturnValueOnce(acknowledged.promise);
  let applied!: Promise<unknown>;
  let retained!: Promise<unknown>;
  try {
    act(() => {
      applied = api.editCellAsync(0, 0, 0, 'applied').catch((error) => error);
      retained = api.editCellAsync(0, 0, 1, 'retained').catch((error) => error);
    });
    await act(async () => {});
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'applied']]);
    await act(async () => publication.reject(failure));
    expect(await applied).toBe(failure);
    expect(await retained).toBe(failure);
    expect(api.failure).toBe(failure);
    expect(host.attached!.state).toBe('failed');
    let saved!: { bytes: Uint8Array; recovery: true };
    await act(async () => { saved = await api.recoverySave(); });
    expect(JSON.parse(new TextDecoder().decode(saved.bytes))).toEqual([['0:0:0', 'applied'], ['0:0:1', 'retained']]);
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'applied'], [0, 0, 1, 'retained']]);
    expect(host.replay).toHaveBeenCalledTimes(1);
    expect(host.session.save).not.toHaveBeenCalled();
  } finally {
    await act(async () => acknowledged.resolve({ sequence: 1, revision: 1, version: 1, result: undefined }));
    suspended.mockRestore();
    await advance();
  }
});

it('admits dependent host edits on worker acknowledgements with RAF suspended', async () => {
  const host = harness(true);
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
    onReady={(value) => { api = value; }} />);
  await opened();
  const frames = host.sessionMethods.frame.mock.calls.length;
  const flush = spyOn(host.attached!, 'flush');
  const acknowledged = deferred<WorkbookReplayReply>();
  const suspended = spyOn(globalThis, 'requestAnimationFrame').mockImplementation(() => ++nextAnimation);
  host.replay.mockReturnValueOnce(acknowledged.promise);
  let first!: Promise<EditResult | null>;
  let second!: Promise<EditResult | null>;
  let settled = false;
  try {
    act(() => {
      first = api.editCellAsync(0, 0, 0, 'first');
      second = api.editCellAsync(0, 0, 1, 'second');
      void first.then(() => { settled = true; });
    });
    await act(async () => {});
    expect(settled).toBe(false);
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'first']]);
    expect(host.replay).toHaveBeenCalledTimes(1);
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('second');
    await act(async () => {
      acknowledged.resolve({ sequence: 1, revision: 1, version: 1, result: undefined });
      expect((await first)?.applied).toBe(true);
      expect((await second)?.applied).toBe(true);
    });
    expect(settled).toBe(true);
    expect(host.attached!.acknowledgedSequence).toBe(2);
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'first'], [0, 0, 1, 'second']]);
    expect(host.sessionMethods.frame).toHaveBeenCalledTimes(frames);
    expect(flush).not.toHaveBeenCalled();
    expect(host.preview).not.toHaveBeenCalled();
    expect(api.failure).toBeNull();
  } finally {
    await act(async () => acknowledged.resolve({ sequence: 1, revision: 1, version: 1, result: undefined }));
    await Promise.allSettled([first, second]);
    suspended.mockRestore();
    flush.mockRestore();
    await advance();
  }
});

it('flushes and saves hidden-tab bulk edits to different cells without any animation frame', async () => {
  const host = harness(true);
  let api!: XlsxWorkerEditorApi;
  render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
    onReady={(value) => { api = value; }} />);
  await opened();
  const workerCells = new Map([['0:0:0', 'initial']]);
  const replay = host.replay.getMockImplementation()!;
  host.replay.mockImplementation(async (envelope) => {
    const reply = await replay(envelope);
    if (envelope.op.method !== 'editCell') throw new Error('Unexpected bulk operation');
    const [sheet, row, col, input] = envelope.op.args;
    workerCells.set(`${sheet}:${row}:${col}`, input);
    return reply;
  });
  host.sessionSave.mockImplementation(async () => new TextEncoder().encode(JSON.stringify([...workerCells])));
  const frames = host.sessionMethods.frame.mock.calls.length;
  const count = painted.length;
  const suspended = spyOn(globalThis, 'requestAnimationFrame').mockImplementation(() => ++nextAnimation);
  const expected = Array.from({ length: 64 }, (_, index): [string, string] => [`0:${Math.floor(index / 8)}:${index % 8}`, `edit-${index}`]);
  try {
    await act(async () => {
      const edits = expected.map(([, input], index) => api.editCellAsync(0, Math.floor(index / 8), index % 8, input));
      const flushing = api.flush();
      const saving = api.save();
      expect((await Promise.all(edits)).every((result) => result?.applied)).toBe(true);
      await flushing;
      const saved = await saving;
      expect(saved).not.toBeNull();
      expect(JSON.parse(new TextDecoder().decode(saved!))).toEqual(expected);
    });
    expect([...host.cells]).toEqual(expected);
    expect(host.attached!.acknowledgedSequence).toBe(64);
    expect(host.replay).toHaveBeenCalledTimes(64);
    expect(host.session.save).toHaveBeenCalledTimes(1);
    expect(host.peerMethods.save).not.toHaveBeenCalled();
    expect(host.preview).not.toHaveBeenCalled();
    expect(host.sessionMethods.frame).toHaveBeenCalledTimes(frames);
    expect(painted).toHaveLength(count);
    expect(api.failure).toBeNull();
  } finally {
    suspended.mockRestore();
    await advance();
  }
});

it('applies three sequential dependent host edits without extra flushes or frames while animation frames are suspended', async () => {
  const host = harness(true);
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
    onReady={(value) => { api = value; }} />);
  await opened();
  const frames = host.sessionMethods.frame.mock.calls.length;
  const flush = spyOn(host.attached!, 'flush');
  const suspended = spyOn(globalThis, 'requestAnimationFrame').mockImplementation(() => ++nextAnimation);
  const count = painted.length;
  try {
    for (const input of ['first', 'second', 'third']) {
      await act(async () => { expect((await api.editCellAsync(0, 0, 0, input))?.applied).toBe(true); });
      expect(view.getByTestId('xlsx-commit-preview').textContent).toBe(input);
      expect(host.sessionMethods.frame).toHaveBeenCalledTimes(frames);
      expect(flush).not.toHaveBeenCalled();
    }
    expect(host.peerMethods.editCell.mock.calls).toEqual([
      [0, 0, 0, 'first'], [0, 0, 0, 'second'], [0, 0, 0, 'third'],
    ]);
    expect(host.replay.mock.calls.map(([envelope]) => [envelope.sequence, envelope.op.args])).toEqual([
      [1, [0, 0, 0, 'first']], [2, [0, 0, 0, 'second']], [3, [0, 0, 0, 'third']],
    ]);
    expect(host.cells.get('0:0:0')).toBe('third');
    expect(painted).toHaveLength(count);
    expect(api.failure).toBeNull();
  } finally {
    suspended.mockRestore();
    flush.mockRestore();
    await advance();
  }
});

it('prepares dependent ready host edits and resolves flush and save with suspended animation frames', async () => {
  const host = harness(true);
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
    onReady={(value) => { api = value; }} />);
  await opened();
  const originalRaf = globalThis.requestAnimationFrame;
  const suspended = spyOn(globalThis, 'requestAnimationFrame').mockImplementation(() => ++nextAnimation);
  let first: Promise<EditResult | null> | undefined;
  let second: Promise<EditResult | null> | undefined;
  let flushing: Promise<void> | undefined;
  let saving: Promise<Uint8Array | null> | undefined;
  try {
    expect(api.hydrated).toBe(true);
    act(() => { first = api.editCellAsync(0, 0, 0, 'first'); });
    await act(async () => { expect((await first)?.applied).toBe(true); });
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('first');
    const count = painted.length;
    act(() => { second = api.editCellAsync(0, 0, 0, 'second'); });
    await act(async () => { expect((await second)?.applied).toBe(true); });
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('second');
    act(() => { flushing = api.flush(); });
    await act(async () => { await flushing; });
    act(() => { saving = api.save(); });
    await act(async () => { expect(await saving).toEqual(new Uint8Array([0])); });
    expect(host.cells.get('0:0:0')).toBe('second');
    expect(host.peerMethods.editCell.mock.calls).toEqual([[0, 0, 0, 'first'], [0, 0, 0, 'second']]);
    expect(host.replay).toHaveBeenCalledTimes(2);
    expect(host.replay.mock.calls.map(([envelope]) => [envelope.sequence, envelope.op.args]))
      .toEqual([[1, [0, 0, 0, 'first']], [2, [0, 0, 0, 'second']]]);
    expect(host.session.save).toHaveBeenCalledTimes(1);
    expect(host.peerMethods.save).not.toHaveBeenCalled();
    expect(host.preview).not.toHaveBeenCalled();
    expect(suspended).toHaveBeenCalled();
    expect(painted).toHaveLength(count);
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('second');
    expect(api.failure).toBeNull();
  } finally {
    suspended.mockRestore();
    globalThis.requestAnimationFrame = originalRaf;
    await advance();
    await Promise.all([first, second, flushing, saving]);
  }
});

for (const successor of ['valid', 'refused'] as const) {
  it(`repairs superseded ready refusals while writes are blocked (${successor})`, async () => {
    const host = harness(true);
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
      onReady={(value) => { api = value; }} />);
    await opened();
    const writes = deferred<void>();
    const adoption = deferred<void>();
    const flush = spyOn(host.attached!, 'flush').mockReturnValue(writes.promise);
    const frame = host.sessionMethods.frame.getMockImplementation()!;
    host.sessionMethods.frame.mockImplementation(async (...args) => {
      await adoption.promise;
      return frame(...args);
    });
    let blocked: Promise<void> | undefined;
    try {
      act(() => { blocked = api.flush(); });
      await act(async () => {});
      expect(flush).toHaveBeenCalledTimes(1);
      host.peerMethods.editCell.mockImplementationOnce(() => { throw new RangeError('Invalid formula'); });
      if (successor === 'refused') {
        host.peerMethods.editCell.mockImplementationOnce(() => { throw new RangeError('Invalid formula'); });
      }
      reviewEdit(view, '=refused()');
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
      reviewEdit(view, successor === 'valid' ? 'valid' : '=refusedAgain()');
      await act(async () => {});
      expect(host.peerMethods.editCell).not.toHaveBeenCalled();
      expect(host.preview).not.toHaveBeenCalled();
      expect(view.getByTestId('xlsx-commit-preview').textContent).toBe(successor === 'valid' ? 'valid' : '=refusedAgain()');
      await act(async () => writes.resolve());
      await advance();
      await blocked;
      expect(api.failure).toBeNull();
      expect(host.peerMethods.editCell).toHaveBeenCalledTimes(2);
      if (successor === 'valid') {
        expect(host.replay).toHaveBeenCalledTimes(1);
        expect(host.replay.mock.calls[0][0].op).toMatchObject({ method: 'editCell', args: [0, 0, 0, 'valid'] });
        expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('valid');
      } else {
        expect(host.replay).not.toHaveBeenCalled();
        expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
        fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Escape' });
        fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
        fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
        expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('initial');
        fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
        await advance();
        expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('initial');
        expect(host.peerMethods.editCell).toHaveBeenCalledTimes(2);
        expect(host.replay).not.toHaveBeenCalled();
      }
      const value = successor === 'valid' ? 'valid' : 'initial';
      expect(host.cells.get('0:0:0')).toBe(value);
      expect(JSON.parse(new TextDecoder().decode(host.peer.save()))).toEqual([['0:0:0', value]]);
      await act(async () => adoption.resolve());
      await advance();
      expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
      expect(view.container.querySelector('[data-paint-source="worker"]')?.getAttribute('data-worker-sequence'))
        .toBe(successor === 'valid' ? '1' : '0');
      expect(host.peerMethods.displayList).not.toHaveBeenCalled();
    } finally {
      try {
        await act(async () => { writes.resolve(); adoption.resolve(); });
        await advance();
        await blocked;
      } finally { flush.mockRestore(); }
    }
  });
}

it('replaces a cold speculative bitmap with visible ready text until adoption', async () => {
  const host = harness(true);
  const hydration = deferred<WorkbookHandle>();
  host.hydrate.mockReturnValue(hydration.promise);
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
    onReady={(value) => { api = value; }} />);
  await opened();
  const adoption = deferred<void>();
  const frame = host.sessionMethods.frame.getMockImplementation()!;
  host.sessionMethods.frame.mockImplementation(async (...args) => {
    await adoption.promise;
    return frame(...args);
  });
  const write = host.peerMethods.editCell.getMockImplementation()!;
  const writing = spyOn(host.peer, 'editCell');
  try {
    reviewEdit(view, 'cold');
    await advance();
    const canvas = view.getByTestId('xlsx-commit-preview') as HTMLCanvasElement;
    expect(canvas.width).toBe(800);
    expect(canvas.dataset.previewReady).toBe('true');
    expect(painted.some((list) => list.commands.some((command) => command.op === 'text' && command.text === 'preview:cold'))).toBe(true);
    expect(host.preview).toHaveBeenCalledTimes(1);
    expect(writing).not.toHaveBeenCalled();
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(api.hydrated).toBe(true);
    expect(host.cells.get('0:0:0')).toBe('cold');
    writing.mockImplementation((sheet, row, col, input) => {
      const preview = view.getByTestId('xlsx-commit-preview');
      expect(preview).not.toBe(canvas);
      expect(canvas.isConnected).toBe(false);
      expect(Number.parseFloat(preview.style.width)).toBeGreaterThan(0);
      expect(Number.parseFloat(preview.style.height)).toBeGreaterThan(0);
      expect(preview.textContent).toBe('ready');
      return write(sheet, row, col, input);
    });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
    reviewEdit(view, 'ready');
    await act(async () => {});
    expect(writing).toHaveBeenCalledTimes(2);
    expect(host.replay).toHaveBeenCalledTimes(2);
    expect(host.preview).toHaveBeenCalledTimes(1);
    expect(host.cells.get('0:0:0')).toBe('ready');
    await advance();
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('ready');
    expect(canvas.isConnected).toBe(false);
    await act(async () => adoption.resolve());
    await advance();
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    expect(view.container.querySelector('[data-paint-source="worker"]')?.getAttribute('data-worker-sequence')).toBe('2');
    expect(host.peerMethods.displayList).not.toHaveBeenCalled();
  } finally {
    try {
      await act(async () => { hydration.resolve(host.peer); adoption.resolve(); });
      await advance();
    } finally { writing.mockRestore(); }
  }
});

it('restores an untouched ready predecessor and retires it after a same-cell refusal', async () => {
  const host = harness(true);
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false}
    onReady={(value) => { api = value; }} />);
  await opened();
  const adoption = deferred<void>();
  const frame = host.sessionMethods.frame.getMockImplementation()!;
  host.sessionMethods.frame.mockImplementation(async (...args) => {
    await adoption.promise;
    return frame(...args);
  });
  try {
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('initial');
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await advance();
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('initial');
    expect(host.replay).not.toHaveBeenCalled();
    expect(host.peerMethods.editCell).not.toHaveBeenCalled();
    host.peerMethods.editCell.mockImplementationOnce(() => { throw new RangeError('Invalid formula'); });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowUp' });
    reviewEdit(view, '=refused()');
    await advance();
    expect(api.failure).toBeNull();
    expect(view.getByTestId('xlsx-input-refusal').textContent).toContain('Invalid formula');
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('=refused()');
    expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('initial');
    expect(host.cells.get('0:0:0')).toBe('initial');
    expect(host.peerMethods.editCell).toHaveBeenCalledTimes(1);
    expect(host.replay).not.toHaveBeenCalled();
    expect(host.attached!.sentSequence).toBe(0);
    await act(async () => adoption.resolve());
    await advance();
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('=refused()');
    expect(host.preview).not.toHaveBeenCalled();
    expect(host.peerMethods.displayList).not.toHaveBeenCalled();
  } finally {
    await act(async () => adoption.resolve());
    await advance();
  }
});

it('keeps the preview until worker adoption and never paints a peer frame', async () => {
  const host = harness();
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
  await opened();
  host.peerMethods.displayList.mockImplementation(() => { throw new Error('Peer grid paint forbidden'); });
  const worker = deferred<xlsx.WorkbookFrame>();
  host.sessionMethods.frame.mockReturnValue(worker.promise);
  reviewEdit(view, 'worker owned');
  await advance();
  expect(host.cells.get('0:0:0')).toBe('worker owned');
  expect(host.preview).not.toHaveBeenCalled();
  expect(view.getByTestId('xlsx-commit-preview').textContent).toBe('worker owned');
  expect(host.peerMethods.displayList).not.toHaveBeenCalled();
  expect(painted.every((list) => list.commands.every((cmd) => cmd.op !== 'text' || !cmd.text.startsWith('peer:')))).toBe(true);
  const [viewport, options] = host.sessionMethods.frame.mock.calls[host.sessionMethods.frame.mock.calls.length - 1]!;
  await act(async () => worker.resolve({ sheet: options?.sheet ?? 0, viewport, version: 'v2', epoch: 2, sequence: 1,
    displayList: { width: 800, height: 600, commands: [{ op: 'text', text: 'worker:worker owned', x: 8, y: 18, fontSize: 11, color: '#000000' }] } }));
  await advance();
  expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
  expect(host.peerMethods.displayList).not.toHaveBeenCalled();
  expect(painted[painted.length - 1]!.commands).toContainEqual({ op: 'text', text: 'worker:worker owned', x: 8, y: 18, fontSize: 11, color: '#000000' });
});

for (const route of ['save', 'saveAsync', 'command', 'shortcut', 'toolbar'] as const) {
  it(`routes every normal XLSX save through the worker and saves the peer only during recovery (${route})`, async () => {
    const host = harness(true);
    let api!: XlsxWorkerEditorApi;
    const delivery = mock((_bytes: Uint8Array) => {});
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen onReady={(value) => { api = value; }} onSave={delivery} />);
    await opened();
    const ack = deferred<WorkbookReplayReply>();
    host.replay.mockReturnValueOnce(ack.promise);
    reviewEdit(view, 'saved');
    let saving: Promise<unknown> | undefined;
    if (route === 'save' || route === 'saveAsync') act(() => { saving = api[route](); });
    else if (route === 'command') act(() => { saving = api.commands.execute('save', null); });
    else if (route === 'shortcut') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 's', ctrlKey: true });
    else fireEvent.click(view.getByTestId('xlsx-save'));
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

for (const [input, normalized] of [['001', '1'], ['true', 'TRUE'], ['false', 'FALSE'], ["'quoted", 'quoted']]) {
  for (const route of ['cell', 'host', 'paste'] as const) {
    it(`retires an engine-normalized preview by operation identity (${route}, ${input})`, async () => {
      const host = harness(false, true);
      let api!: XlsxWorkerEditorApi;
      const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
      await opened();
      const frame = host.sessionMethods.frame.getMockImplementation()!;
      const adoption = deferred<void>();
      host.sessionMethods.frame.mockImplementation(async (viewport, options) => { await adoption.promise; return frame(viewport, options); });
      if (route === 'cell') reviewEdit(view, input);
      else if (route === 'paste') {
        replaceClipboard('readText', mock(async () => input));
        fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      } else void api.editCellAsync(0, 0, 0, input);
      await advance();
      expect(host.cells.get('0:0:0')).toBe(normalized);
      expect(view.getByTestId('xlsx-commit-preview')).toBeTruthy();
      await act(async () => adoption.resolve());
      await advance();
      expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
      expect(view.container.querySelector('[data-paint-source="worker"]')?.getAttribute('data-worker-sequence')).toBe('1');
      expect(host.peerMethods.displayList).not.toHaveBeenCalled();
      expect(painted[painted.length - 1].commands.some((cmd) => cmd.op === 'text' && cmd.text === `worker:${normalized}`)).toBe(true);
    });
  }
}

for (const route of ['draft', 'bulk'] as const) {
  it(`captures local serials for preview and replay (${route})`, async () => {
    const host = harness(true);
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    const clock = spyOn(Date, 'now').mockReturnValue(Date.UTC(2026, 0, 1, 23, 30));
    const timezone = spyOn(Date.prototype, 'getTimezoneOffset').mockReturnValue(-120);
    restorers.push(() => timezone.mockRestore(), () => clock.mockRestore());
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    const input = '=NOW()';
    let pending: Promise<unknown> | undefined;
    if (route === 'draft') reviewEdit(view, input);
    else act(() => { pending = api.applyEdits(reviewBatch(input)); });
    await advance();
    const op = host.preview.mock.calls[0][2][0];
    expect(op.method).toBe(route === 'draft' ? 'editCell' : 'applyEdits');
    if (!op.calculation) throw new Error('Missing preview calculation context');
    expect(op.calculation.nowSerial).toBe(Date.UTC(2026, 0, 2, 1, 30) / 86_400_000 + 25_569);
    const calculation = structuredClone(op.calculation);
    clock.mockReturnValue(Date.UTC(2026, 0, 3));
    await act(async () => hydration.resolve(host.peer));
    await advance();
    await pending;
    expect(host.replay).toHaveBeenCalledTimes(1);
    expect(host.replay.mock.calls[0][0].op).toEqual(op);
    expect(host.replay.mock.calls[0][0].calculation).toEqual(calculation);
    expect(host.peerMethods.setCalculationContext).toHaveBeenCalledWith(calculation);
  });
}

for (const route of ['cell', 'formula'] as const) {
  it(`reuses the volatile draft operation as soon as hydration resolves (${route})`, async () => {
    const host = harness(true);
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen onReady={(value) => { api = value; }} />);
    await opened();
    const input = '=RANDBETWEEN(1,1000000)';
    if (route === 'cell') reviewEdit(view, input);
    else {
      fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value: input } });
      fireEvent.keyDown(view.getByTestId('xlsx-formula-input'), { key: 'Enter' });
    }
    await advance();
    expect(view.getByTestId('xlsx-commit-preview')).toBeTruthy();
    expect(api.hydrated).toBe(false);
    const op = host.preview.mock.calls[0][2][0];
    expect(op).toMatchObject({ method: 'editCell', args: [0, 0, 0, input] });
    expect(op.calculation).toBeTruthy();
    if (!op.calculation) throw new Error('Missing preview calculation context');
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect(host.replay).toHaveBeenCalledTimes(1);
    expect(host.replay.mock.calls[0][0].op).toEqual(op);
    expect(host.replay.mock.calls[0][0].calculation).toEqual(op.calculation);
    expect(host.peerMethods.setCalculationContext).toHaveBeenCalledWith(op.calculation);
    expect(host.cells.get('0:0:0')).toBe(input);
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    expect(view.container.querySelector('[data-paint-source="worker"]')?.getAttribute('data-worker-sequence')).toBe('1');
    expect(host.peerMethods.displayList).not.toHaveBeenCalled();
  });
}

it('reuses the complete volatile host operation and calculation context for preview and replay', async () => {
  const host = harness(true);
  const hydration = deferred<WorkbookHandle>();
  host.hydrate.mockReturnValue(hydration.promise);
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
  await opened();
  const pending = api.editCellAsync(0, 0, 0, '=RANDBETWEEN(1,1000000)');
  await advance();
  expect(view.getByTestId('xlsx-commit-preview')).toBeTruthy();
  const op = host.preview.mock.calls[0][2][0];
  expect(op.calculation).toBeTruthy();
  if (!op.calculation) throw new Error('Missing preview calculation context');
  await act(async () => hydration.resolve(host.peer));
  await advance();
  await pending;
  expect(host.replay.mock.calls[0][0].op).toEqual(op);
  expect(host.replay.mock.calls[0][0].calculation).toEqual(op.calculation);
  expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
});

for (const step of ['patchStyle', 'setFormulas'] as const) {
  it(`omits a host batch preview when complete semantics cannot be materialized (${step})`, async () => {
    const host = harness();
    const hydration = deferred<WorkbookHandle>();
    host.hydrate.mockReturnValue(hydration.promise);
    let api!: XlsxWorkerEditorApi;
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
    await opened();
    const initial = reviewBatch('styled');
    const target = initial.steps[0].target;
    const request: XlsxEditRequest = { ...initial, steps: [...initial.steps,
      step === 'patchStyle' ? { op: 'patchStyle', target, patch: { fontSize: 24 } } :
        { op: 'setFormulas', target, formulas: [['24']] }] };
    const pending = api.applyEdits(request);
    await advance();
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
    expect(host.preview).not.toHaveBeenCalled();
    await act(async () => hydration.resolve(host.peer));
    await advance();
    expect((await pending)?.ok).toBe(true);
    expect(host.editMethods.applyEdits).toHaveBeenCalledWith(request);
    expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
  });
}

it('retires a held preview after scroll and zoom change when the covering worker frame is adopted', async () => {
  const host = harness();
  let api!: XlsxWorkerEditorApi;
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} onReady={(value) => { api = value; }} />);
  await opened();
  const frame = host.sessionMethods.frame.getMockImplementation()!;
  const adoption = deferred<void>();
  host.sessionMethods.frame.mockImplementation(async (viewport, options) => { await adoption.promise; return frame(viewport, options); });
  reviewEdit(view, 'new viewport');
  await advance();
  expect(view.getByTestId('xlsx-commit-preview')).toBeTruthy();
  const scroll = view.getByTestId('xlsx-scroll');
  scroll.scrollTop = 96;
  scroll.scrollLeft = 192;
  fireEvent.scroll(scroll);
  let zoom!: Promise<unknown>;
  act(() => { zoom = api.commands.execute('zoom', { scale: 1.25 }); });
  await advance();
  await zoom;
  expect(view.getByTestId('xlsx-commit-preview')).toBeTruthy();
  await act(async () => adoption.resolve());
  await advance();
  expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
  const canvas = view.container.querySelector('[data-paint-source="worker"]');
  expect(canvas?.getAttribute('data-worker-sequence')).toBe('1');
  expect(canvas?.getAttribute('data-worker-zoom')).toBe('1.25');
  expect(host.sessionMethods.frame.mock.calls[host.sessionMethods.frame.mock.calls.length - 1][0]).toMatchObject({ x: 153.6, y: 76.8 });
  expect(host.peerMethods.displayList).not.toHaveBeenCalled();
  expect(painted[painted.length - 1].commands.some((cmd) => cmd.op === 'text' && cmd.text === 'worker:new viewport')).toBe(true);
});

it('restores a pre-hydration cut preview when clipboard writing is refused', async () => {
  const host = harness();
  const hydration = deferred<WorkbookHandle>();
  host.hydrate.mockReturnValue(hydration.promise);
  const clipboard = promisedClipboard();
  const writing = deferred<void>();
  clipboard.write.mockReturnValue(writing.promise);
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen showToolbar={false} />);
  await opened();
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true });
  await advance();
  expect(view.getByTestId('xlsx-commit-preview')).toBeTruthy();
  expect(host.cells.get('0:0:0')).toBe('initial');
  await act(async () => writing.reject(new DOMException('Clipboard denied', 'NotAllowedError')));
  await advance();
  expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
  expect(host.editMethods.editCells).not.toHaveBeenCalled();
  await act(async () => hydration.resolve(host.peer));
  await advance();
  expect(view.queryByTestId('xlsx-commit-preview')).toBeNull();
  expect(view.getByTestId('xlsx-input-refusal').textContent).toMatch(/clipboard denied/i);
  expect(host.cells.get('0:0:0')).toBe('initial');
  expect(host.editMethods.editCells).not.toHaveBeenCalled();
  expect(host.replay).not.toHaveBeenCalled();
  expect(host.peerMethods.displayList).not.toHaveBeenCalled();
});


describe('peerless keyboard navigation and proposal freshness', () => {
  const refreshes = (host: ReturnType<typeof peerlessHarness>) => host.sent.filter((command) =>
    command.kind === 'applyOps' && 'toolbar' in command.payload && command.payload.toolbar.id === 'refreshProposals');
  const available = () => {
    const proposals = spyOn(xlsx, 'isProposalsAvailable').mockReturnValue(true);
    restorers.push(() => proposals.mockRestore());
  };

  for (const route of ['drag', 'shift-click', 'shift-arrow'] as const) {
    for (const reverse of [false, true]) {
      it(`keeps the name box and formula bar at the ${reverse ? 'reverse' : 'forward'} ${route} anchor`, async () => {
        const host = peerlessHarness(false, true);
        const anchor = reverse ? { row: 3, col: 2 } : { row: 1, col: 1 };
        const focus = reverse ? { row: 1, col: 1 } : { row: 3, col: 2 };
        host.cells.set(`0:${anchor.row}:${anchor.col}`, 'anchor input');
        host.cells.set(`0:${focus.row}:${focus.col}`, 'focus input');
        const view = await host.mount();
        const scroll = view.getByTestId('xlsx-scroll');
        const point = (cell: { row: number; col: number }) => ({ clientX: cell.col * 96 + 10, clientY: cell.row * 24 + 10 });
        fireEvent.mouseDown(scroll, point(anchor));
        if (route === 'drag') fireEvent.mouseMove(scroll, point(focus));
        fireEvent.mouseUp(window);
        if (route === 'shift-click') {
          fireEvent.mouseDown(scroll, { ...point(focus), shiftKey: true });
          fireEvent.mouseUp(window);
        } else if (route === 'shift-arrow') {
          fireEvent.keyDown(scroll, { key: reverse ? 'ArrowLeft' : 'ArrowRight', shiftKey: true });
          for (let index = 0; index < 2; index++) {
            fireEvent.keyDown(scroll, { key: reverse ? 'ArrowUp' : 'ArrowDown', shiftKey: true });
          }
        }
        const selection = { sheetId: 'sheet:0', activeCell: anchor,
          range: { kind: 'rowCol' as const, start: { row: 1, col: 1 }, end: { row: 3, col: 2 } } };
        expect(host.sent[host.sent.length - 1]).toMatchObject({ kind: 'select', payload: { selection } });
        await act(async () => { for (const command of [...host.sent]) host.ack(command, { paint: true }); });
        const name = view.getByTestId('xlsx-name-box') as HTMLInputElement;
        const formula = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
        expect(name.value).toBe(reverse ? 'C4' : 'B2');
        await waitFor(() => expect(formula.value).toBe('anchor input'));
        expect(host.cellInputs).toHaveBeenLastCalledWith(0, reverse ? 'C4' : 'B2');
        expect(view.getByTestId('xlsx-selection').style.width).toBe('192px');
        expect(view.getByTestId('xlsx-selection').style.height).toBe('72px');
        const outline = view.getByTestId('xlsx-overlay-host').children[1] as HTMLElement;
        expect(outline.style.left).toBe(`${anchor.col * 96}px`);
        expect(outline.style.top).toBe(`${anchor.row * 24}px`);
        fireEvent.keyDown(scroll, { key: reverse ? 'ArrowDown' : 'ArrowUp', shiftKey: true });
        expect(name.value).toBe(reverse ? 'C4' : 'B2');
        expect(host.sent[host.sent.length - 1]).toMatchObject({ kind: 'select', payload: { selection: {
          activeCell: anchor, range: { kind: 'rowCol', start: { row: reverse ? 2 : 1, col: 1 },
            end: { row: reverse ? 3 : 2, col: 2 } },
        } } });
      });
    }
  }

  for (const [key, shiftKey, address] of [
    ['ArrowRight', false, 'C2'], ['ArrowLeft', false, 'A2'],
    ['ArrowDown', false, 'B3'], ['ArrowUp', false, 'B1'],
    ['Tab', false, 'C2'], ['Tab', true, 'A2'],
    ['Enter', false, 'B3'], ['Enter', true, 'B1'],
  ] as const) {
    it(`moves from the range anchor with ${shiftKey ? 'Shift+' : ''}${key}`, async () => {
      const host = peerlessHarness();
      const view = await host.mount();
      const scroll = view.getByTestId('xlsx-scroll');
      fireEvent.mouseDown(scroll, { clientX: 106, clientY: 34 });
      fireEvent.mouseMove(scroll, { clientX: 202, clientY: 82 });
      fireEvent.mouseUp(window);
      fireEvent.keyDown(scroll, { key, shiftKey });
      expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe(address);
      expect(view.getByTestId('xlsx-selection').style.width).toBe('96px');
      expect(view.getByTestId('xlsx-selection').style.height).toBe('24px');
    });
  }

  it('restores the edited cell from a redo acknowledgement after Enter navigation', async () => {
    const host = peerlessHarness();
    host.cells.set('0:2:2', '303');
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.mouseDown(scroll, { clientX: 202, clientY: 58 });
    fireEvent.mouseUp(window);
    await act(async () => host.ack(host.sent[0]));
    fireEvent.keyDown(scroll, { key: 'F2' });
    const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    await waitFor(() => expect(input.value).toBe('303'));
    fireEvent.change(input, { target: { value: '4242' } });
    fireEvent.keyDown(input, { key: 'Enter' });
    const edit = host.sent[1];
    const editedCell = { sheetId: 'sheet:0', activeCell: { row: 2, col: 2 },
      range: { kind: 'rowCol' as const, start: { row: 2, col: 2 }, end: { row: 2, col: 2 } } };
    expect(edit).toMatchObject({ kind: 'editCell', payload: { args: [0, 2, 2, '4242'], target: editedCell } });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('C4');
    await act(async () => host.ack(edit, { paint: true, selection: { ...editedCell,
      activeCell: { row: 3, col: 2 }, range: { kind: 'a1', a1: 'C4' } } }));
    fireEvent.mouseDown(scroll, { clientX: 106, clientY: 34 });
    fireEvent.mouseUp(window);
    await act(async () => host.ack(host.sent[2]));
    const undoing = host.commands().execute('undo', null);
    const undo = await host.commandAt(3);
    host.cells.set('0:2:2', '303');
    await act(async () => host.ack(undo, { paint: true, selection: editedCell }));
    await undoing;
    fireEvent.mouseDown(scroll, { clientX: 202, clientY: 58 });
    fireEvent.mouseUp(window);
    await act(async () => host.ack(host.sent[4]));
    const redoing = host.commands().execute('redo', null);
    const redo = await host.commandAt(5);
    expect(redo.kind).toBe('redo');
    host.cells.set('0:2:2', '4242');
    await act(async () => host.ack(redo, { paint: true, selection: editedCell }));
    await redoing;
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('C3');
    await waitFor(() => expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).value).toBe('4242'));
    expect(view.getByTestId('xlsx-selection').style.height).toBe('24px');
  });

  it('outlines the full merge containing the focused cell', async () => {
    const host = peerlessHarness();
    const frame = host.frame.getMockImplementation()!;
    host.frame.mockImplementation(async (...args) => ({ ...await frame(...args),
      mergedRanges: [{ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }] }));
    const view = await host.mount();
    const outline = view.getByTestId('xlsx-overlay-host').children[1] as HTMLElement;
    expect(outline.style.width).toBe('192px');
    expect(outline.style.height).toBe('48px');
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
    expect(outline.style.width).toBe('96px');
    expect(outline.style.height).toBe('24px');
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('C1');
  });

  const mergeA1B2 = [{ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }];

  it('commits typing at a merge anchor and Enter steps past the merge', async () => {
    const host = peerlessHarness();
    host.selectionData.mockImplementation(async (_sheet, row, col) => ({ edge: null, region: null,
      merge: row <= 1 && col <= 1 ? { top: 0, left: 0, bottom: 1, right: 1 } : null }));
    const frame = host.frame.getMockImplementation()!;
    host.frame.mockImplementation(async (...args) => ({ ...await frame(...args), mergedRanges: mergeA1B2 }));
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    const name = view.getByTestId('xlsx-name-box') as HTMLInputElement;
    for (const value of ['v', 'w']) {
      await act(async () => {
        fireEvent.keyDown(scroll, { key: 'Delete' });
        host.ack(host.sent[host.sent.length - 1]!);
      });
      await act(async () => fireEvent.keyDown(scroll, { key: value }));
      fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
      await act(async () => {});
      const edit = host.sent.filter((command) => command.kind === 'editCell').slice(-1)[0]!;
      await act(async () => host.ack(edit));
    }
    expect(host.sent.filter((command) => command.kind === 'editCell').map((command) => command.payload.args))
      .toEqual([[0, 0, 0, 'v'], [0, 2, 0, 'w']]);
    expect(name.value).toBe('A4');
    await act(async () => {
      fireEvent.keyDown(scroll, { key: 'ArrowUp' });
      fireEvent.keyDown(scroll, { key: 'ArrowUp' });
    });
    expect(name.value).toBe('A1');
    await act(async () => { fireEvent.keyDown(scroll, { key: 'Tab' }); });
    expect(name.value).toBe('C1');
  });

  it('routes a click on a covered cell and the typing after it to the merge anchor', async () => {
    const host = peerlessHarness();
    host.selectionData.mockImplementation(async (_sheet, row, col) => ({ edge: null, region: null,
      merge: row <= 1 && col <= 1 ? { top: 0, left: 0, bottom: 1, right: 1 } : null }));
    const frame = host.frame.getMockImplementation()!;
    host.frame.mockImplementation(async (...args) => ({ ...await frame(...args), mergedRanges: mergeA1B2 }));
    const view = await host.mount();
    const scroll = view.getByTestId('xlsx-scroll');
    fireEvent.keyDown(scroll, { key: 'Tab' });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('C1');
    fireEvent.mouseDown(scroll, { clientX: 150, clientY: 40 });
    fireEvent.mouseUp(scroll, { clientX: 150, clientY: 40 });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
    await act(async () => {
      fireEvent.keyDown(scroll, { key: 'Delete' });
      host.ack(host.sent[host.sent.length - 1]!);
    });
    fireEvent.keyDown(scroll, { key: 'x' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    await act(async () => {});
    expect(host.sent.filter((command) => command.kind === 'editCell').map((command) => command.payload.args))
      .toEqual([[0, 0, 0, 'x']]);
  });

  it('anchors a legacy formula IME draft whose focus is a covered cell', async () => {
    const host = harness();
    host.peer.mergedRanges = () => mergeA1B2;
    const frame = host.sessionMethods.frame.getMockImplementation()!;
    host.sessionMethods.frame.mockImplementation(async (...args) => ({ ...await frame(...args), mergedRanges: mergeA1B2 }));
    const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen />);
    await opened();
    await advance();
    const scroll = view.getByTestId('xlsx-scroll');
    for (const key of ['ArrowRight', 'ArrowDown']) fireEvent.keyDown(scroll, { key });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('C2');
    fireEvent.keyDown(scroll, { key: 'ArrowLeft', shiftKey: true });
    const input = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
    act(() => input.focus());
    fireEvent.compositionStart(input);
    fireEvent.change(input, { target: { value: '変' } });
    fireEvent.compositionEnd(input);
    await act(async () => {});
    fireEvent.keyDown(input, { key: 'Enter' });
    await advance();
    expect(host.cells.get('0:0:0')).toBe('変');
    expect(host.cells.has('0:1:1')).toBe(false);
  });

  for (const source of ['cell', 'formula'] as const) {
    it(`refreshes proposals after ${source} input, undo and redo`, async () => {
      available();
      const host = peerlessHarness();
      const view = await host.mount();
      if (source === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'z' });
      const input = view.getByTestId(source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input');
      fireEvent.change(input, { target: { value: 'changed' } });
      fireEvent.keyDown(input, { key: 'Enter' });
      await act(async () => host.ack(host.sent[0], { paint: true }));
      expect(refreshes(host)).toHaveLength(1);
      await act(async () => host.ack(refreshes(host)[0], { result: { proposals: [] } }));
      for (const id of ['undo', 'redo'] as const) {
        const changing = host.commands().execute(id, null);
        await act(async () => {});
        const command = host.sent.slice().reverse().find((command) => command.kind === id)!;
        expect(command).toBeDefined();
        const count = refreshes(host).length;
        await act(async () => host.ack(command, { paint: true }));
        await changing;
        expect(refreshes(host)).toHaveLength(count + 1);
        await act(async () => host.ack(refreshes(host)[count], { result: { proposals: [] } }));
      }
      expect(host.hydrate).not.toHaveBeenCalled();
      expect(host.attach).not.toHaveBeenCalled();
    });
  }

  it('coalesces refreshes behind edits newer than a pending proposal read', async () => {
    available();
    const host = peerlessHarness();
    const view = await host.mount();
    await act(async () => { await host.commands().execute('proposalsPanel', { open: true }); });
    const stale = refreshes(host)[0];
    expect(stale).toBeDefined();
    for (const value of ['a', 'b']) {
      await act(async () => { fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: value }); });
      fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    }
    const edits = host.sent.filter((command) => command.kind === 'editCell');
    await act(async () => {
      host.ack(stale, { result: { proposals: [{ id: 'old', agentId: 'agent', note: 'Old preview', cells: [] }] } });
      for (const command of edits) host.ack(command, { paint: true });
    });
    expect(refreshes(host)).toHaveLength(2);
    expect(view.queryByText('Old preview')).toBeNull();
    const current = refreshes(host)[1];
    expect(current.seq).toBeGreaterThan(edits[1].seq);
    expect(current.frame).toBeUndefined();
    await act(async () => host.ack(current, { result: { proposals: [{ id: 'fresh', agentId: 'agent', note: 'Current preview', cells: [] }] } }));
    expect(view.getByText('Current preview')).toBeDefined();
    expect(refreshes(host)).toHaveLength(2);
  });

  it('prints an acknowledged edit without waiting for proposal reads to paint', async () => {
    available();
    const host = peerlessHarness();
    const view = await host.mount();
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'p' });
    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    const edit = host.sent[0];
    await act(async () => host.ack(edit, { paint: true }));
    expect(refreshes(host)).toHaveLength(1);
    await act(async () => host.ack(refreshes(host)[0], { result: { proposals: [] } }));
    const print = mock(() => {});
    replacePrint(print);
    const printing = host.commands().execute('print', null);
    await act(async () => {});
    expect(host.session.whenPresented).toHaveBeenLastCalledWith(edit.seq);
    await printing;
    expect(print).toHaveBeenCalledTimes(1);
  });
});

describe('peerless composition keyboard focus', () => {
  it('preserves host focus when composed input completes before a queued save can dispatch', async () => {
    let coordinator: ReturnType<typeof peerlessInput.createPeerlessWorkerInputCoordinator> | undefined;
    const original = peerlessInput.createPeerlessWorkerInputCoordinator;
    const creating = spyOn(peerlessInput, 'createPeerlessWorkerInputCoordinator').mockImplementation((hooks) => {
      coordinator = original(hooks);
      return coordinator;
    });
    restorers.push(() => creating.mockRestore());
    const host = peerlessHarness();
    const view = await host.mount();
    const dialog = render(<input aria-label="Pending save dialog" />);
    const hostInput = dialog.getByRole('textbox', { name: 'Pending save dialog' });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    await waitFor(() => expect((input as HTMLInputElement).value).toBe('initial'));
    const entered = deferred<void>();
    const reply = deferred<void>();
    let reading!: Promise<void>;
    act(() => {
      reading = coordinator!.runAfterPendingInput(() => { entered.resolve(); return reply.promise; }, true);
    });
    await entered.promise;
    act(() => input.focus());
    fireEvent.compositionStart(input);
    fireEvent.change(input, { target: { value: 'composed' } });
    let saving!: Promise<Uint8Array>;
    try {
      await act(async () => {
        saving = coordinator!.save();
        void saving.catch(() => {});
        hostInput.focus();
        fireEvent.compositionEnd(input);
      });
      expect(host.sent).toEqual([]);
      expect(document.activeElement === hostInput).toBe(true);
    } finally {
      await act(async () => { reply.resolve(); await reading; });
    }
    const edit = await host.commandAt(0);
    await act(async () => { host.ack(edit, { paint: true }); await saving; });
    expect(edit).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'composed'] } });
    expect(document.activeElement === hostInput).toBe(true);
    expect(host.session.save).toHaveBeenCalledTimes(1);
    expect(host.hydrate).not.toHaveBeenCalled();
    expect(host.attach).not.toHaveBeenCalled();
  });

  it('preserves host focus when F2 waits behind a fence and composed input completes before a queued save', async () => {
    let coordinator: ReturnType<typeof peerlessInput.createPeerlessWorkerInputCoordinator> | undefined;
    const original = peerlessInput.createPeerlessWorkerInputCoordinator;
    const creating = spyOn(peerlessInput, 'createPeerlessWorkerInputCoordinator').mockImplementation((hooks) => {
      coordinator = original(hooks);
      return coordinator;
    });
    restorers.push(() => creating.mockRestore());
    const host = peerlessHarness();
    const view = await host.mount();
    const dialog = render(<input aria-label="Pending save dialog" />);
    const hostInput = dialog.getByRole('textbox', { name: 'Pending save dialog' });
    const entered = deferred<void>();
    const reply = deferred<void>();
    let reading!: Promise<void>;
    act(() => {
      reading = coordinator!.runAfterPendingInput(() => { entered.resolve(); return reply.promise; }, true);
    });
    await entered.promise;
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
    const input = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    expect(input.value).toBe('');
    act(() => input.focus());
    fireEvent.compositionStart(input);
    fireEvent.change(input, { target: { value: 'composed' } });
    const ended = new Event('compositionend', { bubbles: true });
    Object.defineProperty(ended, 'data', { value: 'composed' });
    let saving!: Promise<Uint8Array>;
    try {
      await act(async () => {
        saving = coordinator!.save();
        void saving.catch(() => {});
        hostInput.focus();
        fireEvent(input, ended);
      });
      expect(host.sent).toEqual([]);
      expect(document.activeElement === hostInput).toBe(true);
    } finally {
      await act(async () => { reply.resolve(); await reading; });
    }
    const edit = await host.commandAt(0);
    await act(async () => { host.ack(edit, { paint: true }); await saving; });
    expect(edit).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'initialcomposed'] } });
    expect(document.activeElement === hostInput).toBe(true);
    expect(host.session.save).toHaveBeenCalledTimes(1);
  });

  for (const source of ['cell', 'formula'] as const) {
    it(`restores ${source} keyboard focus after submitting composed text once and running a toolbar command`, async () => {
      const host = peerlessHarness(true);
      const view = await host.mount();
      if (source === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
      const input = view.getByTestId(source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input') as HTMLInputElement;
      if (source === 'cell') await waitFor(() => expect(input.value).toBe('initial'));
      act(() => input.focus());
      fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: 'unfinished' } });
      const blur = spyOn(input, 'blur');
      input.addEventListener('blur', () => {
        input.value = 'final composed text';
        fireEvent.compositionEnd(input);
      }, { once: true });
      restorers.push(() => blur.mockRestore());
      const bold = view.container.querySelector<HTMLButtonElement>('button[aria-label="Bold"]')!;
      await act(async () => fireEvent.click(bold));
      expect(blur).toHaveBeenCalledTimes(1);
      await waitFor(() => expect(host.sent).toHaveLength(2));
      expect(host.sent[0]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'final composed text'] } });
      expect(host.sent[1]).toMatchObject({ kind: 'applyOps', payload: { toolbar: { id: 'bold' } } });
      await act(async () => {
        fireEvent.change(input, { target: { value: 'final composed text' } });
        fireEvent.compositionEnd(input);
        for (const command of host.sent) host.ack(command, { paint: true });
      });
      await advance();
      expect(host.sent.filter(({ kind }) => kind === 'editCell')).toHaveLength(1);
      expect(host.cells.get('0:0:0')).toBe('final composed text');
      expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).value).toBe('final composed text');
      expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
      const sink = view.getByTestId(source === 'cell' ? 'xlsx-scroll' : 'xlsx-formula-input');
      expect(document.activeElement).toBe(sink);
      if (source === 'cell') {
        fireEvent.keyDown(document.activeElement!, { key: 'F2' });
        await waitFor(() => expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('final composed text'));
      } else {
        fireEvent.change(document.activeElement!, { target: { value: 'final composed textx' } });
        expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).value).toBe('final composed textx');
        fireEvent.keyDown(document.activeElement!, { key: 'Enter' });
        await waitFor(() => expect(host.sent.filter(({ kind }) => kind === 'editCell')).toHaveLength(2));
        const edits = host.sent.filter(({ kind }) => kind === 'editCell');
        expect(edits[edits.length - 1]).toMatchObject({ payload: { args: [0, 0, 0, 'final composed textx'] } });
      }
    });

    it(`preserves host dialog focus after saving ${source} IME composition`, async () => {
      const host = peerlessHarness(false, false, true);
      const dialog = render(<input aria-label="Save dialog input" />);
      const hostInput = dialog.getByRole('textbox', { name: 'Save dialog input' });
      const saved = mock(() => hostInput.focus());
      const view = await host.mount(saved);
      if (source === 'cell') fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
      const input = view.getByTestId(source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input') as HTMLInputElement;
      if (source === 'cell') await waitFor(() => expect(input.value).toBe('initial'));
      act(() => input.focus());
      fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: 'unfinished' } });
      input.addEventListener('blur', () => {
        input.value = 'final composed text';
        fireEvent.compositionEnd(input);
      }, { once: true });
      await act(async () => fireEvent.click(view.getByTestId('xlsx-save')));
      await waitFor(() => expect(host.sent).toHaveLength(1));
      expect(host.sent[0]).toMatchObject({ kind: 'editCell', payload: { args: [0, 0, 0, 'final composed text'] } });
      await act(async () => {
        fireEvent.change(input, { target: { value: 'final composed text' } });
        fireEvent.compositionEnd(input);
        host.ack(host.sent[0], { paint: true });
      });
      await advance();
      await waitFor(() => expect(saved).toHaveBeenCalledTimes(1));
      expect(host.session.save).toHaveBeenCalledTimes(1);
      expect(host.sent).toHaveLength(1);
      expect(host.cells.get('0:0:0')).toBe('final composed text');
      expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
      expect(document.activeElement === hostInput).toBe(true);
    });
  }

});

describe('remote proposal freshness', () => {
  for (const status of ['applied', 'noop'] as const) {
    it(`keeps remote proposals ahead of a pending refresh after an ${status} update`, async () => {
      const available = spyOn(xlsx, 'isProposalsAvailable').mockReturnValue(true);
      restorers.push(() => available.mockRestore());
      const host = collaborativePeerlessHarness();
      const view = await host.mount();
      await act(async () => { await host.commands().execute('proposalsPanel', { open: true }); });
      const stale = await host.commandAt(0);
      const updating = host.replica.async!.applyUpdate(Uint8Array.of(3));
      const remote = await host.commandAt(1);
      await act(async () => {
        host.ack(stale, { result: { proposals: [{ id: 'old', agentId: 'agent', note: 'Old remote preview', cells: [] }] } });
        host.ack(remote, { status, result: { remoteUpdate: true,
          proposals: [{ id: 'fresh', agentId: 'agent', note: 'Fresh remote preview', cells: [] }] } });
        await updating;
      });
      expect(view.queryByText('Old remote preview') === null).toBe(true);
      expect(view.queryByText('Fresh remote preview') !== null).toBe(true);
      expect(host.sent).toHaveLength(2);
      expect(host.hydrate).not.toHaveBeenCalled();
      expect(host.attach).not.toHaveBeenCalled();
    });
  }
});

it('applies a later arrow after a pending data jump in keyboard order', async () => {
  const host = peerlessHarness();
  const answer = deferred<xlsx.SelectionData>();
  host.selectionData.mockReturnValueOnce(answer.promise);
  const view = await host.mount();
  const scroll = view.getByTestId('xlsx-scroll');
  fireEvent.keyDown(scroll, { key: 'ArrowDown', ctrlKey: true });
  fireEvent.keyDown(scroll, { key: 'ArrowRight' });
  await act(async () => answer.resolve({ edge: { row: 179, col: 0 } }));
  expect(host.sent.filter((command) => command.kind === 'select').map((command) => command.payload.selection.activeCell))
    .toEqual([{ row: 179, col: 0 }, { row: 179, col: 1 }]);
  expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('B180');
});


it('resolves a legacy name-box destination before typing while hydration waits', async () => {
  const host = harness();
  const hydration = deferred<WorkbookHandle>();
  const answer = deferred<xlsx.SelectionData>();
  host.hydrate.mockReturnValue(hydration.promise);
  host.sessionMethods.selectionData.mockImplementation(async (_sheet, row, col) =>
    row >= 99 && row <= 100 && col >= 3 && col <= 4 ? answer.promise : { merge: null });
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen />);
  await opened();
  const name = view.getByTestId('xlsx-name-box') as HTMLInputElement;
  fireEvent.change(name, { target: { value: 'E101' } });
  fireEvent.keyDown(name, { key: 'Enter' });
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 't' });
  expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
  await act(async () => answer.resolve({ merge: { top: 99, left: 3, bottom: 100, right: 4 } }));
  expect(name.value).toBe('D100');
  expect((view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('t');
  await act(async () => fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' }));
  expect(name.value).toBe('D102');
  await act(async () => hydration.resolve(host.peer));
  await advance();
  expect(host.cells.get('0:99:3')).toBe('t');
  expect(host.cells.has('0:100:4')).toBe(false);
});

it('commits legacy queued a Enter b Enter into two cells while hydration and lookup wait', async () => {
  const host = harness();
  const hydration = deferred<WorkbookHandle>();
  const answer = deferred<xlsx.SelectionData>();
  host.hydrate.mockReturnValue(hydration.promise);
  host.sessionMethods.selectionData.mockImplementation(async (_sheet, row) => row === 99 ? answer.promise : { merge: null });
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen />);
  await opened();
  const name = view.getByTestId('xlsx-name-box');
  fireEvent.change(name, { target: { value: 'D100' } });
  fireEvent.keyDown(name, { key: 'Enter' });
  const scroll = view.getByTestId('xlsx-scroll');
  for (const key of ['a', 'Enter', 'b', 'Enter']) fireEvent.keyDown(scroll, { key });
  expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
  await act(async () => answer.resolve({ merge: { top: 99, left: 3, bottom: 100, right: 4 } }));
  await advance();
  await act(async () => hydration.resolve(host.peer));
  await advance();
  expect([host.cells.get('0:99:3'), host.cells.get('0:101:3')]).toEqual(['a', 'b']);
});

it('keeps painted plain arrows local before and after legacy hydration', async () => {
  const host = harness();
  const hydration = deferred<WorkbookHandle>();
  host.hydrate.mockReturnValue(hydration.promise);
  const region = spyOn(host.peer, 'selectionData');
  const merges = spyOn(host.peer, 'mergedRanges');
  restorers.push(() => region.mockRestore(), () => merges.mockRestore());
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen />);
  await opened();
  const scroll = view.getByTestId('xlsx-scroll');
  const steps = () => {
    for (const key of ['ArrowDown', 'ArrowRight', 'ArrowUp', 'ArrowLeft']) fireEvent.keyDown(scroll, { key });
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1');
  };
  host.sessionMethods.selectionData.mockClear();
  steps();
  expect(host.sessionMethods.selectionData).not.toHaveBeenCalled();
  await act(async () => hydration.resolve(host.peer));
  await advance();
  region.mockClear();
  merges.mockClear();
  steps();
  expect(region).not.toHaveBeenCalled();
  expect(merges.mock.calls.filter(([, range]) => range === 'A1:XFD1048576')).toHaveLength(1);
  expect(host.sessionMethods.selectionData).not.toHaveBeenCalled();
});

it('keeps every legacy formula character visible and commits abc after a pending lookup', async () => {
  const host = harness();
  const hydration = deferred<WorkbookHandle>();
  const answer = deferred<xlsx.SelectionData>();
  host.hydrate.mockReturnValue(hydration.promise);
  host.sessionMethods.selectionData.mockImplementation(async (_sheet, row) => row >= 99 ? answer.promise : { merge: null });
  const view = render(<XlsxEditor legacySyncApi file={file} experimentalWorkerOpen />);
  await opened();
  const name = view.getByTestId('xlsx-name-box');
  fireEvent.change(name, { target: { value: 'E101' } });
  fireEvent.keyDown(name, { key: 'Enter' });
  const input = view.getByTestId('xlsx-formula-input') as HTMLInputElement;
  act(() => input.focus());
  fireEvent.change(input, { target: { value: '' } });
  for (const character of 'abc') fireEvent.change(input, { target: { value: input.value + character } });
  const displayed = input.value;
  fireEvent.keyDown(input, { key: 'Enter' });
  await act(async () => answer.resolve({ merge: { top: 99, left: 3, bottom: 100, right: 4 } }));
  await advance();
  await act(async () => hydration.resolve(host.peer));
  await advance();
  expect({ displayed, committed: host.cells.get('0:99:3') }).toEqual({ displayed: 'abc', committed: 'abc' });
});

for (const status of ['applied', 'noop'] as const) {
  it(`peerless ${status} remote update reports onChange and onDocumentChange like the in-thread route`, async () => {
    const host = collaborativePeerlessHarness();
    const versions: string[] = [];
    let changes = 0;
    await host.mount(undefined, false, { onChange: () => changes++,
      onDocumentChange: ({ version }) => versions.push(version) });
    const updating = host.replica.async!.applyUpdate(new TextEncoder().encode('remote edit'));
    const remote = await host.commandAt(0);
    await act(async () => {
      host.ack(remote, { status, paint: true });
      await updating;
    });
    expect(changes).toBe(status === 'applied' ? 1 : 0);
    expect(versions).toEqual(status === 'applied' ? [host.session.confirmedBasis.version] : []);
  });
}

it('peerless host props set initial zoom and root style and notify versions', async () => {
  const host = peerlessHarness();
  const changes: string[] = [];
  const view = await host.mount(undefined, false, { initialZoom: 2, style: { height: 321 },
    onDocumentChange: ({ version }) => changes.push(version) });
  expect(host.commands().getState('zoom').value).toBe(2);
  expect(view.getByRole('application').style.height).toBe('321px');
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x' });
  fireEvent.change(view.getByTestId('xlsx-cell-editor'), { target: { value: 'host change' } });
  fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
  const command = await host.commandAt(0);
  await act(async () => { host.ack(command, { paint: true }); });
  expect(changes).toEqual([host.session.confirmedBasis.version]);
});

for (const [callback, downloadOnSave, downloads] of [
  [true, undefined, 0], [false, undefined, 1], [true, true, 1], [false, false, 0],
] as const) {
  it(`peerless command save delivers callback=${callback} download=${downloadOnSave}`, async () => {
    const host = peerlessHarness();
    const saves: Uint8Array[] = [];
    const click = spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
    restorers.push(() => click.mockRestore());
    await host.mount(callback ? (bytes) => saves.push(bytes) : undefined, false, { downloadOnSave });
    await act(async () => { await host.commands().execute('save', null); });
    expect(saves).toHaveLength(callback ? 1 : 0);
    expect(click).toHaveBeenCalledTimes(downloads);
  });
}

it('peerless shows the supplied indicator while the workbook opens', async () => {
  const host = peerlessHarness();
  const pending = deferred<WorkbookSession>();
  host.open.mockReturnValue(pending.promise);
  const view = render(host.asyncEditor({ loadingIndicator: <span>Workbook opening</span> }));
  await waitFor(() => expect(view.queryByText('Workbook opening')).not.toBeNull());
  await act(async () => { pending.resolve(host.session); });
  await opened();
  await waitFor(() => expect(view.queryByText('Workbook opening')).toBeNull());
});
