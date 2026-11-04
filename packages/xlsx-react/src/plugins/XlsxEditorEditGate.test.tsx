import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, it, mock, spyOn } from 'bun:test';
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import * as xlsx from '@betteroffice/xlsx';
import { initWasm, type WorkbookSession } from '@betteroffice/xlsx';
import { createTestWorkbookSession } from '../../../xlsx/src/session/testHelpers';
import { editableWorkbookSessionBackend } from '../worker/useEditableSessionWorkbook';
import { XlsxWorkerEditor } from '../worker/XlsxWorkerEditor';
import type { XlsxAsyncEditorApi } from '../worker/createAsyncEditorApi';
import type { XlsxWorkerEditorApi } from '../worker/createWorkerEditorApi';
import {
  XlsxEditor,
  defineXlsxPlugin,
  type XlsxEditorApi,
  type XlsxPlugin,
  type XlsxPluginCellCommit,
  type XlsxPluginCellRef,
  type XlsxPluginCommand,
  type XlsxPluginCommitDecision,
  type XlsxPluginContext,
  type XlsxPluginEditGateContext,
  type XlsxPluginError,
} from '../index';
import { MOUNTED } from '../__fixtures__/mounted';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');

const WASM = resolve(import.meta.dir, '../../../xlsx/src/wasm/generated/xlsx_wasm_bg.wasm');
const VIEWPORT = { width: 800, height: 600 };
const LAYOUT = [
  ['clientWidth', VIEWPORT.width],
  ['clientHeight', VIEWPORT.height],
] as const;
const PROTECTED = 'This cell is on a protected sheet.';
const GATE_CELL_LIMIT = 50_000;
const NOT_ALLOWED = 'This edit is not allowed.';
const STARTING = 'The editor is still starting. Try again in a moment.';
const EDIT_GATE_UNSUPPORTED = 'edit gating is not supported in the legacy worker editor';
const originalGetContext = HTMLCanvasElement.prototype.getContext;
const originalLayout = LAYOUT.map(
  ([property]) => [property, Object.getOwnPropertyDescriptor(HTMLElement.prototype, property)] as const
);

function stubContext(): CanvasRenderingContext2D {
  return new Proxy({} as Record<string, unknown>, {
    get(target, key) {
      if (key === 'measureText') return () => ({ width: 0 });
      if (typeof key === 'string' && key in target) return target[key];
      return () => {};
    },
    set(target, key, value) {
      target[key as string] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
}

beforeAll(async () => {
  HTMLCanvasElement.prototype.getContext = (() => stubContext()) as unknown as HTMLCanvasElement['getContext'];
  for (const [property, value] of LAYOUT) {
    Object.defineProperty(HTMLElement.prototype, property, { configurable: true, get: () => value });
  }
  await initWasm(new Uint8Array(readFileSync(WASM)));
}, 30_000);

const restorers: (() => void)[] = [];
const sessions: WorkbookSession[] = [];

afterEach(async () => {
  await act(async () => { cleanup(); });
  for (const session of sessions.splice(0)) await session.dispose();
  for (const restore of restorers.splice(0).reverse()) restore();
});

afterAll(async () => {
  HTMLCanvasElement.prototype.getContext = originalGetContext;
  for (const [property, descriptor] of originalLayout) {
    if (descriptor) Object.defineProperty(HTMLElement.prototype, property, descriptor);
  }
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

async function book({ protectedSheet = false, sparse = false } = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('xl/styles.xml', '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf applyProtection="1"><protection locked="0"/></xf></cellXfs></styleSheet>');
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sparse ? '<row r="1000000"><c r="A1000000"><v>9</v></c></row>' : '<row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>2</v></c></row>'}</sheetData>${protectedSheet ? '<sheetProtection sheet="1"/>' : ''}</worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}

interface Gate {
  /** The plugin id when it differs from the key, so two entries can share one. */
  id?: string;
  revision?: number;
  initialize?(): unknown;
  /** Runs when the activation ends, after its cleanup was logged. */
  cleanup?(): void;
  onLoad?(): void;
  createState?(): unknown;
  commands?: XlsxPluginCommand<unknown>[];
  cellEditable?(cell: XlsxPluginCellRef, context: XlsxPluginEditGateContext<unknown>): unknown;
  beforeCellCommit?(commit: XlsxPluginCellCommit, context: XlsxPluginEditGateContext<unknown>): unknown;
}

type GateLog = { loaded: Set<string>; cleanups: string[]; contexts: Record<string, XlsxPluginContext<unknown>> };

function gatePlugin(id: string, gate: Gate, log: GateLog): XlsxPlugin {
  return defineXlsxPlugin<unknown>({
    id,
    ...(gate.revision !== undefined && { revision: gate.revision }),
    createState: gate.createState ?? (() => null),
    ...(gate.commands && { commands: gate.commands }),
    initialize(context) {
      log.contexts[id] = context;
      context.onCleanup((reason) => { log.cleanups.push(`${id}:${reason}`); gate.cleanup?.(); });
      return gate.initialize?.() as Promise<void> | undefined;
    },
    onEvent(_context, event) {
      if (event.type !== 'load') return;
      log.loaded.add(id);
      gate.onLoad?.();
    },
    ...(gate.cellEditable && {
      cellEditable: (context, cell) => gate.cellEditable!(cell, context) as boolean,
    }),
    ...(gate.beforeCellCommit && {
      beforeCellCommit: (context, commit) => gate.beforeCellCommit!(commit, context) as XlsxPluginCommitDecision,
    }),
  });
}

type Editor = 'in-thread' | 'peerless';

interface Book {
  view: ReturnType<typeof render>;
  surface: HTMLElement;
  errors: XlsxPluginError[];
  log: GateLog;
  select(col: number, to?: number): Promise<void>;
  /** Selects column A down to its content or the sheet edge. */
  selectColumn(): Promise<void>;
  cellAt(row: number, col: number): Promise<string>;
  editor(): HTMLInputElement | null;
  type(value: string): Promise<void>;
  settle(): Promise<void>;
  input(col: number): Promise<string>;
  saved(): Promise<unknown>;
  /** Re-renders with new definitions for the same plugin ids, keeping the open workbook. */
  replace(gates: Record<string, Gate>): Promise<void>;
  formula(): HTMLInputElement;
  refusal(): string | null;
  /** A real grid mouse-down and mouse-up on a row 1 cell; in-thread editor only. */
  click?(col: number): Promise<void>;
  /** The worker session behind the peerless editor. */
  session?: WorkbookSession;
}

const A1 = ['A1', 'B1', 'C1', 'D1', 'E1'];

type ClipItem = { types: string[]; getType(type: string): Promise<Blob> };

/**
 * A clipboard holding `text` whose writes land only once every part of the item resolved; reads
 * of its text wait for `hold`.
 */
function fakeClipboard(text = '', hold?: Promise<void>) {
  const plain = (value: string): ClipItem => ({ types: ['text/plain'], getType: async () => new Blob([value], { type: 'text/plain' }) });
  let items = [plain(text)];
  const clipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
  const item = Object.getOwnPropertyDescriptor(globalThis, 'ClipboardItem');
  Object.defineProperty(globalThis, 'ClipboardItem', {
    configurable: true,
    value: class {
      constructor(private readonly data: Record<string, Blob | Promise<Blob>>) {}
      get types() { return Object.keys(this.data); }
      async getType(type: string) { return this.data[type]!; }
    },
  });
  Object.defineProperty(navigator, 'clipboard', {
    configurable: true,
    value: {
      readText: async () => {
        const read = items[0]!;
        await hold;
        return (await read.getType('text/plain')).text();
      },
      read: async () => items,
      write: async (next: ClipItem[]) => {
        for (const entry of next) for (const type of entry.types) await entry.getType(type);
        items = next;
      },
      writeText: async (value: string) => { items = [plain(value)]; },
    },
  });
  restorers.push(() => {
    if (clipboard) Object.defineProperty(navigator, 'clipboard', clipboard);
    else Reflect.deleteProperty(navigator, 'clipboard');
    if (item) Object.defineProperty(globalThis, 'ClipboardItem', item);
    else Reflect.deleteProperty(globalThis, 'ClipboardItem');
  });
  return { text: async () => (await items[0]!.getType('text/plain')).text() };
}

const noop: XlsxPluginCommand<unknown> = {
  id: 'noop', label: 'Noop', mutatesDocument: false, execute: () => ({ ok: true, status: 'executed' }),
};
/** A definition error: two commands share one id. */
const DUPLICATE_COMMANDS = [noop, noop];

const reported = (errors: readonly XlsxPluginError[]) => [...new Set(errors.map((error) => `${error.pluginId}:${error.phase}`))];

/** Queues animation frames, stubs painting and returns a pump that runs the queued frames. */
function stubFrames(): () => Promise<void> {
  const frames = new Map<number, FrameRequestCallback>();
  let nextFrame = 0;
  const raf = globalThis.requestAnimationFrame;
  const cancel = globalThis.cancelAnimationFrame;
  const resize = globalThis.ResizeObserver;
  globalThis.requestAnimationFrame = (callback) => { const id = ++nextFrame; frames.set(id, callback); return id; };
  globalThis.cancelAnimationFrame = (id) => { frames.delete(id); };
  globalThis.ResizeObserver = class { observe() {} disconnect() {} } as unknown as typeof ResizeObserver;
  restorers.push(() => {
    globalThis.requestAnimationFrame = raf;
    globalThis.cancelAnimationFrame = cancel;
    globalThis.ResizeObserver = resize;
  });
  const paint = spyOn(xlsx, 'paintDisplayList').mockImplementation(() => {});
  const proposals = spyOn(xlsx, 'isProposalsAvailable').mockReturnValue(false);
  restorers.push(() => paint.mockRestore(), () => proposals.mockRestore());
  return () => act(async () => {
    const callbacks = [...frames.values()];
    frames.clear();
    for (const callback of callbacks) callback(0);
  });
}

async function mountGated(
  editor: Editor, gates: Record<string, Gate>,
  { protectedSheet = false, sparse = false, starting = 0 } = {}
): Promise<Book> {
  const errors: XlsxPluginError[] = [];
  const log: GateLog = { loaded: new Set<string>(), cleanups: [], contexts: {} };
  const plugins = Object.entries(gates).map(([key, gate]) => gatePlugin(gate.id ?? key, gate, log));
  const file = await book({ protectedSheet, sparse });
  const onPluginError = (error: XlsxPluginError) => { errors.push(error); };
  const common = (view: ReturnType<typeof render>) => ({
    view, errors, log,
    surface: view.getByTestId('xlsx-scroll'),
    editor: () => view.queryByTestId('xlsx-cell-editor') as HTMLInputElement | null,
    formula: () => view.getByTestId('xlsx-formula-input') as HTMLInputElement,
    refusal: () => view.queryByTestId('xlsx-input-refusal')?.textContent ?? null,
  });
  if (editor === 'in-thread') {
    let api: XlsxEditorApi | undefined;
    const onReady = (ready: XlsxEditorApi) => { api = ready; };
    const element = (next: XlsxPlugin[]) => <XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={file}
      plugins={next} onPluginError={onPluginError} onReady={onReady} />;
    const view = render(element(plugins));
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    await waitFor(() => expect(log.loaded.size).toBe(plugins.length - starting), MOUNTED);
    const base = common(view);
    return {
      ...base,
      select: (col, to = col) => act(async () => {
        api!.selectCells(0, { anchor: { row: 0, col }, focus: { row: 0, col: to } });
      }),
      selectColumn: () => act(async () => {
        api!.selectCells(0, { anchor: { row: 0, col: 0 }, focus: { row: 1_048_575, col: 0 } });
      }),
      cellAt: async (row, col) => api!.handle.cell(0, row, col).input,
      type: async (value) => {
        fireEvent.keyDown(base.surface, { key: 'F2' });
        const input = await waitFor(() => view.getByTestId('xlsx-cell-editor') as HTMLInputElement);
        fireEvent.change(input, { target: { value } });
        fireEvent.keyDown(input, { key: 'Enter' });
      },
      settle: () => act(async () => { await api!.flushPendingInput().catch(() => {}); }),
      click: async (col) => {
        const probe = xlsx.openWorkbook(file);
        let rect: { x: number; y: number; w: number; h: number } | null;
        try { rect = xlsx.cellRect(probe.displayList({ x: 0, y: 0, ...VIEWPORT }).grid!, 0, col); } finally { probe.dispose(); }
        if (!rect) throw new Error('cell outside the painted window');
        await act(async () => {
          fireEvent.mouseDown(base.surface, { clientX: rect!.x + rect!.w / 2, clientY: rect!.y + rect!.h / 2 });
          fireEvent.mouseUp(window);
        });
      },
      input: async (col) => api!.handle.cell(0, 0, col).input,
      saved: async () => {
        let result: unknown;
        await act(async () => { result = await api!.commands.execute('save', null); });
        return result;
      },
      replace: (next) => act(async () => {
        view.rerender(element(Object.entries(next).map(([id, gate]) => gatePlugin(id, gate, log))));
      }),
    };
  }
  const pump = stubFrames();
  const session = await createTestWorkbookSession(file, undefined, { includeEditState: true });
  sessions.push(session);
  const open = spyOn(editableWorkbookSessionBackend, 'open').mockResolvedValue(session);
  restorers.push(() => open.mockRestore());
  const workerInput = async (col: number) => (await session.call.cellInputs(0, A1[col])).cells[0][0].input;
  const workerCell = async (row: number, col: number) =>
    (await session.call.cellInputs(0, `${String.fromCharCode(65 + col)}${row + 1}`)).cells[0][0].input;
  const ready = mock((_api: XlsxAsyncEditorApi) => {});
  const element = (next: XlsxPlugin[]) => <XlsxWorkerEditor backend="peerless" file={file} plugins={next}
    onPluginError={onPluginError} onReady={ready} />;
  const view = render(element(plugins));
  await waitFor(async () => { await pump(); expect(ready).toHaveBeenCalledTimes(1); }, MOUNTED);
  await waitFor(() => expect(log.loaded.size).toBe(plugins.length - starting), MOUNTED);
  const api = ready.mock.calls[0][0];
  const base = common(view);
  return {
    ...base,
    select: async (col, to = col) => {
      const name = view.getByTestId('xlsx-name-box');
      fireEvent.change(name, { target: { value: A1[col] } });
      fireEvent.keyDown(name, { key: 'Enter' });
      await waitFor(async () => {
        await pump();
        expect(session.selection?.range).toMatchObject({ kind: 'rowCol', start: { row: 0, col }, end: { row: 0, col } });
      });
      for (let next = col + 1; next <= to; next++) {
        fireEvent.keyDown(base.surface, { key: 'ArrowRight', shiftKey: true });
        await waitFor(async () => {
          await pump();
          expect(session.selection?.range).toMatchObject({ kind: 'rowCol', start: { row: 0, col }, end: { row: 0, col: next } });
        });
      }
    },
    selectColumn: async () => {
      const name = view.getByTestId('xlsx-name-box');
      fireEvent.change(name, { target: { value: 'A1' } });
      fireEvent.keyDown(name, { key: 'Enter' });
      await waitFor(async () => { await pump(); expect(session.selection?.range).toMatchObject({ start: { row: 0, col: 0 } }); });
      fireEvent.keyDown(base.surface, { key: 'ArrowDown', ctrlKey: true, shiftKey: true });
      await waitFor(async () => {
        await pump();
        const range = session.selection?.range;
        expect(range).toMatchObject({ kind: 'rowCol', start: { row: 0, col: 0 }, end: { col: 0 } });
        expect(range?.kind === 'rowCol' ? range.end.row : 0).toBeGreaterThanOrEqual(999_999);
      });
    },
    cellAt: workerCell,
    type: async (value) => {
      fireEvent.keyDown(base.surface, { key: 'F2' });
      const input = await waitFor(() => view.getByTestId('xlsx-cell-editor') as HTMLInputElement);
      await waitFor(() => expect(input.value).not.toBe(''));
      fireEvent.change(input, { target: { value } });
      fireEvent.keyDown(input, { key: 'Enter' });
    },
    settle: async () => {
      await act(async () => { await api.version().catch(() => {}); });
      await pump();
    },
    input: workerInput,
    session,
    saved: async () => {
      let result: unknown;
      await act(async () => { result = await api.commands.execute('save', null); });
      return result;
    },
    replace: async (next) => {
      await act(async () => {
        view.rerender(element(Object.entries(next).map(([id, gate]) => gatePlugin(id, gate, log))));
      });
      await pump();
    },
  };
}

const cell = (col: number) => ({ sheetId: expect.any(String), row: 0, col });

for (const editor of ['in-thread', 'peerless'] as const) {
  describe(`XlsxEditor plugin cell-edit gate (${editor})`, () => {
    it('cellEditable false keeps typing, F2, the clear key and the formula bar from editing the cell', async () => {
      const asked: XlsxPluginCellRef[] = [];
      const book = await mountGated(editor, { lock: { cellEditable: (ref) => { asked.push(ref); return ref.col !== 0; } } });
      await book.select(0);
      fireEvent.keyDown(book.surface, { key: 'F2' });
      expect(book.editor()).toBeNull();
      fireEvent.keyDown(book.surface, { key: 'x' });
      expect(book.editor()).toBeNull();
      fireEvent.keyDown(book.surface, { key: 'Delete' });
      await book.settle();
      await waitFor(() => expect(book.refusal()).toBe(NOT_ALLOWED));
      expect(await book.input(0)).toBe('1');
      expect(book.formula().readOnly).toBe(true);
      fireEvent.change(book.formula(), { target: { value: 'typed' } });
      fireEvent.keyDown(book.formula(), { key: 'Enter' });
      await book.settle();
      expect(await book.input(0)).toBe('1');
      expect(asked).toContainEqual(cell(0));
      await book.select(1);
      await waitFor(() => expect(book.formula().readOnly).toBe(false));
      await book.type('open');
      await waitFor(async () => expect(await book.input(1)).toBe('open'));
      expect(book.errors).toEqual([]);
    });

    it('a refused commit leaves the cell unchanged, keeps the draft open and reports host-refused', async () => {
      const commits: XlsxPluginCellCommit[] = [];
      const book = await mountGated(editor, { numbers: { beforeCellCommit: (commit) => {
        commits.push(commit);
        return /^\d+$/.test(commit.input) ? { action: 'accept' } : { action: 'refuse', message: 'Numbers only' };
      } } });
      await book.select(0);
      await book.type('abc');
      await waitFor(() => expect(book.refusal()).toBe('Numbers only'));
      await waitFor(() => expect(book.editor()?.value).toBe('abc'));
      expect(await book.input(0)).toBe('1');
      expect(commits).toEqual([{ ...cell(0), input: 'abc', kind: 'value', source: 'cell-editor' }]);
      expect(await book.saved()).toMatchObject({ ok: false, failure: { code: 'host-refused', message: 'Numbers only' } });
      fireEvent.keyDown(book.editor()!, { key: 'Escape' });
      await book.settle();
      await waitFor(() => expect(book.editor()).toBeNull());
      await book.type('42');
      await waitFor(async () => expect(await book.input(0)).toBe('42'));
      if (editor === 'in-thread') await waitFor(() => expect(book.refusal()).toBeNull());
    });

    it('a definition replaced with the same id and revision gates with its own hooks', async () => {
      const seen: string[] = [];
      const book = await mountGated(editor, { rule: { beforeCellCommit: (commit) => {
        seen.push(`old:${commit.input}`);
        return { action: 'accept' };
      } } });
      await book.replace({ rule: { beforeCellCommit: (commit) => {
        seen.push(`new:${commit.input}`);
        return { action: 'refuse', message: 'Replaced rule' };
      } } });
      await book.select(0);
      await book.type('7');
      await waitFor(() => expect(book.refusal()).toBe('Replaced rule'));
      expect(await book.input(0)).toBe('1');
      expect(seen).toEqual(['new:7']);
      expect(book.errors).toEqual([]);
    });

    it('after a refused commit and Escape, retyping the cell writes the new input', async () => {
      const book = await mountGated(editor, { numbers: { beforeCellCommit: (commit) =>
        /^\d+$/.test(commit.input) ? { action: 'accept' } : { action: 'refuse', message: 'Numbers only' } } });
      await book.select(0);
      await book.type('abc');
      await waitFor(() => expect(book.refusal()).toBe('Numbers only'));
      await waitFor(() => expect(book.editor()?.value).toBe('abc'));
      fireEvent.keyDown(book.editor()!, { key: 'Escape' });
      await book.settle();
      await waitFor(() => expect(book.editor()).toBeNull());
      await book.type('42');
      await waitFor(async () => expect(await book.input(0)).toBe('42'));
    });

    it('a commit while a gating plugin starts is refused, and the same Enter commits once it is ready', async () => {
      let ready!: () => void;
      const held = new Promise<void>((done) => { ready = done; });
      const commits: string[] = [];
      const book = await mountGated(editor, { slow: { initialize: () => held, beforeCellCommit: (commit) => {
        commits.push(commit.input);
        return commit.input === 'no' ? { action: 'refuse', message: 'No' } : { action: 'accept' };
      } } }, { starting: 1 });
      await book.select(0);
      await book.type('9');
      await waitFor(() => expect(book.refusal()).toBe(STARTING));
      await waitFor(() => expect(book.editor()?.value).toBe('9'));
      expect(await book.input(0)).toBe('1');
      expect(commits).toEqual([]);
      await act(async () => { ready(); });
      await waitFor(() => expect(book.log.loaded.has('slow')).toBe(true));
      fireEvent.keyDown(book.editor()!, { key: 'Enter' });
      await waitFor(async () => expect(await book.input(0)).toBe('9'));
      expect(commits).toEqual(['9']);
      await book.select(1);
      await book.type('no');
      await waitFor(() => expect(book.refusal()).toBe('No'));
      expect(await book.input(1)).toBe('2');
    });

    it('a clear while a gating plugin starts is refused, and clears once it is ready', async () => {
      let ready!: () => void;
      const held = new Promise<void>((done) => { ready = done; });
      const commits: XlsxPluginCellCommit[] = [];
      const book = await mountGated(editor, { slow: { initialize: () => held, beforeCellCommit: (commit) => {
        commits.push(commit);
        return { action: 'accept' };
      } } }, { starting: 1 });
      await book.select(0);
      fireEvent.keyDown(book.surface, { key: 'Delete' });
      await book.settle();
      await waitFor(() => expect(book.refusal()).toBe(STARTING));
      expect(await book.input(0)).toBe('1');
      await act(async () => { ready(); });
      await waitFor(() => expect(book.log.loaded.has('slow')).toBe(true));
      fireEvent.keyDown(book.surface, { key: 'Delete' });
      await book.settle();
      await waitFor(async () => expect(await book.input(0)).toBe(''));
      expect(commits).toEqual([{ ...cell(0), input: '', kind: 'clear', source: 'clear-key' }]);
    });

    it('two gating plugins sharing an id are rejected and do not hold edits', async () => {
      const book = await mountGated(editor, {
        first: { id: 'shared', cellEditable: () => false },
        second: { id: 'shared', beforeCellCommit: () => ({ action: 'refuse', message: 'Shared' }) },
      }, { starting: 2 });
      await waitFor(() => expect(book.errors.some((error) => error.pluginId === 'shared' && error.phase === 'definition')).toBe(true));
      await book.select(0);
      await book.type('9');
      await waitFor(async () => expect(await book.input(0)).toBe('9'));
      expect(book.refusal()).toBeNull();
    });

    for (const rejected of [false, true]) {
      const variant = rejected ? ' beside a rejected definition under its id' : '';
      const rejectedGate = (): Record<string, Gate> => rejected ? { rejected: { id: 'x', commands: DUPLICATE_COMMANDS } } : {};

      it(`a gating plugin${variant} refuses a typed commit`, async () => {
        const commits: string[] = [];
        const book = await mountGated(editor, {
          rule: { id: 'x', beforeCellCommit: (commit) => {
            commits.push(commit.input);
            return commit.input === 'no' ? { action: 'refuse', message: 'Rule' } : { action: 'accept' };
          } },
          ...rejectedGate(),
        }, { starting: rejected ? 1 : 0 });
        await book.select(0);
        await book.type('no');
        await waitFor(() => expect(book.refusal()).toBe('Rule'));
        expect(await book.input(0)).toBe('1');
        expect(commits).toEqual(['no']);
        expect(reported(book.errors)).toEqual(rejected ? ['x:definition'] : []);
      });

      it(`a gating plugin${variant} refuses edits while it starts`, async () => {
        let ready!: () => void;
        const held = new Promise<void>((done) => { ready = done; });
        const commits: string[] = [];
        const book = await mountGated(editor, {
          rule: { id: 'x', initialize: () => held, beforeCellCommit: (commit) => {
            commits.push(commit.input);
            return { action: 'accept' };
          } },
          ...rejectedGate(),
        }, { starting: rejected ? 2 : 1 });
        await book.select(0);
        await book.type('9');
        await waitFor(() => expect(book.refusal()).toBe(STARTING));
        await waitFor(() => expect(book.editor()?.value).toBe('9'));
        expect(await book.input(0)).toBe('1');
        expect(commits).toEqual([]);
        await act(async () => { ready(); });
        await waitFor(() => expect(book.log.loaded.has('x')).toBe(true));
        fireEvent.keyDown(book.editor()!, { key: 'Enter' });
        await waitFor(async () => expect(await book.input(0)).toBe('9'));
        expect(commits).toEqual(['9']);
      });
    }

    it('a gating plugin whose initialize throws is ended and does not hold edits', async () => {
      const book = await mountGated(editor, { broken: {
        initialize: () => { throw new Error('init exploded'); },
        beforeCellCommit: () => ({ action: 'refuse', message: 'Broken' }),
      } }, { starting: 1 });
      await waitFor(() => expect(book.errors.some((error) => error.pluginId === 'broken' && error.phase === 'initialize')).toBe(true));
      await book.select(0);
      await book.type('9');
      await waitFor(async () => expect(await book.input(0)).toBe('9'));
      expect(book.refusal()).toBeNull();
    });

    it('a new revision starting after the old one failed to clean up refuses edits until it loads', async () => {
      const book = await mountGated(editor, { rule: { revision: 1, cleanup: () => { throw new Error('cleanup exploded'); },
        beforeCellCommit: () => ({ action: 'accept' }) } });
      let ready!: () => void;
      const held = new Promise<void>((done) => { ready = done; });
      let loaded = false;
      await book.replace({ rule: { revision: 2, initialize: () => held, onLoad: () => { loaded = true; },
        beforeCellCommit: (commit) => commit.input === '9' ? { action: 'refuse', message: 'Second rule' } : { action: 'accept' } } });
      await waitFor(() => expect(book.errors.some((error) => error.pluginId === 'rule' && error.phase === 'cleanup')).toBe(true));
      await book.select(0);
      await book.type('9');
      await waitFor(() => expect(book.refusal()).toBe(STARTING));
      await waitFor(() => expect(book.editor()?.value).toBe('9'));
      expect(await book.input(0)).toBe('1');
      await act(async () => { ready(); });
      await waitFor(() => expect(loaded).toBe(true));
      await act(async () => {});
      fireEvent.keyDown(book.editor()!, { key: 'Enter' });
      await waitFor(() => expect(book.refusal()).toBe('Second rule'));
      expect(await book.input(0)).toBe('1');
    });

    it('a gate added to a plugin that failed to start neither holds edits nor gates them', async () => {
      const book = await mountGated(editor, { late: { initialize: () => { throw new Error('init exploded'); } } }, { starting: 1 });
      await waitFor(() => expect(book.errors.some((error) => error.pluginId === 'late' && error.phase === 'initialize')).toBe(true));
      await book.replace({ late: { beforeCellCommit: () => ({ action: 'refuse', message: 'Late rule' }) } });
      await book.select(0);
      await book.type('9');
      await waitFor(async () => expect(await book.input(0)).toBe('9'));
      expect(book.refusal()).toBeNull();
    });

    it('a refusal notice clears once the corrected commit is written', async () => {
      const book = await mountGated(editor, { numbers: { beforeCellCommit: (commit) =>
        /^\d+$/.test(commit.input) ? { action: 'accept' } : { action: 'refuse', message: 'Numbers only' } } });
      await book.select(0);
      await book.type('abc');
      await waitFor(() => expect(book.refusal()).toBe('Numbers only'));
      await waitFor(() => expect(book.editor()?.value).toBe('abc'));
      fireEvent.change(book.editor()!, { target: { value: '42' } });
      fireEvent.keyDown(book.editor()!, { key: 'Enter' });
      await waitFor(async () => expect(await book.input(0)).toBe('42'));
      await waitFor(() => expect(book.refusal()).toBeNull());
    });

    it('replacements chain in installation order', async () => {
      const seen: string[] = [];
      const book = await mountGated(editor, {
        upper: { beforeCellCommit: (commit) => { seen.push(`upper:${commit.input}`); return { action: 'replace', input: commit.input.toUpperCase() }; } },
        bang: { beforeCellCommit: (commit) => { seen.push(`bang:${commit.input}`); return { action: 'replace', input: `${commit.input}!` }; } },
      });
      await book.select(0);
      await book.type('hi');
      await waitFor(async () => expect(await book.input(0)).toBe('HI!'));
      expect(seen).toEqual(['upper:hi', 'bang:HI']);
    });

    it('any refusal wins and stops the chain', async () => {
      const seen: string[] = [];
      const book = await mountGated(editor, {
        accepts: { beforeCellCommit: () => { seen.push('accepts'); return { action: 'accept' }; } },
        refuses: { beforeCellCommit: () => { seen.push('refuses'); return { action: 'refuse' }; } },
        later: { beforeCellCommit: () => { seen.push('later'); return { action: 'accept' }; } },
      });
      await book.select(0);
      await book.type('nope');
      await waitFor(() => expect(book.refusal()).toBe(NOT_ALLOWED));
      expect(await book.input(0)).toBe('1');
      expect(seen).toEqual(['accepts', 'refuses']);
      expect(await book.saved()).toMatchObject({ ok: false, failure: { code: 'host-refused', message: NOT_ALLOWED } });
    });

    for (const broken of ['throws', 'returns a non-decision'] as const) {
      it(`a plugin that ${broken} is quarantined and the commit proceeds`, async () => {
        let calls = 0;
        const book = await mountGated(editor, {
          broken: { beforeCellCommit: () => {
            calls += 1;
            if (broken === 'throws') throw new Error('gate exploded');
            return 'yes';
          } },
          healthy: { beforeCellCommit: (commit) => ({ action: 'replace', input: `${commit.input}+` }) },
        });
        await book.select(0);
        await book.type('go');
        await waitFor(async () => expect(await book.input(0)).toBe('go+'));
        await waitFor(() => expect(book.errors).toMatchObject([{ pluginId: 'broken', phase: 'edit-gate' }]));
        await waitFor(() => expect(book.log.cleanups).toContain('broken:failed'));
        await book.select(1);
        await book.type('again');
        await waitFor(async () => expect(await book.input(1)).toBe('again+'));
        expect(calls).toBe(1);
      });
    }

    it('a throwing cellEditable is quarantined and the cell stays editable', async () => {
      const book = await mountGated(editor, { broken: { cellEditable: () => { throw new Error('gate exploded'); } } });
      await book.select(0);
      await book.type('fine');
      await waitFor(async () => expect(await book.input(0)).toBe('fine'));
      await waitFor(() => expect(book.errors[0]).toMatchObject({ pluginId: 'broken', phase: 'edit-gate' }));
    });

    it('formula bar commits route through the gate', async () => {
      const commits: XlsxPluginCellCommit[] = [];
      const book = await mountGated(editor, { noFormulas: { beforeCellCommit: (commit) => {
        commits.push(commit);
        return commit.kind === 'formula' ? { action: 'refuse', message: 'No formulas' } : { action: 'accept' };
      } } });
      await book.select(0);
      fireEvent.change(book.formula(), { target: { value: '=B1*2' } });
      fireEvent.keyDown(book.formula(), { key: 'Enter' });
      await waitFor(() => expect(book.refusal()).toBe('No formulas'));
      expect(await book.input(0)).toBe('1');
      expect(commits).toEqual([{ ...cell(0), input: '=B1*2', kind: 'formula', source: 'formula-bar' }]);
      fireEvent.change(book.formula(), { target: { value: '7' } });
      fireEvent.keyDown(book.formula(), { key: 'Enter' });
      await waitFor(async () => expect(await book.input(0)).toBe('7'));
    });

    it('a multi-cell clear with one refused cell changes nothing and asks only cells with content', async () => {
      const commits: XlsxPluginCellCommit[] = [];
      const book = await mountGated(editor, { keepB: { beforeCellCommit: (commit) => {
        commits.push(commit);
        return commit.col === 1 ? { action: 'refuse', message: 'Keep B' } : { action: 'accept' };
      } } });
      await book.select(0, 2);
      fireEvent.keyDown(book.surface, { key: 'Delete' });
      await book.settle();
      await waitFor(() => expect(book.refusal()).toBe('Keep B'));
      expect([await book.input(0), await book.input(1)]).toEqual(['1', '2']);
      expect(commits).toEqual([
        { ...cell(0), input: '', kind: 'clear', source: 'clear-key' },
        { ...cell(1), input: '', kind: 'clear', source: 'clear-key' },
      ]);
      await book.select(0);
      fireEvent.keyDown(book.surface, { key: 'Delete' });
      await book.settle();
      await waitFor(async () => expect([await book.input(0), await book.input(1)]).toEqual(['', '2']));
    });

    it('a multi-cell clear over a cell that is not editable changes nothing', async () => {
      const book = await mountGated(editor, { lock: { cellEditable: (ref) => ref.col !== 1 } });
      await book.select(0, 1);
      fireEvent.keyDown(book.surface, { key: 'Delete' });
      await book.settle();
      await waitFor(() => expect(book.refusal()).toBe(NOT_ALLOWED));
      expect([await book.input(0), await book.input(1)]).toEqual(['1', '2']);
    });

    const nameBox = (book: Book) => (book.view.getByTestId('xlsx-name-box') as HTMLInputElement).value;
    const paste = async (book: Book) => {
      fireEvent.keyDown(book.surface, { key: 'v', ctrlKey: true });
      await book.settle();
    };

    it('a paste over a cell that is not editable is refused whole and keeps the selection', async () => {
      fakeClipboard('p\tq');
      const commits: XlsxPluginCellCommit[] = [];
      const book = await mountGated(editor, { lock: {
        cellEditable: (ref) => ref.col !== 1,
        beforeCellCommit: (commit) => { commits.push(commit); return { action: 'accept' }; },
      } });
      await book.select(0);
      await paste(book);
      await waitFor(() => expect(book.refusal()).toBe(NOT_ALLOWED));
      expect([await book.input(0), await book.input(1)]).toEqual(['1', '2']);
      expect(nameBox(book)).toBe('A1');
      expect(commits).toEqual([]);
    });

    it('a paste asks only the cells it changes', async () => {
      fakeClipboard('1\t2\tz');
      const commits: XlsxPluginCellCommit[] = [];
      const book = await mountGated(editor, { spy: { beforeCellCommit: (commit) => {
        commits.push(commit);
        return { action: 'accept' };
      } } });
      await book.select(0);
      await paste(book);
      await waitFor(async () => expect(await book.input(2)).toBe('z'));
      expect(commits).toEqual([{ ...cell(2), input: 'z', kind: 'value', source: 'paste' }]);
    });

    it('a pasted formula reaches plugins moved to its cell, and a replaced one is written as given', async () => {
      const clipboard = fakeClipboard();
      const commits: XlsxPluginCellCommit[] = [];
      const book = await mountGated(editor, { rewrite: { beforeCellCommit: (commit) => {
        if (commit.source !== 'paste') return { action: 'accept' };
        commits.push(commit);
        return commit.col === 2 ? { action: 'replace', input: '=A1' } : { action: 'accept' };
      } } });
      await book.select(1);
      fireEvent.change(book.formula(), { target: { value: '=A1*2' } });
      fireEvent.keyDown(book.formula(), { key: 'Enter' });
      await waitFor(async () => expect(await book.input(1)).toBe('=A1*2'));
      await book.select(1);
      fireEvent.keyDown(book.surface, { key: 'c', ctrlKey: true });
      await waitFor(async () => expect(await clipboard.text()).toBe('=A1*2'));
      await book.select(2);
      await paste(book);
      await waitFor(async () => expect(await book.input(2)).toBe('=A1'));
      await book.select(3);
      await paste(book);
      await waitFor(async () => expect(await book.input(3)).toBe('=C1*2'));
      expect(commits).toEqual([
        { ...cell(2), input: '=B1*2', kind: 'formula', source: 'paste' },
        { ...cell(3), input: '=C1*2', kind: 'formula', source: 'paste' },
      ]);
    });

    it('a paste changing more cells than plugins may check is refused with the limit', async () => {
      const cols = 100;
      const rows = Math.floor(GATE_CELL_LIMIT / cols) + 1;
      fakeClipboard(Array.from({ length: rows }, () => Array(cols).fill('x').join('\t')).join('\n'));
      let asked = 0;
      const book = await mountGated(editor, { spy: {
        cellEditable: (ref) => { if (ref.row > 0) asked++; return true; },
        beforeCellCommit: () => { asked++; return { action: 'accept' }; },
      } });
      await book.select(0);
      await paste(book);
      await waitFor(() => expect(book.refusal()).toBe(`Too many cells for plugins to check (limit ${GATE_CELL_LIMIT}).`),
        { timeout: 30_000 });
      expect([await book.input(0), await book.input(2)]).toEqual(['1', '']);
      expect(asked).toBe(0);
    }, 60_000);

    it('a cut from a cell that is not editable leaves the clipboard and cells, and an allowed cut clears', async () => {
      const clipboard = fakeClipboard('before');
      const commits: XlsxPluginCellCommit[] = [];
      const book = await mountGated(editor, { lock: {
        cellEditable: (ref) => ref.col !== 0,
        beforeCellCommit: (commit) => { commits.push(commit); return { action: 'accept' }; },
      } });
      await book.select(0);
      fireEvent.keyDown(book.surface, { key: 'x', ctrlKey: true });
      await book.settle();
      await waitFor(() => expect(book.refusal()).toBe(NOT_ALLOWED));
      expect(await book.input(0)).toBe('1');
      expect(await clipboard.text()).toBe('before');
      await book.select(1);
      fireEvent.keyDown(book.surface, { key: 'x', ctrlKey: true });
      await book.settle();
      await waitFor(async () => expect(await book.input(1)).toBe(''));
      expect(await clipboard.text()).toBe('2');
      expect(commits).toEqual([{ ...cell(1), input: '', kind: 'clear', source: 'cut' }]);
    });

    it('a paste while a gating plugin starts is refused, and pastes once it is ready', async () => {
      fakeClipboard('p');
      let ready!: () => void;
      const held = new Promise<void>((done) => { ready = done; });
      const commits: XlsxPluginCellCommit[] = [];
      const book = await mountGated(editor, { slow: { initialize: () => held, beforeCellCommit: (commit) => {
        commits.push(commit);
        return { action: 'accept' };
      } } }, { starting: 1 });
      await book.select(0);
      await paste(book);
      await waitFor(() => expect(book.refusal()).toBe(STARTING));
      expect(await book.input(0)).toBe('1');
      await act(async () => { ready(); });
      await waitFor(() => expect(book.log.loaded.has('slow')).toBe(true));
      await paste(book);
      await waitFor(async () => expect(await book.input(0)).toBe('p'));
      expect(commits).toEqual([{ ...cell(0), input: 'p', kind: 'value', source: 'paste' }]);
    });

    it('a cut while a gating plugin starts is refused with the clipboard and cells unchanged', async () => {
      const clipboard = fakeClipboard('before');
      let ready!: () => void;
      const held = new Promise<void>((done) => { ready = done; });
      const book = await mountGated(editor, { slow: { initialize: () => held,
        beforeCellCommit: () => ({ action: 'accept' }) } }, { starting: 1 });
      await book.select(0);
      fireEvent.keyDown(book.surface, { key: 'x', ctrlKey: true });
      await book.settle();
      await waitFor(() => expect(book.refusal()).toBe(STARTING));
      expect(await book.input(0)).toBe('1');
      expect(await clipboard.text()).toBe('before');
      await act(async () => { ready(); });
      await waitFor(() => expect(book.log.loaded.has('slow')).toBe(true));
      fireEvent.keyDown(book.surface, { key: 'x', ctrlKey: true });
      await book.settle();
      await waitFor(async () => expect(await book.input(0)).toBe(''));
      expect(await clipboard.text()).toBe('1');
    });

    it('a cut queued behind a paste into its empty cell is decided after the paste, with the clipboard unchanged', async () => {
      let release!: () => void;
      const clipboard = fakeClipboard('p', new Promise<void>((done) => { release = done; }));
      const book = await mountGated(editor, { noCut: { beforeCellCommit: (commit) =>
        commit.source === 'cut' ? { action: 'refuse', message: 'No cut' } : { action: 'accept' } } });
      await book.select(2);
      fireEvent.keyDown(book.surface, { key: 'v', ctrlKey: true });
      fireEvent.keyDown(book.surface, { key: 'x', ctrlKey: true });
      await act(async () => { release(); });
      await book.settle();
      await waitFor(async () => expect(await book.input(2)).toBe('p'));
      await waitFor(() => expect(book.refusal()).toBe('No cut'));
      expect(await clipboard.text()).toBe('p');
    });

    if (editor === 'peerless') it('a gated cut of a locked cell on a protected sheet is refused before the clipboard is written', async () => {
      const clipboard = fakeClipboard('before');
      const asked: string[] = [];
      const book = await mountGated(editor, { spy: { beforeCellCommit: (commit) => {
        asked.push(`${commit.source}:${commit.col}`);
        return { action: 'accept' };
      } } }, { protectedSheet: true });
      await book.select(0);
      fireEvent.keyDown(book.surface, { key: 'x', ctrlKey: true });
      await book.settle();
      await waitFor(() => expect(book.refusal()).toBe(PROTECTED));
      expect(await book.input(0)).toBe('1');
      expect(await clipboard.text()).toBe('before');
      expect(asked).toEqual([]);
    });

    if (editor === 'peerless') it('a paste over a locked cell of a protected sheet is refused by protection before plugins are asked', async () => {
      fakeClipboard('p');
      const asked: string[] = [];
      const book = await mountGated(editor, { spy: {
        cellEditable: (ref) => { asked.push(`editable:${ref.col}`); return true; },
        beforeCellCommit: (commit) => { asked.push(`commit:${commit.col}`); return { action: 'accept' }; },
      } }, { protectedSheet: true });
      await book.select(0);
      await paste(book);
      await waitFor(() => expect(book.refusal()).not.toBeNull());
      expect(await book.input(0)).toBe('1');
      expect(asked.filter((entry) => entry.endsWith(':0') || entry.startsWith('commit:'))).toEqual([]);
      await book.select(1);
      await paste(book);
      await waitFor(async () => expect(await book.input(1)).toBe('p'));
      expect(asked).toContain('commit:1');
      expect(asked.filter((entry) => entry.endsWith(':0'))).toEqual([]);
    });

    it('gate hooks get reads only: no edits, commands, setState or run', async () => {
      const seen: string[][] = [];
      const book = await mountGated(editor, { reader: {
        createState: () => ({ config: { locked: true } }),
        beforeCellCommit: (_commit, context) => {
          seen.push(Object.keys(context).sort());
          const loose = context as unknown as Record<string, unknown>;
          seen.push(['edits', 'commands', 'setState', 'run', 'onCleanup', 'navigation']
            .filter((member) => loose[member] !== undefined));
          seen.push([String((context.state as { config: { locked: boolean } }).config.locked)]);
          return { action: 'accept' };
        },
      } });
      await book.select(0);
      await book.type('kept');
      await waitFor(async () => expect(await book.input(0)).toBe('kept'));
      expect(seen).toEqual([['pluginId', 'read', 'snapshot', 'state'], [], ['true']]);
    });

    for (const getter of ['reads once', 'always throws'] as const) {
      it(`a decision getter that ${getter} is read inside the guard`, async () => {
        let reads = 0;
        const book = await mountGated(editor, { getter: { beforeCellCommit: () => ({
          action: 'replace',
          get input() {
            reads += 1;
            if (getter === 'always throws' || reads > 1) throw new Error('getter exploded');
            return 'X';
          },
        }) } });
        await book.select(0);
        await book.type('typed');
        if (getter === 'reads once') {
          await waitFor(async () => expect(await book.input(0)).toBe('X'));
          expect(reads).toBe(1);
          expect(book.errors).toEqual([]);
        } else {
          await waitFor(async () => expect(await book.input(0)).toBe('typed'));
          await waitFor(() => expect(book.errors).toMatchObject([{ pluginId: 'getter', phase: 'edit-gate' }]));
        }
      });
    }

    it('clearing a column whose only content is far down asks the gate once and clears it', async () => {
      const commits: XlsxPluginCellCommit[] = [];
      const book = await mountGated(editor, { spy: { beforeCellCommit: (commit) => {
        commits.push(commit);
        return { action: 'accept' };
      } } }, { sparse: true });
      expect(await book.cellAt(999_999, 0)).toBe('9');
      await book.selectColumn();
      fireEvent.keyDown(book.surface, { key: 'Delete' });
      await book.settle();
      await waitFor(async () => expect(await book.cellAt(999_999, 0)).toBe(''));
      expect(commits).toEqual([{ sheetId: expect.any(String), row: 999_999, col: 0, input: '', kind: 'clear', source: 'clear-key' }]);
    });

    const numbersOnly = () => ({ numbers: { beforeCellCommit: (commit: XlsxPluginCellCommit) =>
      /^\d+$/.test(commit.input) ? { action: 'accept' } : { action: 'refuse', message: 'Numbers only' } } });

    if (editor === 'in-thread') it('a grid press away from a refused draft keeps the selection and focuses the draft', async () => {
      const book = await mountGated(editor, numbersOnly());
      await book.select(0);
      await book.type('abc');
      await waitFor(() => expect(book.refusal()).toBe('Numbers only'));
      await book.click!(1);
      expect(nameBox(book)).toBe('A1');
      expect(book.editor()?.value).toBe('abc');
      expect(document.activeElement).toBe(book.editor());
      expect([await book.input(0), await book.input(1)]).toEqual(['1', '2']);
    });

    if (editor === 'in-thread') it('a refused formula-bar Enter keeps focus in the formula bar and arrows do not move the selection', async () => {
      const book = await mountGated(editor, numbersOnly());
      await book.select(0);
      act(() => { book.formula().focus(); });
      fireEvent.change(book.formula(), { target: { value: 'abc' } });
      fireEvent.keyDown(book.formula(), { key: 'Enter' });
      await waitFor(() => expect(book.refusal()).toBe('Numbers only'));
      expect(document.activeElement).toBe(book.formula());
      fireEvent.keyDown(document.activeElement!, { key: 'ArrowDown' });
      await book.settle();
      expect(nameBox(book)).toBe('A1');
      expect(book.formula().value).toBe('abc');
      expect(await book.input(0)).toBe('1');
    });

    if (editor === 'peerless') it('a second quick commit is decided at the version the first one wrote', async () => {
      const versions: Record<string, string> = {};
      const book = await mountGated(editor, { versioned: { beforeCellCommit: (commit, context) => {
        versions[commit.input] = context.snapshot.version;
        return { action: 'accept' };
      } } });
      const call = book.session!.call;
      const readCells = call.readCells.bind(call);
      let release = () => {};
      const held = new Promise<void>((done) => { release = done; });
      let holding = true;
      const spy = spyOn(call, 'readCells').mockImplementation(async (...args) => {
        if (holding && args[0]?.ranges?.length === 0) {
          holding = false;
          await held;
        }
        return readCells(...args);
      });
      restorers.push(() => spy.mockRestore());
      await book.select(0);
      await book.type('11');
      await waitFor(() => expect(book.editor()).toBeNull());
      fireEvent.keyDown(book.surface, { key: '2' });
      const second = await waitFor(() => book.view.getByTestId('xlsx-cell-editor') as HTMLInputElement);
      fireEvent.change(second, { target: { value: '22' } });
      fireEvent.keyDown(second, { key: 'Enter' });
      await act(async () => { release(); });
      await waitFor(async () => expect([await book.cellAt(0, 0), await book.cellAt(1, 0)]).toEqual(['11', '22']));
      expect(Object.keys(versions).sort()).toEqual(['11', '22']);
      expect(versions['22']).not.toBe(versions['11']);
    });

    if (editor === 'peerless') it('the protected-sheet refusal wins before plugins are asked', async () => {
      const asked: string[] = [];
      const book = await mountGated(editor, { spy: {
        cellEditable: (ref) => { asked.push(`editable:${ref.col}`); return true; },
        beforeCellCommit: (commit) => { asked.push(`commit:${commit.col}`); return { action: 'accept' }; },
      } }, { protectedSheet: true });
      await book.select(0);
      await book.type('changed');
      await waitFor(() => expect(book.refusal()).toBe(PROTECTED));
      await waitFor(() => expect(book.editor()?.value).toBe('changed'));
      expect(await book.input(0)).toBe('1');
      expect(await book.saved()).toMatchObject({ ok: false, failure: { code: 'sheet-protected' } });
      expect(asked.filter((entry) => entry.endsWith(':0'))).toEqual([]);
      fireEvent.keyDown(book.editor()!, { key: 'Escape' });
      await book.select(1);
      await book.type('unlocked');
      await waitFor(async () => expect(await book.input(1)).toBe('unlocked'));
      expect(asked).toContain('commit:1');
      expect(asked.filter((entry) => entry.endsWith(':0'))).toEqual([]);
    });
  });
}

async function mountLegacy(gates: Record<string, Gate>) {
  const errors: XlsxPluginError[] = [];
  const log: GateLog = { loaded: new Set<string>(), cleanups: [], contexts: {} };
  const plugins = Object.entries(gates).map(([id, gate]) => gatePlugin(gate.id ?? id, gate, log));
  const file = await book();
  for (const name of ['Worker', 'OffscreenCanvas'] as const) {
    const original = Object.getOwnPropertyDescriptor(globalThis, name);
    Object.defineProperty(globalThis, name, { configurable: true, writable: true, value: class {} });
    restorers.push(() => {
      if (original) Object.defineProperty(globalThis, name, original);
      else Reflect.deleteProperty(globalThis, name);
    });
  }
  const pump = stubFrames();
  const session = await createTestWorkbookSession(file, undefined,
    { wasm: new Uint8Array(readFileSync(WASM)).buffer, retainPeerHydration: true });
  sessions.push(session);
  const open = spyOn(editableWorkbookSessionBackend, 'open').mockResolvedValue(session);
  restorers.push(() => open.mockRestore());
  let worker: XlsxWorkerEditorApi | undefined;
  const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen file={file} plugins={plugins}
    onPluginError={(error) => { errors.push(error); }}
    onReady={(ready) => { worker = ready as unknown as XlsxWorkerEditorApi; }} />);
  await waitFor(async () => { await pump(); expect(worker?.hydrated).toBe(true); }, MOUNTED);
  const surface = view.getByTestId('xlsx-scroll');
  return {
    errors, log,
    type: async (col: number, value: string) => {
      await act(async () => { worker!.selectCells(0, { anchor: { row: 0, col }, focus: { row: 0, col } }); });
      await pump();
      fireEvent.keyDown(surface, { key: 'F2' });
      const input = await waitFor(() => view.getByTestId('xlsx-cell-editor') as HTMLInputElement);
      fireEvent.change(input, { target: { value } });
      fireEvent.keyDown(input, { key: 'Enter' });
      await act(async () => { await worker!.flush(); });
      await pump();
    },
    input: async (col: number) => (await session.call.cellInputs(0, A1[col])).cells[0][0].input,
  };
}

describe('XlsxEditor plugin cell-edit gate (legacy worker editor)', () => {
  for (const hook of ['beforeCellCommit', 'cellEditable'] as const) {
    it(`a plugin that defines ${hook} is ended with edit-gate and typing commits ungated`, async () => {
      const asked: string[] = [];
      const gating: Gate = hook === 'beforeCellCommit'
        ? { beforeCellCommit: () => { asked.push('commit'); return { action: 'refuse', message: 'never' }; } }
        : { cellEditable: () => { asked.push('editable'); return false; } };
      const legacy = await mountLegacy({ gating, plain: {} });
      await waitFor(() => expect(legacy.log.loaded.has('plain')).toBe(true), MOUNTED);
      await waitFor(() => expect(legacy.errors).toHaveLength(1));
      expect(legacy.errors[0]).toMatchObject({ pluginId: 'gating', phase: 'edit-gate' });
      expect((legacy.errors[0].error as Error).message).toBe(EDIT_GATE_UNSUPPORTED);
      expect(legacy.log.loaded.has('gating')).toBe(false);
      await legacy.type(0, 'abc');
      await waitFor(async () => expect(await legacy.input(0)).toBe('abc'));
      expect(asked).toEqual([]);
      expect(legacy.errors).toHaveLength(1);
    });
  }

  it('a rejected gating definition does not end the accepted plugin under its id', async () => {
    const legacy = await mountLegacy({
      plain: { id: 'x' },
      rejected: { id: 'x', commands: DUPLICATE_COMMANDS, beforeCellCommit: () => ({ action: 'refuse', message: 'never' }) },
    });
    await waitFor(() => expect(legacy.log.loaded.has('x')).toBe(true), MOUNTED);
    await legacy.type(0, 'abc');
    await waitFor(async () => expect(await legacy.input(0)).toBe('abc'));
    expect(reported(legacy.errors)).toEqual(['x:definition']);
    expect(legacy.log.cleanups).toEqual([]);
  });

  it('a plugin without gate hooks still runs and typing commits', async () => {
    const legacy = await mountLegacy({ plain: {} });
    await waitFor(() => expect(legacy.log.loaded.has('plain')).toBe(true), MOUNTED);
    await legacy.type(0, 'abc');
    await waitFor(async () => expect(await legacy.input(0)).toBe('abc'));
    expect(legacy.errors).toEqual([]);
    expect(legacy.log.cleanups).toEqual([]);
  });
});
