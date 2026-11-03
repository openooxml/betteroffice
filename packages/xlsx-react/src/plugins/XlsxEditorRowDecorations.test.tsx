import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, it, mock, spyOn } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { ReactElement } from 'react';
import * as xlsx from '@betteroffice/xlsx';
import { initWasm, openWorkbook, type WorkbookSession } from '@betteroffice/xlsx';
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
  type XlsxPluginContext,
  type XlsxPluginError,
  type XlsxPluginEvent,
  type XlsxPluginGeometry,
  type XlsxPluginRowDecoration,
  type XlsxPluginRowDecorations,
  type XlsxPluginSelection,
  type XlsxPointPosition,
} from '../index';
import { MOUNTED } from '../__fixtures__/mounted';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');

const ROOT = resolve(import.meta.dir, '../../../..');
const WASM = resolve(ROOT, 'packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm');
const VIEWPORT = { width: 800, height: 600 };
const LAYOUT = [
  ['clientWidth', VIEWPORT.width],
  ['clientHeight', VIEWPORT.height],
] as const;
const SHEET = 'sheet:0';
const SLOW = 30_000;
const originalGetContext = HTMLCanvasElement.prototype.getContext;
const originalLayout = LAYOUT.map(
  ([property]) => [property, Object.getOwnPropertyDescriptor(HTMLElement.prototype, property)] as const
);
let plain: Uint8Array;
let shaped: Uint8Array;

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
  plain = new Uint8Array(readFileSync(resolve(ROOT, 'packages/xlsx/test-fixtures/sample.xlsx')));
  const source = openWorkbook(plain.slice());
  source.applyOps([
    { type: 'setFreezePane', sheet: 0, pane: { rows: 2, cols: 0, top_left: { row: 2, col: 0 } } },
    { type: 'setRowHeight', sheet: 0, row: 3, height: 0 },
    { type: 'setRowHeight', sheet: 0, row: 4, height: 40 },
  ]);
  shaped = source.save();
  source.dispose();
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

type Context = XlsxPluginContext<number>;

interface Decorator {
  id: string;
  plugin: XlsxPlugin;
  activations: Extract<XlsxPluginEvent, { type: 'row-decoration-activate' }>[];
  selections: XlsxPluginSelection[];
  contexts: Context[];
  rendered: Context | null;
}

function decorator(id: string, compute: ((context: Context) => unknown) | null, revision?: number): Decorator {
  const record: Decorator = { id, plugin: null as unknown as XlsxPlugin, activations: [], selections: [], contexts: [],
    rendered: null };
  record.plugin = defineXlsxPlugin<number>({
    id,
    ...(revision !== undefined && { revision }),
    createState: () => 0,
    onEvent(context, event) {
      record.contexts.push(context);
      if (event.type === 'selection-change') record.selections.push(event.selection);
      if (event.type === 'row-decoration-activate') record.activations.push(event);
    },
    ...(compute && {
      rowDecorations(context: Context) {
        record.rendered = context;
        return compute(context) as XlsxPluginRowDecorations | null;
      },
    }),
  });
  return record;
}

const sheetOf = (context: Context) => context.snapshot.selection?.sheetId ?? SHEET;

function dots(rows: readonly number[], extra: Partial<XlsxPluginRowDecoration> = {}) {
  return (context: Context): XlsxPluginRowDecorations => ({
    version: context.snapshot.version,
    items: rows.map((row) => ({ id: `d${row}`, sheetId: sheetOf(context), row, kind: 'dot', tooltip: `Row ${row}`, ...extra })),
  });
}

type Editor = 'in-thread' | 'peerless' | 'legacy-worker';

interface Book {
  view: ReturnType<typeof render>;
  surface: HTMLElement;
  errors: XlsxPluginError[];
  until(check: () => boolean): Promise<void>;
  sync(): Promise<void>;
  zoom(scale: number): Promise<void>;
  scroll(top: number): Promise<void>;
  edit(row: number, col: number, value: string): Promise<void>;
  bump(record: Decorator): Promise<void>;
  gutter(): HTMLElement | null;
  row(row: number): HTMLElement | null;
  geometry(record: Decorator): XlsxPluginGeometry | null;
  /** Rerenders the editor with these plugins in place of the mounted ones. */
  replace(records: readonly Decorator[]): Promise<void>;
  /** The editor API's point query; the default worker editor's API has none. */
  point: ((clientX: number, clientY: number) => XlsxPointPosition | null) | null;
}

const a1 = (row: number, col: number) => `${String.fromCharCode(65 + col)}${row + 1}`;

const setCell = (expectVersion: string, row: number, col: number, value: string) => ({
  expectVersion,
  steps: [{ op: 'setCellInputs' as const, target: { sheetId: SHEET, range: { kind: 'a1' as const, a1: a1(row, col) } },
    inputs: [[value]] }],
});

async function mount(editor: Editor, records: readonly Decorator[],
  { file = plain, readOnly = false }: { file?: Uint8Array; readOnly?: boolean } = {}): Promise<Book> {
  const errors: XlsxPluginError[] = [];
  const plugins = records.map((record) => record.plugin);
  const onPluginError = (error: XlsxPluginError) => { errors.push(error); };
  let pump = () => act(async () => {});
  const shared = (view: ReturnType<typeof render>, element: (next: readonly XlsxPlugin[]) => ReactElement) => {
    const surface = view.getByTestId('xlsx-scroll');
    const gutter = () => view.queryByTestId('xlsx-row-gutter');
    const until = (check: () => boolean) => waitFor(async () => {
      await pump();
      expect(check()).toBe(true);
    }, MOUNTED);
    return {
      view, surface, errors, until, sync: () => pump(), gutter,
      row: (row: number) => gutter()?.querySelector<HTMLElement>(`[data-row="${row}"]`) ?? null,
      geometry: (record: Decorator) => last(record.contexts)?.geometry ?? null,
      replace: async (next: readonly Decorator[]) => {
        await act(async () => { view.rerender(element(next.map((record) => record.plugin))); });
        await pump();
      },
      bump: async (record: Decorator) => {
        await act(async () => {
          const context = record.rendered!;
          context.setState((count) => count + 1, await context.read.version().then((result) =>
            'version' in result ? result.version : context.snapshot.version));
        });
        await pump();
      },
      scroll: async (top: number) => {
        surface.scrollTop = top;
        fireEvent.scroll(surface);
        await pump();
      },
    };
  };
  const loaded = async (book: Omit<Book, 'zoom' | 'edit' | 'point'>) => {
    await book.until(() => records.every((record) => record.contexts.length > 0));
    await book.until(() => records.length === 0 ||
      records.some((record) => record.contexts.some((context) => context.geometry !== null)));
  };
  if (editor === 'in-thread') {
    let api: XlsxEditorApi | undefined;
    const element = (next: readonly XlsxPlugin[]) => <XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={file}
      plugins={next} readOnly={readOnly} onPluginError={onPluginError} onReady={(ready) => { api = ready; }} />;
    const view = render(element(plugins));
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    const base = shared(view, element);
    await loaded(base);
    return {
      ...base,
      point: (clientX, clientY) => api!.getPositionAtPoint(clientX, clientY),
      zoom: async (scale) => {
        await act(async () => { await api!.commands.execute('zoom', { scale }); });
        await base.until(() => records.some((record) => base.geometry(record)?.layout.zoom === scale));
      },
      edit: async (row, col, value) => {
        await act(async () => { await api!.applyEdits(setCell(api!.handle.version(), row, col, value)); });
        await pump();
      },
    };
  }
  if (editor === 'legacy-worker') {
    for (const name of ['Worker', 'OffscreenCanvas'] as const) {
      const original = Object.getOwnPropertyDescriptor(globalThis, name);
      Object.defineProperty(globalThis, name, { configurable: true, writable: true, value: class {} });
      restorers.push(() => {
        if (original) Object.defineProperty(globalThis, name, original);
        else Reflect.deleteProperty(globalThis, name);
      });
    }
  }
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
  const session = editor === 'legacy-worker'
    ? await createTestWorkbookSession(file, undefined, { wasm: new Uint8Array(readFileSync(WASM)).buffer, retainPeerHydration: true })
    : await createTestWorkbookSession(file, undefined, { includeEditState: true });
  sessions.push(session);
  const open = spyOn(editableWorkbookSessionBackend, 'open').mockResolvedValue(session);
  restorers.push(() => open.mockRestore());
  pump = () => act(async () => {
    for (let round = 0; round < 3; round++) {
      const callbacks = [...frames.values()];
      frames.clear();
      for (const callback of callbacks) callback(0);
      await new Promise((done) => setTimeout(done, 0));
    }
  });
  if (editor === 'legacy-worker') {
    let worker: XlsxWorkerEditorApi | undefined;
    const element = (next: readonly XlsxPlugin[]) => <XlsxEditor legacySyncApi experimentalWorkerOpen file={file}
      plugins={next} readOnly={readOnly} allowHostProposals={readOnly} onPluginError={onPluginError}
      onReady={(ready) => { worker = ready as unknown as XlsxWorkerEditorApi; }} />;
    const view = render(element(plugins));
    await waitFor(async () => {
      await pump();
      expect(worker?.hydrated).toBe(true);
    }, MOUNTED);
    const base = shared(view, element);
    await loaded(base);
    return {
      ...base,
      point: (clientX, clientY) => worker!.getPositionAtPoint(clientX, clientY),
      zoom: async (scale) => {
        await act(async () => { await worker!.commands.execute('zoom', { scale }); });
        await base.until(() => records.some((record) => base.geometry(record)?.layout.zoom === scale));
      },
      edit: async (row, col, value) => {
        await act(async () => { await worker!.editCellAsync(0, row, col, value); });
        await act(async () => { await worker!.flush().catch(() => {}); });
        await pump();
      },
    };
  }
  const ready = mock((_api: XlsxAsyncEditorApi) => {});
  const element = (next: readonly XlsxPlugin[]) => <XlsxWorkerEditor backend="peerless" file={file} plugins={next}
    readOnly={readOnly} onPluginError={onPluginError} onReady={ready} />;
  const view = render(element(plugins));
  await waitFor(async () => { await pump(); expect(ready).toHaveBeenCalledTimes(1); }, MOUNTED);
  const api = ready.mock.calls[0][0];
  const base = shared(view, element);
  await loaded(base);
  return {
    ...base,
    point: null,
    zoom: async (scale) => {
      await act(async () => { await api.commands.execute('zoom', { scale }); });
      await base.until(() => records.some((record) => base.geometry(record)?.layout.zoom === scale));
    },
    edit: async (row, col, value) => {
      await act(async () => { await api.applyEdits(setCell(await api.version(), row, col, value)); });
      await pump();
    },
  };
}

function last<T>(values: readonly T[]): T | undefined {
  return values[values.length - 1];
}

const px = (value: string | undefined) => Number.parseFloat(value ?? '');

async function expectAligned(book: Book, record: Decorator, row: number) {
  await book.until(() => book.geometry(record)?.getCellRect({ sheetId: SHEET, row, col: 0 }) != null);
  const rect = book.geometry(record)!.getCellRect({ sheetId: SHEET, row, col: 0 })!;
  const element = book.row(row)!;
  expect(element).not.toBeNull();
  expect(px(element.style.top)).toBeCloseTo(rect.y, 5);
  expect(px(element.style.height)).toBeCloseTo(rect.height, 5);
}

type Handlers = { onClick(): void; onKeyDown(event: unknown): void };

function reactHandlers(element: Element): Handlers {
  const key = Object.keys(element).find((name) => name.startsWith('__reactProps$'))!;
  return (element as unknown as Record<string, Handlers>)[key];
}

const staleKey = { key: 'Enter', preventDefault() {}, stopPropagation() {} };

const buttons = (element: HTMLElement | null) =>
  [...(element?.querySelectorAll<HTMLButtonElement>('button') ?? [])];

for (const editor of ['in-thread', 'peerless', 'legacy-worker'] as const) {
  describe(`XlsxEditor plugin row decorations (${editor})`, () => {
    let bare: { scroll: string | null; canvas: string | null } | null = null;

    it('renders no gutter without plugins', async () => {
      const book = await mount(editor, []);
      const canvas = book.surface.querySelector('canvas')!;
      await book.until(() => canvas.style.width !== '');
      expect(book.gutter()).toBeNull();
      bare = { scroll: book.surface.getAttribute('style'), canvas: canvas.getAttribute('style') };
    }, SLOW);

    it('keeps the grid of main while plugins return null or no items', async () => {
      expect(bare).not.toBeNull();
      let shown: 'null' | 'empty' | 'items' = 'null';
      const toggled = decorator('toggled', (context) => shown === 'null' ? null
        : shown === 'empty' ? { version: context.snapshot.version, items: [] } : dots([2])(context));
      const hookless = decorator('hookless', null);
      const book = await mount(editor, [hookless, toggled]);
      const rect = () => book.geometry(hookless)?.getCellRect({ sheetId: SHEET, row: 2, col: 1 }) ?? null;
      await book.until(() => rect() !== null);
      const before = rect();
      const siblings = book.surface.parentElement!.innerHTML.length - book.surface.outerHTML.length;
      const check = () => {
        expect(book.gutter()).toBeNull();
        expect(book.surface.getAttribute('style')).toBe(bare!.scroll);
        expect(book.surface.querySelector('canvas')!.getAttribute('style')).toBe(bare!.canvas);
        expect(book.surface.parentElement!.innerHTML.length - book.surface.outerHTML.length).toBe(siblings);
        expect(rect()).toEqual(before);
      };
      check();
      shown = 'empty';
      await book.bump(toggled);
      check();
      shown = 'items';
      await book.bump(toggled);
      await book.until(() => book.gutter() !== null);
      expect(book.surface.style.marginLeft).toBe('24px');
      shown = 'null';
      await book.bump(toggled);
      await book.until(() => book.gutter() === null);
      check();
      expect(book.errors).toEqual([]);
    }, SLOW);

    it('places items on their rows through scrolling, zoom, frozen, hidden and resized rows', async () => {
      const record = decorator('rows', dots([0, 1, 3, 4, 7, 40, 500]));
      const book = await mount(editor, [record], { file: shaped });
      await book.until(() => book.row(7) !== null);
      const gutter = book.gutter()!;
      expect(gutter.getAttribute('role')).toBe('region');
      expect(gutter.getAttribute('aria-label')).toBe('Plugin row markers');
      expect(gutter.style.width).toBe('24px');
      expect(book.surface.contains(gutter)).toBe(false);
      for (const row of [0, 1, 4, 7]) await expectAligned(book, record, row);
      expect(book.row(3)).toBeNull();
      expect(book.geometry(record)!.getCellRect({ sheetId: SHEET, row: 3, col: 0 })).toBeNull();
      expect(book.row(40)).toBeNull();
      expect(book.row(500)).toBeNull();
      expect(px(book.row(4)!.style.height)).toBeCloseTo((40 * 4) / 3, 1);
      const frozenTop = px(book.row(0)!.style.top);

      await book.scroll(600);
      await book.until(() => book.row(40) !== null && book.row(7) === null);
      expect(px(book.row(0)!.style.top)).toBeCloseTo(frozenTop, 5);
      for (const row of [0, 1, 40]) await expectAligned(book, record, row);

      await book.scroll(0);
      await book.until(() => book.row(7) !== null);
      await book.zoom(2);
      await book.until(() => book.gutter()?.style.width === '48px' && book.row(7) !== null);
      expect(book.surface.style.marginLeft).toBe('48px');
      await book.until(() => Math.abs(px(book.row(7)!.style.top) -
        (book.geometry(record)?.getCellRect({ sheetId: SHEET, row: 7, col: 0 })?.y ?? NaN)) < 0.01);
      for (const row of [0, 1, 4, 7]) await expectAligned(book, record, row);
      expect(book.errors).toEqual([]);
    }, SLOW);

    it('shows two items per row and collapses the rest into a count', async () => {
      const first = decorator('first', (context) => ({ version: context.snapshot.version, items: [
        { id: 'a', sheetId: sheetOf(context), row: 2, kind: 'badge', text: 'OK', tone: 'success' },
        { id: 'b', sheetId: sheetOf(context), row: 2, kind: 'dot', tone: 'danger', tooltip: 'Failed check' },
      ] }));
      const second = decorator('second', (context) => ({ version: context.snapshot.version, items: [
        { id: 'c', sheetId: sheetOf(context), row: 2, kind: 'badge', text: 'NEW', tooltip: 'Added by import' },
      ] }));
      const book = await mount(editor, [first, second]);
      await book.until(() => book.row(2) !== null);
      const row = book.row(2)!;
      expect(buttons(row).map((button) => [button.dataset.pluginId, button.dataset.decorationId,
        button.getAttribute('aria-label'), button.textContent])).toEqual([
        ['first', 'a', 'OK', 'OK'],
        ['first', 'b', 'Failed check', ''],
      ]);
      const more = row.querySelector<HTMLElement>('[data-testid="xlsx-row-decoration-more"]')!;
      expect(more.textContent).toBe('+1');
      expect(more.title).toBe('Added by import');
      expect(more.getAttribute('aria-label')).toBe('Added by import');
    }, SLOW);

    it('hides a stale set until a document change evaluates it at the current version', async () => {
      let pinned: string | null = null;
      const record = decorator('stale', (context) => ({ ...dots([2])(context), version: pinned ?? context.snapshot.version }));
      const book = await mount(editor, [record]);
      await book.until(() => book.row(2) !== null);
      pinned = record.rendered!.snapshot.version;
      await book.edit(1, 1, 'changed');
      await book.until(() => record.rendered!.snapshot.version !== pinned);
      await book.until(() => book.gutter() === null);
      expect(book.surface.style.marginLeft).toBe('');
      pinned = null;
      const stale = record.rendered!.snapshot.version;
      await book.edit(1, 1, 'again');
      await book.until(() => record.rendered!.snapshot.version !== stale);
      await book.until(() => book.row(2) !== null);
      expect(book.errors).toEqual([]);
    }, SLOW);

    it('quarantines malformed or throwing sets with the row-decorations phase', async () => {
      const item = (context: Context, extra: Record<string, unknown> = {}) =>
        ({ id: 'x', sheetId: sheetOf(context), row: 2, kind: 'dot', ...extra });
      const set = (build: (context: Context) => unknown[]) => (context: Context) =>
        ({ version: context.snapshot.version, items: build(context) });
      const bad: Record<string, (context: Context) => unknown> = {
        'wrong-type': (context) => ({ version: 1, items: [item(context)] }),
        'wrong-item': set((context) => [item(context, { row: '2' })]),
        duplicate: set((context) => [item(context), item(context, { row: 3 })]),
        'long-text': set((context) => [item(context, { kind: 'badge', text: 'ABCDE' })]),
        'row-range': set((context) => [item(context, { row: 1_048_576 })]),
        'row-negative': set((context) => [item(context, { row: -1 })]),
        kind: set((context) => [item(context, { kind: 'star' })]),
        tone: set((context) => [item(context, { tone: 'purple' })]),
        'too-many': set((context) => Array.from({ length: 10_001 }, (_, index) => item(context, { id: `x${index}`, row: index }))),
        throws: () => { throw new Error('boom'); },
      };
      let broken = false;
      const flip = decorator('flip', (context) => {
        if (broken) throw new Error('flipped');
        return dots([5])(context);
      });
      const records = [...Object.entries(bad).map(([id, compute]) => decorator(id, compute)), flip];
      const book = await mount(editor, records);
      await book.until(() => book.errors.length === Object.keys(bad).length && book.row(5) !== null);
      expect(book.errors.map((error) => error.phase)).toEqual(Object.keys(bad).map(() => 'row-decorations'));
      expect(book.errors.map((error) => error.pluginId).sort()).toEqual(Object.keys(bad).sort());
      expect(buttons(book.gutter()).map((button) => button.dataset.pluginId)).toEqual(['flip']);
      broken = true;
      await book.bump(flip);
      await book.until(() => book.gutter() === null);
      expect(last(book.errors)).toMatchObject({ pluginId: 'flip', phase: 'row-decorations' });
      expect(String((last(book.errors)!.error as Error).message)).toBe('flipped');
    }, SLOW);

    it('sends row-decoration-activate to the owner on click, Enter and Space', async () => {
      const first = decorator('first', dots([2], { id: 'one' }));
      const second = decorator('second', dots([2], { id: 'two' }));
      const book = await mount(editor, [first, second]);
      await book.until(() => buttons(book.row(2)).length === 2);
      const [one, two] = buttons(book.row(2));
      fireEvent.click(one);
      await book.until(() => first.activations.length === 1);
      expect(first.activations[0]).toEqual({ type: 'row-decoration-activate', generation: expect.any(String),
        id: 'one', sheetId: SHEET, row: 2 });
      two.focus();
      fireEvent.keyDown(two, { key: 'Enter' });
      await book.until(() => second.activations.length === 1);
      fireEvent.keyDown(two, { key: ' ' });
      await book.until(() => second.activations.length === 2);
      await book.sync();
      expect(first.activations).toHaveLength(1);
      expect(second.activations.map((event) => event.id)).toEqual(['two', 'two']);
      expect(book.errors).toEqual([]);
    }, SLOW);

    it('evaluates and routes a replaced definition with the same id and revision', async () => {
      const original = decorator('x', dots([2], { id: 'old' }), 1);
      const book = await mount(editor, [original]);
      await book.until(() => book.row(2) !== null);
      const replacement = decorator('x', dots([5], { id: 'new' }), 1);
      await book.replace([replacement]);
      await book.until(() => book.row(5) !== null && book.row(2) === null);
      expect(buttons(book.row(5)).map((button) => button.dataset.decorationId)).toEqual(['new']);
      fireEvent.click(buttons(book.row(5))[0]);
      await book.until(() => replacement.activations.length === 1);
      expect(replacement.activations[0]).toMatchObject({ id: 'new', row: 5 });
      await book.sync();
      expect(original.activations).toEqual([]);
      expect(book.errors).toEqual([]);
    }, SLOW);

    it('drops activations of removed or stale items', async () => {
      let rows = [2, 5];
      let pinned: string | null = null;
      const record = decorator('owner', (context) => ({ ...dots(rows)(context), version: pinned ?? context.snapshot.version }));
      const book = await mount(editor, [record]);
      await book.until(() => book.row(2) !== null && book.row(5) !== null);
      const removed = reactHandlers(buttons(book.row(2))[0]);
      rows = [5];
      await book.bump(record);
      await book.until(() => book.row(2) === null && book.row(5) !== null);
      await act(async () => { removed.onClick(); removed.onKeyDown(staleKey); });
      await book.sync();
      expect(record.activations).toEqual([]);

      const stale = reactHandlers(buttons(book.row(5))[0]);
      pinned = record.rendered!.snapshot.version;
      await book.edit(1, 1, 'changed');
      await book.until(() => book.gutter() === null);
      await act(async () => { stale.onClick(); stale.onKeyDown(staleKey); });
      await book.sync();
      expect(record.activations).toEqual([]);

      pinned = null;
      await book.edit(1, 1, 'again');
      await book.until(() => book.row(5) !== null);
      fireEvent.click(buttons(book.row(5))[0]);
      await book.until(() => record.activations.length === 1);
      expect(record.activations[0]).toMatchObject({ id: 'd5', row: 5 });
      expect(book.errors).toEqual([]);
    }, SLOW);

    it('builds a crowded row once per set, not on every paint', async () => {
      const record = decorator('crowd', (context) => ({ version: context.snapshot.version,
        items: Array.from({ length: 10_000 }, (_, index) =>
          ({ id: `o${index}`, sheetId: sheetOf(context), row: 0, kind: 'dot', tooltip: `Item ${index}` })) }));
      const book = await mount(editor, [record], { file: shaped });
      await book.until(() => book.row(0) !== null);
      const chip = () => book.row(0)!.querySelector<HTMLElement>('[data-testid="xlsx-row-decoration-more"]')!;
      expect(buttons(book.row(0))).toHaveLength(2);
      expect(chip().textContent).toBe('+9998');
      const tooltip = [...Array.from({ length: 20 }, (_, index) => `Item ${index + 2}`), '+9978'].join('\n');
      expect(chip().title).toBe(tooltip);
      expect(chip().getAttribute('aria-label')).toBe(tooltip);
      const index = await import('./rowDecorations');
      const built = spyOn(index, 'rowDecorationIndex');
      try {
        for (const top of [200, 400, 600]) {
          const before = book.geometry(record)?.layout.id;
          await book.scroll(top);
          await book.until(() => book.geometry(record) !== null && book.geometry(record)!.layout.id !== before);
          expect(chip().title).toBe(tooltip);
        }
        expect(built).not.toHaveBeenCalled();
        await book.bump(record);
        await book.until(() => built.mock.calls.length > 0);
      } finally {
        built.mockRestore();
      }
      expect(book.errors).toEqual([]);
    }, SLOW);

    it('keeps plugin geometry, point queries and the grid hit test on the right cells beside the gutter', async () => {
      let shown = false;
      const record = decorator('shift', (context) => shown ? dots([4])(context) : null);
      const book = await mount(editor, [record]);
      const canvas = book.surface.querySelector('canvas')!;
      const layer = book.view.getByTestId('plugin-overlays');
      let left = 0;
      const placed = () => ({ x: left, y: 0, left, top: 0, right: left + VIEWPORT.width, bottom: VIEWPORT.height,
        width: VIEWPORT.width, height: VIEWPORT.height, toJSON: () => ({}) }) as DOMRect;
      spyOn(canvas, 'getBoundingClientRect').mockImplementation(placed);
      spyOn(layer, 'getBoundingClientRect').mockImplementation(placed);
      const target = { sheetId: SHEET, row: 4, col: 2 };
      await book.until(() => book.geometry(record)?.getCellRect(target) != null);
      const before = book.geometry(record)!.getCellRect(target)!;
      const at = (clientX: number, clientY: number) => book.geometry(record)?.getPositionAtPoint?.(clientX, clientY) ?? null;
      const edge = before.x + 2;
      const middle = before.y + before.height / 2;
      await book.until(() => at(edge, middle) !== null);
      expect(at(edge, middle)).toMatchObject({ sheet: 0, sheetId: SHEET, row: 4, col: 2 });
      if (book.point) expect(book.point(edge, middle)).toEqual({ sheet: 0, row: 4, col: 2 });

      shown = true;
      await book.bump(record);
      await book.until(() => book.row(4) !== null);
      left = 24;
      expect(layer.parentElement).toBe(canvas.parentElement);
      expect(book.surface.contains(layer)).toBe(true);
      expect(book.gutter()!.contains(canvas)).toBe(false);
      expect(book.surface.style.marginLeft).toBe('24px');
      await book.until(() => book.geometry(record)?.getCellRect(target) != null);
      expect(book.geometry(record)!.getCellRect(target)).toEqual(before);
      await book.until(() => at(left + edge, middle) !== null);
      expect(at(edge, middle)).toMatchObject({ sheet: 0, sheetId: SHEET, row: 4, col: 1 });
      expect(at(left + edge, middle)).toMatchObject({ sheet: 0, sheetId: SHEET, row: 4, col: 2 });
      const centre = [left + before.x + before.width / 2, middle] as const;
      const layout = book.geometry(record)!.layout;
      expect(at(...centre)).toEqual({ sheet: 0, sheetId: SHEET, row: 4, col: 2, version: layout.version, layoutId: layout.id });
      if (book.point) {
        expect(at(...centre)).toEqual({ ...book.point(...centre)!, sheetId: SHEET, version: layout.version, layoutId: layout.id });
        expect(book.point(edge, middle)).toEqual({ sheet: 0, row: 4, col: 1 });
      }
      fireEvent.mouseDown(canvas, { clientX: centre[0], clientY: centre[1], button: 0 });
      fireEvent.mouseUp(window, { button: 0 });
      await book.until(() => last(record.selections)?.cells?.anchor.row === 4);
      expect(last(record.selections)?.cells).toMatchObject({ anchor: { row: 4, col: 2 }, focus: { row: 4, col: 2 } });
      expect(book.errors).toEqual([]);
    }, SLOW);

    it('renders only the painted rows of 10,000 items', async () => {
      const rows = Array.from({ length: 10_000 }, (_, row) => row);
      const record = decorator('many', dots(rows));
      const book = await mount(editor, [record]);
      await book.until(() => book.row(0) !== null);
      const painted = rows.slice(0, 200).filter((row) =>
        book.geometry(record)!.getCellRect({ sheetId: SHEET, row, col: 0 }) !== null);
      expect(painted.length).toBeGreaterThan(5);
      expect(painted.length).toBeLessThan(100);
      expect(buttons(book.gutter()).map((button) => Number(button.dataset.decorationId!.slice(1)))).toEqual(painted);
      await book.scroll(4_000);
      await book.until(() => book.row(0) === null && buttons(book.gutter()).length > 0);
      expect(buttons(book.gutter()).length).toBeLessThan(100);
      expect(book.errors).toEqual([]);
    }, SLOW);

    it('renders the gutter in a read-only editor that hosts plugins', async () => {
      const record = decorator('review', dots([2]));
      const book = await mount(editor, [record], { readOnly: true });
      await book.until(() => book.row(2) !== null);
      fireEvent.click(buttons(book.row(2))[0]);
      await book.until(() => record.activations.length === 1);
      expect(book.errors).toEqual([]);
    }, SLOW);
  });
}
