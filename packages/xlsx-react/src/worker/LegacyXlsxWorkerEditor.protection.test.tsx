import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, spyOn } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import * as xlsx from '@betteroffice/xlsx';
import { initWasm } from '../../../xlsx/src/wasm/loader';
import { createTestWorkbookSession } from '../../../xlsx/src/session/testHelpers';
import { XlsxEditor } from '../XlsxEditor';
import type { XlsxCommandArgs, XlsxCommandId, XlsxCommandResult } from '../commands/types';
import type { XlsxWorkerEditorApi } from './createWorkerEditorApi';
import { editableWorkbookSessionBackend } from './useEditableSessionWorkbook';
import { MOUNTED } from '../__fixtures__/mounted';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const PROTECTED = 'This cell is on a protected sheet.';
const restorers: (() => void)[] = [];
let frames: Map<number, FrameRequestCallback>;
let nextFrame = 0;
let sessionToDispose: xlsx.WorkbookSession | undefined;
let wasm: Uint8Array;

beforeAll(async () => {
  wasm = new Uint8Array(await readFile(new URL('../../../xlsx/src/wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  await initWasm(wasm.slice());
});

beforeEach(() => {
  frames = new Map();
  for (const name of ['Worker', 'OffscreenCanvas'] as const) {
    const original = Object.getOwnPropertyDescriptor(globalThis, name);
    Object.defineProperty(globalThis, name, { configurable: true, writable: true, value: class {} });
    restorers.push(() => {
      if (original) Object.defineProperty(globalThis, name, original);
      else Reflect.deleteProperty(globalThis, name);
    });
  }
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
  for (const [key, value] of [['clientWidth', 800], ['clientHeight', 600]] as const) {
    const original = Object.getOwnPropertyDescriptor(HTMLElement.prototype, key);
    Object.defineProperty(HTMLElement.prototype, key, { configurable: true, get: () => value });
    restorers.push(() => {
      if (original) Object.defineProperty(HTMLElement.prototype, key, original);
      else Reflect.deleteProperty(HTMLElement.prototype, key);
    });
  }
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
    save() {}, restore() {}, setTransform() {}, beginPath() {}, rect() {}, clip() {},
  } as unknown as CanvasRenderingContext2D);
  const paint = spyOn(xlsx, 'paintDisplayList').mockImplementation(() => {});
  const proposals = spyOn(xlsx, 'isProposalsAvailable').mockReturnValue(false);
  restorers.push(() => context.mockRestore(), () => paint.mockRestore(), () => proposals.mockRestore());
});

afterEach(async () => {
  await act(async () => { cleanup(); });
  await sessionToDispose?.dispose();
  sessionToDispose = undefined;
  for (const restore of restorers.reverse()) restore();
  restorers.length = 0;
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

async function workbook(formatCells: boolean) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('xl/styles.xml', '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf applyProtection="1"><protection locked="0"/></xf></cellXfs></styleSheet>');
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>2</v></c></row></sheetData><sheetProtection sheet="1"${formatCells ? ' formatCells="0"' : ''}/></worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}

async function pump() {
  await act(async () => {
    const callbacks = [...frames.values()];
    frames.clear();
    for (const callback of callbacks) callback(0);
  });
}

async function mount(formatCells = false) {
  const bytes = await workbook(formatCells);
  const session = await createTestWorkbookSession(bytes, undefined, { wasm: wasm.slice().buffer, retainPeerHydration: true });
  sessionToDispose = session;
  const open = spyOn(editableWorkbookSessionBackend, 'open').mockResolvedValue(session);
  restorers.push(() => open.mockRestore());
  const saves: Uint8Array[] = [];
  let api: XlsxWorkerEditorApi | undefined;
  const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen file={bytes}
    onSave={(saved) => { saves.push(saved); }} onReady={(ready) => { api = ready; }} />);
  await waitFor(async () => {
    await pump();
    expect(api).toBeDefined();
    expect(api!.hydrated).toBe(true);
  }, MOUNTED);
  const surface = view.getByTestId('xlsx-scroll');
  const select = async (col: number, to = col) => {
    await act(async () => { api!.selectCells(0, { anchor: { row: 0, col }, focus: { row: 0, col: to } }); });
    await pump();
  };
  const input = async (col: number) => (await session.call.cellInputs(0, `${String.fromCharCode(65 + col)}1`)).cells[0][0].input;
  const enter = async (value: string) => {
    fireEvent.keyDown(surface, { key: 'F2' });
    const editor = await waitFor(() => view.getByTestId('xlsx-cell-editor') as HTMLInputElement);
    await waitFor(() => expect(editor.value).not.toBe(''));
    fireEvent.change(editor, { target: { value } });
    fireEvent.keyDown(editor, { key: 'Enter' });
  };
  const execute = async <K extends XlsxCommandId>(id: K, args: XlsxCommandArgs[K]) => {
    let result!: XlsxCommandResult;
    await act(async () => { result = await api!.commands.execute(id, args); });
    return result;
  };
  const refusal = () => view.queryByTestId('xlsx-input-refusal')?.textContent ?? null;
  const saved = async () => {
    expect(await execute('save', null)).toEqual({ ok: true, status: 'executed' });
    return xlsx.openWorkbook(saves[saves.length - 1]);
  };
  return { view, session, surface, api: () => api!, select, input, enter, execute, refusal, saved };
}

describe('legacy worker editor on a protected sheet', () => {
  it('keeps a refused locked-cell draft open after an accepted edit, then cancels and saves both', async () => {
    const book = await mount();
    await book.select(1);
    await book.enter('unlocked');
    await waitFor(async () => expect(await book.input(1)).toBe('unlocked'));
    await book.select(0);
    await book.enter('changed');
    await waitFor(() => expect(book.refusal()).toBe(PROTECTED));
    const editor = book.view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    expect(editor.value).toBe('changed');
    expect(await book.input(0)).toBe('1');
    expect(book.api().failure).toBeNull();
    expect(book.session.failure).toBeUndefined();
    fireEvent.keyDown(editor, { key: 'Escape' });
    const reopened = await book.saved();
    try {
      expect(reopened.cell(0, 0, 0).input).toBe('1');
      expect(reopened.cell(0, 0, 1).input).toBe('unlocked');
    } finally { reopened.dispose(); }
  });

  it('refuses Delete on a locked cell and keeps saving available', async () => {
    const book = await mount();
    await book.select(0, 1);
    fireEvent.keyDown(book.surface, { key: 'Delete' });
    await waitFor(() => expect(book.refusal()).toBe(PROTECTED));
    expect([await book.input(0), await book.input(1)]).toEqual(['1', '2']);
    expect(book.api().failure).toBeNull();
    const reopened = await book.saved();
    try { expect(reopened.cell(0, 0, 0).input).toBe('1'); } finally { reopened.dispose(); }
  });

  it('refuses a paste that reaches a locked cell and keeps saving available', async () => {
    const clipboard = navigator.clipboard;
    const { read, readText } = clipboard;
    clipboard.read = undefined as unknown as typeof read;
    clipboard.readText = () => Promise.resolve('p\tq');
    try {
      const book = await mount();
      await book.select(1);
      fireEvent.keyDown(book.surface, { key: 'v', ctrlKey: true });
      await waitFor(() => expect(book.refusal()).toBe(PROTECTED));
      expect([await book.input(1), await book.input(2)]).toEqual(['2', '']);
      expect(book.api().failure).toBeNull();
      const reopened = await book.saved();
      try { expect(reopened.cell(0, 0, 1).input).toBe('2'); } finally { reopened.dispose(); }
    } finally {
      clipboard.read = read;
      clipboard.readText = readText;
    }
  });

  it('refuses merging cells of a protected sheet; the next save reports the refusal once', async () => {
    const book = await mount(true);
    await book.select(0, 1);
    expect(await book.execute('merge', { value: 'all' })).toMatchObject({ ok: false, failure: { code: 'sheet-protected', message: PROTECTED } });
    await waitFor(() => expect(book.refusal()).toBe(PROTECTED));
    expect(book.api().failure).toBeNull();
    expect(await book.execute('save', null)).toMatchObject({ ok: false });
    const reopened = await book.saved();
    try { expect(reopened.mergedRanges(0, 'A1:B1')).toHaveLength(0); } finally { reopened.dispose(); }
  });

  it('refuses formatting a protected sheet that does not allow it; the next save reports the refusal once', async () => {
    const book = await mount();
    await book.select(1);
    expect(await book.execute('bold', null)).toMatchObject({ ok: false, failure: { code: 'sheet-protected', message: PROTECTED } });
    await waitFor(() => expect(book.refusal()).toBe(PROTECTED));
    expect(book.api().failure).toBeNull();
    expect(await book.execute('save', null)).toMatchObject({ ok: false });
    const reopened = await book.saved();
    try { expect(reopened.selectionFormatting(0, 'B1:B1').bold).toBe(false); } finally { reopened.dispose(); }
  });

  it('formats cells of a protected sheet that allows it', async () => {
    const book = await mount(true);
    await book.select(1);
    expect(await book.execute('bold', null)).toEqual({ ok: true, status: 'executed' });
    const reopened = await book.saved();
    try { expect(reopened.selectionFormatting(0, 'B1:B1').bold).toBe(true); } finally { reopened.dispose(); }
  });
});
