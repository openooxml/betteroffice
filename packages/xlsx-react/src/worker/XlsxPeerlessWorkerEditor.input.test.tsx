import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, mock, spyOn } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import * as xlsx from '@betteroffice/xlsx';
import type { Translations } from '@betteroffice/xlsx-i18n';
import { initWasm } from '../../../xlsx/src/wasm/loader';
import { createTestWorkbookSession } from '../../../xlsx/src/session/testHelpers';
import { XlsxWorkerEditor } from './XlsxWorkerEditor';
import { editableWorkbookSessionBackend } from './useEditableSessionWorkbook';
import type { XlsxAsyncEditorApi } from './createAsyncEditorApi';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
let frames: Map<number, FrameRequestCallback>;
let nextFrame = 0;
let sessionToDispose: xlsx.WorkbookSession | undefined;

beforeAll(async () => {
  await initWasm(await readFile(new URL('../../../xlsx/src/wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
});

beforeEach(() => {
  frames = new Map();
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

async function workbook(protectedSheet: boolean, formatCells = false) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('xl/styles.xml', '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf applyProtection="1"><protection locked="0"/></xf></cellXfs></styleSheet>');
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>2</v></c></row></sheetData>${protectedSheet ? `<sheetProtection sheet="1"${formatCells ? ' formatCells="0"' : ''}/>` : ''}</worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}

async function mount(protectedSheet: boolean, formatCells = false, i18n?: Translations) {
  const bytes = await workbook(protectedSheet, formatCells);
  const session = await createTestWorkbookSession(bytes, undefined, { includeEditState: true });
  sessionToDispose = session;
  const open = spyOn(editableWorkbookSessionBackend, 'open').mockResolvedValue(session);
  restorers.push(() => open.mockRestore());
  const ready = mock((_api: XlsxAsyncEditorApi) => {});
  const view = render(<XlsxWorkerEditor backend="peerless" file={bytes} onReady={ready} i18n={i18n} />);
  await waitFor(async () => {
    await act(async () => {
      const callbacks = [...frames.values()];
      frames.clear();
      for (const callback of callbacks) callback(0);
    });
    expect(ready).toHaveBeenCalledTimes(1);
  });
  return { view, session, api: ready.mock.calls[0][0] };
}

async function enter(view: ReturnType<typeof render>, value: string) {
  fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'F2' });
  const input = await waitFor(() => view.getByTestId('xlsx-cell-editor') as HTMLInputElement);
  await waitFor(() => expect(input.value).not.toBe(''));
  fireEvent.change(input, { target: { value } });
  fireEvent.keyDown(input, { key: 'Enter' });
}

describe('peerless protected cell input', () => {
  for (const formatCells of [false, true]) {
    it(`reports protection and preserves a locked cell when formatCells=${formatCells}`, async () => {
      const { view, session, api } = await mount(true, formatCells);
      const before = session.confirmedBasis.version;
      await enter(view, 'changed');
      await waitFor(() => expect(view.getByTestId('xlsx-input-refusal').textContent).toBe('This cell is on a protected sheet.'));
      expect(await api.commands.execute('save', null)).toMatchObject({ ok: false,
        failure: { code: 'sheet-protected', message: 'This cell is on a protected sheet.' } });
      await expect(api.save()).rejects.toMatchObject({ code: 'sheet-protected', message: 'This cell is on a protected sheet.' });
      expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('1');
      expect(session.confirmedBasis.version).toBe(before);
      expect(session.failure).toBeUndefined();
    });
  }

  it('commits an unlocked cell on a protected sheet', async () => {
    const { view, session } = await mount(true);
    const name = view.getByTestId('xlsx-name-box');
    fireEvent.change(name, { target: { value: 'B1' } });
    fireEvent.keyDown(name, { key: 'Enter' });
    await waitFor(() => expect(session.selection?.range).toMatchObject({ kind: 'rowCol', start: { row: 0, col: 1 } }));
    await enter(view, 'unlocked');
    await waitFor(async () => expect((await session.call.cellInputs(0, 'B1')).cells[0][0].input).toBe('unlocked'));
    expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('1');
    expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
    expect(session.failure).toBeUndefined();
  });

  it('commits a cell on an unprotected sheet', async () => {
    const { view, session } = await mount(false);
    await enter(view, 'unprotected');
    await waitFor(async () => expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('unprotected'));
    expect(view.queryByTestId('xlsx-input-refusal')).toBeNull();
    expect(session.failure).toBeUndefined();
  });

  it('localizes the protection refusal', async () => {
    const { view, api } = await mount(true, false, { commands: { reasons: { sheetProtected: 'Diese Zelle liegt auf einem geschützten Blatt.' } } });
    await enter(view, 'changed');
    await waitFor(() => expect(view.getByTestId('xlsx-input-refusal').textContent).toBe('Diese Zelle liegt auf einem geschützten Blatt.'));
    expect(await api.commands.execute('save', null)).toMatchObject({ ok: false,
      failure: { code: 'sheet-protected', message: 'Diese Zelle liegt auf einem geschützten Blatt.' } });
  });
});
