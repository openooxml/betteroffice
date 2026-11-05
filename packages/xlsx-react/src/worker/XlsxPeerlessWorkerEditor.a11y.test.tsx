import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, mock, spyOn } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import * as xlsx from '@betteroffice/xlsx';
import { initWasm } from '../../../xlsx/src/wasm/loader';
import { createTestWorkbookSession } from '../../../xlsx/src/session/testHelpers';
import { XlsxWorkerEditor } from './XlsxWorkerEditor';
import { editableWorkbookSessionBackend } from './useEditableSessionWorkbook';
import type { XlsxAsyncEditorApi } from './createAsyncEditorApi';
import { MOUNTED } from '../__fixtures__/mounted';

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

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function workbook() {
  const zip = new JSZip();
  const sheet = (n: number) => `<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheet(1)}${sheet(2)}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Second" sheetId="2" r:id="rId2"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1"><v>21</v></c><c r="B1"><f>A1*2</f><v>42</v></c><c r="C1" t="inlineStr"><is><t>Total</t></is></c></row></sheetData></worksheet>`);
  zip.file('xl/worksheets/sheet2.xml', `<worksheet xmlns="${MAIN}"><sheetData/></worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}

async function flush() {
  await act(async () => {
    const callbacks = [...frames.values()];
    frames.clear();
    for (const callback of callbacks) callback(0);
  });
}

async function mount() {
  const bytes = await workbook();
  const session = await createTestWorkbookSession(bytes, undefined, { includeEditState: true });
  sessionToDispose = session;
  const open = spyOn(editableWorkbookSessionBackend, 'open').mockResolvedValue(session);
  restorers.push(() => open.mockRestore());
  const ready = mock((_api: XlsxAsyncEditorApi) => {});
  const view = render(<XlsxWorkerEditor backend="peerless" file={bytes} onReady={ready} />);
  await waitFor(async () => {
    await flush();
    expect(ready).toHaveBeenCalledTimes(1);
  }, MOUNTED);
  return { view, session };
}

async function heard(view: ReturnType<typeof render>, text: string) {
  await waitFor(async () => {
    await flush();
    expect(view.getByTestId('xlsx-live-region').textContent?.replace(/ $/, '')).toBe(text);
  }, MOUNTED);
}

describe('peerless grid screen reader support', () => {
  it('labels the grid focus target with the sheet', async () => {
    const { view } = await mount();
    const grid = view.getByTestId('xlsx-scroll');
    expect(grid.getAttribute('role')).toBe('application');
    expect(grid.getAttribute('aria-roledescription')).toBe('worksheet');
    expect(grid.getAttribute('aria-label')).toBe('Data');
    const region = view.getByTestId('xlsx-live-region');
    expect(region.getAttribute('role')).toBe('status');
    expect(region.getAttribute('aria-live')).toBe('polite');
  });

  it('announces the active cell when the grid takes focus and as it moves', async () => {
    const { view } = await mount();
    const grid = view.getByTestId('xlsx-scroll');
    await act(async () => { grid.focus(); });
    await heard(view, 'A1, 21');
    fireEvent.keyDown(grid, { key: 'ArrowRight' });
    await heard(view, 'B1, 42, has formula');
    fireEvent.keyDown(grid, { key: 'ArrowRight' });
    await heard(view, 'C1, Total');
    fireEvent.keyDown(grid, { key: 'ArrowDown' });
    await heard(view, 'C2, empty');
  });

  it('announces a range with its size', async () => {
    const { view } = await mount();
    const grid = view.getByTestId('xlsx-scroll');
    await act(async () => { grid.focus(); });
    await heard(view, 'A1, 21');
    fireEvent.keyDown(grid, { key: 'ArrowRight', shiftKey: true });
    await heard(view, 'A1 to B1 selected, 1 row by 2 columns');
    fireEvent.keyDown(grid, { key: 'ArrowDown', shiftKey: true });
    await heard(view, 'A1 to B2 selected, 2 rows by 2 columns');
  });

  it('stays quiet while focus is outside the grid', async () => {
    const { view, session } = await mount();
    const name = view.getByTestId('xlsx-name-box') as HTMLInputElement;
    await act(async () => { name.focus(); });
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'ArrowRight' });
    await waitFor(async () => {
      await flush();
      expect(session.selection?.range).toMatchObject({ start: { row: 0, col: 1 } });
    }, MOUNTED);
    await new Promise((resolve) => setTimeout(resolve, 800));
    await flush();
    expect(view.getByTestId('xlsx-live-region').textContent).toBe('');
    await act(async () => { view.getByTestId('xlsx-scroll').focus(); });
    await heard(view, 'B1, 42, has formula');
  });

  it('switches sheets with Ctrl+PageDown and names the new sheet', async () => {
    const { view } = await mount();
    const grid = view.getByTestId('xlsx-scroll');
    await act(async () => { grid.focus(); });
    await heard(view, 'A1, 21');
    fireEvent.keyDown(grid, { key: 'PageDown', ctrlKey: true });
    await heard(view, 'Second, sheet 2 of 2, A1, empty');
    await waitFor(() => expect(grid.getAttribute('aria-label')).toBe('Second'));
    fireEvent.keyDown(grid, { key: 'PageUp', ctrlKey: true });
    await heard(view, 'Data, sheet 1 of 2, A1, 21');
  });

  it('labels the cell editor with its address', async () => {
    const { view } = await mount();
    const grid = view.getByTestId('xlsx-scroll');
    await act(async () => { grid.focus(); });
    fireEvent.keyDown(grid, { key: 'ArrowRight' });
    await heard(view, 'B1, 42, has formula');
    fireEvent.keyDown(grid, { key: 'F2' });
    const input = await waitFor(() => view.getByTestId('xlsx-cell-editor'));
    expect(input.getAttribute('aria-label')).toBe('Edit cell B1');
  });

  it('moves between sheet tabs with roving focus', async () => {
    const { view } = await mount();
    const [data, second] = view.getAllByRole('tab');
    const grid = view.getByTestId('xlsx-scroll');
    expect(data.tabIndex).toBe(0);
    expect(second.tabIndex).toBe(-1);
    expect(data.getAttribute('aria-controls')).toBe(grid.id);
    await act(async () => { data.focus(); });
    fireEvent.keyDown(data, { key: 'ArrowRight' });
    await waitFor(async () => {
      await flush();
      expect(second.getAttribute('aria-selected')).toBe('true');
    }, MOUNTED);
    expect(document.activeElement).toBe(second);
    expect(second.tabIndex).toBe(0);
    expect(data.tabIndex).toBe(-1);
    fireEvent.keyDown(second, { key: 'Home' });
    await waitFor(async () => {
      await flush();
      expect(data.getAttribute('aria-selected')).toBe('true');
    }, MOUNTED);
    expect(view.getByTestId('xlsx-live-region').textContent).toBe('');
  });

  it('cycles focus between toolbar, grid and sheet tabs with F6', async () => {
    const { view } = await mount();
    const grid = view.getByTestId('xlsx-scroll');
    const toolbar = view.getByTestId('xlsx-toolbar');
    await act(async () => { grid.focus(); });
    fireEvent.keyDown(grid, { key: 'F6' });
    expect(document.activeElement).toBe(view.getAllByRole('tab')[0]);
    fireEvent.keyDown(document.activeElement!, { key: 'F6' });
    expect(toolbar.contains(document.activeElement)).toBe(true);
    fireEvent.keyDown(document.activeElement!, { key: 'F6' });
    expect(document.activeElement).toBe(grid);
    await heard(view, 'A1, 21');
    fireEvent.keyDown(grid, { key: 'F6', shiftKey: true });
    expect(toolbar.contains(document.activeElement)).toBe(true);
  });

  it('leaves F6 to an open modal dialog', async () => {
    const { view } = await mount();
    const toolbar = view.getByTestId('xlsx-toolbar');
    const dialog = document.createElement('form');
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    const field = document.createElement('input');
    dialog.append(field);
    toolbar.append(dialog);
    await act(async () => { field.focus(); });
    fireEvent.keyDown(field, { key: 'F6' });
    expect(document.activeElement).toBe(field);
    dialog.remove();
  });

  it('leaves F6 to an open menu', async () => {
    const { view } = await mount();
    const menu = document.createElement('div');
    menu.setAttribute('role', 'menu');
    const item = document.createElement('button');
    item.setAttribute('role', 'menuitem');
    menu.append(item);
    view.getByTestId('xlsx-toolbar').append(menu);
    await act(async () => { item.focus(); });
    fireEvent.keyDown(item, { key: 'F6' });
    expect(document.activeElement).toBe(item);
    menu.remove();
  });
});
