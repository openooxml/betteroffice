/**
 * Grid and chart pointer behaviour against the real wasm core: happy-dom has no
 * layout or canvas, so the viewport size and the 2d context are stubbed and
 * click points come from the same display-list geometry the editor hit-tests.
 *
 * Both suites live here on purpose. happy-dom registers into one shared process
 * global, so a second file with its own register/unregister pair would tear the
 * dom down under whichever suite bun happens to run second.
 */

import { hostProposalWorkbook } from './__fixtures__/hostProposalWorkbook';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, it, spyOn } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { useEffect } from 'react';
import {
  cellRect,
  initWasm,
  isPngExportAvailable,
  openWorkbook,
  selectionAt,
} from '@betteroffice/xlsx';
import * as xlsx from '@betteroffice/xlsx';
import { XLSX_MAX_COLUMNS, XLSX_MAX_ROWS } from '@betteroffice/xlsx/collaboration';
import type {
  CellAddr,
  ChartRegion,
  GridMeta,
  Selection,
  WorkbookHandle,
  XlsxEditResult,
} from '@betteroffice/xlsx';
import type { RenderResult } from '@testing-library/react';
import { isMacPlatform } from './commands/descriptors';
import { useXlsxCommands } from './commands/hooks';
import type {
  XlsxCommandArgs,
  XlsxCommandId,
  XlsxCommandResult,
  XlsxCommandStore,
} from './commands/types';
import { EditorToolbar } from './components/EditorToolbar';
import { columnWidthFromPx, rowHeightFromPx } from './components/SheetGrid';
import { ToolbarCommandButton } from './components/toolbar/ToolbarCommand';
import { XlsxCommandAdmissionError } from './index';
import {
  XlsxEditor,
  XlsxSaveRefusedError,
  type XlsxEditorApi,
  type XlsxInThreadEditorProps,
} from './XlsxEditor';
import { editableWorkbookSessionBackend } from './worker/useEditableSessionWorkbook';
import { MOUNTED } from './__fixtures__/mounted';

const WASM = resolve(import.meta.dir, '../../xlsx/src/wasm/generated/xlsx_wasm_bg.wasm');
const FIXTURE = resolve(import.meta.dir, '../../xlsx/test-fixtures/sample.xlsx');
const CHART_FIXTURE = resolve(import.meta.dir, '../../xlsx/test-fixtures/charts.xlsx');
const UNDRAWABLE_FIXTURE = resolve(
  import.meta.dir,
  '../../xlsx/test-fixtures/unsupported-charts.xlsx'
);
const VIEWPORT = { width: 800, height: 600 };
const LINK_TARGET = 'https://example.com/report';
const LINK_CELL: CellAddr = { row: 5, col: 4 };

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor, within } = await import('@testing-library/react');

function stubContext(): CanvasRenderingContext2D {
  const noop = () => {};
  return {
    save: noop,
    restore: noop,
    setTransform: noop,
    clearRect: noop,
    beginPath: noop,
    rect: noop,
    clip: noop,
    setLineDash: noop,
    moveTo: noop,
    lineTo: noop,
    quadraticCurveTo: noop,
    bezierCurveTo: noop,
    closePath: noop,
    fill: noop,
    stroke: noop,
    fillRect: noop,
    fillText: noop,
    measureText: () => ({ width: 0 }),
  } as unknown as CanvasRenderingContext2D;
}

// the stubs live on shared prototypes, so they are installed and restored
// around this file rather than leaking into every other happy-dom suite.
const LAYOUT = [
  ['clientWidth', VIEWPORT.width],
  ['clientHeight', VIEWPORT.height],
] as const;
const originalGetContext = HTMLCanvasElement.prototype.getContext;
const originalOpen = window.open;
const originalLayout = LAYOUT.map(([property]) => {
  return [property, Object.getOwnPropertyDescriptor(HTMLElement.prototype, property)] as const;
});

// no committed fixture carries a hyperlink, so the engine installs one: a
// structural op through the ops escape hatch, saved back out as workbook bytes.
function withHyperlink(bytes: Uint8Array): Uint8Array {
  const handle = openWorkbook(bytes);
  try {
    handle.applyOps([
      {
        type: 'setHyperlinks',
        sheet: 0,
        hyperlinks: [
          { range: { start: LINK_CELL, end: LINK_CELL }, external_target: LINK_TARGET },
        ],
      },
    ]);
    return handle.save();
  } finally {
    handle.dispose();
  }
}

interface Fixture {
  bytes: Uint8Array;
  grid: GridMeta;
  charts: ChartRegion[];
}

function fixtureFrom(bytes: Uint8Array): Fixture {
  const probe = openWorkbook(bytes);
  try {
    const frame = probe.displayList({ x: 0, y: 0, ...VIEWPORT });
    return { bytes, grid: frame.grid as GridMeta, charts: frame.charts ?? [] };
  } finally {
    probe.dispose();
  }
}

let plain: Fixture;
let linked: Fixture;
let charted: Fixture;
let undrawable: Fixture;
let opened: string[] = [];

beforeAll(async () => {
  HTMLCanvasElement.prototype.getContext = (() =>
    stubContext()) as unknown as HTMLCanvasElement['getContext'];
  for (const [property, value] of LAYOUT) {
    Object.defineProperty(HTMLElement.prototype, property, {
      configurable: true,
      get: () => value,
    });
  }
  window.open = ((url?: string | URL) => {
    opened.push(String(url));
    return null;
  }) as typeof window.open;
  await initWasm(new Uint8Array(readFileSync(WASM)));
  const source = new Uint8Array(readFileSync(FIXTURE));
  plain = fixtureFrom(source);
  linked = fixtureFrom(withHyperlink(source));
  charted = fixtureFrom(new Uint8Array(readFileSync(CHART_FIXTURE)));
  undrawable = fixtureFrom(new Uint8Array(readFileSync(UNDRAWABLE_FIXTURE)));
}, 30_000);

afterAll(async () => {
  HTMLCanvasElement.prototype.getContext = originalGetContext;
  window.open = originalOpen;
  for (const [property, descriptor] of originalLayout) {
    if (descriptor) Object.defineProperty(HTMLElement.prototype, property, descriptor);
  }
  // last: bun shares one process across test files, and happy-dom's fetch
  // rejects the file: urls other suites initialise their wasm from.
  await GlobalRegistrator.unregister();
});

afterEach(async () => {
  await act(async () => { cleanup(); });
  opened = [];
});

function pointAt(fixture: Fixture, addr: CellAddr): { clientX: number; clientY: number } {
  const rect = cellRect(fixture.grid, addr.row, addr.col);
  if (!rect) throw new Error(`cell ${addr.row},${addr.col} is outside the painted window`);
  return { clientX: rect.x + rect.w / 2, clientY: rect.y + rect.h / 2 };
}

async function mountEditor(
  fixture: Fixture = plain, onSave?: (bytes: Uint8Array) => void, showSheetHeaders = false, fillHandle = true,
) {
  const ready: { handle: WorkbookHandle | null; api: XlsxEditorApi | null } = { handle: null, api: null };
  const view = render(
    <XlsxEditor legacySyncApi experimentalWorkerOpen={false} showSheetHeaders={showSheetHeaders} fillHandle={fillHandle}
      file={fixture.bytes.slice()}
      onSave={onSave}
      onReady={(api) => {
        ready.handle = api.handle;
        ready.api = api;
      }}
    />
  );
  const nameBox = () => view.getByTestId('xlsx-name-box') as HTMLInputElement;
  await waitFor(() => expect(nameBox().value).toBe('A1'), MOUNTED);
  const surface = view.getByTestId('xlsx-scroll');
  const editor = () => view.queryByTestId('xlsx-cell-editor') as HTMLInputElement | null;
  const press = (addr: CellAddr) => {
    fireEvent.mouseDown(surface, pointAt(fixture, addr));
    fireEvent.mouseUp(surface, pointAt(fixture, addr));
    fireEvent.click(surface, pointAt(fixture, addr));
  };
  return {
    view,
    surface,
    nameBox,
    editor,
    workbook: () => ready.handle!,
    api: () => ready.api!,
    reopenWith: (next: Fixture) =>
      act(async () => {
        view.rerender(
          <XlsxEditor legacySyncApi experimentalWorkerOpen={false} showSheetHeaders={showSheetHeaders} fillHandle={fillHandle}
            file={next.bytes.slice()}
            onSave={onSave}
            onReady={(api) => {
              ready.handle = api.handle;
              ready.api = api;
            }}
          />
        );
      }),
    click: press,
    doubleClick: (addr: CellAddr) => {
      press(addr);
      press(addr);
      fireEvent.doubleClick(surface, pointAt(fixture, addr));
    },
    pressInEditor: (addr: CellAddr) => fireEvent.mouseDown(editor()!, pointAt(fixture, addr)),
    doubleClickInEditor: (addr: CellAddr) => {
      const target = editor()!;
      fireEvent.mouseDown(target, pointAt(fixture, addr));
      fireEvent.mouseUp(target, pointAt(fixture, addr));
      fireEvent.click(target, pointAt(fixture, addr));
      fireEvent.doubleClick(target, pointAt(fixture, addr));
    },
    type: (value: string) => fireEvent.change(editor()!, { target: { value } }),
    formula: () => (view.getByTestId('xlsx-formula-input') as HTMLInputElement).value,
    canUndo: () => {
      const undo = view.getByTestId('xlsx-undo') as HTMLButtonElement;
      return !undo.disabled && undo.getAttribute('aria-disabled') !== 'true';
    },
    outline: () => view.queryByTestId('xlsx-chart-selection'),
    selectionBox: () => view.queryByTestId('xlsx-selection'),
    error: () => view.queryByTestId('xlsx-error'),
    outlineAt: () => {
      const box = view.getByTestId('xlsx-chart-selection');
      return {
        x: Math.round(parseFloat(box.style.left)),
        y: Math.round(parseFloat(box.style.top)),
      };
    },
  };
}

function mountWithHeaders(fixture: Fixture = plain) {
  return mountEditor(fixture, undefined, true);
}

function headerlessActiveCell(surface: HTMLElement): HTMLElement {
  const boxes = Array.from(surface.querySelectorAll<HTMLElement>('div:not([data-testid])'))
    .filter((box) => box.style.border === '2px solid #217346');
  expect(boxes).toHaveLength(1);
  return boxes[0]!;
}

// the centre of a chart's visible region, which is where the editor paints it.
function chartCenter(region: ChartRegion): { clientX: number; clientY: number } {
  return {
    clientX: region.clip.x + region.clip.w / 2,
    clientY: region.clip.y + region.clip.h / 2,
  };
}

function chartRounded(region: ChartRegion): { x: number; y: number } {
  return { x: Math.round(region.rect.x), y: Math.round(region.rect.y) };
}

// past the window an arrow burst stays local in, so it has landed or been
// discarded by the time the assertion runs.
function settle(): Promise<void> {
  return act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 400));
  });
}

async function selectChart(
  view: Awaited<ReturnType<typeof mountEditor>>,
  chart: ChartRegion
): Promise<void> {
  fireEvent.mouseDown(view.surface, chartCenter(chart));
  await waitFor(() => view.outline()!);
  fireEvent.mouseUp(window, chartCenter(chart));
}

describe('XlsxEditor grid pointer handling', () => {
  it('saves source bytes through the API after mounting and painting', async () => {
    const source = new Uint8Array(readFileSync(resolve(
      import.meta.dir,
      '../../xlsx/test-fixtures/no-edit-save.xlsx'
    )));
    let api: XlsxEditorApi | undefined;
    let changes = 0;
    const view = render(
      <XlsxEditor legacySyncApi experimentalWorkerOpen={false}
        file={source.slice()}
        onChange={() => changes++}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    await waitFor(() => expect(view.getAllByRole('gridcell').length).toBeGreaterThan(0), MOUNTED);
    expect(api!.save()).toEqual(source);
    expect(changes).toBe(0);
  });

  it('shows the column letters and row numbers for the painted frame', async () => {
    const view = await mountWithHeaders();
    const grid = view.surface.closest('[data-testid="xlsx-grid"]')!;
    expect(grid.querySelector('[data-testid="xlsx-column-headers"] > div')?.textContent).toBe('A');
    expect(grid.querySelector('[data-testid="xlsx-row-headers"] > div')?.textContent).toBe('1');
  });

  it('hides the header bands by default', async () => {
    const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()} />);
    await waitFor(() => expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A1'), MOUNTED);
    expect(view.queryByTestId('xlsx-grid')).toBeNull();
    expect(view.queryByTestId('xlsx-column-headers')).toBeNull();
    expect(view.queryByTestId('xlsx-row-headers')).toBeNull();
  });

  it('sends no ops when a whole-sheet row boundary is dragged', async () => {
    const view = await mountWithHeaders();
    const grid = view.surface.closest('[data-testid="xlsx-grid"]')!;
    fireEvent.click(grid.querySelector('[data-testid="xlsx-grid-corner"]')!);
    await waitFor(() => expect(selectedAddresses(view))
      .toHaveLength(VIEWPORT_ROWS * VIEWPORT_COLUMNS), MOUNTED);
    const applyOps = spyOn(view.workbook(), 'applyOps');
    const band = grid.querySelector('[data-testid="xlsx-row-headers"]')!;
    const start = plain.grid.rowOffsets[1]!;
    fireEvent.pointerDown(band, { button: 0, clientX: 20, clientY: start, pointerId: 9 });
    fireEvent.pointerMove(band, { clientX: 20, clientY: start + 32, pointerId: 9 });
    fireEvent.pointerUp(band, { clientX: 20, clientY: start + 32, pointerId: 9 });
    expect(applyOps).not.toHaveBeenCalled();
  });

  // the accessible mirror reports every selected cell, so a whole-track
  // selection is visible there however far past the painted window it reaches.
  const VIEWPORT_ROWS = 31;
  const VIEWPORT_COLUMNS = 10;
  const selectedAddresses = (view: Awaited<ReturnType<typeof mountEditor>>) =>
    Array.from(view.view.queryAllByRole('gridcell', { selected: true }) as Element[],
      (cell) => (cell.textContent ?? '').split(',')[0]!.trim());

  const columnBand = (view: { surface: HTMLElement }) =>
    view.surface.closest('[data-testid="xlsx-grid"]')!
      .querySelector('[data-testid="xlsx-column-headers"]')!;

  // happy-dom reports no layout, so the band's rect starts at the origin and a
  // clientX is a band offset; a label's centre is the track's offset plus half
  // its size, which is the same geometry the existing reorder test drives.
  const columnPoint = (col: number) =>
    (plain.grid.colOffsets[col]! + plain.grid.colOffsets[col + 1]!) / 2;

  it('selects the whole column from the column header', async () => {
    const view = await mountWithHeaders();
    const grid = view.surface.closest('[data-testid="xlsx-grid"]')!;

    fireEvent.mouseDown(grid.querySelector('[data-testid="xlsx-column-headers"] [data-col="0"]')!);
    // the selection is the sheet's own row extent, not the painted window's, so
    // every painted cell in the column reads as selected.
    await waitFor(() => expect(selectedAddresses(view).length).toBe(VIEWPORT_ROWS), MOUNTED);
    expect(selectedAddresses(view).every((address) => address.startsWith('A'))).toBe(true);
    expect(selectedAddresses(view).slice(-1)[0]).toBe(`A${VIEWPORT_ROWS}`);
    expect(view.selectionBox()!.style.height).toBe(
      `${Math.round(plain.grid.rowOffsets[VIEWPORT_ROWS]!) - Math.round(plain.grid.rowOffsets[0]!)}px`);
    expect(view.selectionBox()!.style.left).toBe('0px');
    // the selected column fills its own label strongly, so the bands did see
    // the selection; only the opposite axis stays clear.
    const column = grid.querySelector<HTMLElement>('[data-testid="xlsx-column-headers"] [data-col="0"]')!;
    expect(column.style.backgroundColor).toBe('#e3e7ea');
    // a whole-column selection fills its own band, never the other axis: its row
    // extent is the whole sheet, so tinting rows would flood the row header.
    const tintedRows = Array.from(grid.querySelectorAll<HTMLElement>('[data-testid="xlsx-row-headers"] [data-row]'),
      (label) => label.style.backgroundColor).filter((fill) => fill !== '');
    expect(tintedRows).toEqual([]);
  });

  it('draws the cell selection as a snapped edge ring, not a fill', async () => {
    const view = await mountWithHeaders();
    view.click({ row: 1, col: 1 });
    const marquee = await waitFor(() => view.selectionBox()!);
    // a one-pixel inset ring, so the element keeps a zero border. happy-dom keeps the hex as
    // written; a browser serializes it as rgb(33, 115, 70).
    expect(marquee.style.boxShadow).toBe('inset 0 0 0 1px #217346');
    expect(marquee.style.border).toBe('');
    const cell = cellRect(plain.grid, 1, 1)!;
    expect(marquee.style.left).toBe(`${Math.round(cell.x)}px`);
    expect(marquee.style.top).toBe(`${Math.round(cell.y)}px`);
    expect(marquee.style.width).toBe(`${Math.round(cell.x + cell.w) - Math.round(cell.x)}px`);
    expect(marquee.style.height).toBe(`${Math.round(cell.y + cell.h) - Math.round(cell.y)}px`);
    expect(marquee.style.background).toBe('');
    expect(marquee.style.backgroundColor).toBe('');
    // one cell: the two-pixel active-cell ring straddles that same cell's
    // snapped edges, so its box grows one pixel per side, and there is no wash.
    const active = view.view.getByTestId('xlsx-active-cell');
    expect([active.style.left, active.style.top, active.style.width, active.style.height]).toEqual([
      `${Math.round(cell.x) - 1}px`, `${Math.round(cell.y) - 1}px`,
      `${Math.round(cell.x + cell.w) - Math.round(cell.x) + 2}px`,
      `${Math.round(cell.y + cell.h) - Math.round(cell.y) + 2}px`,
    ]);
    expect(active.style.boxShadow).toBe('inset 0 0 0 2px #217346');
    expect(active.style.border).toBe('');
    expect(view.view.queryAllByTestId('xlsx-selection-fill')).toEqual([]);
  });

  it('keeps the bordered selection and active-cell boxes without headers', async () => {
    const view = await mountEditor();
    view.click({ row: 1, col: 1 });
    const marquee = await waitFor(() => view.selectionBox()!);
    const cell = cellRect(plain.grid, 1, 1)!;
    expect(marquee.style.border).toBe('1px solid #217346');
    expect(marquee.style.background).toBe('rgba(33, 115, 70, 0.12)');
    expect(marquee.style.boxShadow).toBe('');
    expect(parseFloat(marquee.style.left)).toBeCloseTo(cell.x);
    expect(parseFloat(marquee.style.width)).toBeCloseTo(cell.w);
    expect(view.view.queryByTestId('xlsx-active-cell')).toBeNull();
    const active = headerlessActiveCell(view.surface);
    expect(active.style.boxShadow).toBe('');
    expect(parseFloat(active.style.left)).toBeCloseTo(cell.x);
    expect(parseFloat(active.style.height)).toBeCloseTo(cell.h);
    expect(view.view.queryAllByTestId('xlsx-selection-fill')).toEqual([]);
  });

  it('washes a range but leaves the active cell clear', async () => {
    const view = await mountWithHeaders();
    // anchor at D4 and drag back to B2, so the active cell is the range's
    // top-left corner: the wash then splits into the band under the active row
    // and the band beside the active column.
    fireEvent.mouseDown(view.surface, pointAt(plain, { row: 3, col: 3 }));
    await act(async () => {
      fireEvent.mouseMove(view.surface, { ...pointAt(plain, { row: 1, col: 1 }), buttons: 1 });
    });
    await act(async () => {
      fireEvent.mouseUp(window, pointAt(plain, { row: 1, col: 1 }));
    });
    const fills = await waitFor(() => {
      const found = view.view.queryAllByTestId('xlsx-selection-fill') as HTMLElement[];
      expect(found.length).toBe(2);
      return found;
    });
    const focus = cellRect(plain.grid, 1, 1)!;
    const range = cellRect(plain.grid, 3, 3)!;
    const [below, beside] = fills.map((fill) => ({
      left: parseFloat(fill.style.left),
      top: parseFloat(fill.style.top),
      width: parseFloat(fill.style.width),
      height: parseFloat(fill.style.height),
      background: fill.style.background,
    }));
    // the first strip starts at the active cell's bottom edge and spans the
    // whole range, so the active row itself stays clear.
    expect(below).toEqual({
      left: Math.round(focus.x),
      top: Math.round(focus.y + focus.h),
      width: Math.round(range.x + range.w) - Math.round(focus.x),
      height: Math.round(range.y + range.h) - Math.round(focus.y + focus.h),
      background: 'rgba(33, 115, 70, 0.12)',
    });
    // the second runs beside the active cell, from its right edge to the range.
    expect(beside).toEqual({
      left: Math.round(focus.x + focus.w),
      top: Math.round(focus.y),
      width: Math.round(range.x + range.w) - Math.round(focus.x + focus.w),
      height: Math.round(focus.y + focus.h) - Math.round(focus.y),
      background: 'rgba(33, 115, 70, 0.12)',
    });
    // swallowed either way: no strip may overlap the active cell's snapped rect.
    const snapped = {
      left: Math.round(focus.x),
      top: Math.round(focus.y),
      right: Math.round(focus.x + focus.w),
      bottom: Math.round(focus.y + focus.h),
    };
    const overlaps = fills.some((fill) => {
      const left = parseFloat(fill.style.left);
      const top = parseFloat(fill.style.top);
      const right = left + parseFloat(fill.style.width);
      const bottom = top + parseFloat(fill.style.height);
      return left < snapped.right && right > snapped.left
        && top < snapped.bottom && bottom > snapped.top;
    });
    expect(overlaps).toBe(false);
    // the active-cell ring snaps to the active cell's own boundaries, not the
    // range's.
    const active = view.view.getByTestId('xlsx-active-cell');
    expect(active.style.left).toBe(`${Math.round(focus.x) - 1}px`);
    expect(active.style.top).toBe(`${Math.round(focus.y) - 1}px`);
  });

  it('selects the whole row from the row header', async () => {
    const view = await mountWithHeaders();
    const grid = view.surface.closest('[data-testid="xlsx-grid"]')!;

    fireEvent.mouseDown(grid.querySelector('[data-testid="xlsx-row-headers"] [data-row="0"]')!);
    await waitFor(() => expect(selectedAddresses(view).length).toBe(VIEWPORT_COLUMNS), MOUNTED);
    expect(selectedAddresses(view)).toEqual(
      ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1', 'J1']);
    expect(view.selectionBox()!.style.width).toBe(
      `${Math.round(plain.grid.colOffsets[VIEWPORT_COLUMNS]!) - Math.round(plain.grid.colOffsets[0]!)}px`);
    expect(view.selectionBox()!.style.height).toBe(
      `${Math.round(plain.grid.rowOffsets[1]!) - Math.round(plain.grid.rowOffsets[0]!)}px`);
  });

  it('selects the whole sheet from the corner above the row numbers', async () => {
    const view = await mountWithHeaders();
    const grid = view.surface.closest('[data-testid="xlsx-grid"]')!;
    const applyOps = spyOn(view.workbook(), 'applyOps');
    const a11y = spyOn(xlsx, 'buildA11yGrid');
    try {
      fireEvent.click(grid.querySelector('[data-testid="xlsx-grid-corner"]')!);
      // the explicit whole-sheet range reaches past the window: every painted
      // cell reads as selected, which a window-sized range would not do here.
      await waitFor(() => expect(selectedAddresses(view))
        .toHaveLength(VIEWPORT_ROWS * VIEWPORT_COLUMNS), MOUNTED);
      // the accessible mirror is handed the raw selection, so its range is the
      // sheet's own extent rather than the painted one.
      const calls = a11y.mock.calls;
      expect(calls[calls.length - 1]![1]).toEqual({
        anchor: { row: 0, col: 0 },
        focus: { row: 0, col: 0 },
        range: { top: 0, left: 0, bottom: XLSX_MAX_ROWS - 1, right: XLSX_MAX_COLUMNS - 1 },
      } satisfies Selection);
      expect(applyOps).not.toHaveBeenCalled();
    } finally {
      a11y.mockRestore();
    }
  });

  it('extends the whole-column selection from the anchor on shift', async () => {
    const view = await mountWithHeaders();
    const columns = view.surface.closest('[data-testid="xlsx-grid"]')!
      .querySelector('[data-testid="xlsx-column-headers"]')!;

    fireEvent.mouseDown(columns.querySelector('[data-col="0"]')!);
    await waitFor(() => expect(selectedAddresses(view).length).toBe(VIEWPORT_ROWS), MOUNTED);
    fireEvent.mouseDown(columns.querySelector('[data-col="2"]')!, { shiftKey: true });
    // the anchor holds at A1 and the focus moves to C1, so A:C spans the sheet.
    await waitFor(() => expect(selectedAddresses(view)[selectedAddresses(view).length - 1])
      .toBe(`C${VIEWPORT_ROWS}`), MOUNTED);
    expect(selectedAddresses(view)).toHaveLength(VIEWPORT_ROWS * 3);
    expect(selectedAddresses(view).slice(0, 3)).toEqual(['A1', 'B1', 'C1']);
  });

  it('dragging a selected column header sends no move op', async () => {
    const view = await mountWithHeaders();
    const band = view.surface.closest('[data-testid="xlsx-grid"]')!
      .querySelector('[data-testid="xlsx-column-headers"]')!;
    const applyOps = spyOn(view.workbook(), 'applyOps');
    fireEvent.mouseDown(band.querySelector('[data-col="5"]')!);
    await waitFor(() => expect(selectedAddresses(view)
      .filter((address) => address.startsWith('F'))).toHaveLength(VIEWPORT_ROWS), MOUNTED);
    // F's label centre; dropping just before H's midpoint lands the gap after G.
    const start = (plain.grid.colOffsets[5]! + plain.grid.colOffsets[6]!) / 2;
    const drop = plain.grid.colOffsets[7]! - 1;
    fireEvent.pointerDown(band.querySelector('[data-col="5"]')!, {
      button: 0, clientX: start, clientY: 5, pointerId: 1 });
    fireEvent.pointerMove(band, { clientX: drop, clientY: 5, pointerId: 1 });
    fireEvent.pointerUp(band, { clientX: drop, clientY: 5, pointerId: 1 });
    expect(applyOps).not.toHaveBeenCalled();
  });

  it('extends the whole-column selection track by track when a label is dragged', async () => {
    const view = await mountWithHeaders();
    const band = columnBand(view);
    fireEvent.mouseDown(band.querySelector('[data-col="0"]')!);
    fireEvent.mouseDown(band.querySelector('[data-col="2"]')!, { shiftKey: true });
    await waitFor(() => expect(selectedAddresses(view)).toHaveLength(VIEWPORT_ROWS * 3), MOUNTED);
    const applyOps = spyOn(view.workbook(), 'applyOps');

    // D is outside A:C, so the press starts a selection drag rather than a
    // reorder; the press alone changes nothing, only crossing a track extends.
    fireEvent.pointerDown(band.querySelector('[data-col="3"]')!, {
      button: 0, clientX: columnPoint(3), clientY: 5, pointerId: 1 });
    expect(selectedAddresses(view)).toHaveLength(VIEWPORT_ROWS * 3);
    fireEvent.pointerMove(band, { clientX: columnPoint(4), clientY: 5, pointerId: 1 });
    await waitFor(() => expect(selectedAddresses(view).slice(-1)[0]).toBe(`E${VIEWPORT_ROWS}`), MOUNTED);
    expect(selectedAddresses(view)).toHaveLength(VIEWPORT_ROWS * 5);
    fireEvent.pointerMove(band, { clientX: columnPoint(5), clientY: 5, pointerId: 1 });
    fireEvent.pointerUp(band, { clientX: columnPoint(5), clientY: 5, pointerId: 1 });

    // extending from the press's own anchor spans A through F, whole tracks.
    const addresses = selectedAddresses(view);
    expect(addresses).toHaveLength(VIEWPORT_ROWS * 6);
    expect(addresses.slice(0, 6)).toEqual(['A1', 'B1', 'C1', 'D1', 'E1', 'F1']);
    expect(addresses.slice(-1)[0]).toBe(`F${VIEWPORT_ROWS}`);
    // a selection drag is not a structural edit.
    expect(applyOps).not.toHaveBeenCalled();
  });

  it('dragging a header inside a selected run sends no move op', async () => {
    const view = await mountWithHeaders();
    const band = columnBand(view);
    fireEvent.mouseDown(band.querySelector('[data-col="0"]')!);
    fireEvent.mouseDown(band.querySelector('[data-col="2"]')!, { shiftKey: true });
    await waitFor(() => expect(selectedAddresses(view)).toHaveLength(VIEWPORT_ROWS * 3), MOUNTED);
    const applyOps = spyOn(view.workbook(), 'applyOps');

    // B lies inside the fully selected A:C run and drops past D's midpoint.
    fireEvent.pointerDown(band.querySelector('[data-col="1"]')!, {
      button: 0, clientX: columnPoint(1), clientY: 5, pointerId: 1 });
    fireEvent.pointerMove(band, { clientX: columnPoint(3) + 1, clientY: 5, pointerId: 1 });
    fireEvent.pointerUp(band, { clientX: columnPoint(3) + 1, clientY: 5, pointerId: 1 });

    expect(applyOps).not.toHaveBeenCalled();
  });

  it('resizes a whole selected column run as one batch of width ops', async () => {
    const view = await mountWithHeaders();
    const band = columnBand(view);
    fireEvent.mouseDown(band.querySelector('[data-col="1"]')!);
    fireEvent.mouseDown(band.querySelector('[data-col="3"]')!, { shiftKey: true });
    await waitFor(() => expect(selectedAddresses(view))
      .toHaveLength(VIEWPORT_ROWS * 3), MOUNTED);
    const applyOps = spyOn(view.workbook(), 'applyOps');

    // the D/E boundary is at E's leading edge; the drag widens D by 35px. The
    // band measures the track from the renderer's full size, not the offset
    // gap, so the expectation starts from the same number.
    const start = plain.grid.colOffsets[4]!;
    const size = plain.grid.colSizes?.[3] ?? start - plain.grid.colOffsets[3]!;
    fireEvent.pointerDown(band, { button: 0, clientX: start, clientY: 5, pointerId: 2 });
    fireEvent.pointerMove(band, { clientX: start + 35, clientY: 5, pointerId: 2 });
    fireEvent.pointerUp(band, { clientX: start + 35, clientY: 5, pointerId: 2 });

    // one edit carries every selected track, so undo takes the gesture back in
    // one step.
    expect(applyOps).toHaveBeenCalledTimes(1);
    const width = columnWidthFromPx(size + 35);
    expect(applyOps.mock.calls[0]![0]).toEqual([
      { type: 'setColWidth', sheet: 0, col: 1, width },
      { type: 'setColWidth', sheet: 0, col: 2, width },
      { type: 'setColWidth', sheet: 0, col: 3, width },
    ]);
  });

  it('resizes a lone column from its boundary with one op', async () => {
    const view = await mountWithHeaders();
    const band = columnBand(view);
    const applyOps = spyOn(view.workbook(), 'applyOps');

    // no whole-track selection: the B/C boundary widens B alone. As above,
    // the band starts from the renderer's full track size.
    const start = plain.grid.colOffsets[2]!;
    const size = plain.grid.colSizes?.[1] ?? start - plain.grid.colOffsets[1]!;
    fireEvent.pointerDown(band, { button: 0, clientX: start, clientY: 5, pointerId: 3 });
    fireEvent.pointerMove(band, { clientX: start + 35, clientY: 5, pointerId: 3 });
    fireEvent.pointerUp(band, { clientX: start + 35, clientY: 5, pointerId: 3 });

    expect(applyOps).toHaveBeenCalledTimes(1);
    expect(applyOps.mock.calls[0]![0]).toEqual([
      { type: 'setColWidth', sheet: 0, col: 1, width: columnWidthFromPx(size + 35) },
    ]);
  });

  it('resizes a whole selected row run as one batch of height ops', async () => {
    const view = await mountWithHeaders();
    const band = view.surface.closest('[data-testid="xlsx-grid"]')!
      .querySelector('[data-testid="xlsx-row-headers"]')!;
    fireEvent.mouseDown(band.querySelector('[data-row="1"]')!);
    fireEvent.mouseDown(band.querySelector('[data-row="3"]')!, { shiftKey: true });
    await waitFor(() => expect(selectedAddresses(view))
      .toHaveLength(VIEWPORT_COLUMNS * 3), MOUNTED);
    const applyOps = spyOn(view.workbook(), 'applyOps');

    // the row 4/5 boundary sits at row 5's leading edge; the drag grows row 4
    // by 32px, in points.
    const start = plain.grid.rowOffsets[4]!;
    const size = start - plain.grid.rowOffsets[3]!;
    fireEvent.pointerDown(band, { button: 0, clientX: 20, clientY: start, pointerId: 4 });
    fireEvent.pointerMove(band, { clientX: 20, clientY: start + 32, pointerId: 4 });
    fireEvent.pointerUp(band, { clientX: 20, clientY: start + 32, pointerId: 4 });

    expect(applyOps).toHaveBeenCalledTimes(1);
    const height = rowHeightFromPx(size + 32);
    expect(applyOps.mock.calls[0]![0]).toEqual([
      { type: 'setRowHeight', sheet: 0, row: 1, height },
      { type: 'setRowHeight', sheet: 0, row: 2, height },
      { type: 'setRowHeight', sheet: 0, row: 3, height },
    ]);
  });

  it('extends from the anchor when a shift-pressed header is dragged', async () => {
    const view = await mountWithHeaders();
    const band = columnBand(view);
    fireEvent.mouseDown(band.querySelector('[data-col="0"]')!);
    await waitFor(() => expect(selectedAddresses(view)).toHaveLength(VIEWPORT_ROWS), MOUNTED);
    // shift-press D keeps the anchor at A while selecting A:D.
    fireEvent.mouseDown(band.querySelector('[data-col="3"]')!, { shiftKey: true });
    await waitFor(() => expect(selectedAddresses(view)).toHaveLength(VIEWPORT_ROWS * 4), MOUNTED);
    const applyOps = spyOn(view.workbook(), 'applyOps');

    // a modifier press always extends, so the drag keeps extending from A
    // instead of picking D up as a reorder.
    fireEvent.pointerDown(band.querySelector('[data-col="3"]')!, {
      button: 0, clientX: columnPoint(3), clientY: 5, pointerId: 1, shiftKey: true });
    fireEvent.pointerMove(band, { clientX: columnPoint(5), clientY: 5, pointerId: 1, shiftKey: true });
    fireEvent.pointerUp(band, { clientX: columnPoint(5), clientY: 5, pointerId: 1 });

    const addresses = selectedAddresses(view);
    expect(addresses).toHaveLength(VIEWPORT_ROWS * 6);
    expect(addresses.slice(0, 6)).toEqual(['A1', 'B1', 'C1', 'D1', 'E1', 'F1']);
    expect(addresses.slice(-1)[0]).toBe(`F${VIEWPORT_ROWS}`);
    expect(applyOps).not.toHaveBeenCalled();
  });

  // the keyboard listener lives on the scroller, so focus has to be there for a
  // key the user sends after a header press to reach the grid at all.
  it('takes focus from a column header press, so the next key moves the selection', async () => {
    const view = await mountWithHeaders();
    const band = columnBand(view);
    // a cell press is what normally focuses the scroller; blur puts focus back
    // on the body, the state a press on a toolbar or sheet control leaves.
    view.click({ row: 2, col: 2 });
    await waitFor(() => expect(document.activeElement).toBe(view.surface), MOUNTED);
    (view.surface as HTMLElement).blur();
    expect(document.activeElement).not.toBe(view.surface);

    // the fixture merges A1:D1, and a whole-column press anchors its focus on
    // row 0, so E is the first column whose top cell is outside that merge and
    // the next column right of it is a plain one-column move.
    fireEvent.mouseDown(band.querySelector('[data-col="4"]')!);
    await waitFor(() => expect(selectedAddresses(view)).toHaveLength(VIEWPORT_ROWS), MOUNTED);
    expect(document.activeElement).toBe(view.surface);
    // keys go to whatever holds focus, so this is the event a real key sends.
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowRight' });
    await waitFor(() => expect(view.nameBox().value).toBe('F1'), MOUNTED);
  });

  it('freezes the columns left of the header the menu opened on', async () => {
    const view = await mountWithHeaders();
    const band = view.surface.closest('[data-testid="xlsx-grid"]')!
      .querySelector('[data-testid="xlsx-column-headers"]')!;
    const applyOps = spyOn(view.workbook(), 'applyOps');
    fireEvent.contextMenu(band.querySelector('[data-col="1"]')!, { button: 2, clientX: 120, clientY: 5 });
    const menu = await waitFor(() => view.view.getByTestId('xlsx-column-header-menu'));
    fireEvent.click(within(menu).getByRole('menuitem', { name: 'Freeze column A' }));
    expect(applyOps).toHaveBeenCalledTimes(1);
    expect(applyOps.mock.calls[0]![0]).toEqual([
      { type: 'setFreezePane', sheet: 0, pane: { rows: 0, cols: 1, top_left: { row: 0, col: 1 } } },
    ]);
  });

  it('commits the open editor and moves the selection when another cell is clicked', async () => {
    const view = await mountEditor();

    view.doubleClick({ row: 2, col: 0 });
    expect(view.editor()?.value).toBe('Line item 1');

    view.type('Edited item');
    view.click({ row: 3, col: 1 });

    expect(view.editor()).toBeNull();
    expect(view.nameBox().value).toBe('B4');
    expect(view.workbook().cell(0, 2, 0).input).toBe('Edited item');
  });

  it('uses synchronous model bounds for Ctrl+End and reveals an empty used-range corner', async () => {
    const view = await mountEditor();
    const workbook = view.workbook();
    workbook.editCell(0, 9_999, 0, 'Far row');
    workbook.editCell(0, 0, 701, 'Far column');
    const before = workbook.save();
    const query = spyOn(workbook, 'selectionData');
    fireEvent.keyDown(view.surface, { key: 'End', ctrlKey: true });
    expect(query).toHaveBeenCalledWith(0, 0, 0, 'usedRange');
    query.mockRestore();
    expect(view.nameBox().value).toBe('ZZ10000');
    expect(view.formula()).toBe('');
    await waitFor(() => {
      expect(view.surface.scrollLeft).toBeGreaterThan(0);
      expect(view.surface.scrollTop).toBeGreaterThan(0);
    });
    fireEvent.keyDown(view.surface, { key: 'Home', ctrlKey: true });
    expect(view.nameBox().value).toBe('A1');
    expect(workbook.save()).toEqual(before);
  });

  it('preserves the selection when a synchronous used-range query fails', async () => {
    const view = await mountEditor();
    fireEvent.keyDown(view.surface, { key: 'ArrowRight' });
    fireEvent.keyDown(view.surface, { key: 'ArrowDown' });
    const before = view.nameBox().value;
    const workbook = view.workbook();
    const selectionData = workbook.selectionData.bind(workbook);
    const query = spyOn(workbook, 'selectionData').mockImplementation((...args) => {
      if (args[3] === 'usedRange') throw new Error('Query unavailable');
      return selectionData(...args);
    });
    fireEvent.keyDown(view.surface, { key: 'End', ctrlKey: true });
    expect(view.nameBox().value).toBe(before);
    query.mockRestore();
  });

  it('commits and reopens on the target when another cell is double-clicked', async () => {
    const view = await mountEditor();

    view.doubleClick({ row: 2, col: 0 });
    view.type('Edited item');
    view.doubleClick({ row: 3, col: 1 });

    expect(view.editor()?.value).toBe('200');
    expect(view.nameBox().value).toBe('B4');
    expect(view.workbook().cell(0, 2, 0).input).toBe('Edited item');
  });

  it('leaves a formula cell unchanged when the pointer moves on', async () => {
    const view = await mountEditor();

    view.doubleClick({ row: 2, col: 3 });
    expect(view.editor()?.value).toBe('=B3+C3');

    view.click({ row: 6, col: 0 });

    expect(view.editor()).toBeNull();
    expect(view.nameBox().value).toBe('A7');
    expect(view.workbook().cell(0, 2, 3).input).toBe('=B3+C3');

    view.doubleClick({ row: 3, col: 3 });
    expect(view.editor()?.value).toBe('=B4+C4');

    view.doubleClick({ row: 7, col: 0 });

    expect(view.editor()?.value).toBe('Line item 6');
    expect(view.nameBox().value).toBe('A8');
    expect(view.workbook().cell(0, 3, 3).input).toBe('=B4+C4');
  });

  it('keeps a press inside the open editor from committing or moving on', async () => {
    const view = await mountEditor();

    view.doubleClick({ row: 2, col: 0 });
    view.type('Edited item');
    view.pressInEditor({ row: 6, col: 0 });

    expect(view.editor()?.value).toBe('Edited item');
    expect(view.nameBox().value).toBe('A3');
    expect(view.workbook().cell(0, 2, 0).input).toBe('Line item 1');
  });

  it('keeps a double-click inside the open editor from reopening it', async () => {
    const view = await mountEditor();

    view.doubleClick({ row: 2, col: 0 });
    view.type('Edited item');
    view.doubleClickInEditor({ row: 2, col: 0 });

    expect(view.editor()?.value).toBe('Edited item');
  });

  it('dismisses the editor without following a hyperlink in the clicked cell', async () => {
    const view = await mountEditor(linked);

    view.doubleClick({ row: 2, col: 0 });
    view.type('Edited item');
    view.click(LINK_CELL);

    expect(view.editor()).toBeNull();
    expect(view.nameBox().value).toBe('E6');
    expect(opened).toEqual([]);

    view.click(LINK_CELL);

    expect(opened).toEqual([LINK_TARGET]);
  });
});

describe('XlsxEditor fill handle', () => {
  function dragHandle(
    view: { view: RenderResult },
    from: CellAddr,
    to: CellAddr
  ) {
    const handle = view.view.getByTestId('xlsx-fill-handle');
    fireEvent.pointerDown(handle, { button: 0, pointerId: 1, ...pointAt(plain, from) });
    fireEvent.pointerMove(handle, { pointerId: 1, ...pointAt(plain, to) });
    fireEvent.pointerUp(handle, { pointerId: 1, ...pointAt(plain, to) });
  }

  it('tiles the source down one column and commits one fillRange call', async () => {
    const view = await mountEditor();
    const workbook = view.workbook();
    view.click({ row: 2, col: 1 });
    expect(workbook.cell(0, 2, 1).input).toBe('100');
    const fill = spyOn(workbook, 'fillRange');

    dragHandle(view, { row: 2, col: 1 }, { row: 4, col: 1 });

    await waitFor(() => expect(fill).toHaveBeenCalledTimes(1));
    expect(fill).toHaveBeenCalledWith(0, 'B3', 'B3:B5');
    expect(workbook.cell(0, 3, 1).input).toBe('100');
    expect(workbook.cell(0, 4, 1).input).toBe('100');
    expect(view.view.queryByTestId('xlsx-fill-preview')).toBeNull();
  });

  it('tiles the source right one row and commits one fillRange call', async () => {
    const view = await mountEditor();
    const workbook = view.workbook();
    view.click({ row: 2, col: 1 });
    const fill = spyOn(workbook, 'fillRange');

    dragHandle(view, { row: 2, col: 1 }, { row: 2, col: 3 });

    await waitFor(() => expect(fill).toHaveBeenCalledTimes(1));
    expect(fill).toHaveBeenCalledWith(0, 'B3', 'B3:D3');
    expect(workbook.cell(0, 2, 2).input).toBe('100');
    expect(workbook.cell(0, 2, 3).input).toBe('100');
  });

  it('cancels the drag on Escape without calling the engine', async () => {
    const view = await mountEditor();
    const workbook = view.workbook();
    view.click({ row: 2, col: 1 });
    const before = workbook.save();
    const fill = spyOn(workbook, 'fillRange');
    const handle = view.view.getByTestId('xlsx-fill-handle');

    fireEvent.pointerDown(handle, { button: 0, pointerId: 1, ...pointAt(plain, { row: 2, col: 1 }) });
    fireEvent.pointerMove(handle, { pointerId: 1, ...pointAt(plain, { row: 4, col: 1 }) });
    expect(view.view.getByTestId('xlsx-fill-preview')).toBeDefined();
    fireEvent.keyDown(window, { key: 'Escape' });
    expect(view.view.queryByTestId('xlsx-fill-preview')).toBeNull();
    fireEvent.pointerUp(handle, { pointerId: 1, ...pointAt(plain, { row: 4, col: 1 }) });

    expect(fill).not.toHaveBeenCalled();
    expect(workbook.save()).toEqual(before);
    expect(view.nameBox().value).toBe('B3');
  });

  it('keeps the workbook and selection when the engine refuses the fill', async () => {
    const view = await mountEditor();
    const workbook = view.workbook();
    view.click({ row: 2, col: 1 });
    const before = workbook.save();
    const message = 'fill cannot safely preserve Excel tables';
    const fill = spyOn(workbook, 'fillRange').mockImplementation(() => {
      throw new Error(message);
    });

    dragHandle(view, { row: 2, col: 1 }, { row: 4, col: 1 });

    await waitFor(() => expect(view.view.getByTestId('xlsx-input-refusal').textContent).toBe(message));
    expect(fill).toHaveBeenCalledTimes(1);
    expect(workbook.save()).toEqual(before);
    expect(view.nameBox().value).toBe('B3');
  });

  it('does not render the handle by default', async () => {
    const editor = await mountEditor(plain, undefined, false, false);
    editor.click({ row: 2, col: 1 });
    expect(editor.view.getByTestId('xlsx-selection')).toBeDefined();
    expect(editor.view.queryByTestId('xlsx-fill-handle')).toBeNull();
  });

  it('renders the handle only for an editable selection', async () => {
    const editable = await mountEditor();
    editable.click({ row: 2, col: 1 });
    expect(editable.view.getByTestId('xlsx-fill-handle')).toBeDefined();
    await act(async () => { editable.view.unmount(); });

    const readOnly = render(
      <XlsxEditor legacySyncApi experimentalWorkerOpen={false} readOnly fillHandle
        file={plain.bytes.slice()} onReady={() => {}} />
    );
    await waitFor(() => expect(readOnly.getByTestId('xlsx-selection')).toBeDefined(), MOUNTED);
    expect(readOnly.queryByTestId('xlsx-fill-handle')).toBeNull();
  });

  it('anchors the handle on the true bottom-right of a partially off-screen selection and fills every row', async () => {
    const view = await mountEditor();
    const canvas = view.view.container.querySelector('canvas') as HTMLCanvasElement;
    canvas.getBoundingClientRect = () => new DOMRect(0, 0, 800, 600);
    const workbook = view.workbook();
    workbook.editCell(0, 99, 1, 'end');
    await act(async () => {
      view.api().selectCells(0, { anchor: { row: 0, col: 0 }, focus: { row: 99, col: 1 } });
    });
    await waitFor(() => expect(view.surface.scrollTop).toBeGreaterThan(0), MOUNTED);
    await act(async () => { fireEvent.scroll(view.surface); });
    // scrolled to B100, only the selection's last row is visible.
    await waitFor(() => expect(parseFloat((view.view.getByTestId('xlsx-selection') as HTMLElement).style.height)).toBeLessThan(100), MOUNTED);
    await waitFor(() => expect(parseFloat(fillHandleElement(view).style.top)).toBeLessThan(200), MOUNTED);
    const endpoint = view.api().getPositionAtPoint(790, 12);
    expect(endpoint).not.toBeNull();
    expect(endpoint!.col).toBeGreaterThan(1);
    const fill = spyOn(workbook, 'fillRange').mockImplementation(() => ({
      applied: true, sheetInfo: workbook.sheetInfo(),
    }));

    await act(async () => {
      fireEvent.pointerDown(fillHandleElement(view), { button: 0, pointerId: 1, clientX: 10, clientY: 12 });
      fireEvent.pointerMove(fillHandleElement(view), { pointerId: 1, clientX: 790, clientY: 12 });
    });
    expect(view.view.getByTestId('xlsx-fill-preview')).toBeDefined();
    await act(async () => {
      fireEvent.pointerUp(fillHandleElement(view), { pointerId: 1, clientX: 790, clientY: 12 });
    });

    await waitFor(() => expect(fill).toHaveBeenCalledTimes(1));
    const [sheet, source, destination] = fill.mock.calls[0];
    expect(sheet).toBe(0);
    expect(source).toMatch(/^A1:[A-Z]{1,3}100$/);
    expect(destination).toMatch(/^A1:[A-Z]{1,3}100$/);
    expect(destination).not.toBe(source);
  });

  it('clamps an off-screen bottom-right to the viewport and still fills the full row count', async () => {
    const view = await mountEditor();
    const canvas = view.view.container.querySelector('canvas') as HTMLCanvasElement;
    canvas.getBoundingClientRect = () => new DOMRect(0, 0, 800, 600);
    const workbook = view.workbook();
    workbook.editCell(0, 99, 1, 'end');
    await act(async () => {
      view.api().selectCells(0, { anchor: { row: 0, col: 0 }, focus: { row: 99, col: 1 } });
    });
    await waitFor(() => expect(view.surface.scrollTop).toBeGreaterThan(0), MOUNTED);
    await act(async () => {
      view.surface.scrollTop = 0;
      view.surface.scrollLeft = 0;
      fireEvent.scroll(view.surface);
    });
    // back at the top the selection's clipped edge spans the viewport, so the
    // handle is clamped there rather than pinned to B100's off-screen corner.
    await waitFor(() => expect(parseFloat((view.view.getByTestId('xlsx-selection') as HTMLElement).style.height)).toBeGreaterThan(500), MOUNTED);
    await waitFor(() => expect(parseFloat(fillHandleElement(view).style.top)).toBeGreaterThan(500), MOUNTED);
    const endpoint = view.api().getPositionAtPoint(790, 36);
    expect(endpoint).not.toBeNull();
    expect(endpoint!.col).toBeGreaterThan(1);
    const fill = spyOn(workbook, 'fillRange').mockImplementation(() => ({
      applied: true, sheetInfo: workbook.sheetInfo(),
    }));

    await act(async () => {
      fireEvent.pointerDown(fillHandleElement(view), { button: 0, pointerId: 1, clientX: 10, clientY: 596 });
      fireEvent.pointerMove(fillHandleElement(view), { pointerId: 1, clientX: 790, clientY: 36 });
    });
    expect(view.view.getByTestId('xlsx-fill-preview')).toBeDefined();
    await act(async () => {
      fireEvent.pointerUp(fillHandleElement(view), { pointerId: 1, clientX: 790, clientY: 36 });
    });

    await waitFor(() => expect(fill).toHaveBeenCalledTimes(1));
    const [sheet, source, destination] = fill.mock.calls[0];
    expect(sheet).toBe(0);
    expect(source).toMatch(/^A1:[A-Z]{1,3}100$/);
    expect(destination).toMatch(/^A1:[A-Z]{1,3}100$/);
    expect(destination).not.toBe(source);
  });

  function selectionSize(view: Awaited<ReturnType<typeof mountEditor>>) {
    return view.api().commands.getState('merge').value;
  }

  it('does not render the handle once Ctrl+A selects the whole sheet', async () => {
    const view = await mountEditor();
    view.click({ row: 2, col: 1 });
    expect(view.view.getByTestId('xlsx-fill-handle')).toBeDefined();

    for (let press = 0; press < 2 && selectionSize(view)?.rows !== 1_048_576; press++) {
      await act(async () => { fireEvent.keyDown(view.surface, { key: 'a', ctrlKey: true }); });
    }

    expect(selectionSize(view)).toEqual({ rows: 1_048_576, columns: 16_384 });
    expect(view.view.getByTestId('xlsx-selection')).toBeDefined();
    expect(view.view.queryByTestId('xlsx-fill-handle')).toBeNull();
  });

  it('does not render the handle once Ctrl+Shift+Down selects a whole empty column', async () => {
    const view = await mountEditor();
    view.click({ row: 0, col: 4 });
    expect(view.nameBox().value).toBe('E1');
    expect(view.view.getByTestId('xlsx-fill-handle')).toBeDefined();

    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowDown', ctrlKey: true, shiftKey: true });
    });
    await act(async () => { fireEvent.scroll(view.surface); });

    expect(selectionSize(view)).toEqual({ rows: 1_048_576, columns: 1 });
    expect(view.view.getByTestId('xlsx-selection')).toBeDefined();
    expect(view.view.queryByTestId('xlsx-fill-handle')).toBeNull();
  });

  it('does not render the handle once Ctrl+Shift+Right selects a whole empty row', async () => {
    const view = await mountEditor();
    await act(async () => { view.api().selectCells(0, selectionAt({ row: 63, col: 0 })); });
    await waitFor(() => expect(view.surface.scrollTop).toBeGreaterThan(0), MOUNTED);
    await act(async () => { fireEvent.scroll(view.surface); });
    expect(view.nameBox().value).toBe('A64');
    await waitFor(() => expect(view.view.getByTestId('xlsx-fill-handle')).toBeDefined(), MOUNTED);

    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowRight', ctrlKey: true, shiftKey: true });
    });
    await act(async () => { fireEvent.scroll(view.surface); });

    expect(selectionSize(view)).toEqual({ rows: 1, columns: 16_384 });
    expect(view.view.getByTestId('xlsx-selection')).toBeDefined();
    expect(view.view.queryByTestId('xlsx-fill-handle')).toBeNull();
  });
});

describe('fill handle with sheet headers', () => {
  it('renders no handle with headers and the fill handle both off', async () => {
    const view = await mountEditor(plain, undefined, false, false);
    view.click({ row: 2, col: 1 });
    await waitFor(() => expect(view.view.getByTestId('xlsx-selection')).toBeDefined());
    expect(view.view.queryByTestId('xlsx-grid')).toBeNull();
    expect(view.view.queryByTestId('xlsx-fill-handle')).toBeNull();
    expect(view.view.queryByTestId('xlsx-focus')).toBeNull();
    expect(view.view.queryByTestId('xlsx-active-cell')).toBeNull();
    expect(headerlessActiveCell(view.surface).style.border).toBe('2px solid #217346');
  });

  it('renders no handle with headers on and the fill handle off', async () => {
    const view = await mountEditor(plain, undefined, true, false);
    view.click({ row: 2, col: 1 });
    await waitFor(() => expect(view.view.getByTestId('xlsx-selection')).toBeDefined());
    expect(view.view.getByTestId('xlsx-grid')).toBeDefined();
    expect(view.view.queryByTestId('xlsx-fill-handle')).toBeNull();
  });

  it('anchors the handle on the range corner and fills B2:C3 down to B2:C6 with headers on', async () => {
    const view = await mountEditor(plain, undefined, true, true);
    const workbook = view.workbook();
    fireEvent.mouseDown(view.surface, pointAt(plain, { row: 1, col: 1 }));
    await act(async () => {
      fireEvent.mouseMove(view.surface, { ...pointAt(plain, { row: 2, col: 2 }), buttons: 1 });
    });
    await act(async () => {
      fireEvent.mouseUp(window, pointAt(plain, { row: 2, col: 2 }));
    });
    await waitFor(() => expect(view.nameBox().value).toBe('C3'));
    const corner = cellRect(plain.grid, 2, 2)!;
    const marquee = view.selectionBox()!;
    expect(marquee.style.left).toBe(`${Math.round(cellRect(plain.grid, 1, 1)!.x)}px`);
    expect(marquee.style.top).toBe(`${Math.round(cellRect(plain.grid, 1, 1)!.y)}px`);
    const handle = fillHandleElement(view);
    expect(parseFloat(handle.style.left)).toBeCloseTo(corner.x + corner.w - 4);
    expect(parseFloat(handle.style.top)).toBeCloseTo(corner.y + corner.h - 4);
    expect(Math.abs(parseFloat(handle.style.left) + 4
      - (parseFloat(marquee.style.left) + parseFloat(marquee.style.width)))).toBeLessThanOrEqual(0.5);
    expect(Math.abs(parseFloat(handle.style.top) + 4
      - (parseFloat(marquee.style.top) + parseFloat(marquee.style.height)))).toBeLessThanOrEqual(0.5);
    const fill = spyOn(workbook, 'fillRange').mockImplementation(() => ({
      applied: true, sheetInfo: workbook.sheetInfo(),
    }));

    fireEvent.pointerDown(handle, { button: 0, pointerId: 1, ...pointAt(plain, { row: 2, col: 2 }) });
    fireEvent.pointerMove(handle, { pointerId: 1, ...pointAt(plain, { row: 5, col: 2 }) });
    fireEvent.pointerUp(handle, { pointerId: 1, ...pointAt(plain, { row: 5, col: 2 }) });

    await waitFor(() => expect(fill).toHaveBeenCalledTimes(1));
    expect(fill).toHaveBeenCalledWith(0, 'B2:C3', 'B2:C6');
  });

  const headerSelectionSize = (view: Awaited<ReturnType<typeof mountEditor>>) =>
    view.api().commands.getState('merge').value;
  const headerGrid = (view: Awaited<ReturnType<typeof mountEditor>>) =>
    view.surface.closest('[data-testid="xlsx-grid"]')!;

  it('does not render the handle once a column header selects the whole column', async () => {
    const view = await mountEditor(plain, undefined, true, true);
    view.click({ row: 2, col: 1 });
    await waitFor(() => expect(view.view.getByTestId('xlsx-fill-handle')).toBeDefined(), MOUNTED);

    fireEvent.mouseDown(headerGrid(view).querySelector('[data-testid="xlsx-column-headers"] [data-col="0"]')!);

    await waitFor(() => expect(headerSelectionSize(view)).toEqual({ rows: 1_048_576, columns: 1 }), MOUNTED);
    expect(view.view.getByTestId('xlsx-selection')).toBeDefined();
    expect(view.view.queryByTestId('xlsx-fill-handle')).toBeNull();
  });

  it('does not render the handle once a row header selects the whole row', async () => {
    const view = await mountEditor(plain, undefined, true, true);
    view.click({ row: 2, col: 1 });
    await waitFor(() => expect(view.view.getByTestId('xlsx-fill-handle')).toBeDefined(), MOUNTED);

    fireEvent.mouseDown(headerGrid(view).querySelector('[data-testid="xlsx-row-headers"] [data-row="0"]')!);

    await waitFor(() => expect(headerSelectionSize(view)).toEqual({ rows: 1, columns: 16_384 }), MOUNTED);
    expect(view.view.getByTestId('xlsx-selection')).toBeDefined();
    expect(view.view.queryByTestId('xlsx-fill-handle')).toBeNull();
  });

  it('does not render the handle once the header corner selects the whole sheet', async () => {
    const view = await mountEditor(plain, undefined, true, true);
    view.click({ row: 2, col: 1 });
    await waitFor(() => expect(view.view.getByTestId('xlsx-fill-handle')).toBeDefined(), MOUNTED);

    fireEvent.click(headerGrid(view).querySelector('[data-testid="xlsx-grid-corner"]')!);

    await waitFor(() => expect(headerSelectionSize(view)).toEqual({ rows: 1_048_576, columns: 16_384 }), MOUNTED);
    expect(view.view.getByTestId('xlsx-selection')).toBeDefined();
    expect(view.view.queryByTestId('xlsx-fill-handle')).toBeNull();
  });
});

function fillHandleElement(view: { view: RenderResult }): HTMLElement {
  return view.view.getByTestId('xlsx-fill-handle');
}
describe('XlsxEditor chart objects', () => {
  it('selects a chart instead of the cells behind it, and deselects off it', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);

    fireEvent.mouseDown(view.surface, chartCenter(chart));
    const outline = await waitFor(() => view.outline()!);
    expect(outline.getAttribute('data-chart-id')).toBe(chart.id);
    // the press must not reach the cells under the chart at all: the grid
    // selection stays exactly where it was, not merely hidden.
    expect(view.nameBox().value).toBe('A1');
    expect(view.selectionBox()).toBeNull();
    fireEvent.mouseUp(window, chartCenter(chart));

    fireEvent.mouseDown(view.surface, { clientX: 2, clientY: 2 });
    await waitFor(() => expect(view.outline()).toBeNull());
    expect(view.selectionBox()).not.toBeNull();
    expect(view.nameBox().value).toBe('A1');
  });

  // the renderer paints a chart it cannot draw as a neutral box rather than
  // failing the frame. it is still an object on the sheet, so it must select
  // and move like any other.
  it('selects and moves a chart the renderer could not draw', async () => {
    const chart = undrawable.charts.find((candidate) => candidate.placeholder);
    expect(chart).toBeDefined();
    expect(chart!.movable).toBe(true);
    const view = await mountEditor(undrawable);
    await selectChart(view, chart!);

    expect(view.outline()!.getAttribute('data-chart-id')).toBe(chart!.id);
    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowRight', shiftKey: true });
    });
    await waitFor(() => expect(view.canUndo()).toBe(true), { timeout: 2000 });
    expect(view.error()).toBeNull();
    await waitFor(() => expect(view.outlineAt().x).toBe(Math.round(chart!.rect.x + 10)));
  });

  it('keeps cell-editing keys off the cells hidden behind a selected chart', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    const before = view.formula();

    fireEvent.mouseDown(view.surface, chartCenter(chart));
    await waitFor(() => view.outline()!);
    fireEvent.mouseUp(window, chartCenter(chart));

    for (const key of ['Delete', 'Backspace', 'x', 'Enter']) {
      await act(async () => {
        fireEvent.keyDown(view.surface, { key });
      });
    }

    expect(view.editor()).toBeNull();
    expect(view.outline()).not.toBeNull();
    expect(view.formula()).toBe(before);
    expect(view.canUndo()).toBe(false);
  });

  it('commits an open edit before taking the press, then selects the chart', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);

    view.doubleClick({ row: 1, col: 1 });
    view.type('Edited item');
    fireEvent.mouseDown(view.surface, chartCenter(chart));

    const outline = await waitFor(() => view.outline()!);
    expect(outline.getAttribute('data-chart-id')).toBe(chart.id);
    expect(view.editor()).toBeNull();
    expect(view.workbook().cell(0, 1, 1).input).toBe('Edited item');
    fireEvent.mouseUp(window, chartCenter(chart));
  });

  it('commits an edit whose input has scrolled out of the window', async () => {
    const SCROLLED_BY = 300;
    const probe = openWorkbook(charted.bytes.slice());
    let target: ChartRegion;
    try {
      target = (probe.displayList({ x: 0, y: SCROLLED_BY, ...VIEWPORT }).charts ?? [])[0];
    } finally {
      probe.dispose();
    }
    const view = await mountEditor(charted);

    view.doubleClick({ row: 1, col: 1 });
    view.type('Edited item');

    await act(async () => {
      view.surface.scrollTop = SCROLLED_BY;
      fireEvent.scroll(view.surface);
    });
    // the input unmounts once its cell leaves the painted window, so there is
    // no blur left to commit through: only the explicit commit saves the edit.
    expect(view.editor()).toBeNull();

    fireEvent.mouseDown(view.surface, chartCenter(target));
    await waitFor(() => view.outline()!);
    fireEvent.mouseUp(window, chartCenter(target));

    expect(view.workbook().cell(0, 1, 1).input).toBe('Edited item');
  });

  it('leaves a press inside the open editor to the editor, not a chart behind it', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);

    view.doubleClick({ row: 1, col: 1 });
    view.type('Edited item');
    // the press lands on the input, at coordinates the chart also covers: the
    // editor is a dom overlay, so without its guard the hit test would reach
    // the chart painted underneath.
    fireEvent.mouseDown(view.editor()!, chartCenter(chart));

    expect(view.editor()?.value).toBe('Edited item');
    expect(view.outline()).toBeNull();
  });

  it('drags a selected chart and repins it through the engine', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    const start = chartCenter(chart);

    fireEvent.mouseDown(view.surface, start);
    await waitFor(() => view.outline()!);
    fireEvent.mouseMove(view.surface, {
      clientX: start.clientX + 40,
      clientY: start.clientY + 24,
      buttons: 1,
    });
    expect(view.outline()!.style.left).toBe(`${chart.rect.x + 40}px`);

    await act(async () => {
      fireEvent.mouseUp(window, { clientX: start.clientX + 40, clientY: start.clientY + 24 });
    });

    await waitFor(() => {
      const at = view.outlineAt();
      expect(at.x).toBe(Math.round(chart.rect.x + 40));
      expect(at.y).toBe(Math.round(chart.rect.y + 24));
    });
  });

  it('lands a run of arrow nudges as one undoable edit', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    await selectChart(view, chart);

    // key repeat: five presses preview locally and touch nothing.
    act(() => {
      for (let press = 0; press < 5; press++) {
        fireEvent.keyDown(view.surface, { key: 'ArrowRight' });
      }
    });
    expect(view.outlineAt().x).toBe(Math.round(chart.rect.x + 5));
    expect(view.canUndo()).toBe(false);

    await waitFor(() => expect(view.canUndo()).toBe(true), { timeout: 2000 });
    expect(view.outlineAt().x).toBe(Math.round(chart.rect.x + 5));

    // one burst is one undo step, not five.
    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'z', ctrlKey: true });
    });
    await waitFor(() => expect(view.outlineAt().x).toBe(Math.round(chart.rect.x)));
    expect(view.canUndo()).toBe(false);
  });

  it('drops a pending burst when the file is swapped under it', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    await selectChart(view, chart);

    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowRight', shiftKey: true });
    });
    expect(view.canUndo()).toBe(false);

    // the editor stays mounted across a file swap, so a burst still in flight
    // would fire against the workbook that replaced the one it was typed on.
    await view.reopenWith(charted);
    await waitFor(() => expect(view.nameBox().value).toBe('A1'));
    await settle();

    expect(view.error()).toBeNull();
    expect(view.canUndo()).toBe(false);
    const still = (view.workbook().displayList({ x: 0, y: 0, ...VIEWPORT }).charts ?? []).find(
      (candidate) => candidate.id === chart.id
    );
    expect(Math.round(still!.rect.x)).toBe(chartRounded(chart).x);
  });

  it('discards a burst that returns to where it started', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    await selectChart(view, chart);

    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowRight', shiftKey: true });
      fireEvent.keyDown(view.surface, { key: 'ArrowLeft', shiftKey: true });
    });
    expect(view.outlineAt().x).toBe(Math.round(chart.rect.x));

    await settle();
    expect(view.canUndo()).toBe(false);
    expect(view.outlineAt().x).toBe(Math.round(chart.rect.x));
  });

  it('lands a pending burst before a save reads the workbook', async () => {
    const [chart] = charted.charts;
    const saved: Uint8Array[] = [];
    const view = await mountEditor(charted, (bytes) => saved.push(bytes));
    await selectChart(view, chart);

    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowRight', shiftKey: true });
    });
    expect(view.canUndo()).toBe(false);

    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 's', ctrlKey: true });
    });

    expect(view.canUndo()).toBe(true);
    expect(saved).toHaveLength(1);
    const reopened = openWorkbook(saved[0]);
    try {
      const moved = (reopened.displayList({ x: 0, y: 0, ...VIEWPORT }).charts ?? []).find(
        (candidate) => candidate.id === chart.id
      );
      expect(Math.round(moved!.rect.x)).toBe(Math.round(chart.rect.x + 10));
    } finally {
      reopened.dispose();
    }
  });

  it('drops a selection and its pending burst on escape', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    await selectChart(view, chart);

    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowRight', shiftKey: true });
    });
    fireEvent.keyDown(view.surface, { key: 'Escape' });

    await waitFor(() => expect(view.outline()).toBeNull());
    await settle();
    expect(view.canUndo()).toBe(false);
  });

  it('cancels an armed drag on escape instead of landing it on release', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    const start = chartCenter(chart);

    fireEvent.mouseDown(view.surface, start);
    await waitFor(() => view.outline()!);
    fireEvent.mouseMove(view.surface, {
      clientX: start.clientX + 40,
      clientY: start.clientY + 24,
      buttons: 1,
    });
    fireEvent.keyDown(view.surface, { key: 'Escape' });
    await waitFor(() => expect(view.outline()).toBeNull());

    await act(async () => {
      fireEvent.mouseUp(window, { clientX: start.clientX + 40, clientY: start.clientY + 24 });
    });

    expect(view.canUndo()).toBe(false);
  });

  it('never arms a drag from a non-primary press', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    const start = chartCenter(chart);

    fireEvent.mouseDown(view.surface, { ...start, button: 2 });
    const outline = await waitFor(() => view.outline()!);
    expect(outline.getAttribute('data-chart-id')).toBe(chart.id);
    // the context menu swallows the matching release, so the next unrelated
    // primary release anywhere must not land a move.
    await act(async () => {
      fireEvent.mouseUp(window, { clientX: start.clientX + 200, clientY: start.clientY + 150 });
    });

    expect(view.canUndo()).toBe(false);
    expect(view.outlineAt()).toEqual(chartRounded(chart));
  });

  it('disarms a drag whose release was lost when the next press arrives', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    const start = chartCenter(chart);

    fireEvent.mouseDown(view.surface, start);
    await waitFor(() => view.outline()!);
    // no mouseup: the release went somewhere this window never saw.
    fireEvent.mouseDown(view.surface, { clientX: 2, clientY: 2 });
    await act(async () => {
      fireEvent.mouseUp(window, { clientX: start.clientX + 300, clientY: start.clientY + 300 });
    });

    expect(view.canUndo()).toBe(false);
  });

  it('hit-tests the frame it painted, not a scroll offset it has not drawn yet', async () => {
    const SCROLLED_BY = 400;
    const target = charted.charts[0];
    const point = chartCenter(target);
    const probe = openWorkbook(charted.bytes.slice());
    try {
      // the fixture must genuinely disagree at this point across the two
      // viewports, or this proves nothing.
      const stale = probe.chartAtPoint(
        { x: 0, y: SCROLLED_BY, ...VIEWPORT },
        point.clientX,
        point.clientY
      );
      expect(stale?.id).not.toBe(target.id);
    } finally {
      probe.dispose();
    }

    const view = await mountEditor(charted);
    // freeze the repaint: scrolling now advances the scroll offset while the
    // canvas still shows the frame painted at the old one.
    const scheduled = window.requestAnimationFrame;
    window.requestAnimationFrame = (() => 0) as typeof window.requestAnimationFrame;
    try {
      view.surface.scrollTop = SCROLLED_BY;
      fireEvent.scroll(view.surface);
      fireEvent.mouseDown(view.surface, point);
      const outline = await waitFor(() => view.outline()!);
      expect(outline.getAttribute('data-chart-id')).toBe(target.id);
      fireEvent.mouseUp(window, point);
    } finally {
      window.requestAnimationFrame = scheduled;
    }
  });

  it('selects a chart pinned to the sheet but never drags it', async () => {
    const pinned = charted.charts.find((chart) => !chart.movable);
    expect(pinned).toBeDefined();
    const view = await mountEditor(charted);
    const start = chartCenter(pinned!);

    fireEvent.mouseDown(view.surface, start);
    const outline = await waitFor(() => view.outline()!);
    expect(outline.getAttribute('data-chart-id')).toBe(pinned!.id);

    await act(async () => {
      fireEvent.mouseMove(view.surface, {
        clientX: start.clientX + 40,
        clientY: start.clientY + 24,
        buttons: 1,
      });
      fireEvent.mouseUp(window, { clientX: start.clientX + 40, clientY: start.clientY + 24 });
    });

    expect(view.outlineAt().x).toBe(chartRounded(pinned!).x);

    await act(async () => {
      fireEvent.keyDown(view.surface, { key: 'ArrowRight', shiftKey: true });
    });
    await settle();
    expect(view.nameBox().value).toBe('A1');
    expect(view.outlineAt().x).toBe(chartRounded(pinned!).x);
    // the engine refuses to repin an absolute anchor, so a nudge that reached
    // it would surface as an error overlay rather than doing nothing.
    expect(view.error()).toBeNull();
    expect(view.canUndo()).toBe(false);

    fireEvent.mouseMove(view.surface, start);
    expect(view.surface.style.cursor).toBe('pointer');
  });

  it('abandons a drag whose release the window never saw', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);
    const start = chartCenter(chart);

    fireEvent.mouseDown(view.surface, start);
    await waitFor(() => view.outline()!);
    // the button comes back up off-window: the next move reports none held.
    fireEvent.mouseMove(view.surface, {
      clientX: start.clientX + 60,
      clientY: start.clientY + 60,
      buttons: 0,
    });

    await act(async () => {
      fireEvent.mouseUp(window, { clientX: start.clientX + 500, clientY: start.clientY + 500 });
    });

    expect(view.outlineAt().x).toBe(chartRounded(chart).x);
    expect(view.canUndo()).toBe(false);
  });

  it('drops a selection that scrolls out of the painted frame', async () => {
    const [chart] = charted.charts;
    const view = await mountEditor(charted);

    fireEvent.mouseDown(view.surface, chartCenter(chart));
    await waitFor(() => view.outline()!);
    fireEvent.mouseUp(window, chartCenter(chart));

    await act(async () => {
      view.surface.scrollTop = 4000;
      fireEvent.scroll(view.surface);
    });

    // no invisible selection left swallowing the keyboard.
    await waitFor(() => expect(view.outline()).toBeNull());
    await waitFor(() => expect(view.selectionBox()).not.toBeNull());
  });
});

describe('XlsxEditor host integration', () => {
  it('keeps viewing mode navigable without exposing user mutations', async () => {
    let api: XlsxEditorApi | undefined;
    let changes = 0;
    const view = render(
      <XlsxEditor legacySyncApi experimentalWorkerOpen={false}
        file={plain.bytes.slice()}
        onChange={() => changes++}
        onReady={(ready) => {
          api = ready;
        }}
        readOnly
      />
    );
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    const surface = view.getByTestId('xlsx-scroll');
    const target = { row: 2, col: 0 };
    const before = api!.handle.cell(0, target.row, target.col).input;

    fireEvent.doubleClick(surface, pointAt(plain, target));
    fireEvent.keyDown(surface, { key: 'x' });
    fireEvent.keyDown(surface, { key: 'Delete' });

    expect(view.queryByTestId('xlsx-toolbar')).toBeNull();
    expect(view.queryByTestId('xlsx-cell-editor')).toBeNull();
    expect(api!.handle.cell(0, target.row, target.col).input).toBe(before);
    expect(changes).toBe(0);

    await act(async () => {
      expect(api!.selectCells(0, selectionAt({ row: 3, col: 1 }))).toBe(true);
    });
    await waitFor(() => {
      const selected = view.getByRole('gridcell', { selected: true });
      expect(selected.textContent).toContain('B4');
    });
    await act(async () => api!.clearSelection());
    await waitFor(() =>
      expect(view.queryAllByRole('gridcell', { selected: true })).toHaveLength(0)
    );
    expect(api!.selectCells(99, selectionAt({ row: 0, col: 0 }))).toBe(false);
  });

  it('notifies on applied edits and saves through the host API', async () => {
    let api: XlsxEditorApi | undefined;
    let changes = 0;
    const view = render(
      <XlsxEditor legacySyncApi experimentalWorkerOpen={false}
        file={plain.bytes.slice()}
        onChange={() => changes++}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    const surface = view.getByTestId('xlsx-scroll');
    const target = { row: 2, col: 0 };

    await act(async () => {
      expect(api!.selectCells(0, selectionAt(target))).toBe(true);
    });
    fireEvent.doubleClick(surface, pointAt(plain, target));
    const editor = await waitFor(() => view.getByTestId('xlsx-cell-editor'));
    fireEvent.change(editor, { target: { value: 'Host edit' } });
    fireEvent.keyDown(editor, { key: 'Enter' });
    expect(changes).toBe(1);

    await act(async () => {
      api!.selectCells(0, selectionAt({ row: 3, col: 1 }));
      api!.clearSelection();
    });
    let saved!: Uint8Array;
    await act(async () => {
      saved = api!.save();
    });
    expect(changes).toBe(1);

    const reopened = openWorkbook(saved);
    try {
      expect(reopened.cell(0, target.row, target.col).input).toBe('Host edit');
    } finally {
      reopened.dispose();
    }
  });

  type GatedItem = { types: string[]; getType(type: string): Promise<Blob> };
  const gatedClipboard = () => {
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const state = { items: [] as GatedItem[] };
    let first = true;
    const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    const originalItem = Object.getOwnPropertyDescriptor(globalThis, 'ClipboardItem');
    Object.defineProperty(globalThis, 'ClipboardItem', {
      configurable: true,
      value: class {
        constructor(private readonly data: Record<string, Blob>) {}
        get types() { return Object.keys(this.data); }
        async getType(type: string) { return this.data[type]!; }
      },
    });
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        readText: async () => {
          const items = state.items;
          if (first) { first = false; await gate; }
          return items[0] ? (await items[0].getType('text/plain')).text() : '';
        },
        read: async () => state.items,
        write: async (next: GatedItem[]) => { state.items = next; },
        writeText: async (text: string) => {
          state.items = [{ types: ['text/plain'], getType: async () => new Blob([text], { type: 'text/plain' }) }];
        },
      },
    });
    return {
      state,
      release: () => release(),
      restore: () => {
        if (originalClipboard) Object.defineProperty(navigator, 'clipboard', originalClipboard);
        else Reflect.deleteProperty(navigator, 'clipboard');
        if (originalItem) Object.defineProperty(globalThis, 'ClipboardItem', originalItem);
        else Reflect.deleteProperty(globalThis, 'ClipboardItem');
      },
    };
  };

  for (const escape of [true, false]) it(`${escape ? 'escape cancels' : 'without escape, runs'} a cut queued behind a pending paste`, async () => {
    const file = plain.bytes.slice();
    let api: XlsxEditorApi | undefined;
    const clipboard = gatedClipboard();
    const { state } = clipboard;

    try {
      const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={file}
        onReady={(ready) => { api = ready; }} />);
      await waitFor(() => expect(api).toBeDefined(), MOUNTED);
      const surface = view.getByTestId('xlsx-scroll');
      const source = { row: 40, col: 10 };
      const destination = { row: 40, col: 12 };
      const value = 'Moved value';
      await act(async () => { api!.handle.editCells(0, [{ ...source, input: value }]); });
      expect(api!.handle.cell(0, source.row, source.col).input).toBe(value);
      expect(api!.handle.cell(0, destination.row, destination.col).input).toBe('');

      await act(async () => { api!.selectCells(0, selectionAt(destination)); });
      fireEvent.keyDown(surface, { key: 'v', ctrlKey: true });
      await act(async () => { api!.selectCells(0, selectionAt(source)); });
      fireEvent.keyDown(surface, { key: 'x', ctrlKey: true });
      if (escape) fireEvent.keyDown(surface, { key: 'Escape' });
      await act(async () => clipboard.release());
      if (escape) {
        await act(async () => {});
        expect(state.items).toHaveLength(1);
        expect(api!.handle.cell(0, source.row, source.col).input).toBe(value);
        return;
      }
      await waitFor(() => expect(api!.handle.cell(0, source.row, source.col).input).toBe(''));
      expect(state.items).toHaveLength(1);
      await act(async () => { api!.selectCells(0, selectionAt(destination)); });
      fireEvent.keyDown(surface, { key: 'v', ctrlKey: true });
      await waitFor(() => expect(api!.handle.cell(0, destination.row, destination.col).input).toBe(value));
    } finally {
      clipboard.restore();
    }
  });

  it('keeps a cut made behind a pending paste on the clipboard after an earlier copy', async () => {
    let api: XlsxEditorApi | undefined;
    const clipboard = gatedClipboard();
    try {
      const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()}
        onReady={(ready) => { api = ready; }} />);
      await waitFor(() => expect(api).toBeDefined(), MOUNTED);
      const surface = view.getByTestId('xlsx-scroll');
      const copied = { row: 40, col: 10 };
      const cut = { row: 40, col: 11 };
      const destination = { row: 40, col: 12 };
      await act(async () => {
        api!.handle.editCells(0, [{ ...copied, input: 'Copied value' }, { ...cut, input: 'Cut value' }]);
      });

      await act(async () => { api!.selectCells(0, selectionAt(destination)); });
      fireEvent.keyDown(surface, { key: 'v', ctrlKey: true });
      await act(async () => { api!.selectCells(0, selectionAt(copied)); });
      fireEvent.keyDown(surface, { key: 'c', ctrlKey: true });
      await act(async () => { api!.selectCells(0, selectionAt(cut)); });
      fireEvent.keyDown(surface, { key: 'x', ctrlKey: true });
      await act(async () => clipboard.release());

      await waitFor(() => expect(api!.handle.cell(0, cut.row, cut.col).input).toBe(''));
      expect(await (await clipboard.state.items[0]!.getType('text/plain')).text()).toBe('Cut value');
      await act(async () => { api!.selectCells(0, selectionAt(destination)); });
      fireEvent.keyDown(surface, { key: 'v', ctrlKey: true });
      await waitFor(() => expect(api!.handle.cell(0, destination.row, destination.col).input).toBe('Cut value'));
      expect(api!.handle.cell(0, copied.row, copied.col).input).toBe('Copied value');
    } finally {
      clipboard.restore();
    }
  });

  it('rebases a pending paste from the copy it read after a newer copy', async () => {
    let api: XlsxEditorApi | undefined;
    const clipboard = gatedClipboard();
    try {
      const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()}
        onReady={(ready) => { api = ready; }} />);
      await waitFor(() => expect(api).toBeDefined(), MOUNTED);
      const surface = view.getByTestId('xlsx-scroll');
      await act(async () => {
        api!.handle.editCells(0, [{ row: 40, col: 10, input: '=K40' }, { row: 40, col: 14, input: 'Newer' }]);
      });

      await act(async () => { api!.selectCells(0, selectionAt({ row: 40, col: 10 })); });
      await act(async () => { fireEvent.keyDown(surface, { key: 'c', ctrlKey: true }); });
      await act(async () => { api!.selectCells(0, selectionAt({ row: 41, col: 10 })); });
      fireEvent.keyDown(surface, { key: 'v', ctrlKey: true });
      await act(async () => { api!.selectCells(0, selectionAt({ row: 40, col: 14 })); });
      await act(async () => { fireEvent.keyDown(surface, { key: 'c', ctrlKey: true }); });
      await act(async () => clipboard.release());

      await waitFor(() => expect(api!.handle.cell(0, 41, 10).input).toBe('=K41'));
    } finally {
      clipboard.restore();
    }
  });

  it('does not finish an asynchronous paste after entering viewing mode', async () => {
    const file = plain.bytes.slice();
    let api: XlsxEditorApi | undefined;
    let resolveClipboard!: (text: string) => void;
    const clipboardText = new Promise<string>((resolve) => {
      resolveClipboard = resolve;
    });
    const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { readText: () => clipboardText },
    });

    try {
      const onReady = (ready: XlsxEditorApi) => {
        api = ready;
      };
      const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={file} onReady={onReady} />);
      await waitFor(() => expect(api).toBeDefined(), MOUNTED);
      const target = { row: 2, col: 0 };
      const before = api!.handle.cell(0, target.row, target.col).input;
      await act(async () => {
        api!.selectCells(0, selectionAt(target));
      });

      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      view.rerender(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={file} onReady={onReady} readOnly />);
      await act(async () => resolveClipboard('late paste'));

      expect(api!.handle.cell(0, target.row, target.col).input).toBe(before);
    } finally {
      if (originalClipboard) {
        Object.defineProperty(navigator, 'clipboard', originalClipboard);
      } else {
        Reflect.deleteProperty(navigator, 'clipboard');
      }
    }
  });

  it('applies an accepted paste and sends no move for a header drag made while it reads', async () => {
    // the fixture merges A1:D1, which refuses a column move across it.
    const prepared = openWorkbook(plain.bytes.slice());
    let file: Uint8Array;
    try {
      prepared.applyOps([{
        type: 'unmergeCells', sheet: 0,
        range: { start: { row: 0, col: 0 }, end: { row: 0, col: 3 } },
      }]);
      prepared.editCells(0, [
        { row: 0, col: 0, input: '10' },
        { row: 0, col: 1, input: '20' },
        { row: 0, col: 2, input: '' },
        { row: 0, col: 3, input: '40' },
        { row: 0, col: 4, input: '50' },
      ]);
      file = prepared.save();
    } finally {
      prepared.dispose();
    }

    let resolveClipboard!: (text: string) => void;
    const clipboardText = new Promise<string>((resolve) => {
      resolveClipboard = resolve;
    });
    const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { readText: () => clipboardText },
    });
    try {
      const view = await mountWithHeaders({ bytes: file, grid: plain.grid, charts: plain.charts });
      const handle = view.workbook();
      const row = () => [0, 1, 2, 3, 4].map((col) => handle.cell(0, 0, col).input);
      const applyOps = spyOn(handle, 'applyOps');
      expect(row()).toEqual(['10', '20', '', '40', '50']);
      view.click({ row: 0, col: 1 });

      // the paste is accepted and its clipboard read is held open...
      fireEvent.keyDown(view.surface, { key: 'v', ctrlKey: true });

      // ...then B is selected by its header and dragged past D.
      const band = view.surface.closest('[data-testid="xlsx-grid"]')!
        .querySelector('[data-testid="xlsx-column-headers"]')!;
      const start = (plain.grid.colOffsets[1]! + plain.grid.colOffsets[2]!) / 2;
      const drop = plain.grid.colOffsets[4]! - 1;
      fireEvent.mouseDown(band.querySelector('[data-col="1"]')!);
      fireEvent.pointerDown(band.querySelector('[data-col="1"]')!, {
        button: 0, clientX: start, clientY: 5, pointerId: 1 });
      fireEvent.pointerMove(band, { clientX: drop, clientY: 5, pointerId: 1 });
      fireEvent.pointerUp(band, { clientX: drop, clientY: 5, pointerId: 1 });

      expect(applyOps).not.toHaveBeenCalled();
      expect(row()).toEqual(['10', '20', '', '40', '50']);

      await act(async () => resolveClipboard('99'));

      await waitFor(() => expect(row()).toEqual(['10', '99', '', '40', '50']));
      expect(applyOps).not.toHaveBeenCalled();
    } finally {
      if (originalClipboard) Object.defineProperty(navigator, 'clipboard', originalClipboard);
      else Reflect.deleteProperty(navigator, 'clipboard');
    }
  });
});

describe('XlsxEditor proposal review', () => {
  it('reviews a staged proposal through accept, undo, reject, stale warning, and force apply', async () => {
    let api: XlsxEditorApi | undefined;
    const view = render(
      <XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()} onReady={(ready) => { api = ready; }} />
    );
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    const workbook = api!.handle;
    const before = workbook.cell(0, 6, 4).input;
    const stage = async (input: string) => {
      await act(async () => {
        workbook.propose('Audit agent', 'Review this change', [
          { sheet: 0, row: 6, col: 4, input },
        ]);
        api!.refreshProposals();
      });
    };

    await stage('12');
    expect(workbook.cell(0, 6, 4).input).toBe(before);
    fireEvent.click(view.getByTestId('xlsx-proposals-button'));
    expect(view.getByTestId('xlsx-proposal').textContent).toContain('Audit agent');
    fireEvent.click(view.getByTestId('xlsx-proposal-accept'));
    await waitFor(() => expect(workbook.cell(0, 6, 4).input).toBe('12'));
    expect(workbook.listProposals()).toHaveLength(0);
    fireEvent.click(view.getByTestId('xlsx-undo'));
    await waitFor(() => expect(workbook.cell(0, 6, 4).input).toBe(before));

    await stage('24');
    fireEvent.click(view.getByTestId('xlsx-proposal-reject'));
    await waitFor(() => expect(workbook.listProposals()).toHaveLength(0));
    expect(workbook.cell(0, 6, 4).input).toBe(before);

    await stage('42');
    await act(async () => {
      workbook.editCell(0, 6, 4, '99');
      api!.refreshProposals();
    });
    fireEvent.click(view.getByTestId('xlsx-proposal-accept'));
    await waitFor(() =>
      expect(view.getByTestId('xlsx-proposal-stale').textContent).toContain('E7')
    );
    expect(workbook.cell(0, 6, 4).input).toBe('99');
    fireEvent.click(view.getByTestId('xlsx-proposal-force'));
    await waitFor(() => expect(workbook.cell(0, 6, 4).input).toBe('42'));
    expect(workbook.listProposals()).toHaveLength(0);

    await act(async () => {
      workbook.editCell(0, 1, 6, '10');
    });
    await stage('=G2*2');
    expect(view.getByTestId('xlsx-proposal-cell-new').textContent).toBe('20');
    await act(async () => {
      workbook.editCell(0, 1, 6, '99');
      api!.refreshProposals();
    });
    fireEvent.click(view.getByTestId('xlsx-proposal-accept'));
    await waitFor(() =>
      expect(view.getByTestId('xlsx-proposal-cell-new').textContent).toBe('198')
    );
    expect(workbook.cell(0, 6, 4).input).toBe('42');
    fireEvent.click(view.getByTestId('xlsx-proposal-accept'));
    await waitFor(() => expect(workbook.cell(0, 6, 4).input).toBe('=G2*2'));
    expect(workbook.listProposals()).toHaveLength(0);
  });
});

describe('XlsxEditor pending host edits', () => {
  it('settles a pending chart move before selecting another sheet', async () => {
    const source = openWorkbook(charted.bytes);
    source.applyOps([{ type: 'addSheet', index: 1, name: 'Extra' }]);
    const file = source.save();
    source.dispose();
    let api: XlsxEditorApi | undefined;
    const view = render(
      <XlsxEditor legacySyncApi experimentalWorkerOpen={false}
        file={file}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    const surface = view.getByTestId('xlsx-scroll');
    const chart = api!.handle.displayList({ x: 0, y: 0, ...VIEWPORT }).charts![0];
    await waitFor(() => {
      fireEvent.mouseDown(surface, chartCenter(chart));
      fireEvent.mouseUp(window, chartCenter(chart));
      expect(view.queryByTestId('xlsx-chart-selection')).not.toBeNull();
    }, MOUNTED);
    await act(async () => {
      fireEvent.keyDown(surface, { key: 'ArrowRight' });
    });
    expect(
      Math.round(parseFloat(view.getByTestId('xlsx-chart-selection').style.left))
    ).toBe(Math.round(chart.rect.x + 1));
    await act(async () => {
      expect(api!.selectCells(1, selectionAt({ row: 0, col: 0 }))).toBe(true);
    });
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 350));
    });
    await act(async () => {
      api!.selectCells(0, selectionAt({ row: 0, col: 0 }));
    });
    const after = api!.handle
      .displayList({ x: 0, y: 0, ...VIEWPORT })
      .charts!.find((c) => c.id === chart.id)!;
    expect(after.rect.x).toBe(chart.rect.x + 1);
  });

  it('commits the current cell draft before selecting another cell', async () => {
    let api: XlsxEditorApi | undefined;
    const view = render(
      <XlsxEditor legacySyncApi experimentalWorkerOpen={false}
        file={plain.bytes.slice()}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    const target = { row: 2, col: 0 };
    await act(async () => {
      api!.selectCells(0, selectionAt(target));
    });
    fireEvent.doubleClick(view.getByTestId('xlsx-scroll'), pointAt(plain, target));
    const editor = await waitFor(() => view.getByTestId('xlsx-cell-editor'));
    fireEvent.change(editor, { target: { value: 'Draft that must survive' } });
    await act(async () => {
      api!.selectCells(0, selectionAt({ row: 3, col: 1 }));
    });
    expect(api!.handle.cell(0, target.row, target.col).input).toBe(
      'Draft that must survive'
    );
  });
  for (const source of ['cell', 'formula'] as const) {
    for (const action of ['save', 'clear', 'select'] as const) {
      it(`commits a ${source} draft before host ${action}`, async () => {
        let api: XlsxEditorApi | undefined;
        let changes = 0;
        const view = render(
          <XlsxEditor legacySyncApi experimentalWorkerOpen={false}
            file={plain.bytes.slice()}
            onChange={() => {
              changes += 1;
            }}
            onReady={(ready) => {
              api = ready;
            }}
          />
        );
        await waitFor(() => expect(api).toBeDefined(), MOUNTED);
        const target = { row: 2, col: 0 };
        await act(async () => {
          api!.selectCells(0, selectionAt(target));
        });
        if (source === 'cell') {
          fireEvent.doubleClick(view.getByTestId('xlsx-scroll'), pointAt(plain, target));
        }
        const input = view.getByTestId(
          source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input'
        );
        fireEvent.change(input, { target: { value: 'Saved draft' } });
        expect(api!.selectCells(-1, selectionAt(target))).toBe(false);
        expect(api!.handle.cell(0, target.row, target.col).input).toBe('Line item 1');
        let saved: Uint8Array | undefined;
        await act(async () => {
          if (action === 'save') saved = api!.save();
          else if (action === 'clear') api!.clearSelection();
          else api!.selectCells(0, selectionAt({ row: 3, col: 1 }));
        });
        expect(api!.handle.cell(0, target.row, target.col).input).toBe('Saved draft');
        expect(changes).toBe(1);
        if (saved) {
          const reopened = openWorkbook(saved);
          try {
            expect(reopened.cell(0, target.row, target.col).input).toBe('Saved draft');
          } finally {
            reopened.dispose();
          }
        }
      });
    }
  }
  it('commits the next cell edit on blur after a host save', async () => {
    let api: XlsxEditorApi | undefined;
    const view = render(
      <XlsxEditor legacySyncApi experimentalWorkerOpen={false}
        file={plain.bytes.slice()}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    const target = { row: 2, col: 0 };
    await act(async () => {
      api!.selectCells(0, selectionAt(target));
    });
    fireEvent.change(view.getByTestId('xlsx-formula-input'), {
      target: { value: 'First draft' },
    });
    await act(async () => {
      api!.save();
    });
    fireEvent.doubleClick(view.getByTestId('xlsx-scroll'), pointAt(plain, target));
    const editor = view.getByTestId('xlsx-cell-editor');
    fireEvent.change(editor, { target: { value: 'Second draft' } });
    fireEvent.blur(editor);
    expect(api!.handle.cell(0, target.row, target.col).input).toBe('Second draft');
  });
});

describe('XlsxEditor commands', () => {
  const MOD = isMacPlatform() ? { metaKey: true } : { ctrlKey: true };
  const target = { row: 2, col: 0 };

  async function mountCommands(props: Partial<XlsxInThreadEditorProps> = {}) {
    let api: XlsxEditorApi | undefined;
    let changes = 0;
    const saves: Uint8Array[] = [];
    const onReady = (ready: XlsxEditorApi) => {
      api = ready;
    };
    const element = (next: Partial<XlsxInThreadEditorProps> = {}) => (
      <XlsxEditor legacySyncApi experimentalWorkerOpen={false}
        file={plain.bytes}
        onChange={() => {
          changes += 1;
        }}
        onSave={(bytes) => {
          saves.push(bytes);
        }}
        onReady={onReady}
        {...props}
        {...next}
      />
    );
    const view = render(element());
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    await act(async () => {
      api!.selectCells(0, selectionAt(target));
    });
    const execute = async <K extends XlsxCommandId>(id: K, args: XlsxCommandArgs[K]) => {
      let result!: XlsxCommandResult;
      await act(async () => {
        result = await api!.commands.execute(id, args);
      });
      return result;
    };
    const typeInCell = (value: string) => {
      fireEvent.doubleClick(view.getByTestId('xlsx-scroll'), pointAt(plain, target));
      const editor = view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
      fireEvent.change(editor, { target: { value } });
      return editor;
    };
    return {
      view,
      api: () => api!,
      changes: () => changes,
      saves,
      execute,
      typeInCell,
      rerender: (next: Partial<XlsxInThreadEditorProps>) => view.rerender(element(next)),
      input: () => api!.handle.cell(0, target.row, target.col).input,
    };
  }

  const failure = (result: XlsxCommandResult) => (result.ok ? null : result.failure.code);

  it('exposes one store for the default toolbar and host calls', async () => {
    const editor = await mountCommands();
    const commands = editor.api().commands;
    expect(commands.getState('bold')).toEqual({ enabled: true, active: false });
    expect(await editor.execute('bold', null)).toEqual({ ok: true, status: 'executed' });
    await waitFor(() => expect(commands.getState('bold').active).toBe(true));
    expect(editor.api().handle.selectionFormatting(0, 'A3:A3').bold).toBe(true);
    expect(failure(await editor.execute('searchMenus', null))).toBe('unsupported-command');
    expect(commands.getState('exportPng').enabled).toBe(isPngExportAvailable());
  });

  it('writes a cell draft once before a command, without a second commit on blur', async () => {
    const editor = await mountCommands();
    const input = editor.typeInCell('Bold draft');
    expect(await editor.execute('bold', null)).toEqual({ ok: true, status: 'executed' });
    expect(editor.input()).toBe('Bold draft');
    expect(editor.view.queryByTestId('xlsx-cell-editor')).toBeNull();
    expect(editor.api().handle.selectionFormatting(0, 'A3:A3').bold).toBe(true);
    fireEvent.blur(input);
    expect(editor.changes()).toBe(2);
    expect((editor.view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A3');
  });

  it('undoes a formula draft it committed, then refuses a second undo and redoes it', async () => {
    const editor = await mountCommands();
    const before = editor.input();
    fireEvent.change(editor.view.getByTestId('xlsx-formula-input'), {
      target: { value: 'Typed' },
    });
    expect(editor.api().commands.getState('undo').enabled).toBe(true);
    expect(await editor.execute('undo', null)).toEqual({ ok: true, status: 'executed' });
    expect(editor.input()).toBe(before);
    expect(failure(await editor.execute('undo', null))).toBe('nothing-to-undo');
    expect(await editor.execute('redo', null)).toEqual({ ok: true, status: 'executed' });
    expect(editor.input()).toBe('Typed');
    fireEvent.change(editor.view.getByTestId('xlsx-formula-input'), {
      target: { value: 'Later' },
    });
    expect(editor.input()).toBe('Typed');
  });

  it('merges, saves and exports after a pending draft', async () => {
    const editor = await mountCommands();
    await act(async () => {
      editor.api().selectCells(0, { anchor: { row: 2, col: 0 }, focus: { row: 2, col: 1 } });
    });
    fireEvent.change(editor.view.getByTestId('xlsx-formula-input'), {
      target: { value: 'Merged draft' },
    });
    expect(await editor.execute('merge', { value: 'all' })).toEqual({ ok: true, status: 'executed' });
    expect(editor.api().handle.cell(0, 2, 1).input).toBe('Merged draft');
    expect(editor.api().handle.mergedRanges(0, 'A3:B3')).toHaveLength(1);
    await waitFor(() =>
      expect(editor.api().commands.getState('merge', { value: 'unmerge' }).enabled).toBe(true)
    );
    fireEvent.change(editor.view.getByTestId('xlsx-formula-input'), {
      target: { value: 'Saved draft' },
    });
    expect(await editor.execute('save', null)).toEqual({ ok: true, status: 'executed' });
    const reopened = openWorkbook(editor.saves[0]);
    try {
      expect(reopened.cell(0, 2, 0).input).toBe('Saved draft');
      expect(reopened.cell(0, 2, 1).input).toBe('Merged draft');
      expect(reopened.mergedRanges(0, 'A3:B3')).toHaveLength(1);
    } finally {
      reopened.dispose();
    }
  });

  it('keeps a draft that could not be written and does not run the command', async () => {
    const editor = await mountCommands();
    editor.typeInCell('Refused');
    const handle = editor.api().handle;
    const editCell = handle.editCell;
    handle.editCell = () => {
      throw new Error('cell is locked');
    };
    try {
      expect(failure(await editor.execute('bold', null))).toBe('input-failed');
    } finally {
      handle.editCell = editCell;
    }
    const input = editor.view.getByTestId('xlsx-cell-editor') as HTMLInputElement;
    expect(input.value).toBe('Refused');
    expect(handle.selectionFormatting(0, 'A3:A3').bold).toBe(false);
    expect(failure(await editor.execute('bold', null))).toBe('input-failed');
    fireEvent.keyDown(input, { key: 'Enter' });
    expect(editor.input()).toBe('Refused');
    await act(async () => {
      editor.api().selectCells(0, selectionAt(target));
    });
    expect(await editor.execute('bold', null)).toEqual({ ok: true, status: 'executed' });
  });

  it('ends a composition before the command and writes its final text', async () => {
    const editor = await mountCommands();
    const input = editor.typeInCell('日本');
    act(() => input.focus());
    fireEvent.compositionStart(input);
    fireEvent.keyDown(input, { key: 'Enter', keyCode: 229 });
    expect(editor.view.getByTestId('xlsx-cell-editor')).toBe(input);
    let result: Promise<XlsxCommandResult> | undefined;
    let settled = false;
    await act(async () => {
      result = editor.api().commands.execute('italic', null);
      void result.then(() => (settled = true));
    });
    expect(document.activeElement).not.toBe(input);
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 700));
    });
    expect(settled).toBe(false);
    expect(editor.input()).not.toBe('日本語');
    fireEvent.change(input, { target: { value: '日本語' } });
    await act(async () => {
      fireEvent.compositionEnd(input);
      expect(await result!).toEqual({ ok: true, status: 'executed' });
    });
    expect(editor.input()).toBe('日本語');
    expect(editor.api().handle.selectionFormatting(0, 'A3:A3').italic).toBe(true);
    expect(editor.view.queryByTestId('xlsx-cell-editor')).toBeNull();
  });

  it('saves a composed formula draft even when Enter commits it while Save waits', async () => {
    const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    let resolveClipboard!: (text: string) => void;
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { readText: () => new Promise<string>((resolve) => (resolveClipboard = resolve)) },
    });
    try {
      const editor = await mountCommands();
      await act(async () => {
        editor.api().selectCells(0, selectionAt({ row: 6, col: 4 }));
      });
      fireEvent.keyDown(editor.view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      await act(async () => {
        editor.api().selectCells(0, selectionAt(target));
      });
      const formula = editor.view.getByTestId('xlsx-formula-input') as HTMLInputElement;
      act(() => formula.focus());
      fireEvent.compositionStart(formula);
      fireEvent.change(formula, { target: { value: '日本' } });
      const save = editor.api().commands.execute('save', null);
      fireEvent.change(formula, { target: { value: '日本語' } });
      await act(async () => {
        fireEvent.compositionEnd(formula);
        await new Promise((resolve) => setTimeout(resolve, 0));
      });
      act(() => formula.focus());
      fireEvent.keyDown(formula, { key: 'Enter' });
      await act(async () => resolveClipboard('Pasted'));
      expect(await save).toEqual({ ok: true, status: 'executed' });
      const saved = openWorkbook(editor.saves[0]);
      try {
        expect(saved.cell(0, target.row, target.col).input).toBe('日本語');
        expect(saved.cell(0, 6, 4).input).toBe('Pasted');
      } finally {
        saved.dispose();
      }
    } finally {
      if (originalClipboard) Object.defineProperty(navigator, 'clipboard', originalClipboard);
      else Reflect.deleteProperty(navigator, 'clipboard');
    }
  });

  function holdClipboard() {
    const original = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    const reads: ((text: string) => void)[] = [];
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { readText: () => new Promise<string>((resolve) => reads.push(resolve)) },
    });
    return {
      resolve: (text: string) => reads.shift()!(text),
      restore: () => {
        if (original) Object.defineProperty(navigator, 'clipboard', original);
        else Reflect.deleteProperty(navigator, 'clipboard');
      },
    };
  }

  const oversized = 'x'.repeat(32_768);

  const blocksUntilCorrected = async (leave: 'Enter' | 'clearSelection') => {
    const clipboard = holdClipboard();
    try {
      const editor = await mountCommands();
      fireEvent.keyDown(editor.view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      const formula = editor.view.getByTestId('xlsx-formula-input') as HTMLInputElement;
      const typeAt = async (row: number, col: number, value: string) => {
        await act(async () => {
          editor.api().selectCells(0, selectionAt({ row, col }));
        });
        fireEvent.change(formula, { target: { value } });
      };
      await typeAt(3, 1, oversized);
      fireEvent.keyDown(formula, { key: 'Enter' });
      await typeAt(5, 2, 'Newer');
      await act(async () => clipboard.resolve('Pasted'));
      expect(editor.input()).toBe('Pasted');
      expect(formula.value).toBe('Newer');

      const blocked = await editor.execute('save', null);
      expect(blocked.ok ? null : blocked.failure.code).toBe('input-failed');
      const zoom = await editor.execute('zoom', { scale: 2 });
      expect(zoom.ok ? null : zoom.failure.code).toBe('input-failed');
      expect(editor.saves).toHaveLength(0);
      expect(() => editor.api().save()).toThrow(XlsxSaveRefusedError);

      if (leave === 'Enter') fireEvent.keyDown(formula, { key: 'Enter' });
      else await act(async () => editor.api().clearSelection());
      expect(editor.api().handle.cell(0, 5, 2).input).toBe('Newer');
      await waitFor(() => expect(formula.value).toBe(oversized));
      expect((editor.view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('B4');
      const stillBlocked = await editor.execute('save', null);
      expect(stillBlocked.ok ? null : stillBlocked.failure.code).toBe('input-failed');

      fireEvent.change(formula, { target: { value: 'Corrected' } });
      fireEvent.keyDown(formula, { key: 'Enter' });
      expect(editor.api().handle.cell(0, 3, 1).input).toBe('Corrected');
      expect(await editor.execute('save', null)).toEqual({ ok: true, status: 'executed' });
      const saved = openWorkbook(editor.saves[0]);
      try {
        expect(saved.cell(0, 3, 1).input).toBe('Corrected');
        expect(saved.cell(0, 5, 2).input).toBe('Newer');
      } finally {
        saved.dispose();
      }
    } finally {
      clipboard.restore();
    }
  };

  it('blocks every command while a rejected entry waits for correction on its own cell', () =>
    blocksUntilCorrected('Enter'));

  it('reopens a rejected entry at its own cell when a host clears the selection', () =>
    blocksUntilCorrected('clearSelection'));

  it('refuses a synchronous save while accepted input waits, and saves it through the command', async () => {
    const clipboard = holdClipboard();
    try {
      const editor = await mountCommands();
      await act(async () => {
        editor.api().selectCells(0, selectionAt({ row: 0, col: 0 }));
      });
      const surface = editor.view.getByTestId('xlsx-scroll');
      fireEvent.keyDown(surface, { key: 'v', ctrlKey: true });
      await act(async () => {
        editor.api().selectCells(0, selectionAt({ row: 3, col: 1 }));
      });
      fireEvent.keyDown(surface, { key: 'F2' });
      const cell = editor.view.getByTestId('xlsx-cell-editor');
      fireEvent.change(cell, { target: { value: 'Entered' } });
      fireEvent.keyDown(cell, { key: 'Enter' });
      expect(editor.view.queryByTestId('xlsx-cell-editor')).toBeNull();

      let refusal: unknown;
      try {
        editor.api().save();
      } catch (error) {
        refusal = error;
      }
      expect(refusal).toBeInstanceOf(XlsxSaveRefusedError);
      expect((refusal as XlsxSaveRefusedError).code).toBe('input-pending');

      const saving = editor.api().commands.execute('save', null);
      await act(async () => clipboard.resolve('Pasted'));
      expect(await saving).toEqual({ ok: true, status: 'executed' });
      const saved = openWorkbook(editor.saves[0]);
      try {
        expect(saved.cell(0, 0, 0).input).toBe('Pasted');
        expect(saved.cell(0, 3, 1).input).toBe('Entered');
      } finally {
        saved.dispose();
      }
      await act(async () => {
        await new Promise((resolve) => setTimeout(resolve, 0));
      });
      const bytes = editor.api().save();
      const reopened = openWorkbook(bytes);
      try {
        expect(reopened.cell(0, 3, 1).input).toBe('Entered');
      } finally {
        reopened.dispose();
      }
    } finally {
      clipboard.restore();
    }
  });

  for (const caller of ['selectCells', 'clearSelection', 'sheet switch'] as const) {
    it(`queues a correction closed by ${caller} behind the entry it corrects`, async () => {
      const clipboard = holdClipboard();
      try {
        const editor = await mountCommands();
        const surface = editor.view.getByTestId('xlsx-scroll');
        await act(async () => {
          editor.api().selectCells(0, selectionAt({ row: 0, col: 0 }));
        });
        fireEvent.keyDown(surface, { key: 'v', ctrlKey: true });
        const typeAt = async (value: string, commit: boolean) => {
          await act(async () => {
            editor.api().selectCells(0, selectionAt({ row: 3, col: 1 }));
          });
          fireEvent.keyDown(surface, { key: 'F2' });
          const cell = editor.view.getByTestId('xlsx-cell-editor');
          act(() => cell.focus());
          fireEvent.change(cell, { target: { value } });
          if (commit) fireEvent.keyDown(cell, { key: 'Enter' });
        };
        await typeAt('first', true);
        await typeAt('corrected', false);
        await act(async () => {
          if (caller === 'selectCells') {
            expect(editor.api().selectCells(0, selectionAt({ row: 5, col: 2 }))).toBe(true);
          } else if (caller === 'clearSelection') {
            editor.api().clearSelection();
          } else {
            fireEvent.click(editor.view.getAllByRole('tab')[1]);
          }
        });
        expect(editor.view.queryByTestId('xlsx-cell-editor')).toBeNull();
        expect(() => editor.api().save()).toThrow(XlsxSaveRefusedError);
        await act(async () => clipboard.resolve('Pasted'));
        await act(async () => {
          await new Promise((resolve) => setTimeout(resolve, 0));
        });
        expect(editor.api().handle.cell(0, 0, 0).input).toBe('Pasted');
        expect(editor.api().handle.cell(0, 3, 1).input).toBe('corrected');
      } finally {
        clipboard.restore();
      }
    });
  }

  it('runs a command on the selection a host set in the same handler', async () => {
    const editor = await mountCommands();
    const api = editor.api();
    let result!: XlsxCommandResult;
    await act(async () => {
      expect(api.selectCells(0, selectionAt({ row: 3, col: 1 }))).toBe(true);
      expect(api.commands.getState('merge')).toMatchObject({ value: { rows: 1, columns: 1 } });
      result = await api.commands.execute('bold', null);
    });
    expect(result).toEqual({ ok: true, status: 'executed' });
    expect(api.handle.selectionFormatting(0, 'B4:B4').bold).toBe(true);
    expect(api.handle.selectionFormatting(0, 'A3:A3').bold).toBe(false);
  });

  it('runs a command on the sheet a host switched to in the same handler', async () => {
    const editor = await mountCommands();
    const api = editor.api();
    let result!: XlsxCommandResult;
    await act(async () => {
      expect(api.selectCells(1, selectionAt({ row: 0, col: 0 }))).toBe(true);
      result = await api.commands.execute('italic', null);
    });
    expect(result).toEqual({ ok: true, status: 'executed' });
    expect(api.handle.selectionFormatting(1, 'A1:A1').italic).toBe(true);
    expect(api.handle.selectionFormatting(0, 'A3:A3').italic).toBe(false);
    expect(api.handle.selectionFormatting(0, 'A1:A1').italic).toBe(false);
  });

  it('keeps a queued entry on its cell and runs a command on the newly set selection', async () => {
    const clipboard = holdClipboard();
    try {
      const editor = await mountCommands();
      const api = editor.api();
      const surface = editor.view.getByTestId('xlsx-scroll');
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 0, col: 0 }));
      });
      fireEvent.keyDown(surface, { key: 'v', ctrlKey: true });
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 3, col: 1 }));
      });
      fireEvent.keyDown(surface, { key: 'F2' });
      const cell = editor.view.getByTestId('xlsx-cell-editor');
      fireEvent.change(cell, { target: { value: 'Queued entry' } });
      fireEvent.keyDown(cell, { key: 'Enter' });
      let bold!: Promise<XlsxCommandResult>;
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 5, col: 2 }));
        bold = api.commands.execute('bold', null);
      });
      await act(async () => clipboard.resolve('Pasted'));
      expect(await bold).toEqual({ ok: true, status: 'executed' });
      expect(api.handle.cell(0, 3, 1).input).toBe('Queued entry');
      expect(api.handle.selectionFormatting(0, 'C6:C6').bold).toBe(true);
      expect(api.handle.selectionFormatting(0, 'B4:B4').bold).toBe(false);
      expect(api.handle.selectionFormatting(0, 'B5:B5').bold).toBe(false);
    } finally {
      clipboard.restore();
    }
  });

  function holdClipboardWrite() {
    const original = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    const writes: { text: string; resolve: () => void }[] = [];
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        writeText: (text: string) =>
          new Promise<void>((resolve) => writes.push({ text, resolve })),
      },
    });
    return {
      writes,
      restore: () => {
        if (original) Object.defineProperty(navigator, 'clipboard', original);
        else Reflect.deleteProperty(navigator, 'clipboard');
      },
    };
  }

  it('clears only the range and sheet a cut started on', async () => {
    const clipboard = holdClipboardWrite();
    try {
      const editor = await mountCommands();
      const api = editor.api();
      const surface = editor.view.getByTestId('xlsx-scroll');
      api.handle.editCells(0, [
        { row: 0, col: 0, input: 'Cut me' },
        { row: 3, col: 1, input: 'Keep me' },
      ]);
      api.handle.editCells(1, [{ row: 0, col: 0, input: 'Other sheet' }]);
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 0, col: 0 }));
      });
      fireEvent.keyDown(surface, { key: 'x', ctrlKey: true });
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 3, col: 1 }));
      });
      await act(async () => clipboard.writes[0].resolve());
      expect(clipboard.writes[0].text).toBe('Cut me');
      expect(api.handle.cell(0, 0, 0).input).toBe('');
      expect(api.handle.cell(0, 3, 1).input).toBe('Keep me');

      fireEvent.keyDown(surface, { key: 'x', ctrlKey: true });
      await act(async () => {
        api.selectCells(1, selectionAt({ row: 0, col: 0 }));
      });
      await act(async () => clipboard.writes[1].resolve());
      expect(clipboard.writes[1].text).toBe('Keep me');
      expect(api.handle.cell(0, 3, 1).input).toBe('');
      expect(api.handle.cell(1, 0, 0).input).toBe('Other sheet');
    } finally {
      clipboard.restore();
    }
  });

  it('copies the range a host selected in the same handler', async () => {
    const clipboard = holdClipboardWrite();
    try {
      const editor = await mountCommands();
      const api = editor.api();
      const surface = editor.view.getByTestId('xlsx-scroll');
      api.handle.editCells(0, [
        { row: 0, col: 0, input: 'First' },
        { row: 3, col: 1, input: 'Second' },
      ]);
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 0, col: 0 }));
      });
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 3, col: 1 }));
        fireEvent.keyDown(surface, { key: 'c', ctrlKey: true });
      });
      expect(clipboard.writes.map((write) => write.text)).toEqual(['Second']);
    } finally {
      clipboard.restore();
    }
  });

  for (const key of ['c', 'Delete'] as const) it(`${key === 'c' ? 'copies' : 'clears'} a Ctrl+A sheet selection through its used cells`, async () => {
    const clipboard = holdClipboardWrite();
    try {
      const editor = await mountCommands();
      const api = editor.api();
      const surface = editor.view.getByTestId('xlsx-scroll');
      api.handle.editCells(0, [
        { row: 0, col: 0, input: 'First' },
        { row: 3, col: 1, input: 'Second' },
      ]);
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 40, col: 20 }));
      });
      await act(async () => {
        fireEvent.keyDown(surface, { key: 'a', ctrlKey: true });
      });
      await act(async () => {
        fireEvent.keyDown(surface, key === 'c' ? { key, ctrlKey: true } : { key });
      });
      if (key === 'Delete') {
        await waitFor(() => expect(api.handle.cell(0, 3, 1).input).toBe(''));
        expect(api.handle.cell(0, 0, 0).input).toBe('');
        return;
      }
      const rows = clipboard.writes[0]!.text.split(/\r?\n/).map((row) => row.split('\t'));
      expect(rows[0]![0]).toBe('First');
      expect(rows[3]![1]).toBe('Second');
      expect(rows.length).toBeLessThan(1_000);
    } finally {
      clipboard.restore();
    }
  });

  const keepsRejectedOnItsSheet = async (back: 'tab' | 'selectCells') => {
    const source = openWorkbook(plain.bytes.slice());
    source.applyOps([{ type: 'addSheet', index: 1, name: 'Second' }]);
    const file = source.save();
    source.dispose();
    const clipboard = holdClipboard();
    try {
      const editor = await mountCommands({ file });
      fireEvent.keyDown(editor.view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      const formula = editor.view.getByTestId('xlsx-formula-input') as HTMLInputElement;
      await act(async () => {
        editor.api().selectCells(0, selectionAt({ row: 3, col: 1 }));
      });
      fireEvent.change(formula, { target: { value: oversized } });
      fireEvent.keyDown(formula, { key: 'Enter' });
      const tabs = editor.view.getAllByRole('tab');
      await act(async () => {
        fireEvent.click(tabs[1]);
      });
      fireEvent.change(formula, { target: { value: 'On the second sheet' } });
      await act(async () => clipboard.resolve('Pasted'));
      expect(formula.value).toBe('On the second sheet');

      fireEvent.keyDown(formula, { key: 'Enter' });
      expect(editor.api().handle.cell(1, 0, 0).input).toBe('On the second sheet');
      expect(formula.value).not.toBe(oversized);
      expect(editor.view.queryByTestId('xlsx-cell-editor')).toBeNull();
      const blocked = await editor.execute('save', null);
      expect(blocked.ok ? null : blocked.failure.code).toBe('input-failed');
      expect(editor.api().handle.cell(0, 3, 1).input).not.toBe('On the second sheet');
      expect(editor.api().handle.cell(0, 3, 1).input).not.toBe(oversized);

      await act(async () => {
        if (back === 'tab') fireEvent.click(editor.view.getAllByRole('tab')[0]);
        else editor.api().selectCells(0, selectionAt({ row: 0, col: 0 }));
      });
      await waitFor(() => expect(formula.value).toBe(oversized));
      expect((editor.view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('B4');
      fireEvent.keyDown(formula, { key: 'Escape' });
      expect(await editor.execute('save', null)).toEqual({ ok: true, status: 'executed' });
      const saved = openWorkbook(editor.saves[0]);
      try {
        expect(saved.cell(0, 3, 1).input).not.toBe(oversized);
        expect(saved.cell(1, 0, 0).input).toBe('On the second sheet');
      } finally {
        saved.dispose();
      }
    } finally {
      clipboard.restore();
    }
  };

  it('keeps a rejected entry on its own sheet and lets Escape discard it', () =>
    keepsRejectedOnItsSheet('tab'));

  it('reopens a rejected entry at its own cell when a host selects its sheet', () =>
    keepsRejectedOnItsSheet('selectCells'));

  it('never reopens a rejected entry whose correction a host selection queued', async () => {
    const clipboard = holdClipboard();
    try {
      const editor = await mountCommands();
      const formula = editor.view.getByTestId('xlsx-formula-input') as HTMLInputElement;
      const nameBox = editor.view.getByTestId('xlsx-name-box') as HTMLInputElement;
      await act(async () => {
        editor.api().selectCells(0, selectionAt({ row: 3, col: 1 }));
      });
      fireEvent.change(formula, { target: { value: oversized } });
      fireEvent.keyDown(formula, { key: 'Enter' });
      expect(formula.value).toBe(oversized);
      fireEvent.keyDown(editor.view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      fireEvent.change(formula, { target: { value: 'Corrected' } });
      let selected: boolean | undefined;
      await act(async () => {
        selected = editor.api().selectCells(0, selectionAt({ row: 0, col: 0 }));
      });
      expect(selected).toBe(true);
      expect(nameBox.value).toBe('A1');
      expect(formula.value).not.toBe(oversized);

      await act(async () => clipboard.resolve('Pasted'));
      expect(editor.api().handle.cell(0, 3, 1).input).toBe('Corrected');
      expect(nameBox.value).toBe('A1');
      expect(formula.value).not.toBe(oversized);
      expect(await editor.execute('save', null)).toEqual({ ok: true, status: 'executed' });
    } finally {
      clipboard.restore();
    }
  });

  it('fires onReady only once the workbook painted its first frame', async () => {
    const target = cellRect(plain.grid, 3, 0)!;
    const point = { x: target.x + target.w / 2, y: target.y + target.h / 2 };
    const getRect = HTMLCanvasElement.prototype.getBoundingClientRect;
    HTMLCanvasElement.prototype.getBoundingClientRect = () => new DOMRect(0, 0, VIEWPORT.width, VIEWPORT.height);
    const hits: unknown[] = [];
    const props = { onReady: (ready: XlsxEditorApi) => { hits.push(ready.getPositionAtPoint(point.x, point.y)); } };
    try {
      const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} {...props} file={plain.bytes.slice()} />);
      await waitFor(() => expect(hits).toHaveLength(1), MOUNTED);
      view.rerender(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} {...props} file={plain.bytes.slice()} />);
      await waitFor(() => expect(hits).toHaveLength(2), MOUNTED);
      expect(hits).toEqual([{ sheet: 0, row: 3, col: 0 }, { sheet: 0, row: 3, col: 0 }]);
    } finally {
      HTMLCanvasElement.prototype.getBoundingClientRect = getRect;
    }
  });

  it('fires onReady once and shows the error when the first frame cannot be painted', async () => {
    const paint = spyOn(xlsx, 'paintDisplayList').mockImplementation(() => {
      throw new Error('pattern tile has no context');
    });
    const ready: XlsxEditorApi[] = [];
    try {
      const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()}
        onReady={(api) => { ready.push(api); }} />);
      await waitFor(() => expect(ready).toHaveLength(1), MOUNTED);
      expect(view.getByTestId('xlsx-render-error').textContent).toContain('pattern tile has no context');
      paint.mockRestore();
      fireEvent.scroll(view.getByTestId('xlsx-scroll'));
      await act(async () => { await new Promise((resolve) => setTimeout(resolve, 50)); });
      expect(ready).toHaveLength(1);
    } finally {
      paint.mockRestore();
    }
  });

  it('prints a replacement workbook only after it painted', async () => {
    let paints = 0;
    const getContext = HTMLCanvasElement.prototype.getContext;
    HTMLCanvasElement.prototype.getContext = (() => {
      paints += 1;
      return stubContext();
    }) as unknown as HTMLCanvasElement['getContext'];
    const print = window.print;
    let paintsAtPrint: number | null = null;
    window.print = () => {
      paintsAtPrint = paints;
    };
    try {
      let api: XlsxEditorApi | undefined;
      const view = render(
        <XlsxEditor legacySyncApi experimentalWorkerOpen={false}
          file={plain.bytes.slice()}
          onReady={(ready) => {
            api = ready;
          }}
        />
      );
      await waitFor(() => expect(api).toBeDefined(), MOUNTED);
      let printing: Promise<XlsxCommandResult> | undefined;
      paints = 0;
      view.rerender(
        <XlsxEditor legacySyncApi experimentalWorkerOpen={false}
          file={charted.bytes.slice()}
          onReady={(ready) => {
            printing = ready.commands.execute('print', null);
          }}
        />
      );
      await waitFor(() => expect(printing).toBeDefined());
      expect(await printing!).toEqual({ ok: true, status: 'executed' });
      expect(paintsAtPrint ?? 0).toBeGreaterThan(0);
    } finally {
      HTMLCanvasElement.prototype.getContext = getContext;
      window.print = print;
    }
  });

  it('fails a command waiting on a composition whose input went away', async () => {
    const editor = await mountCommands();
    const input = editor.typeInCell('日本');
    fireEvent.compositionStart(input);
    let result: Promise<XlsxCommandResult> | undefined;
    await act(async () => {
      result = editor.api().commands.execute('italic', null);
    });
    await act(async () => editor.rerender({ readOnly: true }));
    await act(async () => {
      const outcome = await result!;
      expect(outcome.ok ? null : outcome.failure.code).toBe('input-failed');
    });
    expect(editor.api().handle.selectionFormatting(0, 'A3:A3').italic).toBe(false);
  });

  it('fails a command whose preceding chart move could not land', async () => {
    let api: XlsxEditorApi | undefined;
    const saves: Uint8Array[] = [];
    const view = render(
      <XlsxEditor legacySyncApi experimentalWorkerOpen={false}
        file={charted.bytes.slice()}
        onSave={(bytes) => {
          saves.push(bytes);
        }}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    const surface = view.getByTestId('xlsx-scroll');
    const [chart] = charted.charts;
    await waitFor(() => {
      fireEvent.mouseDown(surface, chartCenter(chart));
      fireEvent.mouseUp(window, chartCenter(chart));
      expect(view.queryByTestId('xlsx-chart-selection')).not.toBeNull();
    }, MOUNTED);
    await act(async () => {
      fireEvent.keyDown(surface, { key: 'ArrowRight' });
    });
    const moveChart = api!.handle.moveChart;
    api!.handle.moveChart = () => {
      throw new Error('anchor refused');
    };
    let result!: XlsxCommandResult;
    try {
      await act(async () => {
        result = await api!.commands.execute('save', null);
      });
    } finally {
      api!.handle.moveChart = moveChart;
    }
    expect(result.ok ? null : result.failure.code).toBe('input-failed');
    expect(saves).toHaveLength(0);
  });

  for (const source of ['cell', 'formula'] as const) {
    it(`writes a ${source} draft before formatting, borders and export`, async () => {
      const clicks: string[] = [];
      const click = HTMLAnchorElement.prototype.click;
      HTMLAnchorElement.prototype.click = function (this: HTMLAnchorElement) {
        clicks.push(this.download);
      };
      try {
        const editor = await mountCommands({ fileName: 'report.xlsx', defaultFonts: false });
        const type = (value: string) =>
          source === 'cell'
            ? editor.typeInCell(value)
            : fireEvent.change(editor.view.getByTestId('xlsx-formula-input'), {
                target: { value },
              });
        type('Styled');
        expect(await editor.execute('borderStyle', { value: 'dashed' })).toEqual({
          ok: true,
          status: 'noop',
        });
        expect(editor.input()).toBe('Styled');
        type('Framed');
        expect(await editor.execute('borderColor', { color: '#ff0000' })).toEqual({
          ok: true,
          status: 'noop',
        });
        expect(editor.input()).toBe('Framed');
        expect(await editor.execute('borderPreset', { value: 'outer' })).toEqual({
          ok: true,
          status: 'executed',
        });
        const formatting = editor.api().handle.selectionFormatting(0, 'A3:A3');
        expect([editor.input(), formatting.borderStyle, formatting.borderColor]).toEqual([
          'Framed',
          'dashed',
          '#ff0000',
        ]);
        expect(editor.api().commands.getState('borderStyle').value).toBe('dashed');
        type('Exported');
        const exported: unknown[] = [];
        const handle = editor.api().handle;
        const renderPng = handle.renderPng;
        handle.renderPng = (target) => {
          exported.push({ target, input: handle.cell(0, 2, 0).input });
          return Promise.resolve(Uint8Array.of(137, 80, 78, 71));
        };
        try {
          expect(await editor.execute('exportPng', null)).toEqual({ ok: true, status: 'executed' });
        } finally { handle.renderPng = renderPng; }
        expect(editor.input()).toBe('Exported');
        expect(clicks).toEqual(['report.png']);
        expect(exported).toEqual([{ target: { sheet: 0, viewport: expect.any(Object) }, input: 'Exported' }]);
      } finally {
        HTMLAnchorElement.prototype.click = click;
      }
    });
  }

  it('captures a paint format, applies it to the next selection and resets on read-only or reopen', async () => {
    const editor = await mountCommands();
    await editor.execute('bold', null);
    expect(await editor.execute('paintFormat', null)).toEqual({ ok: true, status: 'executed' });
    await waitFor(() => expect(editor.api().commands.getState('paintFormat').active).toBe(true));
    const surface = editor.view.getByTestId('xlsx-scroll');
    const next = pointAt(plain, { row: 3, col: 1 });
    await act(async () => {
      fireEvent.mouseDown(surface, next);
      fireEvent.mouseUp(window, next);
    });
    const painted = (editor.view.getByTestId('xlsx-name-box') as HTMLInputElement).value;
    expect(painted).not.toBe('A3');
    await waitFor(() =>
      expect(editor.api().handle.selectionFormatting(0, `${painted}:${painted}`).bold).toBe(true)
    );
    expect(editor.api().commands.getState('paintFormat').active).toBe(false);

    await editor.execute('paintFormat', null);
    await act(async () => editor.rerender({ readOnly: true }));
    await act(async () => editor.rerender({ readOnly: false }));
    expect(editor.api().commands.getState('paintFormat').active).toBe(false);

    await editor.execute('paintFormat', null);
    await act(async () => editor.rerender({ file: plain.bytes.slice() }));
    await waitFor(() => expect(editor.api().commands.getState('bold').enabled).toBe(true));
    expect(editor.api().commands.getState('paintFormat').active).toBe(false);
  });

  it('reports the workbook lifecycle and fails commands of a replaced workbook', async () => {
    let api: XlsxEditorApi | undefined;
    let committed: XlsxCommandStore | undefined;
    function Probe() {
      const store = useXlsxCommands();
      useEffect(() => { committed = store; }, [store]);
      return null;
    }
    const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} toolbar={<Probe />} />);
    await waitFor(() => expect(committed).toBeDefined());
    const store = committed!;
    const code = (state: { enabled: boolean; disabledReason?: { code: string } }) =>
      state.enabled ? null : state.disabledReason!.code;
    expect(code(store.getState('save'))).toBe('no-document');
    view.rerender(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()} onReady={(ready) => void (api = ready)} />);
    expect(code(store.getState('save'))).toBe('document-loading');
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    await waitFor(() => expect(store.getState('save').enabled).toBe(true));
    expect(api!.commands).toBe(store);
  });

  it('keeps a failed Enter commit open in the editor, also through host selection calls', async () => {
    const editor = await mountCommands();
    const input = editor.typeInCell('Locked');
    const handle = editor.api().handle;
    const editCell = handle.editCell;
    handle.editCell = () => {
      throw new Error('cell is locked');
    };
    let selected: boolean | undefined;
    try {
      fireEvent.keyDown(input, { key: 'Enter' });
      await act(async () => {
        selected = editor.api().selectCells(0, selectionAt({ row: 5, col: 2 }));
        editor.api().clearSelection();
      });
    } finally {
      handle.editCell = editCell;
    }
    expect(selected).toBe(false);
    expect((editor.view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('Locked');
    expect((editor.view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe('A3');
    fireEvent.keyDown(editor.view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    expect(editor.input()).toBe('Locked');
  });

  it('writes pastes, Enter commits and commands in the order they were accepted', async () => {
    const reads: ((text: string) => void)[] = [];
    const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        readText: () => new Promise<string>((resolve) => reads.push(resolve)),
      },
    });
    try {
      const editor = await mountCommands();
      const surface = editor.view.getByTestId('xlsx-scroll');
      fireEvent.keyDown(surface, { key: 'v', ctrlKey: true });
      const save = editor.api().commands.execute('save', null);
      await act(async () => {
        editor.api().selectCells(0, selectionAt({ row: 3, col: 1 }));
      });
      fireEvent.keyDown(surface, { key: 'v', ctrlKey: true });
      await act(async () => reads[1]('Second'));
      expect(editor.api().handle.cell(0, 3, 1).input).not.toBe('Second');
      await act(async () => reads[0]('First'));
      expect(await save).toEqual({ ok: true, status: 'executed' });
      const saved = openWorkbook(editor.saves[0]);
      try {
        expect(saved.cell(0, 2, 0).input).toBe('First');
        expect(saved.cell(0, 3, 1).input).not.toBe('Second');
      } finally {
        saved.dispose();
      }
      await waitFor(() => expect(editor.api().handle.cell(0, 3, 1).input).toBe('Second'));

      fireEvent.keyDown(surface, { key: 'v', ctrlKey: true });
      const input = editor.typeInCell('Queued');
      fireEvent.keyDown(input, { key: 'Enter' });
      expect(editor.view.queryByTestId('xlsx-cell-editor')).toBeNull();
      const handle = editor.api().handle;
      const editCell = handle.editCell;
      handle.editCell = () => {
        throw new Error('cell is locked');
      };
      const bold = editor.api().commands.execute('bold', null);
      try {
        await act(async () => reads[2]('Third'));
        const result = await bold;
        expect(result.ok ? null : result.failure.code).toBe('input-failed');
      } finally {
        handle.editCell = editCell;
      }
      await waitFor(() =>
        expect((editor.view.getByTestId('xlsx-cell-editor') as HTMLInputElement).value).toBe('Queued')
      );
    } finally {
      if (originalClipboard) Object.defineProperty(navigator, 'clipboard', originalClipboard);
      else Reflect.deleteProperty(navigator, 'clipboard');
    }
  });

  it('refuses a queued cell command whose selection moved, and a color picked for another selection', async () => {
    const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    let resolveClipboard!: (text: string) => void;
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { readText: () => new Promise<string>((resolve) => (resolveClipboard = resolve)) },
    });
    try {
      const editor = await mountCommands();
      fireEvent.keyDown(editor.view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      const bold = editor.api().commands.execute('bold', null);
      await act(async () => resolveClipboard('A\tB'));
      const result = await bold;
      expect(result.ok ? null : result.failure.code).toBe('target-changed');
      expect(editor.api().handle.selectionFormatting(0, 'A3:B3').bold).toBe(false);

      const picker = editor.view.getByLabelText('Text color', { selector: 'input' });
      fireEvent.click(picker);
      await act(async () => {
        editor.api().selectCells(0, selectionAt({ row: 5, col: 2 }));
      });
      await act(async () => {
        fireEvent.change(picker, { target: { value: '#ff0000' } });
      });
      expect(editor.api().handle.selectionFormatting(0, 'C6:C6').textColor).toBe('#000000');
      fireEvent.click(picker);
      await act(async () => {
        fireEvent.change(picker, { target: { value: '#00ff00' } });
      });
      await waitFor(() =>
        expect(editor.api().handle.selectionFormatting(0, 'C6:C6').textColor).toBe('#00ff00')
      );
    } finally {
      if (originalClipboard) Object.defineProperty(navigator, 'clipboard', originalClipboard);
      else Reflect.deleteProperty(navigator, 'clipboard');
    }
  });

  it('prints only once the canvas shows the text written before it', async () => {
    const painted: string[] = [];
    const getContext = HTMLCanvasElement.prototype.getContext;
    HTMLCanvasElement.prototype.getContext = (() => ({
      ...stubContext(),
      fillText: (text: string) => painted.push(text),
    })) as unknown as HTMLCanvasElement['getContext'];
    const print = window.print;
    let printedWith: boolean | null = null;
    window.print = () => {
      printedWith = painted.includes('Printed draft');
    };
    try {
      const editor = await mountCommands();
      fireEvent.change(editor.view.getByTestId('xlsx-formula-input'), {
        target: { value: 'Printed draft' },
      });
      painted.length = 0;
      const printing = editor.api().commands.execute('print', null);
      expect(printedWith).toBeNull();
      await waitFor(() => expect(printedWith).not.toBeNull());
      expect(printedWith === true).toBe(true);
      expect(await printing).toEqual({ ok: true, status: 'executed' });

      printedWith = null;
      const handle = editor.api().handle;
      const displayList = handle.displayList;
      handle.displayList = () => {
        throw new Error('render failed');
      };
      try {
        fireEvent.change(editor.view.getByTestId('xlsx-formula-input'), {
          target: { value: 'Unpainted' },
        });
        const failed = await editor.api().commands.execute('print', null);
        expect(failed.ok ? null : failed.failure.code).toBe('render-failed');
        expect(printedWith).toBeNull();
      } finally {
        handle.displayList = displayList;
      }
    } finally {
      HTMLCanvasElement.prototype.getContext = getContext;
      window.print = print;
    }
  });

  it('keeps replacement formula bars read-only without a writable cell', async () => {
    const host = (
      <EditorToolbar mode="commands">
        <EditorToolbar.FormulaBar />
      </EditorToolbar>
    );
    const editor = await mountCommands({ toolbar: host, readOnly: true });
    const formula = editor.view.getByTestId('xlsx-formula-input') as HTMLInputElement;
    expect(formula.readOnly).toBe(true);
    const before = editor.input();
    fireEvent.change(formula, { target: { value: 'Blocked' } });
    await act(async () => editor.rerender({ toolbar: host, readOnly: false }));
    expect(formula.readOnly).toBe(false);
    expect(formula.value).toBe(before);
    expect(await editor.execute('save', null)).toEqual({ ok: true, status: 'executed' });
    expect(editor.input()).toBe(before);
  });

  it('dispatches shortcuts from the grid and keeps text undo in the cell editor', async () => {
    const editor = await mountCommands();
    const surface = editor.view.getByTestId('xlsx-scroll');
    await act(async () => {
      fireEvent.keyDown(surface, { key: 'b', ...MOD });
    });
    await waitFor(() =>
      expect(editor.api().handle.selectionFormatting(0, 'A3:A3').bold).toBe(true)
    );
    const input = editor.typeInCell('Shortcut draft');
    await act(async () => {
      fireEvent.keyDown(input, { key: 'z', ...MOD });
    });
    expect(editor.view.getByTestId('xlsx-cell-editor')).toBe(input);
    expect(editor.input()).not.toBe('Shortcut draft');
    await act(async () => {
      fireEvent.keyDown(input, { key: 's', ...MOD });
    });
    await waitFor(() => expect(editor.saves).toHaveLength(1));
    expect(editor.input()).toBe('Shortcut draft');
    const reopened = openWorkbook(editor.saves[0]);
    try {
      expect(reopened.cell(0, target.row, target.col).input).toBe('Shortcut draft');
    } finally {
      reopened.dispose();
    }
  });

  it('waits for an accepted paste, and fails once the workbook is replaced', async () => {
    const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    let resolveClipboard!: (text: string) => void;
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        readText: () =>
          new Promise<string>((resolve) => {
            resolveClipboard = resolve;
          }),
      },
    });
    try {
      const editor = await mountCommands();
      const before = editor.input();
      const surface = editor.view.getByTestId('xlsx-scroll');
      fireEvent.keyDown(surface, { key: 'v', ...{ ctrlKey: true } });
      const undo = editor.api().commands.execute('undo', null);
      await act(async () => resolveClipboard('Pasted'));
      expect(await undo).toEqual({ ok: true, status: 'executed' });
      expect(editor.input()).toBe(before);

      fireEvent.keyDown(surface, { key: 'v', ctrlKey: true });
      const bold = editor.api().commands.execute('bold', null);
      await act(async () => {
        editor.rerender({ file: plain.bytes.slice() });
      });
      await act(async () => resolveClipboard('Late'));
      expect(failure(await bold)).toBe('document-replaced');
    } finally {
      if (originalClipboard) Object.defineProperty(navigator, 'clipboard', originalClipboard);
      else Reflect.deleteProperty(navigator, 'clipboard');
    }
  });

  it('refuses a command during an unfinished chart drag and lands a nudge before undo', async () => {
    let api: XlsxEditorApi | undefined;
    const view = render(
      <XlsxEditor legacySyncApi experimentalWorkerOpen={false}
        file={charted.bytes.slice()}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    const surface = view.getByTestId('xlsx-scroll');
    const [chart] = charted.charts;
    await waitFor(() => {
      fireEvent.mouseDown(surface, chartCenter(chart));
      expect(view.queryByTestId('xlsx-chart-selection')).not.toBeNull();
    }, MOUNTED);
    let result!: XlsxCommandResult;
    await act(async () => {
      result = await api!.commands.execute('save', null);
    });
    expect(failure(result)).toBe('gesture-active');
    expect(api!.commands.getState('bold').enabled).toBe(false);
    fireEvent.mouseUp(window, chartCenter(chart));

    await act(async () => {
      fireEvent.keyDown(surface, { key: 'ArrowRight', shiftKey: true });
    });
    await act(async () => {
      result = await api!.commands.execute('undo', null);
    });
    expect(result).toEqual({ ok: true, status: 'executed' });
    const after = api!.handle
      .displayList({ x: 0, y: 0, ...VIEWPORT })
      .charts!.find((candidate) => candidate.id === chart.id)!;
    expect(after.rect.x).toBe(chart.rect.x);
    expect(api!.commands.getState('undo').enabled).toBe(false);
  });

  it('gates writes when read-only while saving stays available', async () => {
    const editor = await mountCommands();
    await act(async () => editor.rerender({ readOnly: true }));
    const bold = editor.api().commands.getState('bold');
    expect(bold.enabled ? null : bold.disabledReason.code).toBe('read-only');
    expect(failure(await editor.execute('undo', null))).toBe('read-only');
    expect(await editor.execute('save', null)).toEqual({ ok: true, status: 'executed' });
  });

  it('enables merging while collaborating', async () => {
    let api: XlsxEditorApi | undefined;
    render(
      <XlsxEditor legacySyncApi experimentalWorkerOpen={false}
        file={plain.bytes.slice()}
        collaboration={{}}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    await act(async () => {
      api!.selectCells(0, { anchor: { row: 0, col: 0 }, focus: { row: 1, col: 1 } });
    });
    const merge = api!.commands.getState('merge', { value: 'all' });
    expect(merge.enabled ? null : merge.disabledReason.code).toBeNull();
  });

  it('allows host proposal review in a read-only editor only when opted in', async () => {
    const editor = await mountCommands({ file: await hostProposalWorkbook(), readOnly: true, allowHostProposals: true });
    expect(editor.view.getByTestId('xlsx-proposals-button')).toBeDefined();
    const api = editor.api();
    let result!: Awaited<ReturnType<NonNullable<XlsxEditorApi['proposeChanges']>>>;
    await act(async () => {
      result = await api.proposeChanges!({ agentId: 'Host agent',
        edits: [{ sheet: 0, row: 6, col: 4, input: '12' }] });
    });
    if (!result.ok) throw new Error(result.failure.message);
    const id = result.proposal.id;
    await editor.execute('proposalsPanel', { open: true });
    expect(editor.view.getByTestId('xlsx-proposals-panel').textContent).toContain('Host agent');
    expect(failure(await editor.execute('bold', null))).toBe('read-only');
    expect(failure(await editor.execute('undo', null))).toBe('read-only');
    const batch = { expectVersion: api.handle.version(), steps: [{ op: 'setCellInputs' as const,
      target: { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'E7' } }, inputs: [['99']] }] };
    await expect(api.applyEdits(batch)).resolves.toMatchObject({ ok: false, failure: { code: 'read-only' } });
    await expect(api.validateEdits(batch)).resolves.toMatchObject({ ok: false, failure: { code: 'read-only' } });
    expect(await editor.execute('proposalAccept', { proposalId: id })).toEqual({ ok: true, status: 'executed' });
    expect(api.handle.cell(0, 6, 4).input).toBe('12');
    await act(async () => editor.rerender({ readOnly: true, allowHostProposals: false }));
    await expect(api.proposeChanges!({ agentId: 'Host agent', edits: [] }))
      .resolves.toMatchObject({ ok: false, failure: { code: 'read-only' } });
  });

  it('accepts, refuses stale and force-applies proposals through commands', async () => {
    const editor = await mountCommands();
    const workbook = editor.api().handle;
    const stage = async (input: string) => {
      let id = '';
      await act(async () => {
        id = workbook.propose('Audit agent', null, [{ sheet: 0, row: 6, col: 4, input }]).id;
        editor.api().refreshProposals();
      });
      return id;
    };
    const first = await stage('12');
    expect(editor.api().commands.getState('proposalsPanel').value).toBe(1);
    expect(await editor.execute('proposalAccept', { proposalId: first })).toEqual({
      ok: true,
      status: 'executed',
    });
    expect(workbook.cell(0, 6, 4).input).toBe('12');
    expect(failure(await editor.execute('proposalReject', { proposalId: first }))).toBe(
      'proposal-not-found'
    );

    const second = await stage('42');
    await act(async () => {
      workbook.editCell(0, 6, 4, '99');
      editor.api().refreshProposals();
    });
    expect(failure(await editor.execute('proposalAccept', { proposalId: second }))).toBe(
      'proposal-stale'
    );
    expect(await editor.execute('proposalsPanel', null)).toEqual({ ok: true, status: 'executed' });
    await waitFor(() => editor.view.getByTestId('xlsx-proposal-stale'));
    expect(workbook.listProposals()[0]?.forceBlocked).not.toBe(true);
    expect(editor.view.getByTestId('xlsx-proposal-force').getAttribute('aria-disabled')).toBeNull();
    expect(editor.view.queryByTestId('xlsx-proposal-force-blocked')).toBeNull();
    expect(
      await editor.execute('proposalAccept', { proposalId: second, force: true })
    ).toEqual({ ok: true, status: 'executed' });
    expect(workbook.cell(0, 6, 4).input).toBe('42');
  });

  it('shows a deleted proposal as stale before acceptance and retains it through undo', async () => {
    const editor = await mountCommands();
    const workbook = editor.api().handle;
    let id = '';
    await act(async () => {
      id = workbook.propose('Audit agent', null, [{ sheet: 0, row: 6, col: 4, input: '42' }]).id;
      workbook.applyOps([{ type: 'deleteRows', sheet: 0, at: 6, count: 1 }]);
      editor.api().refreshProposals();
    });
    await editor.execute('proposalsPanel', null);
    await waitFor(() => expect(editor.view.getByTestId('xlsx-proposal-stale').textContent).toContain('E7'));
    expect(editor.view.getByTestId('xlsx-proposal').textContent).toContain('Audit agent');
    expect(workbook.listProposals()[0]?.forceBlocked).toBe(true);
    expect(editor.view.queryByTestId('xlsx-proposal-force')).toBeNull();
    expect(editor.view.getByTestId('xlsx-proposal-force-blocked').textContent).toBe(
      'This proposal cannot be applied after a structural edit.'
    );
    await act(async () => {
      workbook.undo();
      editor.api().refreshProposals();
    });
    await waitFor(() => expect(editor.view.queryByTestId('xlsx-proposal-stale')).toBeNull());
    expect(workbook.listProposals()[0]?.forceBlocked).not.toBe(true);
    expect(editor.view.queryByTestId('xlsx-proposal-force-blocked')).toBeNull();
    expect(await editor.execute('proposalAccept', { proposalId: id })).toEqual({ ok: true, status: 'executed' });
    expect(workbook.cell(0, 6, 4).input).toBe('42');
  });

  it('hides force for an unrewritable formula even when another proposed cell only drifted', async () => {
    const editor = await mountCommands();
    const workbook = editor.api().handle;
    let id = '';
    await act(async () => {
      id = workbook.propose('Audit agent', null, [
        { sheet: 0, row: 6, col: 4, input: '=SUM(Table1[Amount])' },
        { sheet: 0, row: 9, col: 4, input: '42' },
      ]).id;
      workbook.editCell(0, 9, 4, '99');
      workbook.applyOps([{ type: 'insertRows', sheet: 0, at: 0, count: 1 }]);
      editor.api().refreshProposals();
    });
    await editor.execute('proposalsPanel', null);
    await waitFor(() => expect(editor.view.getByTestId('xlsx-proposal-stale').textContent).toContain('E8'));
    expect(workbook.listProposals()[0]?.forceBlocked).toBe(true);
    expect(editor.view.queryByTestId('xlsx-proposal-force')).toBeNull();
    expect(editor.view.getByTestId('xlsx-proposal-force-blocked').textContent).toBe(
      'This proposal cannot be applied after a structural edit.'
    );
    expect(failure(await editor.execute('proposalAccept', { proposalId: id, force: true }))).toBe(
      'proposal-stale'
    );
    expect(workbook.listProposals().map((proposal) => proposal.id)).toContain(id);
    expect(workbook.cell(0, 10, 4).input).toBe('99');
  });

  it('places host chrome by the toolbar and showToolbar props', async () => {
    const host = (
      <EditorToolbar mode="commands">
        <EditorToolbar.Toolbar>
          <ToolbarCommandButton id="undo" />
        </EditorToolbar.Toolbar>
        <EditorToolbar.FormulaBar />
      </EditorToolbar>
    );
    const editor = await mountCommands({ toolbar: host, readOnly: true });
    expect(editor.view.getByTestId('xlsx-toolbar')).toBeDefined();
    expect(editor.view.getByTestId('xlsx-formula-input')).toBeDefined();
    expect(editor.view.queryByTestId('xlsx-save')).toBeNull();
    await editor.execute('proposalsPanel', { open: true });
    const workspace = editor.view.getByTestId('xlsx-workspace');
    const panel = editor.view.getByTestId('xlsx-proposals-panel');
    expect(workspace.contains(panel)).toBe(true);
    expect(panel.style.top).toBe('4px');
    expect(
      editor.view.getByTestId('xlsx-toolbar').compareDocumentPosition(workspace) &
        Node.DOCUMENT_POSITION_FOLLOWING
    ).toBeTruthy();

    await act(async () => editor.rerender({ toolbar: null, readOnly: false }));
    expect(editor.view.queryByTestId('xlsx-toolbar')).toBeNull();
    expect(workspace.contains(editor.view.getByTestId('xlsx-proposals-panel'))).toBe(true);

    const bar = (
      <EditorToolbar mode="commands">
        <EditorToolbar.Toolbar>
          <ToolbarCommandButton id="bold" />
        </EditorToolbar.Toolbar>
      </EditorToolbar>
    );
    await act(async () => editor.rerender({ toolbar: bar, readOnly: false }));
    expect(editor.view.queryByTestId('xlsx-formula-input')).toBeNull();
    await act(async () => editor.rerender({ toolbar: bar, showToolbar: false, readOnly: false }));
    expect(editor.view.queryByTestId('xlsx-toolbar')).toBeNull();
    await act(async () => editor.rerender({ toolbar: undefined, readOnly: true }));
    expect(editor.view.queryByTestId('xlsx-toolbar')).toBeNull();
    await act(async () => editor.rerender({ toolbar: undefined, readOnly: false }));
    expect(editor.view.getByTestId('xlsx-save')).toBeDefined();
  });
});

describe('XlsxEditor edit batches', () => {
  async function mountApi(props: { readOnly?: boolean; onChange?: () => void } = {}) {
    let api: XlsxEditorApi | undefined;
    const file = plain.bytes.slice();
    const onReady = (ready: XlsxEditorApi) => {
      api = ready;
    };
    const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={file} onReady={onReady} {...props} />);
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    return { view, api: api!, file, onReady };
  }

  async function settled<T>(call: () => Promise<T>): Promise<T> {
    let value!: T;
    await act(async () => {
      value = await call();
    });
    return value;
  }

  const failed = (call: () => Promise<unknown>) =>
    settled(() => call().then(() => null, (error: unknown) => error));

  const setB3 = (expectVersion: string, value: string) => ({
    expectVersion,
    steps: [
      {
        op: 'setCellInputs' as const,
        target: { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'B3' } },
        inputs: [[value]],
      },
    ],
  });

  function deferredClipboard() {
    let resolve!: (text: string) => void;
    const text = new Promise<string>((done) => {
      resolve = done;
    });
    const original = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { readText: () => text },
    });
    return {
      resolve,
      restore: () => {
        if (original) Object.defineProperty(navigator, 'clipboard', original);
        else Reflect.deleteProperty(navigator, 'clipboard');
      },
    };
  }

  it('commits drafts first, applies against the caller version and notifies once', async () => {
    let changes = 0;
    const { view, api } = await mountApi({ onChange: () => changes++ });
    const target = { row: 2, col: 1 };
    await act(async () => {
      api.selectCells(0, selectionAt(target));
    });
    fireEvent.doubleClick(view.getByTestId('xlsx-scroll'), pointAt(plain, target));
    fireEvent.change(await waitFor(() => view.getByTestId('xlsx-cell-editor')), {
      target: { value: '321' },
    });
    const beforeDraft = api.handle.version();

    const version = await settled(() => api.version());
    expect(api.handle.cell(0, 2, 1).input).toBe('321');
    expect(version).not.toBe(beforeDraft);
    expect(changes).toBe(1);

    const stale = await settled(() => api.applyEdits(setB3(beforeDraft, '999')));
    expect(stale.ok).toBe(false);
    if (!stale.ok) expect(stale.failure.code).toBe('stale-version');
    expect(changes).toBe(1);

    const result = await settled(() => api.applyEdits(setB3(version, '4242')));
    expect(result.ok && result.applied).toBe(true);
    expect(changes).toBe(2);
    await waitFor(() =>
      expect((view.getByTestId('xlsx-formula-input') as HTMLInputElement).value).toBe('4242')
    );
    const read = await settled(() =>
      api.readCells({ ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'D3' } }] })
    );
    expect(read.ok && read.ranges[0].cells[0][0].displayText).toBe('4299');
  });

  it('waits for a paste in flight before applying', async () => {
    const clipboard = deferredClipboard();
    try {
      const { view, api } = await mountApi();
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 2, col: 1 }));
      });
      const version = await settled(() => api.version());
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      const result = await settled(async () => {
        const pending = api.applyEdits(setB3(version, 'batch'));
        clipboard.resolve('pasted');
        return pending;
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.failure.code).toBe('stale-version');
      expect(api.handle.cell(0, 2, 1).input).toBe('pasted');
    } finally {
      clipboard.restore();
    }
  });

  it('keeps an unsettled paste on the workbook it was accepted for', async () => {
    const clipboard = deferredClipboard();
    try {
      const { view, api } = await mountApi();
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 2, col: 1 }));
      });
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      const outcome = api.version().then(
        () => null,
        (error: unknown) => error
      );
      let current = api;
      view.rerender(
        <XlsxEditor legacySyncApi experimentalWorkerOpen={false}
          file={plain.bytes.slice()}
          onReady={(ready) => {
            current = ready;
          }}
        />
      );
      await waitFor(() => expect(current).not.toBe(api));
      const applied = await settled(() =>
        current.applyEdits(setB3(current.handle.version(), 'fresh'))
      );
      expect(applied).toMatchObject({ ok: true, applied: true });
      expect(current.handle.cell(0, 2, 1).input).toBe('fresh');
      expect(await settled(() => outcome)).toMatchObject({ code: 'document-replaced' });
    } finally {
      clipboard.restore();
    }
  });

  it('rejects when the workbook is replaced while input is flushing', async () => {
    const clipboard = deferredClipboard();
    try {
      const { view, api, onReady } = await mountApi();
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 2, col: 1 }));
      });
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      const outcome = api.applyEdits(setB3(api.handle.version(), 'batch')).then(
        () => null,
        (error: unknown) => error
      );
      view.rerender(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()} onReady={onReady} />);
      await act(async () => clipboard.resolve('late'));
      const replaced = await settled(() => outcome);
      expect(replaced).toBeInstanceOf(XlsxCommandAdmissionError);
      expect(replaced).toMatchObject({
        name: 'XlsxCommandAdmissionError',
        code: 'document-replaced',
        message: 'The workbook was replaced',
      });
    } finally {
      clipboard.restore();
    }
  });

  it('lands a batch issued behind a queued paste after it and before later input', async () => {
    const clipboard = deferredClipboard();
    try {
      const { view, api } = await mountApi();
      const surface = view.getByTestId('xlsx-scroll');
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 2, col: 1 }));
      });
      const before = await settled(() => api.version());
      fireEvent.keyDown(surface, { key: 'v', ctrlKey: true });
      const pending = api.applyEdits(setB3(before, 'batch'));
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 3, col: 1 }));
      });
      fireEvent.keyDown(surface, { key: 'F2' });
      const cell = view.getByTestId('xlsx-cell-editor');
      fireEvent.change(cell, { target: { value: 'Later' } });
      fireEvent.keyDown(cell, { key: 'Enter' });
      expect(api.handle.cell(0, 3, 1).input).toBe('200');

      const result = await settled(async () => {
        clipboard.resolve('pasted');
        return pending;
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.failure.code).toBe('stale-version');
        expect([before, api.handle.version()]).not.toContain(result.version);
      }
      expect(api.handle.cell(0, 2, 1).input).toBe('pasted');
      expect(api.handle.cell(0, 3, 1).input).toBe('Later');
    } finally {
      clipboard.restore();
    }
  });

  it('fails reads and batches behind a refused entry with input-failed, as commands', async () => {
    const { view, api } = await mountApi();
    const target = { row: 2, col: 1 };
    await act(async () => {
      api.selectCells(0, selectionAt(target));
    });
    const version = await settled(() => api.version());
    fireEvent.doubleClick(view.getByTestId('xlsx-scroll'), pointAt(plain, target));
    const editor = await waitFor(() => view.getByTestId('xlsx-cell-editor'));
    fireEvent.change(editor, { target: { value: 'Refused' } });
    const editCell = api.handle.editCell;
    api.handle.editCell = () => {
      throw new Error('cell is locked');
    };
    try {
      fireEvent.keyDown(editor, { key: 'Enter' });
    } finally {
      api.handle.editCell = editCell;
    }

    const command = await settled(() => api.commands.execute('bold', null));
    expect(command.ok ? null : command.failure.code).toBe('input-failed');
    for (const call of [
      () => api.applyEdits(setB3(version, 'batch')),
      () => api.readCells({ ranges: [] }),
    ]) {
      const refused = await failed(call);
      expect(refused).toBeInstanceOf(XlsxCommandAdmissionError);
      expect((refused as XlsxCommandAdmissionError).code).toBe('input-failed');
    }
    expect(api.handle.cell(0, 2, 1).input).toBe('100');

    fireEvent.keyDown(view.getByTestId('xlsx-cell-editor'), { key: 'Enter' });
    expect(api.handle.cell(0, 2, 1).input).toBe('Refused');
    const current = await settled(() => api.version());
    const applied = await settled(() => api.applyEdits(setB3(current, 'batch')));
    expect(applied.ok && applied.applied).toBe(true);
  });

  function deferredClipboardWrite() {
    let settle!: () => void;
    const written = new Promise<void>((done) => {
      settle = done;
    });
    const original = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { writeText: () => written },
    });
    return {
      settle,
      restore: () => {
        if (original) Object.defineProperty(navigator, 'clipboard', original);
        else Reflect.deleteProperty(navigator, 'clipboard');
      },
    };
  }

  it('clears the accepted cut target once the clipboard write settles', async () => {
    const clipboard = deferredClipboardWrite();
    try {
      const { view, api } = await mountApi();
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 2, col: 1 }));
      });
      const version = await settled(() => api.version());
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true });
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 3, col: 1 }));
      });
      const result = await settled(async () => {
        const pending = api.applyEdits(setB3(version, 'batch'));
        clipboard.settle();
        return pending;
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.failure.code).toBe('stale-version');
      expect(api.handle.cell(0, 2, 1).input).toBe('');
      expect(api.handle.cell(0, 3, 1).input).toBe('200');
    } finally {
      clipboard.restore();
    }
  });

  it('drops a deferred cut when the workbook is replaced or turns read-only', async () => {
    for (const change of ['replace', 'readOnly'] as const) {
      const clipboard = deferredClipboardWrite();
      try {
        const { view, api, file, onReady } = await mountApi();
        await act(async () => {
          api.selectCells(0, selectionAt({ row: 2, col: 1 }));
        });
        fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true });
        let current = api;
        const swap = (ready: XlsxEditorApi) => {
          current = ready;
        };
        if (change === 'replace') {
          view.rerender(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()} onReady={swap} />);
          await waitFor(() => expect(current).not.toBe(api));
        } else {
          view.rerender(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={file} onReady={onReady} readOnly />);
        }
        await act(async () => clipboard.settle());
        expect(current.handle.cell(0, 2, 1).input).toBe('100');
      } finally {
        clipboard.restore();
        cleanup();
      }
    }
  });

  it('refuses a write when the editor turns read-only while input flushes', async () => {
    const clipboard = deferredClipboard();
    try {
      const { view, api, file, onReady } = await mountApi();
      await act(async () => {
        api.selectCells(0, selectionAt({ row: 2, col: 1 }));
      });
      const version = await settled(() => api.version());
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      const outcome = api.applyEdits(setB3(version, 'batch'));
      view.rerender(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={file} onReady={onReady} readOnly />);
      await act(async () => clipboard.resolve('late'));
      const result = await settled(() => outcome);
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.failure.code).toBe('read-only');
      expect(api.handle.cell(0, 2, 1).input).toBe('100');
    } finally {
      clipboard.restore();
    }
  });

  it('refuses writes while read-only and ends a composition before a batch', async () => {
    const readOnly = await mountApi({ readOnly: true });
    const readVersion = await settled(() => readOnly.api.version());
    const refusals = [
      await settled(() => readOnly.api.validateEdits(setB3(readVersion, '1'))),
      await settled(() => readOnly.api.applyEdits(setB3(readVersion, '1'))),
    ];
    for (const refused of refusals) {
      expect(refused.ok).toBe(false);
      if (!refused.ok) expect(refused.failure.code).toBe('read-only');
    }
    const found = await settled(() => readOnly.api.findText({ text: 'Quarterly' }));
    expect(found.ok).toBe(true);
    cleanup();

    const { view, api } = await mountApi();
    const target = { row: 2, col: 1 };
    await act(async () => {
      api.selectCells(0, selectionAt(target));
    });
    fireEvent.doubleClick(view.getByTestId('xlsx-scroll'), pointAt(plain, target));
    const editor = await waitFor(() => view.getByTestId('xlsx-cell-editor'));
    act(() => editor.focus());
    fireEvent.compositionStart(editor);
    fireEvent.change(editor, { target: { value: '5' } });
    const before = api.handle.version();
    let waited = true;
    let composing!: Promise<XlsxEditResult>;
    await act(async () => {
      composing = api.applyEdits(setB3(before, '1'));
      void composing.then(() => (waited = false));
      await new Promise((resolve) => setTimeout(resolve, 0));
    });
    expect(waited).toBe(true);
    expect(api.handle.cell(0, 2, 1).input).toBe('100');
    fireEvent.change(editor, { target: { value: '55' } });
    const composed = await settled(async () => {
      fireEvent.compositionEnd(editor);
      return composing;
    });
    expect(composed.ok).toBe(false);
    if (!composed.ok) expect(composed.failure.code).toBe('stale-version');
    expect(api.handle.cell(0, 2, 1).input).toBe('55');
    const version = await settled(() => api.version());
    const validated = await settled(() => api.validateEdits(setB3(version, '66')));
    expect(validated.ok && validated.wouldApply).toBe(true);
  });
});

describe('XlsxEditor host controls', () => {
  it('returns the saved bytes from the sync save() under legacySyncApi alone', async () => {
    const worker = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
    Object.defineProperty(globalThis, 'Worker', { configurable: true, writable: true, value: class {} });
    const workerOpen = spyOn(editableWorkbookSessionBackend, 'open').mockImplementation(() => {
      throw new Error('opened the worker editor');
    });
    try {
      let api: XlsxEditorApi | undefined;
      render(<XlsxEditor legacySyncApi file={plain.bytes.slice()} onReady={(ready) => { api = ready; }} />);
      await waitFor(() => expect(api).toBeDefined(), MOUNTED);
      expect(workerOpen).not.toHaveBeenCalled();
      expect(api!.save().byteLength).toBeGreaterThan(0);
    } finally {
      workerOpen.mockRestore();
      if (worker) Object.defineProperty(globalThis, 'Worker', worker);
      else Reflect.deleteProperty(globalThis, 'Worker');
    }
  });

  it('lets a host flush and serialize inside the save callback without re-entering it', async () => {
    let api!: XlsxEditorApi;
    let requests = 0;
    let saved!: Uint8Array;
    const builtIn: Uint8Array[] = [];
    const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()} onReady={(ready) => { api = ready; }}
      onSave={(bytes) => builtIn.push(bytes)} onSaveRequest={async () => {
        requests++;
        await api.flushPendingInput();
        saved = api.save();
      }} />);
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value: 'Host-owned save' } });
    let result!: XlsxCommandResult;
    await act(async () => { result = await api.commands.execute('save', null); });
    expect(result).toEqual({ ok: true, status: 'noop' });
    expect(requests).toBe(1);
    expect(builtIn).toHaveLength(0);
    const reopened = openWorkbook(saved);
    try { expect(reopened.cell(0, 0, 0).input).toBe('Host-owned save'); } finally { reopened.dispose(); }
  });

  it('allows another save after the host callback rejects', async () => {
    let api!: XlsxEditorApi;
    let requests = 0;
    const saved: Uint8Array[] = [];
    render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()} onReady={(ready) => { api = ready; }}
      onSave={(bytes) => saved.push(bytes)} onSaveRequest={async () => {
        if (++requests === 1) throw new Error('Host save failed');
        return true;
      }} />);
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    await act(async () => {
      const failed = await api.commands.execute('save', null);
      expect(failed).toMatchObject({ ok: false, failure: { code: 'command-failed' } });
      expect(await api.commands.execute('save', null)).toEqual({ ok: true, status: 'executed' });
    });
    expect(saved).toHaveLength(1);
    expect(requests).toBe(2);
  });

  for (const decision of [true, false, undefined] as const) {
    it(`awaits one save request returning ${String(decision)} before exporting`, async () => {
      let api: XlsxEditorApi | undefined;
      let requests = 0;
      let release!: (value: boolean | void) => void;
      const pending = new Promise<boolean | void>((resolve) => { release = resolve; });
      const saved: Uint8Array[] = [];
      const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()} onReady={(ready) => { api = ready; }}
        onSaveRequest={() => { requests++; return pending; }} onSave={(bytes) => saved.push(bytes)} />);
      await waitFor(() => expect(api).toBeDefined(), MOUNTED);
      fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value: 'Host save draft' } });
      fireEvent.keyDown(view.getByTestId('xlsx-formula-input'), { key: 's', ctrlKey: true });
      fireEvent.keyDown(view.getByTestId('xlsx-formula-input'), { key: 's', ctrlKey: true });
      await waitFor(() => expect(requests).toBe(1));
      expect(saved).toHaveLength(0);
      await act(async () => { release(decision); });
      expect(saved).toHaveLength(decision === true ? 1 : 0);
      if (decision === true) {
        const reopened = openWorkbook(saved[0]);
        try { expect(reopened.cell(0, 0, 0).input).toBe('Host save draft'); } finally { reopened.dispose(); }
      }
    });
  }

  for (const source of ['cell', 'formula'] as const) {
    it(`waits for ${source} composition before making the draft authoritative`, async () => {
      let api: XlsxEditorApi | undefined;
      const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()} onReady={(ready) => { api = ready; }} />);
      await waitFor(() => expect(api).toBeDefined(), MOUNTED);
      const target = { row: 2, col: 0 };
      await act(async () => { api!.selectCells(0, selectionAt(target)); });
      if (source === 'cell') fireEvent.doubleClick(view.getByTestId('xlsx-scroll'), pointAt(plain, target));
      const input = view.getByTestId(source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input');
      fireEvent.compositionStart(input);
      fireEvent.change(input, { target: { value: '日本' } });
      let finished = false;
      const flushed = api!.flushPendingInput().then(() => { finished = true; });
      await Promise.resolve();
      expect(finished).toBe(false);
      expect(() => api!.save()).toThrow(XlsxSaveRefusedError);
      fireEvent.keyDown(input, { key: 'Enter', isComposing: true });
      expect(api!.handle.cell(0, target.row, target.col).input).toBe('Line item 1');
      fireEvent.change(input, { target: { value: '日本語' } });
      await act(async () => { fireEvent.compositionEnd(input, { data: '日本語' }); await flushed; });
      expect(api!.handle.cell(0, target.row, target.col).input).toBe('日本語');
      let bytes!: Uint8Array;
      await act(async () => { bytes = api!.save(); });
      const reopened = openWorkbook(bytes);
      try { expect(reopened.cell(0, target.row, target.col).input).toBe('日本語'); } finally { reopened.dispose(); }
    });
  }

  it('flushes an accepted asynchronous paste', async () => {
    let api: XlsxEditorApi | undefined;
    let release!: (value: string) => void;
    const clipboard = navigator.clipboard.readText;
    navigator.clipboard.readText = () => new Promise<string>((resolve) => { release = resolve; });
    try {
      const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()} onReady={(ready) => { api = ready; }} />);
      await waitFor(() => expect(api).toBeDefined(), MOUNTED);
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      let finished = false;
      const flushed = api!.flushPendingInput().then(() => { finished = true; });
      await Promise.resolve();
      expect(finished).toBe(false);
      await act(async () => { release('Pasted\tPair'); await flushed; });
      expect(api!.handle.cell(0, 0, 0).input).toBe('Pasted');
      expect(api!.handle.cell(0, 0, 1).input).toBe('Pair');
    } finally { navigator.clipboard.readText = clipboard; }
  });

  it('rejects stale flushes and discards a save waiting on the replaced workbook', async () => {
    const opened: XlsxEditorApi[] = [];
    let requests = 0;
    let release!: (value: boolean) => void;
    const pending = new Promise<boolean>((resolve) => { release = resolve; });
    const saved: Uint8Array[] = [];
    const props = { onReady: (ready: XlsxEditorApi) => { opened.push(ready); },
      onSaveRequest: () => { requests++; return pending; }, onSave: (bytes: Uint8Array) => { saved.push(bytes); } };
    const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} {...props} file={plain.bytes.slice()} />);
    await waitFor(() => expect(opened).toHaveLength(1), MOUNTED);
    fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 's', ctrlKey: true });
    await waitFor(() => expect(requests).toBe(1));
    fireEvent.compositionStart(view.getByTestId('xlsx-formula-input'));
    const flushed = opened[0].flushPendingInput();
    void flushed.catch(() => {});
    await act(async () => { view.rerender(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} {...props} file={plain.bytes.slice()} />); });
    await expect(flushed).rejects.toThrow(XlsxCommandAdmissionError);
    await waitFor(() => expect(opened).toHaveLength(2));
    await act(async () => { release(true); });
    expect(saved).toHaveLength(0);
    expect(() => opened[0].save()).toThrow(XlsxCommandAdmissionError);
    expect(opened[0].getPositionAtPoint(1, 1)).toBeNull();
  });

  it('queries painted cells without moving selection and returns null over charts', async () => {
    let api: XlsxEditorApi | undefined;
    const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={charted.bytes.slice()} onReady={(ready) => { api = ready; }} />);
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    const canvas = view.container.querySelector('canvas')!;
    canvas.getBoundingClientRect = () => new DOMRect(40, 60, VIEWPORT.width, VIEWPORT.height);
    const surface = view.getByTestId('xlsx-scroll');
    const live = api!.handle.displayList({ x: surface.scrollLeft, y: surface.scrollTop, ...VIEWPORT });
    const target = cellRect(live.grid!, 3, 0)!;
    const point = { clientX: target.x + target.w / 2, clientY: target.y + target.h / 2 };
    await waitFor(() => expect(api!.getPositionAtPoint(40 + point.clientX, 60 + point.clientY)).toEqual({ sheet: 0, row: 3, col: 0 }));
    for (const scale of [0.75, 1.5]) {
      canvas.getBoundingClientRect = () => new DOMRect(40, 60, VIEWPORT.width * scale, VIEWPORT.height * scale);
      expect(api!.getPositionAtPoint(40 + point.clientX * scale, 60 + point.clientY * scale)).toEqual({ sheet: 0, row: 3, col: 0 });
    }
    canvas.getBoundingClientRect = () => new DOMRect(40, 60, VIEWPORT.width, VIEWPORT.height);
    const name = (view.getByTestId('xlsx-name-box') as HTMLInputElement).value;
    const focused = document.activeElement;
    const chart = charted.charts![0];
    expect(api!.getPositionAtPoint(40 + chart.rect.x + chart.rect.w / 2, 60 + chart.rect.y + chart.rect.h / 2)).toBeNull();
    expect(api!.getPositionAtPoint(39, 70)).toBeNull();
    expect(api!.getPositionAtPoint(NaN, 70)).toBeNull();
    expect((view.getByTestId('xlsx-name-box') as HTMLInputElement).value).toBe(name);
    expect(document.activeElement).toBe(focused);
  });
});

describe('clipboard access failure recovery', () => {
  for (const unavailable of [false, true]) {
    it(`keeps saving available when clipboard reads are ${unavailable ? 'unavailable' : 'denied'}`, async () => {
      const read = navigator.clipboard.readText;
      navigator.clipboard.readText = unavailable ? undefined as unknown as typeof read
        : () => Promise.reject(new Error('Clipboard access denied'));
      try {
        let api: XlsxEditorApi | undefined;
        const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()} onReady={(ready) => { api = ready; }} />);
        await waitFor(() => expect(api).toBeDefined(), MOUNTED);
        fireEvent.change(view.getByTestId('xlsx-formula-input'), { target: { value: 'Keep my edits' } });
        fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
        let bytes!: Uint8Array;
        await act(async () => { await api!.flushPendingInput(); bytes = api!.save(); });
        const reopened = openWorkbook(bytes);
        try { expect(reopened.cell(0, 0, 0).input).toBe('Keep my edits'); } finally { reopened.dispose(); }
      } finally { navigator.clipboard.readText = read; }
    });
  }

  it('keeps the original cell when a cut cannot write to the clipboard', async () => {
    const write = navigator.clipboard.writeText;
    navigator.clipboard.writeText = () => Promise.reject(new Error('Clipboard access denied'));
    try {
      let api: XlsxEditorApi | undefined;
      const view = render(<XlsxEditor legacySyncApi experimentalWorkerOpen={false} file={plain.bytes.slice()} onReady={(ready) => { api = ready; }} />);
      await waitFor(() => expect(api).toBeDefined(), MOUNTED);
      const before = api!.handle.cell(0, 0, 0).input;
      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'x', ctrlKey: true });
      await act(async () => { await api!.flushPendingInput(); });
      expect(api!.handle.cell(0, 0, 0).input).toBe(before);
    } finally { navigator.clipboard.writeText = write; }
  });
});

describe('async editor host props', () => {
  it('opens an ArrayBuffer once across rerenders', async () => {
    const file = plain.bytes.slice().buffer;
    const ready: unknown[] = [];
    const onReady = (api: unknown) => { ready.push(api); };
    const view = render(<XlsxEditor experimentalWorkerOpen={false} file={file} onReady={onReady} />);
    await waitFor(() => expect(ready).toHaveLength(1), MOUNTED);
    view.rerender(<XlsxEditor experimentalWorkerOpen={false} file={file} onReady={onReady} className="updated" />);
    await act(async () => {});
    expect(ready).toHaveLength(1);
    const api = ready[0] as import('./worker/createAsyncEditorApi').XlsxAsyncEditorApi;
    expect(await api.readCells({ ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }] })).toMatchObject({ ok: true });
  });

  it('reports remote document changes to onChange and onDocumentChange once, like the worker route', async () => {
    type Api = import('./worker/createAsyncEditorApi').XlsxAsyncEditorApi;
    let source!: Api;
    let target!: Api;
    const sourceVersions: string[] = [];
    const targetVersions: string[] = [];
    let sourceChanges = 0;
    let targetChanges = 0;
    const first = render(<XlsxEditor experimentalWorkerOpen={false} file={plain.bytes}
      collaboration={{ clientId: 101 }} onReady={(api) => { source = api; }}
      onChange={() => sourceChanges++} onDocumentChange={({ version }) => sourceVersions.push(version)} />);
    const second = render(<XlsxEditor experimentalWorkerOpen={false} file={plain.bytes}
      collaboration={{ clientId: 202 }} onReady={(api) => { target = api; }}
      onChange={() => targetChanges++} onDocumentChange={({ version }) => targetVersions.push(version)} />);
    await waitFor(() => { expect(source).toBeDefined(); expect(target).toBeDefined(); }, MOUNTED);
    const before = await target.version();
    const vector = await target.collaboration!.encodeStateVector();
    await act(async () => {
      await source.applyEdits({ expectVersion: await source.version(), steps: [{ op: 'setCellInputs',
        target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B3' } }, inputs: [['remote edit']] }] });
    });
    const update = await source.collaboration!.encodeStateAsUpdate(vector);
    await act(async () => { await target.collaboration!.applyUpdate(update); });
    const version = await target.version();
    expect(version).not.toBe(before);
    expect(targetVersions).toEqual([version]);
    expect(sourceVersions).toEqual([await source.version()]);
    expect(sourceChanges).toBe(1);
    expect(targetChanges).toBe(1);
    await act(async () => { await target.collaboration!.applyUpdate(update); });
    expect(targetVersions).toEqual([version]);
    expect(targetChanges).toBe(1);
    await act(async () => {
      await target.applyEdits({ expectVersion: version, steps: [{ op: 'setCellInputs',
        target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B4' } }, inputs: [['local edit']] }] });
    });
    expect(targetVersions).toEqual([version, await target.version()]);
    expect(targetChanges).toBe(2);
    first.unmount();
    second.unmount();
  });

  it('reports a whole-document remote update into a pristine replica to onChange once', async () => {
    type Api = import('./worker/createAsyncEditorApi').XlsxAsyncEditorApi;
    let source!: Api;
    let target!: Api;
    const targetVersions: string[] = [];
    let targetChanges = 0;
    const first = render(<XlsxEditor experimentalWorkerOpen={false} file={plain.bytes}
      collaboration={{ clientId: 303 }} onReady={(api) => { source = api; }} />);
    const second = render(<XlsxEditor experimentalWorkerOpen={false} file={plain.bytes}
      collaboration={{ clientId: 404 }} onReady={(api) => { target = api; }}
      onChange={() => targetChanges++} onDocumentChange={({ version }) => targetVersions.push(version)} />);
    await waitFor(() => { expect(source).toBeDefined(); expect(target).toBeDefined(); }, MOUNTED);
    const before = await target.version();
    await act(async () => {
      await source.applyEdits({ expectVersion: await source.version(), steps: [{ op: 'setCellInputs',
        target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B3' } }, inputs: [['remote edit']] }] });
    });
    const update = await source.collaboration!.encodeStateAsUpdate();
    await act(async () => { await target.collaboration!.applyUpdate(update); });
    const version = await target.version();
    expect(version).not.toBe(before);
    expect(targetVersions).toEqual([version]);
    expect(targetChanges).toBe(1);
    await act(async () => {
      await target.applyEdits({ expectVersion: version, steps: [{ op: 'setCellInputs',
        target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B4' } }, inputs: [['local edit']] }] });
    });
    expect(targetVersions).toEqual([version, await target.version()]);
    expect(targetChanges).toBe(2);
    first.unmount();
    second.unmount();
  });

  it('reports the committed version with onChange', async () => {
    let api!: import('./worker/createAsyncEditorApi').XlsxAsyncEditorApi;
    const versions: string[] = [];
    let changes = 0;
    const view = render(<XlsxEditor experimentalWorkerOpen={false} file={plain.bytes}
      onReady={(value) => { api = value; }} onChange={() => changes++}
      onDocumentChange={(change) => versions.push(change.version)} />);
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    await act(async () => {
      await api.applyEdits({ expectVersion: await api.version(), steps: [{ op: 'setCellInputs',
        target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'B3' } }, inputs: [['host edit']] }] });
    });
    expect(changes).toBe(1);
    expect(versions).toEqual([await api.version()]);
    view.unmount();
  });

  for (const [initialZoom, scale] of [[2, 2], [0, 0.1], [99, 5], [NaN, 1]] as const) {
    it(`uses initial zoom ${initialZoom} and merges root style`, async () => {
      let api!: import('./worker/createAsyncEditorApi').XlsxAsyncEditorApi;
      const onReady = (value: typeof api) => { api = value; };
      const element = (zoom: number) => <XlsxEditor experimentalWorkerOpen={false} file={plain.bytes}
        initialZoom={zoom} style={{ height: 345, background: 'red' }} onReady={onReady} />;
      const view = render(element(initialZoom));
      await waitFor(() => expect(api).toBeDefined(), MOUNTED);
      expect(api.commands.getState('zoom').value).toBe(scale);
      expect((view.container.firstElementChild as HTMLElement).style.height).toBe('345px');
      view.rerender(element(3));
      expect(api.commands.getState('zoom').value).toBe(scale);
    });
  }

  for (const [callback, downloadOnSave, downloads] of [
    [true, undefined, 0], [false, undefined, 1], [true, true, 1], [false, false, 0],
  ] as const) {
    it(`delivers command saves with callback=${callback} download=${downloadOnSave}`, async () => {
      let api!: import('./worker/createAsyncEditorApi').XlsxAsyncEditorApi;
      const saves: Uint8Array[] = [];
      const click = spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
      try {
        render(<XlsxEditor experimentalWorkerOpen={false} file={plain.bytes} downloadOnSave={downloadOnSave}
          onSave={callback ? (bytes) => saves.push(bytes) : undefined} onReady={(value) => { api = value; }} />);
        await waitFor(() => expect(api).toBeDefined(), MOUNTED);
        await act(async () => { expect(await api.commands.execute('save', null)).toMatchObject({ ok: true }); });
        expect(saves).toHaveLength(callback ? 1 : 0);
        expect(click).toHaveBeenCalledTimes(downloads);
      } finally { click.mockRestore(); }
    });
  }
});

it('in-thread renders the supplied loading indicator while opening a file', async () => {
  let ready = false;
  const view = render(<XlsxEditor experimentalWorkerOpen={false} file={plain.bytes}
    loadingIndicator={<span>Opening in thread</span>} onReady={() => { ready = true; }} />);
  expect(view.queryByText('Opening in thread')).not.toBeNull();
  await waitFor(() => expect(ready).toBe(true), MOUNTED);
  expect(view.queryByText('Opening in thread')).toBeNull();
});

it('pins the editor root to ltr', async () => {
  const view = render(<XlsxEditor experimentalWorkerOpen={false} file={plain.bytes} onReady={() => {}} />);
  await waitFor(() => expect(view.getByTestId('xlsx-name-box')).toBeDefined(), MOUNTED);
  expect((view.container.firstElementChild as HTMLElement).getAttribute('dir')).toBe('ltr');
});

it('async host save calls the latest onSave once without a request or download', async () => {
  let api!: import('./worker/createAsyncEditorApi').XlsxAsyncEditorApi;
  const ready = (value: typeof api) => { api = value; };
  const first: Uint8Array[] = [];
  const latest: Uint8Array[] = [];
  const click = spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
  const request = () => { throw new Error('Host save must bypass the request'); };
  try {
    const view = render(<XlsxEditor experimentalWorkerOpen={false} file={plain.bytes} onReady={ready}
      onSave={(bytes) => first.push(bytes)} onSaveRequest={request} downloadOnSave />);
    await waitFor(() => expect(api).toBeDefined(), MOUNTED);
    view.rerender(<XlsxEditor experimentalWorkerOpen={false} file={plain.bytes} onReady={ready}
      onSave={(bytes) => latest.push(bytes)} onSaveRequest={request} downloadOnSave />);
    let bytes!: Uint8Array;
    await act(async () => { bytes = await api.save(); });
    expect(first).toHaveLength(0);
    expect(latest).toEqual([bytes]);
    expect(latest[0]).toBe(bytes);
    expect(click).not.toHaveBeenCalled();
  } finally { click.mockRestore(); }
});
