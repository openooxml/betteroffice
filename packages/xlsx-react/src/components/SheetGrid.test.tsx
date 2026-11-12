import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { DisplayList, Selection } from '@betteroffice/xlsx';
import { XLSX_MAX_COLUMNS, XLSX_MAX_ROWS } from '@betteroffice/xlsx/collaboration';
import { fireEvent } from '@testing-library/react';
import {
  HEADER_RESIZE_ZONE_PX,
  MAX_COLUMN_WIDTH_PX,
  MAX_ROW_HEIGHT_PX,
  SheetGrid,
  clampTrackSize,
  columnWidthFromPx,
  headerTracks,
  reorderRun,
  resizeGestureAt,
  resizePreview,
  resizeRun,
  rowHeightFromPx,
  type SheetGridHeaderHandlers,
} from './SheetGrid';
import { columnLetter } from '../columnLetter';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const frame = (grid: DisplayList['grid']): DisplayList => ({ width: 320, height: 120, commands: [], grid });
// the resize strips are band children too, so a label is a band child that
// carries a track attribute.
const labels = (root: Element, testId: string) =>
  Array.from(root.querySelectorAll(`[data-testid="${testId}"] > div[data-col], [data-testid="${testId}"] > div[data-row]`),
    (node) => node.textContent);

const geometry = { startRow: 0, startCol: 0,
  colOffsets: [0, 80, 160, 240], rowOffsets: [0, 20, 40] };
// each test owns a fresh record of what the surface was told.
const mount = (handlers: SheetGridHeaderHandlers, zoom = 1,
  grid: DisplayList['grid'] = geometry, selection: Selection | null = null,
  sheet?: number | null) => {
  presses.length = 0;
  resizes.length = 0;
  reorders.length = 0;
  selects.length = 0;
  return render(<SheetGrid frame={frame(grid)} zoom={zoom} showSheetHeaders headerHandlers={handlers}
    selection={selection} sheet={sheet}><div /></SheetGrid>);
};
// a whole-track selection on one axis: the only state a label press reorders from.
const wholeColumn = (col: number): Selection => ({
  anchor: { row: 0, col }, focus: { row: XLSX_MAX_ROWS - 1, col },
});
const wholeRow = (row: number): Selection => ({
  anchor: { row, col: 0 }, focus: { row, col: XLSX_MAX_COLUMNS - 1 },
});
// a run of whole tracks: the selection a press moves as one block.
const wholeColumns = (left: number, right: number): Selection => ({
  anchor: { row: 0, col: left }, focus: { row: XLSX_MAX_ROWS - 1, col: right },
});
const wholeRows = (top: number, bottom: number): Selection => ({
  anchor: { row: top, col: 0 }, focus: { row: bottom, col: XLSX_MAX_COLUMNS - 1 },
});
const columnLabel = (root: Element, index: number) =>
  root.querySelector(`[data-testid="xlsx-column-headers"] [data-col="${index}"]`)!;
const rowLabel = (root: Element, index: number) =>
  root.querySelector(`[data-testid="xlsx-row-headers"] [data-row="${index}"]`)!;
const presses: { axis: string; index: number; shift: boolean }[] = [];
const resizes: { axis: string; index: number; size: number; run?: { at: number; count: number } }[] = [];
const reorders: { axis: string; at: number; before: number; count: number }[] = [];
const selects: { axis: string; index: number }[] = [];
const handlers: SheetGridHeaderHandlers = {
  onHeaderMouseDown: (axis, index, event) => presses.push({ axis, index, shift: event.shiftKey }),
  onHeaderResize: (axis, index, size, run) => resizes.push({ axis, index, size, run }),
  onHeaderReorder: (axis, at, before, count) => reorders.push({ axis, at, before, count }),
  onHeaderDragSelect: (axis, index) => selects.push({ axis, index }),
};
// happy-dom has no layout, so the band's rect is the origin and every pointer
// coordinate is read straight off the painted frame geometry.
const stubRect = (element: Element, left = 0, top = 0) => {
  const rect = { left, top, width: 0, height: 0, right: left, bottom: top, x: left, y: top };
  element.getBoundingClientRect = () => rect as DOMRect;
  return rect;
};

test('column labels use absolute sheet indices, not viewport positions', () => {
  const { container } = render(<SheetGrid frame={frame({ startRow: 0, startCol: 26,
    colOffsets: [0, 80, 160], rowOffsets: [0, 20] })} zoom={1} showSheetHeaders><div /></SheetGrid>);
  expect(labels(container, 'xlsx-column-headers')).toEqual(['AA', 'AB']);
});

test('row labels are one-based sheet rows', () => {
  const { container } = render(<SheetGrid frame={frame({ startRow: 9, startCol: 0,
    colOffsets: [0, 80], rowOffsets: [0, 20, 40] })} zoom={1} showSheetHeaders><div /></SheetGrid>);
  expect(labels(container, 'xlsx-row-headers')).toEqual(['10', '11']);
});

test('compact track indices keep their labels, and zero-sized tracks are skipped', () => {
  const { container } = render(<SheetGrid frame={frame({ startRow: 0, startCol: 0,
    colIndices: [0, 2, 5], rowIndices: [0, 3], colOffsets: [0, 80, 80, 160],
    rowOffsets: [0, 20, 40] })} zoom={1} showSheetHeaders><div /></SheetGrid>);
  expect(labels(container, 'xlsx-column-headers')).toEqual(['A', 'F']);
  expect(labels(container, 'xlsx-row-headers')).toEqual(['1', '4']);
});

test('header geometry follows zoom while the band height stays fixed', () => {
  const { container } = render(<SheetGrid frame={frame({ startRow: 0, startCol: 0,
    colOffsets: [0, 80, 160], rowOffsets: [0, 20, 40] })} zoom={2} showSheetHeaders><div /></SheetGrid>);
  const [first, second] = Array.from(
    container.querySelectorAll('[data-testid="xlsx-column-headers"] > div'),
  ) as HTMLElement[];
  expect([first.style.left, first.style.width, first.style.height]).toEqual(['0px', '160px', '24px']);
  expect(second.style.left).toBe('160px');
});

test('a frozen column paints the split where the canvas draws it', () => {
  const { container } = render(<SheetGrid frame={frame({ startRow: 0, startCol: 0,
    frozenCols: 1, dividerX: 80, colOffsets: [0, 80, 160], rowOffsets: [0, 20] })} zoom={1} showSheetHeaders><div /></SheetGrid>);
  const divider = container.querySelector<HTMLElement>(
    '[data-testid="xlsx-column-headers"] [data-testid="xlsx-freeze-divider"]')!;
  expect(divider.getAttribute('data-axis')).toBe('col');
  // the canvas strokes 2px centred on 80, so the band copies 79..81.
  expect([divider.style.left, divider.style.width, divider.style.height]).toEqual(['79px', '2px', '24px']);
  expect(divider.style.background).toBe('#8a8a8a');
  expect(divider.style.pointerEvents).toBe('none');
  expect(container.querySelector('[data-testid="xlsx-row-headers"] [data-testid="xlsx-freeze-divider"]')).toBeNull();
});

test('a frozen row paints the split across the row band', () => {
  const { container } = render(<SheetGrid frame={frame({ startRow: 0, startCol: 0,
    frozenRows: 1, dividerY: 20, colOffsets: [0, 80], rowOffsets: [0, 20, 40] })} zoom={1} showSheetHeaders><div /></SheetGrid>);
  const divider = container.querySelector<HTMLElement>(
    '[data-testid="xlsx-row-headers"] [data-testid="xlsx-freeze-divider"]')!;
  expect(divider.getAttribute('data-axis')).toBe('row');
  expect([divider.style.top, divider.style.width, divider.style.height]).toEqual(['19px', '44px', '2px']);
  expect(divider.style.background).toBe('#8a8a8a');
  expect(divider.style.pointerEvents).toBe('none');
});

test('a frame without a freeze paints no band divider', () => {
  const { container } = mount(handlers);
  expect(container.querySelectorAll('[data-testid="xlsx-freeze-divider"]')).toHaveLength(0);
});

test('a compacted hidden column span draws a doubled separator instead of the gridline', () => {
  // columns 1 and 2 are hidden, so the boundary between 0 and 3 is where they
  // collapse: it gets the pair of hairlines, not the one-pixel divider.
  const { container } = mount(handlers, 1, { startRow: 0, startCol: 0,
    colIndices: [0, 3, 4], colOffsets: [0, 80, 160, 240], rowOffsets: [0, 20],
    hiddenColSpans: [[1, 2]] });
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  const markers = Array.from(band.querySelectorAll<HTMLElement>('[data-testid="xlsx-hidden-divider"]'));
  expect(markers.map((marker) => [marker.getAttribute('data-axis'), marker.style.left])).toEqual([
    ['col', '78px'], ['col', '82px'],
  ]);
  expect(markers.map((marker) => [marker.style.width, marker.style.background])).toEqual([
    ['1px', '#9aa0a6'], ['1px', '#9aa0a6'],
  ]);
  // the marked boundary keeps no plain divider; the other boundaries keep theirs.
  const dividers = Array.from(band.querySelectorAll<HTMLElement>('[data-track-divider]'),
    (divider) => divider.style.left);
  expect(dividers).toEqual(['159px', '239px']);
});

test('a hidden span at either end of the window renders no separator', () => {
  // nothing visible follows the span, so there is no boundary inside the window.
  const trailing = mount(handlers, 1, { startRow: 0, startCol: 0,
    colIndices: [0, 1, 2], colOffsets: [0, 80, 160, 240], rowOffsets: [0, 20],
    hiddenColSpans: [[3, 4]] });
  expect(trailing.container.querySelectorAll('[data-testid="xlsx-hidden-divider"]')).toHaveLength(0);
  // the window opens on a hidden span: no visible track precedes it to separate from.
  const leading = mount(handlers, 1, { startRow: 0, startCol: 0,
    colIndices: [2, 3], colOffsets: [0, 80, 160], rowOffsets: [0, 20],
    hiddenColSpans: [[0, 1]] });
  expect(leading.container.querySelectorAll('[data-testid="xlsx-hidden-divider"]')).toHaveLength(0);
});

test('a frozen lead drops a collapsed span that lies above the scrolling body', () => {
  // row 1 is pinned, the body is scrolled to row 12, and row 8 is hidden above
  // it: the span collapsed onto row 12's leading edge, between the frozen lead
  // and the body, where the reader can never scroll to it.
  const { container } = mount(handlers, 1, { startRow: 0, startCol: 0,
    frozenRows: 1, dividerY: 20, colOffsets: [0, 80], rowIndices: [0, 11, 12], rowOffsets: [0, 20, 40, 60],
    hiddenRowSpans: [[7, 7]] });
  expect(container.querySelectorAll('[data-testid="xlsx-hidden-divider"]')).toHaveLength(0);
});

test('a frozen lead keeps a collapsed span that lies inside the scrolling body', () => {
  // the same pinned row 1, but the body starts at row 3 and the hidden row 7
  // collapsed onto row 8, which the body paints: the separator stays.
  const { container } = mount(handlers, 1, { startRow: 0, startCol: 0,
    frozenRows: 1, dividerY: 20, colOffsets: [0, 80], rowIndices: [0, 3, 4, 8],
    rowOffsets: [0, 20, 40, 60, 80],
    hiddenRowSpans: [[7, 7]] });
  const markers = Array.from(container.querySelectorAll<HTMLElement>('[data-testid="xlsx-hidden-divider"]'));
  expect(markers.map((marker) => [marker.getAttribute('data-axis'), marker.style.top]))
    .toEqual([['row', '58px'], ['row', '62px']]);
});

test('a compacted hidden row span draws the doubled separator across the row band', () => {
  const { container } = mount(handlers, 1, { startRow: 0, startCol: 0,
    colOffsets: [0, 80], rowIndices: [0, 3, 4], rowOffsets: [0, 20, 40, 60],
    hiddenRowSpans: [[1, 2]] });
  const band = container.querySelector('[data-testid="xlsx-row-headers"]')!;
  const markers = Array.from(band.querySelectorAll<HTMLElement>('[data-testid="xlsx-hidden-divider"]'));
  expect(markers.map((marker) => [marker.getAttribute('data-axis'), marker.style.top, marker.style.width])).toEqual([
    ['row', '18px', '44px'], ['row', '22px', '44px'],
  ]);
  const dividers = Array.from(band.querySelectorAll<HTMLElement>('[data-track-divider]'),
    (divider) => divider.style.top);
  expect(dividers).toEqual(['39px', '59px']);
});

// A frame hiding sheet column 1..2 and row 1..2, so each band has one collapsed
// boundary: column 3's leading edge at 80, row 3's leading edge at 20.
const hiddenGrid = {
  startRow: 0, startCol: 0, colIndices: [0, 3], colOffsets: [0, 80, 160],
  rowIndices: [0, 3], rowOffsets: [0, 20, 40],
  hiddenColSpans: [[1, 2]] as [number, number][], hiddenRowSpans: [[1, 2]] as [number, number][],
};
test('spans that collapse onto the same boundary share one separator', () => {
  const { container } = mount(handlers, 1, { startRow: 0, startCol: 0,
    colIndices: [0, 3], colOffsets: [0, 80, 160], rowOffsets: [0, 20],
    hiddenColSpans: [[1, 1], [2, 2]] });
  expect(container.querySelectorAll('[data-testid="xlsx-hidden-divider"]')).toHaveLength(2);
});

test('a collapsed hidden span paints the doubled separator and no reveal handle', () => {
  const { container } = mount(handlers, 1, hiddenGrid);
  expect(container.querySelectorAll('[data-testid="xlsx-hidden-divider"]')).toHaveLength(4);
  expect(container.querySelectorAll('[role="button"]')).toHaveLength(0);
});

test('the select-all corner is live only with an onSelectAll handler', () => {
  let calls = 0;
  const live = mount({ ...handlers, onSelectAll: () => { calls += 1; } });
  const corner = live.container.querySelector('[data-testid="xlsx-grid-corner"]')!;
  expect(corner.getAttribute('aria-hidden')).toBeNull();
  expect(corner.getAttribute('role')).toBe('button');
  // the triangle is the affordance; it exists only where the corner works.
  expect(corner.querySelector('span')).not.toBeNull();
  fireEvent.click(corner);
  expect(calls).toBe(1);

  const inert = render(<SheetGrid frame={frame(geometry)} zoom={1} showSheetHeaders
    headerHandlers={{ onHeaderMouseDown: handlers.onHeaderMouseDown }}><div /></SheetGrid>);
  const dead = inert.container.querySelector('[data-testid="xlsx-grid-corner"]')!;
  expect(dead.getAttribute('aria-hidden')).toBe('true');
  expect(dead.getAttribute('role')).toBeNull();
  expect(dead.querySelector('span')).toBeNull();
  fireEvent.click(dead);
  expect(calls).toBe(1);
});

test('a press on the select-all corner arms no band gesture', () => {
  const { container } = mount({ ...handlers, onSelectAll: () => {} });
  const corner = container.querySelector('[data-testid="xlsx-grid-corner"]')!;
  const columnBand = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(columnBand);
  fireEvent.mouseDown(corner, { button: 0 });
  fireEvent.pointerDown(corner, { button: 0, clientX: 0, clientY: 0, pointerId: 9 });
  fireEvent.pointerMove(columnBand, { clientX: 200, clientY: 0, pointerId: 9 });
  fireEvent.pointerUp(columnBand, { clientX: 200, clientY: 0, pointerId: 9 });
  expect(presses).toEqual([]);
  expect(resizes).toEqual([]);
  expect(selects).toEqual([]);
  expect(reorders).toEqual([]);
});

test('a frozen label still takes its press and a frozen boundary still resizes', () => {
  const { container } = mount(handlers, 1, { startRow: 0, startCol: 0,
    frozenRows: 1, frozenCols: 1, dividerX: 80, dividerY: 20,
    colOffsets: [0, 80, 160], rowOffsets: [0, 20, 40] });
  fireEvent.mouseDown(columnLabel(container, 0), {});
  expect(presses).toEqual([{ axis: 'col', index: 0, shift: false }]);
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 80, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 120, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 120, clientY: 0, pointerId: 1 });
  expect(resizes).toEqual([{ axis: 'col', index: 0, size: 120 }]);
});

test('A1 letters cover bijective base-26 boundaries', () => {
  expect([0, 1, 25, 26, 27, 51, 52, 701, 702, 16383].map(columnLetter))
    .toEqual(['A', 'B', 'Z', 'AA', 'AB', 'AZ', 'BA', 'ZZ', 'AAA', 'XFD']);
});

test('headerTracks skips zero width and keeps the absolute index', () => {
  expect(headerTracks({ startRow: 0, startCol: 0, rowOffsets: [0, 10],
    colIndices: [0, 4, 5], colOffsets: [0, 80, 80, 160] }, 'col'))
    .toEqual([{ index: 0, offset: 0, size: 80 }, { index: 5, offset: 80, size: 80 }]);
});

test('headerTracks sizes a frozen track from the renderer, not the clipped offset', () => {
  // the frozen pane cuts the leading 15px off column A: the offset gap is 65
  // but the track is still 80 wide, so a boundary drag must start from 80
  expect(headerTracks({ startRow: 0, startCol: 0, rowOffsets: [0, 20],
    colIndices: [0, 1], colOffsets: [0, 65, 145], colSizes: [80, 80] }, 'col'))
    .toEqual([{ index: 0, offset: 0, size: 80 }, { index: 1, offset: 65, size: 80 }]);
});

test('reorderRun takes the whole-track selection, else only the pressed track', () => {
  const columns = { top: 0, bottom: XLSX_MAX_ROWS - 1, left: 1, right: 3 };
  const rows = { top: 1, bottom: 3, left: 0, right: XLSX_MAX_COLUMNS - 1 };
  const cells = { top: 1, bottom: 3, left: 1, right: 3 };
  // pressed inside the run: the run is the whole selection, both axes.
  expect(reorderRun('col', 2, columns)).toEqual({ from: 1, to: 3 });
  expect(reorderRun('row', 2, rows)).toEqual({ from: 1, to: 3 });
  // a cell range that merely reaches the track is never a whole-track run.
  expect(reorderRun('col', 2, cells)).toEqual({ from: 2, to: 2 });
  expect(reorderRun('row', 2, cells)).toEqual({ from: 2, to: 2 });
  expect(reorderRun('col', 7, null)).toEqual({ from: 7, to: 7 });
});

test('resizeRun hands back the whole-track run only from a boundary inside it', () => {
  const columns = { top: 0, bottom: XLSX_MAX_ROWS - 1, left: 1, right: 3 };
  const rows = { top: 1, bottom: 3, left: 0, right: XLSX_MAX_COLUMNS - 1 };
  const cells = { top: 1, bottom: 3, left: 1, right: 3 };
  expect(resizeRun('col', 3, columns)).toEqual({ at: 1, count: 3 });
  expect(resizeRun('row', 3, rows)).toEqual({ at: 1, count: 3 });
  // the boundary's leading track sits outside the run: single-track resize.
  expect(resizeRun('col', 0, columns)).toBeUndefined();
  expect(resizeRun('row', 0, rows)).toBeUndefined();
  // a cell range is never a whole-track run, and a non-whole selection of one
  // track is today's single-track resize.
  expect(resizeRun('col', 2, cells)).toBeUndefined();
  expect(resizeRun('col', 7, null)).toBeUndefined();
});

test('pressing a column label selects that column, and shift extends', () => {
  const { container } = mount(handlers);
  fireEvent.mouseDown(columnLabel(container, 1), { shiftKey: false });
  fireEvent.mouseDown(columnLabel(container, 2), { shiftKey: true });
  expect(presses).toEqual([{ axis: 'col', index: 1, shift: false }, { axis: 'col', index: 2, shift: true }]);
});

test('pressing a row label selects that row', () => {
  const { container } = mount(handlers);
  fireEvent.mouseDown(rowLabel(container, 1), {});
  expect(presses).toEqual([{ axis: 'row', index: 1, shift: false }]);
});

test('a hidden track renders no label, so no press can reach it', () => {
  // column 4 is zero-sized between two visible columns: it has no element.
  const { container } = mount(handlers, 1, { startRow: 0, startCol: 0, colIndices: [0, 4, 5],
    rowOffsets: [0, 20, 40], colOffsets: [0, 80, 80, 160] });
  expect(labels(container, 'xlsx-column-headers')).toEqual(['A', 'F']);
  expect(container.querySelector('[data-col="4"]')).toBeNull();
  expect(container.querySelectorAll('[data-resize-index="4"]')).toHaveLength(0);
  fireEvent.mouseDown(container.querySelector('[data-testid="xlsx-column-headers"]')!, {});
  expect(presses).toEqual([]);
});

test('a boundary press and drag resizes the column to the left of it', () => {
  const { container } = mount(handlers);
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 80, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 120, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 120, clientY: 0, pointerId: 1 });
  expect(resizes).toEqual([{ axis: 'col', index: 0, size: 120 }]);
  // the press never reached the label underneath.
  expect(presses).toEqual([]);
});

test('a boundary press and drag resizes the row above it', () => {
  const { container } = mount(handlers);
  const band = container.querySelector('[data-testid="xlsx-row-headers"]')!;
  stubRect(band);
  // the boundary between rows 0 and 1 sits at 20 sheet units.
  fireEvent.pointerDown(band, { button: 0, clientX: 0, clientY: 20, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 0, clientY: 35, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 0, clientY: 35, pointerId: 1 });
  expect(resizes).toEqual([{ axis: 'row', index: 0, size: 35 }]);
});

test('a press that never moves past the slop commits no resize', () => {
  const { container } = mount(handlers);
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 80, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 82, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 82, clientY: 0, pointerId: 1 });
  expect(resizes).toEqual([]);
});

test('a press away from any boundary selects the column under the pointer', () => {
  const { container } = mount(handlers);
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 100, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 100, clientY: 0, pointerId: 1 });
  fireEvent.mouseDown(columnLabel(container, 1), {});
  expect(resizes).toEqual([]);
  expect(presses).toEqual([{ axis: 'col', index: 1, shift: false }]);
});

// five 80px columns: A 0..80, B 80..160, C 160..240, D 240..320, E 320..400.
const wideColumns = { startRow: 0, startCol: 0,
  colOffsets: [0, 80, 160, 240, 320, 400], rowOffsets: [0, 20] };
// five 20pt rows: 1 0..20, 2 20..40, 3 40..60, 4 60..80, 5 80..100.
const tallRows = { startRow: 0, startCol: 0,
  colOffsets: [0, 80], rowOffsets: [0, 20, 40, 60, 80, 100] };

test('a column boundary inside a whole-column selection resizes the whole run', () => {
  // B:D selected: dragging the D/E boundary sets all three columns to the
  // dragged width, so the run rides along and the preview stays the track only.
  const { container } = mount(handlers, 1, wideColumns, wholeColumns(1, 3));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 320, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 340, clientY: 0, pointerId: 1 });
  // the live preview is still the dragged track alone.
  expect((columnLabel(container, 1) as HTMLElement).style.width).toBe('80px');
  expect((columnLabel(container, 3) as HTMLElement).style.width).toBe('100px');
  fireEvent.pointerUp(band, { clientX: 340, clientY: 0, pointerId: 1 });
  expect(resizes).toEqual([{ axis: 'col', index: 3, size: 100, run: { at: 1, count: 3 } }]);
});

test('a column boundary outside the whole-column run resizes one track', () => {
  const { container } = mount(handlers, 1, wideColumns, wholeColumns(1, 3));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  // the A/B boundary's leading track (A) is left of the run.
  fireEvent.pointerDown(band, { button: 0, clientX: 80, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 100, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 100, clientY: 0, pointerId: 1 });
  expect(resizes).toEqual([{ axis: 'col', index: 0, size: 100 }]);
  expect(resizes[0]!.run).toBeUndefined();
});

test('a row boundary inside a whole-row selection resizes the whole run', () => {
  // rows 2:4 selected: dragging the 4/5 boundary sets all three rows.
  const { container } = mount(handlers, 1, tallRows, wholeRows(1, 3));
  const band = container.querySelector('[data-testid="xlsx-row-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 0, clientY: 80, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 0, clientY: 90, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 0, clientY: 90, pointerId: 1 });
  expect(resizes).toEqual([{ axis: 'row', index: 3, size: 30, run: { at: 1, count: 3 } }]);
});

test('a row boundary outside the whole-row run resizes one track', () => {
  const { container } = mount(handlers, 1, tallRows, wholeRows(1, 3));
  const band = container.querySelector('[data-testid="xlsx-row-headers"]')!;
  stubRect(band);
  // the 1/2 boundary's leading row (1) is above the run.
  fireEvent.pointerDown(band, { button: 0, clientX: 0, clientY: 20, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 0, clientY: 30, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 0, clientY: 30, pointerId: 1 });
  expect(resizes).toEqual([{ axis: 'row', index: 0, size: 30 }]);
  expect(resizes[0]!.run).toBeUndefined();
});

test('the resize hit zone resolves the boundary track for both axes', () => {
  const columns = headerTracks(geometry, 'col');
  const rows = headerTracks(geometry, 'row');
  expect(resizeGestureAt(columns, 76)).toEqual({ index: 0, offset: 0, size: 80 });
  expect(resizeGestureAt(columns, 84)).toEqual({ index: 0, offset: 0, size: 80 });
  expect(resizeGestureAt(columns, 80 + HEADER_RESIZE_ZONE_PX + 1)).toBeNull();
  expect(resizeGestureAt(columns, 40)).toBeNull();
  // the last track has no successor boundary.
  expect(resizeGestureAt(columns, 240)).toBeNull();
  expect(resizeGestureAt(rows, 20)).toEqual({ index: 0, offset: 0, size: 20 });
  expect(resizeGestureAt(rows, 21)).toEqual({ index: 0, offset: 0, size: 20 });
  expect(resizeGestureAt(rows, 40)).toBeNull();
});

test('resize strips cover each boundary and stop at the last track', () => {
  const { container } = mount(handlers);
  const strips = Array.from(container.querySelectorAll<HTMLElement>('[data-resize-axis="col"]'));
  expect(strips.map((strip) => strip.getAttribute('data-resize-index'))).toEqual(['0', '1']);
  expect(strips.map((strip) => strip.style.left)).toEqual(['76px', '156px']);
  // the strip starts below the hidden handle (y 1..9) and covers the rest of
  // the 24px band, so boundary presses in that area still arm a resize.
  expect(strips.map((strip) => [strip.style.top, strip.style.height]))
    .toEqual([['9px', '15px'], ['9px', '15px']]);
  const [rowStrip] = Array.from(container.querySelectorAll<HTMLElement>('[data-resize-axis="row"]'));
  // the row strip starts past the handle (x 2..10) and keeps the rest of the
  // 44px band; the boundary at y=20 gives a strip from y=16.
  expect([rowStrip!.style.left, rowStrip!.style.width, rowStrip!.style.top, rowStrip!.style.height])
    .toEqual(['10px', '34px', '16px', '8px']);
});

test('resize sizes are clamped in pixels to the engine bounds', () => {
  expect(MAX_COLUMN_WIDTH_PX).toBe(1790);
  expect(MAX_ROW_HEIGHT_PX).toBe(546);
  expect(clampTrackSize('col', 5000)).toBe(MAX_COLUMN_WIDTH_PX);
  expect(clampTrackSize('col', -5)).toBe(0);
  expect(clampTrackSize('row', 5000)).toBe(MAX_ROW_HEIGHT_PX);
  expect(clampTrackSize('row', 22.5)).toBe(22.5);
});

test('a dragged pixel size converts to the characters and points the engine stores', () => {
  // the sample sheet's 34-character column paints 243.004px.
  expect(columnWidthFromPx(243.004)).toBeCloseTo(34, 2);
  // excel's default 15pt row paints 20px.
  expect(rowHeightFromPx(20)).toBe(15);
  expect(columnWidthFromPx(0)).toBe(0);
  expect(rowHeightFromPx(-4)).toBe(0);
});

test('header bands are absent by default', () => {
  const { container } = render(<SheetGrid frame={frame(geometry)} zoom={1}><div data-testid="child" /></SheetGrid>);
  expect(container.querySelector('[data-testid="xlsx-grid"]')).toBeNull();
  expect(container.querySelector('[data-testid="xlsx-column-headers"]')).toBeNull();
  expect(container.querySelector('[data-testid="xlsx-row-headers"]')).toBeNull();
  expect(container.firstElementChild?.getAttribute('data-testid')).toBe('child');
});

test('the resize hit zone scales with zoom', () => {
  const columns = headerTracks(geometry, 'col');
  expect(resizeGestureAt(columns, 87, 0.5)).toEqual({ index: 0, offset: 0, size: 80 });
  expect(resizeGestureAt(columns, 89, 0.5)).toBeNull();
  expect(resizeGestureAt(columns, 82, 2)).toEqual({ index: 0, offset: 0, size: 80 });
  expect(resizeGestureAt(columns, 83, 2)).toBeNull();
});

test('a drag previews the new size and shifts the tracks after it', () => {
  const tracks = headerTracks(geometry, 'col');
  expect(resizePreview(tracks, 0, 40)).toEqual([
    { index: 0, offset: 0, size: 120 },
    { index: 1, offset: 120, size: 80 },
    { index: 2, offset: 200, size: 80 },
  ]);
  // shrinking clamps at zero and pulls later tracks back with it.
  expect(resizePreview(tracks, 1, -100)).toEqual([
    { index: 0, offset: 0, size: 80 },
    { index: 1, offset: 80, size: 0 },
    { index: 2, offset: 60, size: 80 },
  ]);
});

test('a zoomed frame keeps the hit zone and the committed size in sheet units', () => {
  const { container } = mount(handlers, 2);
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  // the boundary sits at 80 sheet units = 160 css px at zoom 2.
  fireEvent.pointerDown(band, { button: 0, clientX: 160, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 240, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 240, clientY: 0, pointerId: 1 });
  expect(resizes).toEqual([{ axis: 'col', index: 0, size: 120 }]);
});

// the band tints are plain neutral greys: hover, a selection that merely
// reaches a track, then a whole-track selection.
const tint = (element: Element) => (element as HTMLElement).style.backgroundColor;

test('the split paints over the gridline it replaces', () => {
  // A two-pixel split on a track boundary must read as one line, not as a split
  // plus the one-pixel gridline peeking out from under its edge.
  const { container } = mount(handlers, 1, { startRow: 0, startCol: 0,
    colOffsets: [0, 103, 178, 190], rowOffsets: [0, 20], frozenCols: 2, dividerX: 178 });
  const split = container.querySelector<HTMLElement>('[data-testid="xlsx-freeze-divider"][data-axis="col"]')!;
  expect(split.style.zIndex).toBe('1');
  expect(split.style.width).toBe('2px');
});

test('both bands keep an edge on the side that faces the cells', () => {
  const { container } = mount(handlers);
  const edges = Array.from(container.querySelectorAll<HTMLElement>('[data-testid="xlsx-header-edge"]'));
  expect(edges.map((edge) => edge.dataset.axis)).toEqual(['col', 'row']);
  const col = edges[0]!;
  expect([col.style.height, col.style.bottom, col.style.background]).toEqual(['1px', '0px', '#e0e0e0']);
  const row = edges[1]!;
  expect([row.style.width, row.style.right, row.style.background]).toEqual(['1px', '0px', '#e0e0e0']);
});

test('hover tints only the label under the pointer and clears on leave', () => {
  const { container } = mount(handlers);
  const hovered = columnLabel(container, 1);
  fireEvent.pointerEnter(hovered);
  expect(tint(hovered)).toBe('#f1f3f4');
  expect(tint(columnLabel(container, 0))).toBe('');
  expect(tint(columnLabel(container, 2))).toBe('');
  fireEvent.pointerLeave(hovered);
  expect(tint(hovered)).toBe('');
});

test('a whole-column selection highlights exactly that column label', () => {
  const column = 1;
  const { container } = mount(handlers, 1, geometry, {
    anchor: { row: 0, col: column },
    focus: { row: XLSX_MAX_ROWS - 1, col: column },
  });
  const selected = columnLabel(container, column) as HTMLElement;
  expect(tint(selected)).toBe('#e3e7ea');
  expect(selected.style.fontWeight).toBe('600');
  expect(selected.style.color).toBe('#3c4043');
  expect(selected.style.boxShadow).toBe('');
  // the accent is a band overlay that replaces the edge under the label.
  const underline = container.querySelector<HTMLElement>('[data-testid="xlsx-header-underline"][data-axis="col"]')!;
  expect([underline.style.height, underline.style.bottom, underline.style.background])
    .toEqual(['2px', '0px', '#9aa0a6']);
  expect(tint(columnLabel(container, 0))).toBe('');
  expect(tint(columnLabel(container, 2))).toBe('');
  expect((columnLabel(container, 0) as HTMLElement).style.fontWeight).toBe('');
  expect(container.querySelectorAll('[data-testid="xlsx-header-underline"]').length).toBe(1);
});

test('a cell-range selection tints every track it reaches, without selecting them', () => {
  const { container } = mount(handlers, 1, geometry, {
    anchor: { row: 1, col: 1 },
    focus: { row: 1, col: 2 },
  });
  // B:C and row 2 are reached: light fill, never bold or underlined.
  for (const index of [1, 2]) {
    const label = columnLabel(container, index) as HTMLElement;
    expect(tint(label)).toBe('#eceff1');
    expect(label.style.fontWeight).toBe('');
    expect(label.style.boxShadow).toBe('');
  }
  expect(tint(columnLabel(container, 0))).toBe('');
  expect(tint(container.querySelector<HTMLElement>('[data-row="1"]')!)).toBe('#eceff1');
  expect(tint(container.querySelector<HTMLElement>('[data-row="0"]')!)).toBe('');
});

test('a whole-column selection leaves row labels alone', () => {
  const { container } = mount(handlers, 1, geometry, {
    anchor: { row: 0, col: 1 },
    focus: { row: XLSX_MAX_ROWS - 1, col: 1 },
  });
  // the column's row extent is the whole sheet; tinting every row label would
  // flood the band, which is why only the columns are filled.
  const tintedRows = Array.from(container.querySelectorAll<HTMLElement>('[data-testid="xlsx-row-headers"] [data-row]'))
    .filter((label) => label.style.backgroundColor !== '');
  expect(tintedRows).toEqual([]);
});

test('a boundary drag draws the guide at the previewed boundary and drops it on release', () => {
  const { container } = mount(handlers);
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 80, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 120, clientY: 0, pointerId: 1 });
  const guide = container.querySelector<HTMLElement>('[data-testid="xlsx-resize-guide"]')!;
  expect(guide).not.toBeNull();
  expect([guide.style.left, guide.style.width, guide.style.height]).toEqual(['120px', '1px', '24px']);
  expect(guide.style.pointerEvents).toBe('none');
  expect(guide.style.background).toBe('#5f6368');
  // the live preview is still painted alongside it.
  expect((columnLabel(container, 0) as HTMLElement).style.width).toBe('120px');
  fireEvent.pointerUp(band, { clientX: 120, clientY: 0, pointerId: 1 });
  expect(container.querySelector('[data-testid="xlsx-resize-guide"]')).toBeNull();
});

test('a row boundary drag draws a horizontal guide at the previewed edge', () => {
  const { container } = mount(handlers);
  const band = container.querySelector('[data-testid="xlsx-row-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 0, clientY: 20, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 0, clientY: 35, pointerId: 1 });
  const guide = container.querySelector<HTMLElement>('[data-testid="xlsx-resize-guide"]')!;
  expect([guide.style.top, guide.style.width, guide.style.height]).toEqual(['35px', '44px', '1px']);
});

test('dragging a label reorders it past the next midpoint and shows the insertion line', () => {
  // 30px tracks so a 40px drag clears a midpoint: mids 15, 45, 75. The band
  // only picks up a label the selection already covers in full.
  const narrow = { startRow: 0, startCol: 0, colOffsets: [0, 30, 60, 90], rowOffsets: [0, 20] };
  const { container } = mount(handlers, 1, narrow, wholeColumn(0));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(container, 0), { button: 0, clientX: 15, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 55, clientY: 0, pointerId: 1 });
  const guide = container.querySelector<HTMLElement>('[data-testid="xlsx-reorder-guide"]')!;
  expect(guide).not.toBeNull();
  expect([guide.style.left, guide.style.width, guide.style.height]).toEqual(['60px', '2px', '24px']);
  expect(guide.style.pointerEvents).toBe('none');
  // the dragged label lifts: raised above its neighbours, riding the pointer,
  // with no shadow of its own (adjacent labels would shadow each other).
  const lifted = columnLabel(container, 0) as HTMLElement;
  expect(lifted.style.boxShadow).toBe('');
  expect(lifted.style.zIndex).toBe('3');
  expect(lifted.style.transform).toBe('translateX(40px)');
  // a single-track run lifts only that label.
  const neighbour = columnLabel(container, 1) as HTMLElement;
  expect(neighbour.style.boxShadow).toBe('');
  expect(neighbour.style.transform).toBe('');
  fireEvent.pointerUp(band, { clientX: 55, clientY: 0, pointerId: 1 });
  // column 0 lands in gap 2 (between B and C), which `to` reports raw.
  expect(reorders).toEqual([{ axis: 'col', at: 0, before: 2, count: 1 }]);
  expect(container.querySelector('[data-testid="xlsx-reorder-guide"]')).toBeNull();
  expect((columnLabel(container, 0) as HTMLElement).style.transform).toBe('');
  // the gesture never committed a track press.
  expect(presses).toEqual([]);
});

test('the lifted label is matched by sheet index, not by its slot in the visible band', () => {
  // a scrolled frame starts at AA: the label index (26) differs from its slot (0).
  const { container } = mount(handlers, 1, { startRow: 0, startCol: 26,
    colOffsets: [0, 30, 60, 90], rowOffsets: [0, 20] }, wholeColumn(26));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(container, 26), { button: 0, clientX: 15, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 55, clientY: 0, pointerId: 1 });
  expect((columnLabel(container, 26) as HTMLElement).style.transform).toBe('translateX(40px)');
  expect((columnLabel(container, 27) as HTMLElement).style.transform).toBe('');
});

test('a row label drag reorders rows and draws the line across the band', () => {
  const { container } = mount(handlers, 1, { startRow: 0, startCol: 0,
    colOffsets: [0, 80], rowOffsets: [0, 20, 40, 60] }, wholeRow(0));
  const band = container.querySelector('[data-testid="xlsx-row-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(rowLabel(container, 0), { button: 0, clientX: 0, clientY: 10, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 0, clientY: 45, pointerId: 1 });
  const guide = container.querySelector<HTMLElement>('[data-testid="xlsx-reorder-guide"]')!;
  expect([guide.style.top, guide.style.width, guide.style.height]).toEqual(['40px', '44px', '2px']);
  fireEvent.pointerUp(band, { clientX: 0, clientY: 45, pointerId: 1 });
  expect(reorders).toEqual([{ axis: 'row', at: 0, before: 2, count: 1 }]);
});

test('without onHeaderReorder no label drag gesture exists', () => {
  const { container } = mount({ onHeaderMouseDown: handlers.onHeaderMouseDown });
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 40, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 300, clientY: 0, pointerId: 1 });
  expect(container.querySelector('[data-testid="xlsx-reorder-guide"]')).toBeNull();
  fireEvent.pointerUp(band, { clientX: 300, clientY: 0, pointerId: 1 });
  expect(reorders).toEqual([]);
});

test('a boundary press resizes even when reordering is available', () => {
  const { container } = mount(handlers);
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 80, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 120, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 120, clientY: 0, pointerId: 1 });
  expect(reorders).toEqual([]);
  expect(resizes).toEqual([{ axis: 'col', index: 0, size: 120 }]);
});

test('the insertion gap reads track geometry, not the width of the selected range', () => {
  const narrow = { startRow: 0, startCol: 0, colOffsets: [0, 30, 60, 90], rowOffsets: [0, 20] };
  const drag = (root: Element) => {
    const band = root.querySelector('[data-testid="xlsx-column-headers"]')!;
    stubRect(band);
    fireEvent.pointerDown(columnLabel(root, 0), { button: 0, clientX: 15, clientY: 0, pointerId: 1 });
    fireEvent.pointerMove(band, { clientX: 55, clientY: 0, pointerId: 1 });
    const left = root.querySelector<HTMLElement>('[data-testid="xlsx-reorder-guide"]')!.style.left;
    fireEvent.pointerUp(band, { clientX: 55, clientY: 0, pointerId: 1 });
    return left;
  };
  const single = drag(mount(handlers, 1, narrow, wholeColumn(0)).container);
  const singleCalls = reorders.slice();
  expect(singleCalls).toEqual([{ axis: 'col', at: 0, before: 2, count: 1 }]);
  // a wider whole-track selection still leaves the gap at the same midpoint;
  // only the run's own slots change what the drop means (here it lands right
  // after A:B, so it moves nothing).
  expect(drag(mount(handlers, 1, narrow, wholeColumns(0, 1)).container)).toBe(single);
  expect(reorders).toEqual([]);
});

test('a click still selects a label while reordering is available, but a started reorder does not', () => {
  const narrow = { startRow: 0, startCol: 0, colOffsets: [0, 30, 60, 90], rowOffsets: [0, 20] };
  const { container } = mount(handlers, 1, narrow, wholeColumn(0));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  // press without moving: the browser's mousedown lands before the slop, so the
  // track is still selected — here on a track the selection does not cover.
  fireEvent.pointerDown(columnLabel(container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 1 });
  fireEvent.mouseDown(columnLabel(container, 1), {});
  fireEvent.pointerUp(band, { clientX: 45, clientY: 0, pointerId: 1 });
  expect(presses).toEqual([{ axis: 'col', index: 1, shift: false }]);
  expect(reorders).toEqual([]);
  expect(selects).toEqual([]);
  // once the drag is a reorder of the selected track, a press on a label is not
  // a selection.
  fireEvent.pointerDown(columnLabel(container, 0), { button: 0, clientX: 15, clientY: 0, pointerId: 2 });
  fireEvent.pointerMove(band, { clientX: 55, clientY: 0, pointerId: 2 });
  fireEvent.mouseDown(columnLabel(container, 1), {});
  expect(presses).toEqual([{ axis: 'col', index: 1, shift: false }]);
  fireEvent.pointerUp(band, { clientX: 55, clientY: 0, pointerId: 2 });
  expect(reorders).toEqual([{ axis: 'col', at: 0, before: 2, count: 1 }]);
  expect(selects).toEqual([]);
});

// 30px tracks, so a drag crosses a whole track every 30px: A 0..30, B 30..60,
// C 60..90.
const narrow = { startRow: 0, startCol: 0, colOffsets: [0, 30, 60, 90], rowOffsets: [0, 20] };

test('a drag from an unselected label extends the selection and never reorders', () => {
  const { container } = mount(handlers, 1, narrow);
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(container, 0), { button: 0, clientX: 15, clientY: 0, pointerId: 1 });
  // the press itself still selects A; that is the anchor the drag grows from.
  fireEvent.mouseDown(columnLabel(container, 0), {});
  fireEvent.pointerMove(band, { clientX: 45, clientY: 0, pointerId: 1 });
  // moving inside the same track is not a new extend.
  fireEvent.pointerMove(band, { clientX: 50, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 75, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 75, clientY: 0, pointerId: 1 });
  expect(selects).toEqual([{ axis: 'col', index: 1 }, { axis: 'col', index: 2 }]);
  expect(reorders).toEqual([]);
  expect(presses).toEqual([{ axis: 'col', index: 0, shift: false }]);
  expect(container.querySelector('[data-testid="xlsx-reorder-guide"]')).toBeNull();
});

test('a drag on a whole-track selected label still reorders it', () => {
  const { container } = mount(handlers, 1, narrow, wholeColumn(0));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(container, 0), { button: 0, clientX: 15, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 55, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 55, clientY: 0, pointerId: 1 });
  expect(reorders).toEqual([{ axis: 'col', at: 0, before: 2, count: 1 }]);
  expect(selects).toEqual([]);
});

// --- a reorder moves the whole selected run --------------------------------

// four 30px columns: A 0..30, B 30..60, C 60..90, D 90..120, midpoints
// 15, 45, 75, 105.
const fourTracks = { startRow: 0, startCol: 0,
  colOffsets: [0, 30, 60, 90, 120], rowOffsets: [0, 20] };

test('a drag from inside a whole-column selection moves the whole run', () => {
  const { container } = mount(handlers, 1, fourTracks, wholeColumns(0, 2));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 115, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 115, clientY: 0, pointerId: 1 });
  // pressed B, moved A:C: `from` is the run's first slot and the drop is past D.
  expect(reorders).toEqual([{ axis: 'col', at: 0, before: 4, count: 3 }]);
  expect(selects).toEqual([]);
});

test('the whole selected run lifts while a multi-track reorder is in progress', () => {
  // five 30px columns so the run A:C has neighbours on both sides of the band.
  const fiveTracks = { startRow: 0, startCol: 0,
    colOffsets: [0, 30, 60, 90, 120, 150], rowOffsets: [0, 20] };
  const { container } = mount(handlers, 1, fiveTracks, wholeColumns(0, 2));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  // pressed B (centre 45) and dragged to 115: the block travels 70px.
  fireEvent.pointerDown(columnLabel(container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 115, clientY: 0, pointerId: 1 });
  // every label of A:C rides the pointer as one block, keeping its spacing and
  // with no shadow of its own: adjacent labels would cast onto each other.
  for (const index of [0, 1, 2]) {
    const label = columnLabel(container, index) as HTMLElement;
    expect(label.style.transform).toBe('translateX(70px)');
    expect(label.style.backgroundColor).toBe('#ffffff');
    expect(label.style.fontWeight).toBe('600');
    expect(label.style.color).toBe('#3c4043');
    expect(label.style.boxShadow).toBe('');
    expect(label.style.cursor).toBe('grabbing');
    expect(label.style.zIndex).toBe('3');
  }
  // labels outside the captured run keep the band's normal style.
  for (const index of [3, 4]) {
    const label = columnLabel(container, index) as HTMLElement;
    expect(label.style.transform).toBe('');
    expect(label.style.boxShadow).toBe('');
    expect(label.style.zIndex).toBe('');
  }
});

test('a multi-track run lifts by sheet index, not by its slots on a scrolled band', () => {
  // a scrolled frame starts at AA (26): labels 26..28 sit in slots 0..2.
  const { container } = mount(handlers, 1, { startRow: 0, startCol: 26,
    colOffsets: [0, 30, 60, 90, 120], rowOffsets: [0, 20] }, wholeColumns(26, 28));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(container, 27), { button: 0, clientX: 45, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 115, clientY: 0, pointerId: 1 });
  for (const index of [26, 27, 28]) {
    expect((columnLabel(container, index) as HTMLElement).style.transform).toBe('translateX(70px)');
  }
  // AA's neighbour past the run does not lift.
  expect((columnLabel(container, 29) as HTMLElement).style.transform).toBe('');
});

test('a drop inside the selected run moves nothing', () => {
  const { container } = mount(handlers, 1, fourTracks, wholeColumns(0, 2));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  // press B, drop between B and C: the gap lies strictly inside A:C.
  fireEvent.pointerDown(columnLabel(container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 65, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 65, clientY: 0, pointerId: 1 });
  // press B, drop right after C: the run is already there.
  fireEvent.pointerDown(columnLabel(container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 2 });
  fireEvent.pointerMove(band, { clientX: 95, clientY: 0, pointerId: 2 });
  fireEvent.pointerUp(band, { clientX: 95, clientY: 0, pointerId: 2 });
  // press B, drop before A: the run already starts there.
  fireEvent.pointerDown(columnLabel(container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 3 });
  fireEvent.pointerMove(band, { clientX: 10, clientY: 0, pointerId: 3 });
  fireEvent.pointerUp(band, { clientX: 10, clientY: 0, pointerId: 3 });
  expect(reorders).toEqual([]);
});

test('a click inside a selected run collapses it to the pressed track on release', () => {
  const { container } = mount(handlers, 1, fourTracks, wholeColumns(0, 2));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  // pressing B arms the run's reorder, so the label's press is held back and
  // the whole run keeps its highlight while the pointer is down.
  fireEvent.pointerDown(columnLabel(container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 1 });
  fireEvent.mouseDown(columnLabel(container, 1), {});
  expect(presses).toEqual([]);
  expect([0, 1, 2].map((index) => tint(columnLabel(container, index))))
    .toEqual(['#e3e7ea', '#e3e7ea', '#e3e7ea']);
  expect(tint(columnLabel(container, 3))).toBe('');
  // released without a drag: the held press reaches the surface once, so it
  // collapses A:C to B, and no move is posted.
  fireEvent.pointerUp(band, { clientX: 45, clientY: 0, pointerId: 1 });
  expect(presses).toEqual([{ axis: 'col', index: 1, shift: false }]);
  expect(reorders).toEqual([]);
});

test('a click held on one sheet does not collapse the selection after a sheet change', () => {
  const view = mount(handlers, 1, fourTracks, wholeColumns(0, 2), 0);
  const band = view.container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(view.container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 1 });
  fireEvent.mouseDown(columnLabel(view.container, 1), {});
  // the sheet changes under the held press: the click belonged to the old one,
  // so the release must not select column B on the new sheet.
  view.rerender(<SheetGrid frame={frame(fourTracks)} zoom={1} showSheetHeaders headerHandlers={handlers}
    selection={wholeColumns(0, 2)} sheet={1}><div /></SheetGrid>);
  stubRect(band);
  fireEvent.pointerUp(band, { clientX: 45, clientY: 0, pointerId: 1 });
  expect(presses).toEqual([]);
  expect(reorders).toEqual([]);
});

test('a drag out of a selected run reorders without collapsing the selection', () => {
  const { container } = mount(handlers, 1, fourTracks, wholeColumns(0, 2));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 1 });
  // the browser's mousedown still lands; the armed drag holds it back.
  fireEvent.mouseDown(columnLabel(container, 1), {});
  fireEvent.pointerMove(band, { clientX: 115, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 115, clientY: 0, pointerId: 1 });
  // the drop is the gesture's only outcome: the run moves as a block and the
  // release is not replayed as a press.
  expect(reorders).toEqual([{ axis: 'col', at: 0, before: 4, count: 3 }]);
  expect(presses).toEqual([]);
});

test('a press on a header outside the selected run selects it once', () => {
  const { container } = mount(handlers, 1, fourTracks, wholeColumns(0, 2));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  // D lies outside A:C, so no reorder is armed and the label's own press runs;
  // the band's release must not replay it a second time.
  fireEvent.pointerDown(columnLabel(container, 3), { button: 0, clientX: 105, clientY: 0, pointerId: 1 });
  fireEvent.mouseDown(columnLabel(container, 3), {});
  fireEvent.pointerUp(band, { clientX: 105, clientY: 0, pointerId: 1 });
  expect(presses).toEqual([{ axis: 'col', index: 3, shift: false }]);
});

test('a cancelled press inside a selected run collapses nothing', () => {
  const { container } = mount(handlers, 1, fourTracks, wholeColumns(0, 2));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 1 });
  fireEvent.mouseDown(columnLabel(container, 1), {});
  fireEvent.pointerCancel(band, { pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 45, clientY: 0, pointerId: 1 });
  expect(presses).toEqual([]);
  expect(reorders).toEqual([]);
});

test('a gap names the sheet track it lands before, hidden tracks included', () => {
  // A, B, C, (D hidden), E: the gap right after C sits at E's leading edge, so
  // it names sheet column 4, not C's successor. The band compares that with the
  // run's own sheet bounds, which are the coordinates the move op takes.
  const grid = { startRow: 0, startCol: 0, colIndices: [0, 1, 2, 4],
    colOffsets: [0, 30, 60, 90, 120], rowOffsets: [0, 20], hiddenColSpans: [[3, 3]] as [number, number][] };
  const { container } = mount(handlers, 1, grid, wholeColumns(0, 2));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 95, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 95, clientY: 0, pointerId: 1 });
  // the run A:C lands before sheet column 4, past the hidden D.
  expect(reorders).toEqual([{ axis: 'col', at: 0, before: 4, count: 3 }]);
  // past E the gap names one past the last painted track.
  fireEvent.pointerDown(columnLabel(container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 2 });
  fireEvent.pointerMove(band, { clientX: 115, clientY: 0, pointerId: 2 });
  fireEvent.pointerUp(band, { clientX: 115, clientY: 0, pointerId: 2 });
  expect(reorders).toEqual([{ axis: 'col', at: 0, before: 4, count: 3 },
    { axis: 'col', at: 0, before: 5, count: 3 }]);
});

test('a run whose first track is hidden still moves the whole run', () => {
  // the selection starts at hidden column 1 and ends at hidden column 6: the
  // band names the run from the selection, not from the tracks it paints, so
  // the leading hidden tracks travel with it instead of being dropped.
  const grid = { startRow: 0, startCol: 0, colIndices: [0, 3, 4, 7],
    colOffsets: [0, 80, 160, 240, 320], rowOffsets: [0, 20],
    hiddenColSpans: [[1, 2], [5, 6]] as [number, number][] };
  const { container } = mount(handlers, 1, grid, wholeColumns(1, 6));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(container, 4), { button: 0, clientX: 200, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 300, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 300, clientY: 0, pointerId: 1 });
  // A:G is nine tracks wide (1..6 with the hidden ones inside); the drop past
  // the last painted track names one past sheet column 7.
  expect(reorders).toEqual([{ axis: 'col', at: 1, before: 8, count: 6 }]);
});

test('a run whose first track is hidden still moves the whole run of rows', () => {
  // the row mirror of the column case: rows 2..7 move even though rows 2, 3, 6
  // and 7 are hidden and the painted band starts at row 1.
  const grid = { startRow: 0, startCol: 0, colOffsets: [0, 80], rowIndices: [0, 3, 4, 7],
    rowOffsets: [0, 80, 160, 240, 320],
    hiddenRowSpans: [[1, 2], [5, 6]] as [number, number][] };
  const { container } = mount(handlers, 1, grid, wholeRows(1, 6));
  const band = container.querySelector('[data-testid="xlsx-row-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(rowLabel(container, 4), { button: 0, clientX: 0, clientY: 200, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 0, clientY: 300, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 0, clientY: 300, pointerId: 1 });
  expect(reorders).toEqual([{ axis: 'row', at: 1, before: 8, count: 6 }]);
});

test('a drag from inside a whole-row selection moves the whole run', () => {
  const { container } = mount(handlers, 1, { startRow: 0, startCol: 0,
    colOffsets: [0, 80], rowOffsets: [0, 20, 40, 60, 80] }, wholeRows(0, 2));
  const band = container.querySelector('[data-testid="xlsx-row-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(rowLabel(container, 1), { button: 0, clientX: 0, clientY: 30, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 0, clientY: 75, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 0, clientY: 75, pointerId: 1 });
  // midpoints 10, 30, 50, 70: 75 is past the fourth, so the gap is 4.
  expect(reorders).toEqual([{ axis: 'row', at: 0, before: 4, count: 3 }]);
});

test('a drag on a label a cell range only touches extends, never reorders', () => {
  // B2:C3 is not a whole-track selection: its column extent stops at C.
  const selection: Selection = {
    anchor: { row: 1, col: 1 }, focus: { row: 2, col: 2 },
    range: { top: 1, left: 1, bottom: 2, right: 2 },
  };
  const { container } = mount(handlers, 1, narrow, selection);
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 75, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 75, clientY: 0, pointerId: 1 });
  expect(selects).toEqual([{ axis: 'col', index: 2 }]);
  expect(reorders).toEqual([]);
});

test('a shift press on a selected label extends instead of reordering', () => {
  const { container } = mount(handlers, 1, narrow, wholeColumn(0));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(container, 0), {
    button: 0, clientX: 15, clientY: 0, pointerId: 1, shiftKey: true });
  // the shifted press extends to the pressed track; the anchor stays at A.
  fireEvent.mouseDown(columnLabel(container, 0), { shiftKey: true });
  fireEvent.pointerMove(band, { clientX: 45, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 75, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 75, clientY: 0, pointerId: 1 });
  expect(selects).toEqual([{ axis: 'col', index: 1 }, { axis: 'col', index: 2 }]);
  expect(reorders).toEqual([]);
  expect(presses).toEqual([{ axis: 'col', index: 0, shift: true }]);
});

test('a row selection drag extends down the row band', () => {
  const { container } = mount(handlers, 1, { startRow: 0, startCol: 0,
    colOffsets: [0, 80], rowOffsets: [0, 20, 40, 60] });
  const band = container.querySelector('[data-testid="xlsx-row-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(rowLabel(container, 0), { button: 0, clientX: 0, clientY: 10, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 0, clientY: 45, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 0, clientY: 45, pointerId: 1 });
  expect(selects).toEqual([{ axis: 'row', index: 2 }]);
  expect(reorders).toEqual([]);
});

// --- header band defects reproduced from the reviewer's report ---------------

test('track dividers snap to the canvas gridline at fractional zoom', () => {
  // 102 * 1.25 = 127.5: the canvas rounds the boundary to 128 and strokes
  // [127, 128], so the band hairline must not sit at 126.5.
  const { container } = mount(handlers, 1.25, { startRow: 0, startCol: 0,
    colOffsets: [0, 102, 204], rowOffsets: [0, 20] });
  const dividers = Array.from(container.querySelectorAll<HTMLElement>(
    '[data-testid="xlsx-column-headers"] [data-track-divider]'), (divider) => divider.style.left);
  expect(dividers).toEqual(['127px', '254px']);
});

test('a frozen split snaps to the canvas gridline at fractional zoom', () => {
  const { container } = mount(handlers, 1.25, { startRow: 0, startCol: 0,
    frozenCols: 1, dividerX: 102, colOffsets: [0, 102, 204], rowOffsets: [0, 20] });
  const split = container.querySelector<HTMLElement>(
    '[data-testid="xlsx-column-headers"] [data-testid="xlsx-freeze-divider"]')!;
  // the 2px split is centred on the snapped 128: [127, 129].
  expect([split.style.left, split.style.width]).toEqual(['127px', '2px']);
});

test('hidden-span markers snap to whole pixels at fractional zoom', () => {
  const { container } = mount(handlers, 1.25, { startRow: 0, startCol: 0,
    colIndices: [0, 3], colOffsets: [0, 102, 204], rowOffsets: [0, 20],
    hiddenColSpans: [[1, 2]] });
  const markers = Array.from(container.querySelectorAll<HTMLElement>(
    '[data-testid="xlsx-column-headers"] [data-testid="xlsx-hidden-divider"]'), (marker) => marker.style.left);
  expect(markers).toEqual(['126px', '130px']);
});

test('the resize guide snaps to a whole pixel at fractional zoom', () => {
  const { container } = mount(handlers, 1.25, { startRow: 0, startCol: 0,
    colOffsets: [0, 102, 204], rowOffsets: [0, 20] });
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 127.5, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 137.5, clientY: 0, pointerId: 1 });
  const guide = container.querySelector<HTMLElement>('[data-testid="xlsx-resize-guide"]')!;
  // the dragged boundary is 110 sheet px = 137.5 css px, painted at 138.
  expect(guide.style.left).toBe('138px');
});

test('the reorder guide snaps to a whole pixel at fractional zoom', () => {
  const { container } = mount(handlers, 1.25, { startRow: 0, startCol: 0,
    colOffsets: [0, 25, 50, 75], rowOffsets: [0, 20] }, wholeColumn(0));
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(container, 0), { button: 0, clientX: 12.5, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 56.25, clientY: 0, pointerId: 1 });
  const guide = container.querySelector<HTMLElement>('[data-testid="xlsx-reorder-guide"]')!;
  // the gap after column B is 50 sheet px = 62.5 css px, painted at 63.
  expect(guide.style.left).toBe('63px');
});

test('a resize armed on one sheet does not commit after the sheet changes', () => {
  const view = mount(handlers, 1, geometry, null, 0);
  const band = view.container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 80, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 120, clientY: 0, pointerId: 1 });
  view.rerender(<SheetGrid frame={frame(geometry)} zoom={1} showSheetHeaders headerHandlers={handlers} sheet={1}>
    <div /></SheetGrid>);
  stubRect(band);
  fireEvent.pointerUp(band, { clientX: 120, clientY: 0, pointerId: 1 });
  expect(resizes).toEqual([]);
});

test('a resize armed without a sheet identity does not commit against a new frame', () => {
  // the band cannot name the painted sheet, so the armed frame itself is the
  // identity: a repaint mid-drag cancels instead of resizing another sheet.
  const view = mount(handlers);
  const band = view.container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 80, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 120, clientY: 0, pointerId: 1 });
  view.rerender(<SheetGrid frame={frame(geometry)} zoom={1} showSheetHeaders headerHandlers={handlers}>
    <div /></SheetGrid>);
  stubRect(band);
  fireEvent.pointerUp(band, { clientX: 120, clientY: 0, pointerId: 1 });
  expect(resizes).toEqual([]);
});

test('a reorder armed on one frame does not commit after the frame scrolls', () => {
  // the slot the press landed on maps to a different column once the band
  // repaints a scrolled window, so the release must not move anything.
  const view = mount(handlers, 1, { startRow: 0, startCol: 0,
    colOffsets: [0, 30, 60, 90], rowOffsets: [0, 20] }, wholeColumn(1), 0);
  const band = view.container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(view.container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 85, clientY: 0, pointerId: 1 });
  view.rerender(<SheetGrid frame={frame({ startRow: 0, startCol: 1, colIndices: [1, 2, 3],
    colOffsets: [0, 30, 60, 90], rowOffsets: [0, 20] })} zoom={1} showSheetHeaders headerHandlers={handlers} sheet={0}>
    <div /></SheetGrid>);
  stubRect(band);
  fireEvent.pointerUp(band, { clientX: 85, clientY: 0, pointerId: 1 });
  expect(reorders).toEqual([]);
});

test('a reorder armed on one sheet does not commit after the sheet changes', () => {
  const view = mount(handlers, 1, { startRow: 0, startCol: 0,
    colOffsets: [0, 30, 60, 90], rowOffsets: [0, 20] }, wholeColumn(1), 0);
  const band = view.container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(view.container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 85, clientY: 0, pointerId: 1 });
  view.rerender(<SheetGrid frame={frame({ startRow: 0, startCol: 0,
    colOffsets: [0, 30, 60, 90], rowOffsets: [0, 20] })} zoom={1} showSheetHeaders headerHandlers={handlers} sheet={1}>
    <div /></SheetGrid>);
  stubRect(band);
  fireEvent.pointerUp(band, { clientX: 85, clientY: 0, pointerId: 1 });
  expect(reorders).toEqual([]);
});

test('a reorder armed on one sheet survives an unrelated repaint', () => {
  // every repaint is a new frame object: hover, a selection echo and an op echo
  // all repaint without moving a single track. The sheet and the band geometry
  // decide whether the drag still belongs, so the gesture must not die on the
  // first pointer move after one of them.
  const grid = { startRow: 0, startCol: 0, colOffsets: [0, 30, 60, 90], rowOffsets: [0, 20] };
  const view = mount(handlers, 1, grid, wholeColumn(0), 0);
  const band = view.container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(view.container, 0), { button: 0, clientX: 15, clientY: 0, pointerId: 1 });
  view.rerender(
    <SheetGrid frame={frame({ ...grid })} zoom={1} showSheetHeaders headerHandlers={handlers}
      selection={wholeColumn(0)} sheet={0}>
      <div />
    </SheetGrid>
  );
  stubRect(band);
  fireEvent.pointerMove(band, { clientX: 55, clientY: 0, pointerId: 1 });
  expect(view.container.querySelector('[data-testid="xlsx-reorder-guide"]')).not.toBeNull();
  fireEvent.pointerUp(band, { clientX: 55, clientY: 0, pointerId: 1 });
  expect(reorders).toEqual([{ axis: 'col', at: 0, before: 2, count: 1 }]);
});

test('a reorder armed on one sheet is dropped once the band scrolls', () => {
  // the repaint keeps the sheet but moves the tracks: slot 0 is now column B,
  // so the lifted slot no longer names the pressed track.
  const view = mount(handlers, 1, { startRow: 0, startCol: 0,
    colOffsets: [0, 30, 60, 90], rowOffsets: [0, 20] }, wholeColumn(1), 0);
  const band = view.container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(view.container, 1), { button: 0, clientX: 45, clientY: 0, pointerId: 1 });
  view.rerender(
    <SheetGrid frame={frame({ startRow: 0, startCol: 1, colIndices: [1, 2, 3],
      colOffsets: [0, 30, 60, 90], rowOffsets: [0, 20] })} zoom={1} showSheetHeaders headerHandlers={handlers}
      selection={wholeColumn(1)} sheet={0}>
      <div />
    </SheetGrid>
  );
  stubRect(band);
  fireEvent.pointerMove(band, { clientX: 85, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 85, clientY: 0, pointerId: 1 });
  expect(reorders).toEqual([]);
  expect(view.container.querySelector('[data-testid="xlsx-reorder-guide"]')).toBeNull();
});

test('a resize armed on one sheet is dropped once the band scrolls', () => {
  // a scroll under a boundary drag would otherwise commit the armed drag's size
  // onto whichever track now starts at that boundary.
  const view = mount(handlers, 1, { startRow: 0, startCol: 0,
    colOffsets: [0, 80, 160, 240], rowOffsets: [0, 20] }, null, 0);
  const band = view.container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 80, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 120, clientY: 0, pointerId: 1 });
  view.rerender(
    <SheetGrid frame={frame({ startRow: 0, startCol: 1, colIndices: [1, 2, 3],
      colOffsets: [0, 80, 160, 240], rowOffsets: [0, 20] })} zoom={1} showSheetHeaders headerHandlers={handlers} sheet={0}>
      <div />
    </SheetGrid>
  );
  stubRect(band);
  fireEvent.pointerUp(band, { clientX: 120, clientY: 0, pointerId: 1 });
  expect(resizes).toEqual([]);
});

test('a selection drag is dropped once the band scrolls', () => {
  // the pointer reaches the slot the scroll moved, not the track the press was
  // measured against, so no extend is reported for it.
  const view = mount(handlers, 1, { startRow: 0, startCol: 0,
    colOffsets: [0, 80, 160, 240], rowOffsets: [0, 20] });
  const band = view.container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(view.container, 0), { button: 0, clientX: 20, clientY: 0, pointerId: 1 });
  view.rerender(
    <SheetGrid frame={frame({ startRow: 0, startCol: 1, colIndices: [1, 2, 3],
      colOffsets: [0, 80, 160, 240], rowOffsets: [0, 20] })} zoom={1} showSheetHeaders headerHandlers={handlers}>
      <div />
    </SheetGrid>
  );
  stubRect(band);
  fireEvent.pointerMove(band, { clientX: 200, clientY: 0, pointerId: 1 });
  fireEvent.pointerUp(band, { clientX: 200, clientY: 0, pointerId: 1 });
  expect(selects).toEqual([]);
});

test('a resize armed on one sheet commits after that sheet repaints', () => {
  // a repaint is a new frame object for the same sheet: the armed sheet, not
  // the frame's identity, decides whether the gesture still belongs. This is
  // the path a surface's `sheet` prop feeds; without it the repaint cancels.
  const view = mount(handlers, 1, geometry, null, 0);
  const band = view.container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 80, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 120, clientY: 0, pointerId: 1 });
  view.rerender(<SheetGrid frame={frame(geometry)} zoom={1} showSheetHeaders headerHandlers={handlers} sheet={0}>
    <div /></SheetGrid>);
  stubRect(band);
  fireEvent.pointerUp(band, { clientX: 120, clientY: 0, pointerId: 1 });
  expect(resizes).toEqual([{ axis: 'col', index: 0, size: 120 }]);
});

test('a reorder armed on one sheet commits after a re-render of that sheet', () => {
  // the sheet identity outlives the render pass. The frame is held identical on
  // purpose: a reorder's gap is measured against the armed tracks, so the band
  // still drops the move when the painted frame itself changes (the scroll
  // test above), even on the same sheet.
  const grid = { startRow: 0, startCol: 0, colOffsets: [0, 30, 60, 90], rowOffsets: [0, 20] };
  const painted = frame(grid);
  const tree = (
    <SheetGrid frame={painted} zoom={1} showSheetHeaders headerHandlers={handlers} selection={wholeColumn(0)} sheet={0}>
      <div />
    </SheetGrid>
  );
  presses.length = 0;
  resizes.length = 0;
  reorders.length = 0;
  selects.length = 0;
  const view = render(tree);
  const band = view.container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(columnLabel(view.container, 0), { button: 0, clientX: 15, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 55, clientY: 0, pointerId: 1 });
  view.rerender(tree);
  stubRect(band);
  fireEvent.pointerUp(band, { clientX: 55, clientY: 0, pointerId: 1 });
  expect(reorders).toEqual([{ axis: 'col', at: 0, before: 2, count: 1 }]);
});

test('without onHeaderResize a press near a boundary selects the track', () => {
  // a viewer supplies only the press and the selection drag: a press inside the
  // resize slop must not be swallowed by a gesture nothing can commit.
  const viewer: SheetGridHeaderHandlers = {
    onHeaderMouseDown: handlers.onHeaderMouseDown,
    onHeaderDragSelect: handlers.onHeaderDragSelect,
  };
  const { container } = mount(viewer);
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 79, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 120, clientY: 0, pointerId: 1 });
  expect(container.querySelector('[data-testid="xlsx-resize-guide"]')).toBeNull();
  fireEvent.mouseDown(columnLabel(container, 0), {});
  fireEvent.pointerUp(band, { clientX: 120, clientY: 0, pointerId: 1 });
  expect(presses).toEqual([{ axis: 'col', index: 0, shift: false }]);
  expect(resizes).toEqual([]);
});

test('a resize bases on the header track span, not the label box', () => {
  // the frozen split leaves only a narrow slice of a wide column on screen; the
  // gesture must still start from the track's size, and preview from it too.
  const { container } = mount(handlers, 1, { startRow: 0, startCol: 0,
    frozenCols: 1, dividerX: 80, colOffsets: [0, 80, 280, 380], rowOffsets: [0, 20] });
  const label = columnLabel(container, 1) as HTMLElement;
  const clipped = { left: 80, top: 0, width: 40, height: 24, right: 120, bottom: 24, x: 80, y: 0 };
  label.getBoundingClientRect = () => clipped as DOMRect;
  const band = container.querySelector('[data-testid="xlsx-column-headers"]')!;
  stubRect(band);
  fireEvent.pointerDown(band, { button: 0, clientX: 280, clientY: 0, pointerId: 1 });
  fireEvent.pointerMove(band, { clientX: 290, clientY: 0, pointerId: 1 });
  expect(label.style.width).toBe('210px');
  fireEvent.pointerUp(band, { clientX: 290, clientY: 0, pointerId: 1 });
  expect(resizes).toEqual([{ axis: 'col', index: 1, size: 210 }]);
});
