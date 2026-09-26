/**
 * Grid and chart pointer behaviour against the real wasm core: happy-dom has no
 * layout or canvas, so the viewport size and the 2d context are stubbed and
 * click points come from the same display-list geometry the editor hit-tests.
 *
 * Both suites live here on purpose. happy-dom registers into one shared process
 * global, so a second file with its own register/unregister pair would tear the
 * dom down under whichever suite bun happens to run second.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  cellRect,
  initWasm,
  isPngExportAvailable,
  openWorkbook,
  selectionAt,
} from '@betteroffice/xlsx';
import type {
  CellAddr,
  ChartRegion,
  GridMeta,
  WorkbookHandle,
  XlsxEditResult,
} from '@betteroffice/xlsx';
import { isMacPlatform } from './commands/descriptors';
import { useXlsxCommands } from './commands/hooks';
import type {
  XlsxCommandArgs,
  XlsxCommandId,
  XlsxCommandResult,
  XlsxCommandStore,
} from './commands/types';
import { EditorToolbar } from './components/EditorToolbar';
import { ToolbarCommandButton } from './components/toolbar/ToolbarCommand';
import { XlsxCommandAdmissionError } from './index';
import {
  XlsxEditor,
  XlsxSaveRefusedError,
  type XlsxEditorApi,
  type XlsxEditorProps,
} from './XlsxEditor';

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
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');

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
});

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

afterEach(() => {
  cleanup();
  opened = [];
});

function pointAt(fixture: Fixture, addr: CellAddr): { clientX: number; clientY: number } {
  const rect = cellRect(fixture.grid, addr.row, addr.col);
  if (!rect) throw new Error(`cell ${addr.row},${addr.col} is outside the painted window`);
  return { clientX: rect.x + rect.w / 2, clientY: rect.y + rect.h / 2 };
}

async function mountEditor(fixture: Fixture = plain, onSave?: (bytes: Uint8Array) => void) {
  const ready: { handle: WorkbookHandle | null } = { handle: null };
  const view = render(
    <XlsxEditor
      file={fixture.bytes.slice()}
      onSave={onSave}
      onReady={(api) => {
        ready.handle = api.handle;
      }}
    />
  );
  const nameBox = () => view.getByTestId('xlsx-name-box') as HTMLInputElement;
  await waitFor(() => expect(nameBox().value).toBe('A1'));
  const surface = view.getByTestId('xlsx-scroll');
  const editor = () => view.queryByTestId('xlsx-cell-editor') as HTMLInputElement | null;
  const press = (addr: CellAddr) => {
    fireEvent.mouseDown(surface, pointAt(fixture, addr));
    fireEvent.mouseUp(surface, pointAt(fixture, addr));
    fireEvent.click(surface, pointAt(fixture, addr));
  };
  return {
    surface,
    nameBox,
    editor,
    workbook: () => ready.handle!,
    reopenWith: (next: Fixture) =>
      act(async () => {
        view.rerender(
          <XlsxEditor
            file={next.bytes.slice()}
            onSave={onSave}
            onReady={(api) => {
              ready.handle = api.handle;
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
    for (let press = 0; press < 5; press++) {
      await act(async () => {
        fireEvent.keyDown(view.surface, { key: 'ArrowRight' });
      });
    }
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
      <XlsxEditor
        file={plain.bytes.slice()}
        onChange={() => changes++}
        onReady={(ready) => {
          api = ready;
        }}
        readOnly
      />
    );
    await waitFor(() => expect(api).toBeDefined());
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
      <XlsxEditor
        file={plain.bytes.slice()}
        onChange={() => changes++}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined());
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
      const view = render(<XlsxEditor file={file} onReady={onReady} />);
      await waitFor(() => expect(api).toBeDefined());
      const target = { row: 2, col: 0 };
      const before = api!.handle.cell(0, target.row, target.col).input;
      await act(async () => {
        api!.selectCells(0, selectionAt(target));
      });

      fireEvent.keyDown(view.getByTestId('xlsx-scroll'), { key: 'v', ctrlKey: true });
      view.rerender(<XlsxEditor file={file} onReady={onReady} readOnly />);
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
});

describe('XlsxEditor proposal review', () => {
  it('reviews a staged proposal through accept, undo, reject, stale warning, and force apply', async () => {
    let api: XlsxEditorApi | undefined;
    const view = render(
      <XlsxEditor file={plain.bytes.slice()} onReady={(ready) => { api = ready; }} />
    );
    await waitFor(() => expect(api).toBeDefined());
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
      <XlsxEditor
        file={file}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined());
    const surface = view.getByTestId('xlsx-scroll');
    const chart = api!.handle.displayList({ x: 0, y: 0, ...VIEWPORT }).charts![0];
    fireEvent.mouseDown(surface, chartCenter(chart));
    fireEvent.mouseUp(window, chartCenter(chart));
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
      <XlsxEditor
        file={plain.bytes.slice()}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined());
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
          <XlsxEditor
            file={plain.bytes.slice()}
            onChange={() => {
              changes += 1;
            }}
            onReady={(ready) => {
              api = ready;
            }}
          />
        );
        await waitFor(() => expect(api).toBeDefined());
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
      <XlsxEditor
        file={plain.bytes.slice()}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined());
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

  async function mountCommands(props: Partial<XlsxEditorProps> = {}) {
    let api: XlsxEditorApi | undefined;
    let changes = 0;
    const saves: Uint8Array[] = [];
    const onReady = (ready: XlsxEditorApi) => {
      api = ready;
    };
    const element = (next: Partial<XlsxEditorProps> = {}) => (
      <XlsxEditor
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
    await waitFor(() => expect(api).toBeDefined());
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
      rerender: (next: Partial<XlsxEditorProps>) => view.rerender(element(next)),
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
      expect(reopened.cell(0, 2, 1).input).toBe('Saved draft');
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

  it('blocks every command while a rejected entry waits for correction on its own cell', async () => {
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

      fireEvent.keyDown(formula, { key: 'Enter' });
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
  });

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

  it('keeps a rejected entry on its own sheet and lets Escape discard it', async () => {
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
        fireEvent.click(editor.view.getAllByRole('tab')[0]);
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
        <XlsxEditor
          file={plain.bytes.slice()}
          onReady={(ready) => {
            api = ready;
          }}
        />
      );
      await waitFor(() => expect(api).toBeDefined());
      let printing: Promise<XlsxCommandResult> | undefined;
      view.rerender(
        <XlsxEditor
          file={charted.bytes.slice()}
          onReady={(ready) => {
            paints = 0;
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
      <XlsxEditor
        file={charted.bytes.slice()}
        onSave={(bytes) => {
          saves.push(bytes);
        }}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined());
    const surface = view.getByTestId('xlsx-scroll');
    const [chart] = charted.charts;
    fireEvent.mouseDown(surface, chartCenter(chart));
    fireEvent.mouseUp(window, chartCenter(chart));
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
        const editor = await mountCommands({ fileName: 'report.xlsx' });
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
        expect(await editor.execute('exportPng', null)).toEqual({ ok: true, status: 'executed' });
        expect(editor.input()).toBe('Exported');
        expect(clicks).toEqual(['report.png']);
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
    const view = render(<XlsxEditor onReady={(ready) => void (api = ready)} />);
    const store = await new Promise<XlsxCommandStore>((resolve) => {
      function Probe() {
        resolve(useXlsxCommands());
        return null;
      }
      view.rerender(<XlsxEditor toolbar={<Probe />} />);
    });
    const code = (state: { enabled: boolean; disabledReason?: { code: string } }) =>
      state.enabled ? null : state.disabledReason!.code;
    expect(code(store.getState('save'))).toBe('no-document');
    view.rerender(<XlsxEditor file={plain.bytes.slice()} onReady={(ready) => void (api = ready)} />);
    expect(code(store.getState('save'))).toBe('document-loading');
    await waitFor(() => expect(api).toBeDefined());
    await waitFor(() => expect(store.getState('save').enabled).toBe(true));
    expect(api!.commands).toBe(store);
  });

  it('keeps a failed Enter commit open in the editor', async () => {
    const editor = await mountCommands();
    const input = editor.typeInCell('Locked');
    const handle = editor.api().handle;
    const editCell = handle.editCell;
    handle.editCell = () => {
      throw new Error('cell is locked');
    };
    try {
      fireEvent.keyDown(input, { key: 'Enter' });
    } finally {
      handle.editCell = editCell;
    }
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
      <XlsxEditor
        file={charted.bytes.slice()}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined());
    const surface = view.getByTestId('xlsx-scroll');
    const [chart] = charted.charts;
    fireEvent.mouseDown(surface, chartCenter(chart));
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

  it('disables merging while collaborating', async () => {
    let api: XlsxEditorApi | undefined;
    render(
      <XlsxEditor
        file={plain.bytes.slice()}
        collaboration={{}}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await waitFor(() => expect(api).toBeDefined());
    await act(async () => {
      api!.selectCells(0, { anchor: { row: 0, col: 0 }, focus: { row: 1, col: 1 } });
    });
    const merge = api!.commands.getState('merge', { value: 'all' });
    expect(merge.enabled ? null : merge.disabledReason.code).toBe('collaboration-unsupported');
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
    expect(
      await editor.execute('proposalAccept', { proposalId: second, force: true })
    ).toEqual({ ok: true, status: 'executed' });
    expect(workbook.cell(0, 6, 4).input).toBe('42');
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
    const view = render(<XlsxEditor file={file} onReady={onReady} {...props} />);
    await waitFor(() => expect(api).toBeDefined());
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
        <XlsxEditor
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
      view.rerender(<XlsxEditor file={plain.bytes.slice()} onReady={onReady} />);
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
          view.rerender(<XlsxEditor file={plain.bytes.slice()} onReady={swap} />);
          await waitFor(() => expect(current).not.toBe(api));
        } else {
          view.rerender(<XlsxEditor file={file} onReady={onReady} readOnly />);
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
      view.rerender(<XlsxEditor file={file} onReady={onReady} readOnly />);
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
