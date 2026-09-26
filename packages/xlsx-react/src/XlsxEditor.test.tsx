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
import { cellRect, initWasm, openWorkbook, selectionAt } from '@betteroffice/xlsx';
import type { CellAddr, ChartRegion, GridMeta, WorkbookHandle } from '@betteroffice/xlsx';
import { XlsxEditor, type XlsxEditorApi } from './XlsxEditor';

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
    canUndo: () => !(view.getByTestId('xlsx-undo') as HTMLButtonElement).disabled,
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
      expect(String(await settled(() => outcome))).toMatch(/changed while flushing/);
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
      expect(await settled(() => current.version())).toBe(current.handle.version());
      expect(String(await settled(() => outcome))).toMatch(/changed while flushing/);
    } finally {
      clipboard.restore();
    }
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

  it('refuses writes while read-only and waits out text composition', async () => {
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
    fireEvent.compositionStart(editor);
    const composing = await failed(() => api.applyEdits(setB3(api.handle.version(), '1')));
    expect(String(composing)).toMatch(/composition/);
    fireEvent.compositionEnd(editor);
    fireEvent.change(editor, { target: { value: '55' } });
    const version = await settled(() => api.version());
    expect(api.handle.cell(0, 2, 1).input).toBe('55');
    const validated = await settled(() => api.validateEdits(setB3(version, '66')));
    expect(validated.ok && validated.wouldApply).toBe(true);
  });
});
