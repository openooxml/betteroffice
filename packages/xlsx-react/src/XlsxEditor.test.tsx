/**
 * Grid pointer behaviour against the real wasm core: happy-dom has no layout or
 * canvas, so the viewport size and the 2d context are stubbed and click points
 * come from the same display-list geometry the editor hit-tests.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { cellRect, initWasm, openWorkbook } from '@betteroffice/xlsx';
import type { CellAddr, GridMeta, WorkbookHandle } from '@betteroffice/xlsx';
import { XlsxEditor } from './XlsxEditor';

const WASM = resolve(import.meta.dir, '../../xlsx/src/wasm/generated/xlsx_wasm_bg.wasm');
const FIXTURE = resolve(import.meta.dir, '../../xlsx/test-fixtures/sample.xlsx');
const VIEWPORT = { width: 800, height: 600 };
const LINK_TARGET = 'https://example.com/report';
const LINK_CELL: CellAddr = { row: 5, col: 4 };

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const { cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');

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
}

function fixtureFrom(bytes: Uint8Array): Fixture {
  const probe = openWorkbook(bytes);
  try {
    return { bytes, grid: probe.displayList({ x: 0, y: 0, ...VIEWPORT }).grid as GridMeta };
  } finally {
    probe.dispose();
  }
}

let plain: Fixture;
let linked: Fixture;
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

async function mountEditor(fixture: Fixture = plain) {
  const ready: { handle: WorkbookHandle | null } = { handle: null };
  const view = render(
    <XlsxEditor
      file={fixture.bytes.slice()}
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
    nameBox,
    editor,
    workbook: () => ready.handle!,
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
  };
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
