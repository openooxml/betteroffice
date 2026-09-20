import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import type { DiagramHandle, DiagramSnapshot } from '@betteroffice/vsdx';
import { createT, en } from '@betteroffice/vsdx-i18n';
import { useState } from 'react';
import type { RefObject } from 'react';
import { ShapeContextMenu } from './components/ribbon/ShapeContextMenu';
import { CanvasContextMenu } from './components/ribbon/CanvasContextMenu';
import { RibbonCommandsContext, RibbonCommandsProvider } from './components/ribbon/commands';
import type { RibbonCommand, RibbonCommandId, RibbonCommands } from './components/ribbon/commands';
import { ShapeMiniToolbar, miniToolbarPosition } from './components/ribbon/ShapeMiniToolbar';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const { cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');

type Selection = { pageId: string; shapeId: string; hit: { kind: 'shape'; shapeId: string } };

interface Calls {
  deletes: unknown[][];
  reorders: unknown[][];
  formulas: Array<{ pageId: string; shapeId: string; cell: string; formula: string }>;
}

function cell(name: string, value: string) {
  return { locator: { sheet: { page: 1 }, shapeId: 1, section: null, row: null, cellName: name }, name, formula: value, value };
}

function snapshot(cells: Array<ReturnType<typeof cell>>): DiagramSnapshot {
  return { pages: [{ id: 'page', sourcePartPath: 'page', name: 'Page', shapes: [{ id: 'three', sourceId: 1, name: 'three', children: [], cells }] }] };
}

function stubHandle(state: DiagramSnapshot, calls: Calls, refused: readonly string[] = []): DiagramHandle {
  return {
    snapshot: () => state,
    canUndo: () => false,
    canRedo: () => false,
    deleteShape: (...args: [string, string]) => { calls.deletes.push([...args]); return {}; },
    deleteShapes: (deletes: ReadonlyArray<{ pageId: string; shapeId: string }>) => { calls.deletes.push([...deletes]); return []; },
    reorderShape: (...args: [string, string, number]) => { calls.reorders.push([...args]); return {}; },
    setCellFormula: (pageId: string, shapeId: string, locator: { cellName: string }, formula: string) => {
      calls.formulas.push({ pageId, shapeId, cell: locator.cellName, formula });
      return {};
    },
    probeCellWrites: (_pageId: string, _shapeId: string, queries: ReadonlyArray<{ cellName: string }>) =>
      queries.map(({ cellName }) => refused.includes(cellName)
        ? { cellName, allowed: false, targetCellName: null, refusal: 'guard', reason: 'GUARD protects the requested cell' }
        : { cellName, allowed: true, targetCellName: cellName, refusal: null, reason: null }),
    setCellFormulas: (writes: ReadonlyArray<{ pageId: string; shapeId: string; cellName: string; formula: string }>) => {
      for (const write of writes) calls.formulas.push({ pageId: write.pageId, shapeId: write.shapeId, cell: write.cellName, formula: write.formula });
      return [];
    },
  } as unknown as DiagramHandle;
}

function renderShapeMenu(options: { cells?: Array<ReturnType<typeof cell>>; refused?: readonly string[]; position?: { top: number; left: number }; withFocusTarget?: boolean } = {}) {
  cleanup();
  const calls: Calls = { deletes: [], reorders: [], formulas: [] };
  const state = snapshot(options.cells ?? [cell('FillForegnd', 'RGB(255,0,0)'), cell('LineColor', 'RGB(0,0,255)'), cell('Angle', '0'), cell('FlipX', '0'), cell('FlipY', '0')]);
  const diagram = stubHandle(state, calls, options.refused ?? []);
  const closed: string[] = [];
  const selection: Selection[] = [{ pageId: 'page', shapeId: 'three', hit: { kind: 'shape', shapeId: 'three' } }];
  let focusTarget: HTMLElement | undefined;
  if (options.withFocusTarget) {
    focusTarget = document.createElement('button');
    document.body.appendChild(focusTarget);
  }
  const rendered = render(<MenuHost diagram={diagram} selection={selection} position={options.position ?? { top: 200, left: 100 }} closed={closed} focusTarget={focusTarget} />);
  return { view: rendered, calls, closed, focusTarget, selection };
}

function MenuHost({ diagram, selection, position, closed, focusTarget }: { diagram: DiagramHandle; selection: Selection[]; position: { top: number; left: number }; closed: string[]; focusTarget?: HTMLElement }) {
  const [open, setOpen] = useState(true);
  if (!open) return null;
  return (
    <RibbonCommandsProvider handle={diagram} snapshot={diagram.snapshot()} pageId="page" selection={selection} onMutation={() => {}} onError={() => {}} onDownload={() => {}}>
      <ShapeContextMenu t={createT(en)} position={position} onClose={() => { closed.push('close'); setOpen(false); }} onCloseAndFocus={() => { closed.push('focus'); setOpen(false); focusTarget?.focus(); }} />
    </RibbonCommandsProvider>
  );
}

function toolbar(): HTMLElement | null {
  return document.querySelector('[role="toolbar"]') as HTMLElement | null;
}

function shapeMenu(): HTMLElement | null {
  const menus = Array.from(document.querySelectorAll('[role="menu"]')) as HTMLElement[];
  return menus.find((menu) => menu.getAttribute('aria-label') === en.contextMenu.label) ?? null;
}

afterEach(() => {
  cleanup();
  document.body.innerHTML = '';
});

test('the shape menu carries a fill and line colour picker above it', () => {
  const originalRect = HTMLElement.prototype.getBoundingClientRect;
  HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
    if (this.getAttribute?.('role') === 'toolbar') return { x: 0, y: 0, width: 80, height: 40, top: 0, left: 0, right: 80, bottom: 40, toJSON: () => ({}) } as DOMRect;
    if (this.getAttribute?.('role') === 'menu' && !this.hasAttribute('data-submenu')) return { x: 0, y: 0, width: 220, height: 200, top: 200, left: 100, right: 320, bottom: 400, toJSON: () => ({}) } as DOMRect;
    return originalRect.call(this);
  };
  const { view } = renderShapeMenu();
  try {
    const bar = toolbar();
    expect(bar).not.toBeNull();
    expect(bar?.getAttribute('aria-label')).toBe(en.contextMenu.miniToolbarLabel);
    const fill = bar?.querySelector('[data-command-id="fillColor"]') as HTMLInputElement;
    const line = bar?.querySelector('[data-command-id="lineColor"]') as HTMLInputElement;
    expect(fill?.getAttribute('type')).toBe('color');
    expect(line?.getAttribute('type')).toBe('color');
    expect(fill?.value).toBe('#ff0000');
    expect(line?.value).toBe('#0000ff');
    expect(bar?.getAttribute('data-below')).toBe('false');
    expect(bar?.style.top).toBe('156px');
    expect(bar?.style.left).toBe('100px');
    const menu = shapeMenu() as HTMLElement;
    expect(menu).not.toBeNull();
    expect(Number.parseFloat(bar?.style.top ?? '0')).toBeLessThan(Number.parseFloat(menu.style.top));
  } finally {
    HTMLElement.prototype.getBoundingClientRect = originalRect;
    view.unmount();
  }
});

test('the toolbar flips below the menu near the top edge', () => {
  const originalRect = HTMLElement.prototype.getBoundingClientRect;
  HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
    if (this.getAttribute?.('role') === 'toolbar') return { x: 0, y: 0, width: 80, height: 40, top: 0, left: 0, right: 80, bottom: 40, toJSON: () => ({}) } as DOMRect;
    if (this.getAttribute?.('role') === 'menu' && !this.hasAttribute('data-submenu')) return { x: 0, y: 0, width: 220, height: 200, top: 4, left: 100, right: 320, bottom: 204, toJSON: () => ({}) } as DOMRect;
    return originalRect.call(this);
  };
  const { view } = renderShapeMenu({ position: { top: 10, left: 100 } });
  try {
    const bar = toolbar() as HTMLElement;
    expect(bar).not.toBeNull();
    expect(bar.getAttribute('data-below')).toBe('true');
    expect(bar.style.top).toBe('208px');
    const menu = shapeMenu() as HTMLElement;
    expect(Number.parseFloat(bar.style.top)).toBeGreaterThanOrEqual(Number.parseFloat(menu.style.top));
  } finally {
    HTMLElement.prototype.getBoundingClientRect = originalRect;
    view.unmount();
  }
});

test('the toolbar follows the menu after the menu clamps into the viewport', async () => {
  const originalRect = HTMLElement.prototype.getBoundingClientRect;
  HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
    if (this.getAttribute?.('role') === 'toolbar') return { x: 0, y: 0, width: 80, height: 40, top: 0, left: 0, right: 80, bottom: 40, toJSON: () => ({}) } as DOMRect;
    if (this.getAttribute?.('role') === 'menu' && !this.hasAttribute('data-submenu')) {
      const top = Number.parseFloat(this.style.top || '0');
      const left = Number.parseFloat(this.style.left || '0');
      return { x: left, y: top, width: 220, height: 340, top, left, right: left + 220, bottom: top + 340, toJSON: () => ({}) } as DOMRect;
    }
    return originalRect.call(this);
  };
  const { view } = renderShapeMenu({ position: { top: 700, left: 100 } });
  try {
    const clamped = window.innerHeight - 340 - 4;
    await waitFor(() => expect((shapeMenu() as HTMLElement).style.top).toBe(`${clamped}px`));
    await waitFor(() => expect((toolbar() as HTMLElement).style.top).toBe(`${clamped - 44}px`));
    expect(toolbar()?.getAttribute('data-below')).toBe('false');
  } finally {
    HTMLElement.prototype.getBoundingClientRect = originalRect;
    view.unmount();
  }
});

test('a colour pick runs the ribbon command and keeps the menu open', () => {
  const { view, calls, closed } = renderShapeMenu();
  try {
    const fill = toolbar()?.querySelector('[data-command-id="fillColor"]') as HTMLInputElement;
    expect(fill).not.toBeNull();
    fireEvent.change(fill, { target: { value: '#00ff00' } });
    expect(calls.formulas).toEqual([{ pageId: 'page', shapeId: 'three', cell: 'FillForegnd', formula: 'RGB(0,255,0)' }]);
    expect(shapeMenu()).not.toBeNull();
    expect(toolbar()).not.toBeNull();
    expect(closed).toEqual([]);
  } finally {
    view.unmount();
  }
});

test('a press on the toolbar does not dismiss the menu', () => {
  const { view, closed } = renderShapeMenu();
  try {
    const fill = toolbar()?.querySelector('[data-command-id="fillColor"]') as HTMLInputElement;
    fireEvent.mouseDown(fill);
    expect(shapeMenu()).not.toBeNull();
    expect(toolbar()).not.toBeNull();
    expect(closed).toEqual([]);
  } finally {
    view.unmount();
  }
});

test('escape closes the toolbar together with the menu', () => {
  const { view, closed, focusTarget } = renderShapeMenu({ withFocusTarget: true });
  try {
    const fill = toolbar()?.querySelector('[data-command-id="fillColor"]') as HTMLInputElement;
    fireEvent.keyDown(fill, { key: 'Escape' });
    expect(shapeMenu()).toBeNull();
    expect(toolbar()).toBeNull();
    expect(closed).toEqual(['focus']);
    expect(document.activeElement).toBe(focusTarget as HTMLElement);
  } finally {
    view.unmount();
  }
});

test('an outside press and a menu action close the toolbar with the menu', () => {
  const first = renderShapeMenu();
  try {
    fireEvent.mouseDown(document.body);
    expect(shapeMenu()).toBeNull();
    expect(toolbar()).toBeNull();
    expect(first.closed).toEqual(['close']);
  } finally {
    first.view.unmount();
  }
  const second = renderShapeMenu();
  try {
    const item = document.querySelector('[role="menu"] [data-command-id="delete"]') as HTMLElement;
    fireEvent.click(item);
    expect(second.calls.deletes).toEqual([[{ pageId: 'page', shapeId: 'three' }]]);
    expect(shapeMenu()).toBeNull();
    expect(toolbar()).toBeNull();
    expect(second.closed).toEqual(['focus']);
  } finally {
    second.view.unmount();
  }
});

test('guarded colour cells hide the mini toolbar instead of refusing on pick', () => {
  const { view } = renderShapeMenu({ cells: [cell('FillForegnd', 'GUARD(RGB(255,0,0))'), cell('LineColor', 'GUARD(RGB(0,0,255))'), cell('Angle', '0'), cell('FlipX', '0'), cell('FlipY', '0')], refused: ['FillForegnd', 'LineColor'] });
  try {
    expect(shapeMenu()).not.toBeNull();
    expect(toolbar()).toBeNull();
  } finally {
    view.unmount();
  }
});

test('a SETATREF redirect to a guarded cell hides only that swatch', () => {
  const redirected = renderShapeMenu({ cells: [cell('FillForegnd', 'RGB(255,0,0)'), cell('LineColor', 'SETATREF(LineTarget)'), cell('LineTarget', 'GUARD(RGB(0,0,255))')], refused: ['LineColor'] });
  try {
    const bar = toolbar();
    expect(bar).not.toBeNull();
    expect(bar?.querySelector('[data-command-id="fillColor"]')).not.toBeNull();
    expect(bar?.querySelector('[data-command-id="lineColor"]')).toBeNull();
  } finally {
    redirected.view.unmount();
  }
});

test('commands that do not exist or are disabled are not shown', () => {
  function PartialHost({ commands }: { commands: Partial<Record<RibbonCommandId, RibbonCommand | undefined>> }) {
    const ref = { current: null } as RefObject<HTMLDivElement | null>;
    return (
      <RibbonCommandsContext.Provider value={commands as RibbonCommands}>
        <ShapeMiniToolbar t={createT(en)} toolbarRef={ref} />
      </RibbonCommandsContext.Provider>
    );
  }
  const fillOnly: Partial<Record<RibbonCommandId, RibbonCommand | undefined>> = {
    fillColor: { id: 'fillColor', enabled: true, value: '#ff0000', run: () => {} },
  };
  const first = render(<PartialHost commands={fillOnly} />);
  try {
    const bar = toolbar();
    expect(bar).not.toBeNull();
    expect(bar?.querySelector('[data-command-id="fillColor"]')).not.toBeNull();
    expect(bar?.querySelector('[data-command-id="lineColor"]')).toBeNull();
    expect(bar?.querySelector('[disabled]')).toBeNull();
  } finally {
    first.unmount();
  }
  const disabled: Partial<Record<RibbonCommandId, RibbonCommand | undefined>> = {
    fillColor: { id: 'fillColor', enabled: false, value: '#ff0000', run: () => {} },
  };
  const second = render(<PartialHost commands={disabled} />);
  try {
    expect(toolbar()).toBeNull();
  } finally {
    second.unmount();
  }
});

test('the canvas menu has no mini toolbar', () => {
  cleanup();
  const calls: Calls = { deletes: [], reorders: [], formulas: [] };
  const state = snapshot([cell('Angle', '0')]);
  const diagram = stubHandle(state, calls);
  const view = render(
    <RibbonCommandsProvider handle={diagram} snapshot={diagram.snapshot()} pageId="page" selection={[]} onMutation={() => {}} onError={() => {}} onDownload={() => {}}>
      <CanvasContextMenu t={createT(en)} position={{ top: 300, left: 300 }} onClose={() => {}} onCloseAndFocus={() => {}} />
    </RibbonCommandsProvider>,
  );
  try {
    expect(toolbar()).toBeNull();
    expect(document.querySelector('[role="menu"]')).not.toBeNull();
  } finally {
    view.unmount();
  }
});

test('miniToolbarPosition clamps to the viewport and flips only when needed', () => {
  const viewport = { width: 1024, height: 768 };
  expect(miniToolbarPosition({ top: 200, left: 100, bottom: 400, right: 320 }, { width: 80, height: 40 }, viewport)).toEqual({ top: 156, left: 100, below: false });
  expect(miniToolbarPosition({ top: 4, left: 100, bottom: 204, right: 320 }, { width: 80, height: 40 }, viewport)).toEqual({ top: 208, left: 100, below: true });
  expect(miniToolbarPosition({ top: 200, left: 1000, bottom: 400, right: 1220 }, { width: 80, height: 40 }, viewport)).toEqual({ top: 156, left: 940, below: false });
  expect(miniToolbarPosition({ top: 4, left: 100, bottom: 760, right: 320 }, { width: 80, height: 40 }, viewport)).toEqual({ top: 724, left: 100, below: true });
});
