import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import type { DiagramHandle, DiagramSnapshot } from '@betteroffice/vsdx';
import { createT, en } from '@betteroffice/vsdx-i18n';
import { useState } from 'react';
import { SHAPE_CONTEXT_ENTRIES, ShapeContextMenu } from './components/ribbon/ShapeContextMenu';
import { RibbonCommandsProvider } from './components/ribbon/commands';
import type { RibbonCommandId } from './components/ribbon/commands';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const { cleanup, fireEvent, render } = await import('@testing-library/react');

const LEAF_IDS: readonly RibbonCommandId[] = ['delete', 'bringToFront', 'bringForward', 'sendBackward', 'sendToBack', 'rotateRight', 'rotateLeft', 'flipHorizontal', 'flipVertical'];

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
  const ids = ['one', 'two', 'three', 'four', 'five'];
  return { pages: [{ id: 'page', sourcePartPath: 'page', name: 'Page', shapes: ids.map((id) => ({ id, sourceId: 1, name: id, children: [], cells })) }] };
}

function stubHandle(state: DiagramSnapshot, calls: Calls): DiagramHandle {
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
    setCellFormulas: (writes: ReadonlyArray<{ pageId: string; shapeId: string; cellName: string; formula: string }>) => {
      for (const write of writes) calls.formulas.push({ pageId: write.pageId, shapeId: write.shapeId, cell: write.cellName, formula: write.formula });
      return [];
    },
  } as unknown as DiagramHandle;
}

function Host({ diagram, selection, position, closed, focusTarget }: { diagram: DiagramHandle; selection: Selection[]; position: { top: number; left: number }; closed: string[]; focusTarget?: HTMLElement }) {
  const [open, setOpen] = useState(true);
  if (!open) return null;
  return (
    <RibbonCommandsProvider handle={diagram} snapshot={diagram.snapshot()} pageId="page" selection={selection} onMutation={() => {}} onError={() => {}} onDownload={() => {}}>
      <ShapeContextMenu t={createT(en)} position={position} onClose={() => { closed.push('close'); setOpen(false); }} onCloseAndFocus={() => { closed.push('focus'); setOpen(false); focusTarget?.focus(); }} />
    </RibbonCommandsProvider>
  );
}

function renderMenu(options: { cells?: Array<ReturnType<typeof cell>>; position?: { top: number; left: number }; withFocusTarget?: boolean; shapeId?: string } = {}) {
  cleanup();
  const calls: Calls = { deletes: [], reorders: [], formulas: [] };
  const state = snapshot(options.cells ?? [cell('Angle', '0'), cell('FlipX', '0'), cell('FlipY', '0')]);
  const diagram = stubHandle(state, calls);
  const closed: string[] = [];
  const shapeId = options.shapeId ?? 'three';
  const selection: Selection[] = [{ pageId: 'page', shapeId, hit: { kind: 'shape', shapeId } }];
  let focusTarget: HTMLElement | undefined;
  if (options.withFocusTarget) {
    focusTarget = document.createElement('button');
    document.body.appendChild(focusTarget);
  }
  const view = render(<Host diagram={diagram} selection={selection} position={options.position ?? { top: 100, left: 100 }} closed={closed} focusTarget={focusTarget} />);
  return { view, calls, closed, focusTarget, selection };
}

function parentMenu(): HTMLElement {
  const menus = Array.from(document.querySelectorAll('[role="menu"]')) as HTMLElement[];
  const found = menus.find((menu) => menu.getAttribute('aria-label') === en.contextMenu.label);
  expect(found).not.toBeUndefined();
  return found as HTMLElement;
}

function topLevelButtons(): HTMLElement[] {
  const menu = parentMenu();
  return Array.from(menu.querySelectorAll('button')).filter((button) => button.closest('[data-submenu]') === null) as HTMLElement[];
}

function topLevelId(button: HTMLElement): string {
  return button.getAttribute('data-submenu-id') ?? button.getAttribute('data-command-id') ?? '';
}

function openSubmenu(id: string): HTMLElement {
  const trigger = parentMenu().querySelector(`[data-submenu-id="${id}"]`) as HTMLElement;
  expect(trigger).not.toBeNull();
  fireEvent.click(trigger);
  const submenu = document.querySelector(`[data-submenu="${id}"]`) as HTMLElement;
  expect(submenu).not.toBeNull();
  return submenu;
}

afterEach(() => {
  cleanup();
  document.body.innerHTML = '';
});

test('the shape menu follows Visio order with a single divider and focuses delete', () => {
  const { view } = renderMenu();
  try {
    expect(SHAPE_CONTEXT_ENTRIES.map((entry) => entry.id)).toEqual(['delete', 'bringToFront', 'sendToBack', 'rotateRight']);
    expect(SHAPE_CONTEXT_ENTRIES[1].children?.map((child) => child.id)).toEqual(['bringToFront', 'bringForward']);
    expect(SHAPE_CONTEXT_ENTRIES[2].children?.map((child) => child.id)).toEqual(['sendBackward', 'sendToBack']);
    expect(SHAPE_CONTEXT_ENTRIES[3].children?.map((child) => child.id)).toEqual(['rotateRight', 'rotateLeft', 'flipHorizontal', 'flipVertical']);
    const menu = parentMenu();
    expect(menu.getAttribute('aria-label')).toBe(en.contextMenu.label);
    expect(topLevelButtons().map(topLevelId)).toEqual(['delete', 'bringToFront', 'sendToBack', 'rotateRight']);
    expect(menu.querySelectorAll('[role="separator"]')).toHaveLength(1);
    expect(menu.querySelector('[data-command-id="delete"]')).not.toBeNull();
    expect(menu.querySelector('[data-submenu-id="bringToFront"]')).not.toBeNull();
    expect(menu.querySelector('[data-submenu-id="sendToBack"]')).not.toBeNull();
    expect(menu.querySelector('[data-submenu-id="rotateRight"]')).not.toBeNull();
    for (const trigger of Array.from(menu.querySelectorAll('[data-submenu-id]'))) expect(trigger.textContent).toContain('▸');
    expect(menu.querySelector('[aria-keyshortcuts]')).toBeNull();
    expect(document.activeElement?.getAttribute('data-command-id')).toBe('delete');
    const submenu = openSubmenu('rotateRight');
    for (const id of ['rotateRight', 'rotateLeft', 'flipHorizontal', 'flipVertical']) expect(submenu.querySelector(`[data-command-id="${id}"]`)).not.toBeNull();
  } finally {
    view.unmount();
  }
});

test('each entry runs its command and closes the menu', () => {
  const last = (values: unknown[]): unknown => values[values.length - 1];
  const runLeaf = (submenu: string | null, id: RibbonCommandId) => {
    const { view, calls, closed } = renderMenu();
    try {
      let item: HTMLElement | null;
      if (submenu === null) {
        item = document.querySelector(`[role="menu"] [data-command-id="${id}"]`) as HTMLElement;
      } else {
        openSubmenu(submenu);
        item = document.querySelector(`[data-submenu="${submenu}"] [data-command-id="${id}"]`) as HTMLElement;
      }
      expect(item).not.toBeNull();
      fireEvent.click(item);
      if (id === 'delete') expect(calls.deletes).toEqual([[{ pageId: 'page', shapeId: 'three' }]]);
      else if (id === 'bringToFront') expect(last(calls.reorders)).toEqual(['page', 'three', 4]);
      else if (id === 'bringForward') expect(last(calls.reorders)).toEqual(['page', 'three', 3]);
      else if (id === 'sendBackward') expect(last(calls.reorders)).toEqual(['page', 'three', 1]);
      else if (id === 'sendToBack') expect(last(calls.reorders)).toEqual(['page', 'three', 0]);
      else if (id === 'rotateRight') expect(calls.formulas).toEqual([{ pageId: 'page', shapeId: 'three', cell: 'Angle', formula: String(Math.PI / 2) }]);
      else if (id === 'rotateLeft') expect(calls.formulas).toEqual([{ pageId: 'page', shapeId: 'three', cell: 'Angle', formula: String(-Math.PI / 2) }]);
      else if (id === 'flipHorizontal') expect(calls.formulas).toEqual([{ pageId: 'page', shapeId: 'three', cell: 'FlipX', formula: '1' }]);
      else expect(calls.formulas).toEqual([{ pageId: 'page', shapeId: 'three', cell: 'FlipY', formula: '1' }]);
      expect(document.querySelector('[role="menu"]')).toBeNull();
      expect(closed[0]).toBe('focus');
    } finally {
      view.unmount();
    }
  };
  runLeaf(null, 'delete');
  runLeaf('bringToFront', 'bringToFront');
  runLeaf('bringToFront', 'bringForward');
  runLeaf('sendToBack', 'sendBackward');
  runLeaf('sendToBack', 'sendToBack');
  runLeaf('rotateRight', 'rotateRight');
  runLeaf('rotateRight', 'rotateLeft');
  runLeaf('rotateRight', 'flipHorizontal');
  runLeaf('rotateRight', 'flipVertical');
  expect(LEAF_IDS).toEqual(['delete', 'bringToFront', 'bringForward', 'sendBackward', 'sendToBack', 'rotateRight', 'rotateLeft', 'flipHorizontal', 'flipVertical']);
});

test('escape closes the menu and returns focus', () => {
  const { view, closed, focusTarget } = renderMenu({ withFocusTarget: true });
  try {
    expect(parentMenu()).not.toBeNull();
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Escape' });
    expect(document.querySelector('[role="menu"]')).toBeNull();
    expect(closed[0]).toBe('focus');
    expect(document.activeElement).toBe(focusTarget as HTMLElement);
  } finally {
    view.unmount();
  }
});

test('a pointer press outside the menu closes it without running a command', () => {
  const { view, calls, closed } = renderMenu();
  try {
    expect(parentMenu()).not.toBeNull();
    fireEvent.mouseDown(document.body);
    expect(document.querySelector('[role="menu"]')).toBeNull();
    expect(closed).toEqual(['close']);
    expect(calls.deletes).toEqual([]);
    expect(calls.reorders).toEqual([]);
    expect(calls.formulas).toEqual([]);
  } finally {
    view.unmount();
  }
});

test('arrow keys wrap and home and end jump across the top level', () => {
  const { view, calls, closed } = renderMenu();
  try {
    const t = createT(en);
    const labelOf = (id: RibbonCommandId) => t(`ribbon.commands.${id}`);
    expect(document.activeElement?.getAttribute('aria-label')).toBe(labelOf('delete'));
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
    expect(document.activeElement?.getAttribute('data-submenu-id')).toBe('bringToFront');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
    expect(document.activeElement?.getAttribute('data-submenu-id')).toBe('sendToBack');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
    expect(document.activeElement?.getAttribute('data-submenu-id')).toBe('rotateRight');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
    expect(document.activeElement?.getAttribute('data-command-id')).toBe('delete');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowUp' });
    expect(document.activeElement?.getAttribute('data-submenu-id')).toBe('rotateRight');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'End' });
    expect(document.activeElement?.getAttribute('data-submenu-id')).toBe('rotateRight');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Home' });
    expect(document.activeElement?.getAttribute('data-command-id')).toBe('delete');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
    expect(document.activeElement?.getAttribute('data-submenu-id')).toBe('bringToFront');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowRight' });
    expect(document.activeElement?.getAttribute('data-command-id')).toBe('bringToFront');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
    expect(document.activeElement?.getAttribute('aria-label')).toBe(labelOf('bringForward'));
    fireEvent.click(document.activeElement as HTMLElement);
    expect(calls.reorders).toEqual([['page', 'three', 3]]);
    expect(document.querySelector('[role="menu"]')).toBeNull();
    expect(closed[0]).toBe('focus');
  } finally {
    view.unmount();
  }
});

test('arrow right opens a submenu and escape steps back out', () => {
  const { view, closed } = renderMenu({ withFocusTarget: true });
  try {
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
    expect(document.activeElement?.getAttribute('data-submenu-id')).toBe('bringToFront');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Enter' });
    const submenu = document.querySelector('[data-submenu="bringToFront"]') as HTMLElement;
    expect(submenu).not.toBeNull();
    expect(document.activeElement?.getAttribute('data-command-id')).toBe('bringToFront');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Escape' });
    expect(document.querySelector('[data-submenu="bringToFront"]')).toBeNull();
    expect(document.querySelector('[role="menu"]')).not.toBeNull();
    expect(document.activeElement?.getAttribute('data-submenu-id')).toBe('bringToFront');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowRight' });
    expect(document.querySelector('[data-submenu="bringToFront"]')).not.toBeNull();
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowLeft' });
    expect(document.querySelector('[data-submenu="bringToFront"]')).toBeNull();
    expect(document.activeElement?.getAttribute('data-submenu-id')).toBe('bringToFront');
    expect(closed).toEqual([]);
  } finally {
    view.unmount();
  }
});

test('the menu flips inside the viewport near an edge', () => {
  const originalRect = HTMLElement.prototype.getBoundingClientRect;
  HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
    if (this.getAttribute?.('role') === 'menu' && !this.hasAttribute('data-submenu')) return { x: 0, y: 0, width: 220, height: 340, top: 0, left: 0, right: 220, bottom: 340, toJSON: () => ({}) } as DOMRect;
    return originalRect.call(this);
  };
  const { view } = renderMenu({ position: { top: 1000, left: 740 } });
  try {
    const menu = parentMenu() as HTMLElement;
    expect(menu).not.toBeNull();
    expect(menu.style.left).toBe(`${Math.min(740, window.innerWidth - 220 - 4)}px`);
    expect(menu.style.top).toBe(`${Math.min(1000, window.innerHeight - 340 - 4)}px`);
  } finally {
    HTMLElement.prototype.getBoundingClientRect = originalRect;
    view.unmount();
  }
});

test('the menu re-clamps when the viewport shrinks while it is open', () => {
  const originalRect = HTMLElement.prototype.getBoundingClientRect;
  HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
    if (this.getAttribute?.('role') === 'menu' && !this.hasAttribute('data-submenu')) return { x: 0, y: 0, width: 220, height: 340, top: 0, left: 0, right: 220, bottom: 340, toJSON: () => ({}) } as DOMRect;
    return originalRect.call(this);
  };
  const originalWidth = window.innerWidth;
  const originalHeight = window.innerHeight;
  const { view } = renderMenu({ position: { top: 100, left: 100 } });
  try {
    const menu = parentMenu() as HTMLElement;
    expect(menu.style.top).toBe('100px');
    expect(menu.style.left).toBe('100px');
    Object.defineProperty(window, 'innerWidth', { value: 300, configurable: true });
    Object.defineProperty(window, 'innerHeight', { value: 200, configurable: true });
    fireEvent(window, new window.Event('resize'));
    expect(menu.style.left).toBe('76px');
    expect(menu.style.top).toBe('4px');
  } finally {
    Object.defineProperty(window, 'innerWidth', { value: originalWidth, configurable: true });
    Object.defineProperty(window, 'innerHeight', { value: originalHeight, configurable: true });
    HTMLElement.prototype.getBoundingClientRect = originalRect;
    view.unmount();
  }
});

test('a locked shape disables its refused operation and focuses the first allowed entry', () => {
  const { view } = renderMenu({ cells: [cell('LockDelete', '1'), cell('Angle', 'GUARD(0)'), cell('FlipX', '0'), cell('FlipY', '0')] });
  try {
    const menu = parentMenu() as HTMLElement;
    expect(menu).not.toBeNull();
    expect((menu.querySelector('[data-command-id="delete"]') as HTMLButtonElement).disabled).toBe(true);
    expect((menu.querySelector('[data-submenu-id="bringToFront"]') as HTMLButtonElement).disabled).toBe(false);
    expect((menu.querySelector('[data-submenu-id="sendToBack"]') as HTMLButtonElement).disabled).toBe(false);
    expect(document.activeElement?.getAttribute('data-submenu-id')).toBe('bringToFront');
  } finally {
    view.unmount();
  }
});

test('a click on a hover-opened submenu trigger keeps the submenu open', () => {
  const { view } = renderMenu();
  try {
    const trigger = parentMenu().querySelector('[data-submenu-id="bringToFront"]') as HTMLElement;
    expect(trigger).not.toBeNull();
    fireEvent.mouseEnter(trigger);
    expect(document.querySelector('[data-submenu="bringToFront"]')).not.toBeNull();
    fireEvent.click(trigger);
    expect(document.querySelector('[data-submenu="bringToFront"]')).not.toBeNull();
    expect(document.querySelector('[role="menu"]')).not.toBeNull();
  } finally {
    view.unmount();
  }
});

test('a submenu with no enabled child stays closed to keyboard and pointer', () => {
  const { view } = renderMenu({ shapeId: 'five' });
  try {
    const menu = parentMenu() as HTMLElement;
    expect((menu.querySelector('[data-submenu-id="bringToFront"]') as HTMLButtonElement).disabled).toBe(true);
    expect((menu.querySelector('[data-submenu-id="sendToBack"]') as HTMLButtonElement).disabled).toBe(false);
    expect((menu.querySelector('[data-submenu-id="rotateRight"]') as HTMLButtonElement).disabled).toBe(false);
    expect(document.activeElement?.getAttribute('data-command-id')).toBe('delete');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
    expect(document.activeElement?.getAttribute('data-submenu-id')).toBe('sendToBack');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
    expect(document.activeElement?.getAttribute('data-submenu-id')).toBe('rotateRight');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
    expect(document.activeElement?.getAttribute('data-command-id')).toBe('delete');
  } finally {
    view.unmount();
  }
});

test('a fully guarded rotate submenu stays closed to keyboard and pointer', () => {
  const { view } = renderMenu({ cells: [cell('Angle', 'GUARD(0)'), cell('FlipX', 'GUARD(0)'), cell('FlipY', 'GUARD(0)')] });
  try {
    const menu = parentMenu() as HTMLElement;
    expect((menu.querySelector('[data-submenu-id="rotateRight"]') as HTMLButtonElement).disabled).toBe(true);
    expect((menu.querySelector('[data-submenu-id="sendToBack"]') as HTMLButtonElement).disabled).toBe(false);
    expect(document.activeElement?.getAttribute('data-command-id')).toBe('delete');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
    expect(document.activeElement?.getAttribute('data-submenu-id')).toBe('bringToFront');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
    expect(document.activeElement?.getAttribute('data-submenu-id')).toBe('sendToBack');
    fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
    expect(document.activeElement?.getAttribute('data-command-id')).toBe('delete');
  } finally {
    view.unmount();
  }
});

test('a guarded Angle leaves the rotate submenu open with only the flips enabled', () => {
  const { view, calls } = renderMenu({ cells: [cell('Angle', 'GUARD(0)'), cell('FlipX', '0'), cell('FlipY', '0')] });
  try {
    expect((parentMenu().querySelector('[data-submenu-id="rotateRight"]') as HTMLButtonElement).disabled).toBe(false);
    const submenu = openSubmenu('rotateRight');
    const disabled = (id: string) => (submenu.querySelector(`[data-command-id="${id}"]`) as HTMLButtonElement).disabled;
    expect(disabled('rotateRight')).toBe(true);
    expect(disabled('rotateLeft')).toBe(true);
    expect(disabled('flipHorizontal')).toBe(false);
    expect(disabled('flipVertical')).toBe(false);
    expect(document.activeElement?.getAttribute('data-command-id')).toBe('flipHorizontal');
    fireEvent.click(submenu.querySelector('[data-command-id="rotateRight"]') as HTMLElement);
    expect(calls.formulas).toEqual([]);
  } finally {
    view.unmount();
  }
});

for (const submenuId of ['bringToFront', 'rotateRight']) for (const nearRight of [false, true]) test(`${submenuId} submenu ${nearRight ? 'flips left' : 'opens right'} and stays above the viewport bottom`, () => {
  const originalRect = HTMLElement.prototype.getBoundingClientRect;
  const left = nearRight ? window.innerWidth - 224 : 20;
  HTMLElement.prototype.getBoundingClientRect = function () {
    if (this.hasAttribute('data-submenu-id')) return { left, right: left + 220, top: window.innerHeight - 20, width: 220, height: 30 } as DOMRect;
    if (this.hasAttribute('data-submenu')) return { left: 0, top: 0, width: 200, height: 80 } as DOMRect;
    return originalRect.call(this);
  };
  try {
    renderMenu();
    const submenu = openSubmenu(submenuId);
    expect(submenu.style.position).toBe('fixed');
    expect(Number.parseFloat(submenu.style.left)).toBe(nearRight ? left - 200 : left + 220);
    expect(Number.parseFloat(submenu.style.top)).toBe(window.innerHeight - 84);
  } finally { cleanup(); HTMLElement.prototype.getBoundingClientRect = originalRect; }
});
