import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import type { DiagramHandle, DiagramSnapshot } from '@betteroffice/vsdx';
import { createT, en } from '@betteroffice/vsdx-i18n';
import { useState } from 'react';
import { CANVAS_CONTEXT_ENTRIES, CanvasContextMenu } from './components/ribbon/CanvasContextMenu';
import { RibbonCommandsProvider } from './components/ribbon/commands';
import type { RibbonCommandId } from './components/ribbon/commands';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const { cleanup, fireEvent, render } = await import('@testing-library/react');

interface Calls {
  undos: number;
  redos: number;
  added: unknown[][];
}

function stubHandle(calls: Calls, options: { canUndo: boolean; canRedo: boolean } = { canUndo: true, canRedo: true }): DiagramHandle {
  const state: DiagramSnapshot = { pages: [{ id: 'page', sourcePartPath: 'page', name: 'Page', shapes: [] }] };
  return {
    snapshot: () => state,
    canUndo: () => options.canUndo,
    canRedo: () => options.canRedo,
    undo: () => { calls.undos += 1; },
    redo: () => { calls.redos += 1; },
    addShape: (...args: [string, unknown]) => { calls.added.push([...args]); return { id: 'new', pageId: 'page' }; },
  } as unknown as DiagramHandle;
}

function Host({ diagram, position, closed, focusTarget }: { diagram: DiagramHandle; position: { top: number; left: number }; closed: string[]; focusTarget?: HTMLElement }) {
  const [open, setOpen] = useState(true);
  if (!open) return null;
  return (
    <RibbonCommandsProvider handle={diagram} snapshot={diagram.snapshot()} pageId="page" selection={[]} onMutation={() => {}} onError={() => {}} onDownload={() => {}}>
      <CanvasContextMenu t={createT(en)} position={position} onClose={() => { closed.push('close'); setOpen(false); }} onCloseAndFocus={() => { closed.push('focus'); setOpen(false); focusTarget?.focus(); }} />
    </RibbonCommandsProvider>
  );
}

function renderMenu(options: { position?: { top: number; left: number }; withFocusTarget?: boolean; canUndo?: boolean; canRedo?: boolean } = {}) {
  cleanup();
  const calls: Calls = { undos: 0, redos: 0, added: [] };
  const diagram = stubHandle(calls, { canUndo: options.canUndo ?? true, canRedo: options.canRedo ?? true });
  const closed: string[] = [];
  let focusTarget: HTMLElement | undefined;
  if (options.withFocusTarget) {
    focusTarget = document.createElement('button');
    document.body.appendChild(focusTarget);
  }
  const view = render(<Host diagram={diagram} position={options.position ?? { top: 100, left: 100 }} closed={closed} focusTarget={focusTarget} />);
  return { view, calls, closed, focusTarget };
}

function canvasMenu(): HTMLElement {
  const menus = Array.from(document.querySelectorAll('[role="menu"]')) as HTMLElement[];
  const found = menus.find((menu) => menu.getAttribute('aria-label') === en.contextMenu.canvasLabel);
  expect(found).not.toBeUndefined();
  return found as HTMLElement;
}

afterEach(() => {
  cleanup();
  document.body.innerHTML = '';
});

test('the canvas menu offers only existing ribbon commands with a single divider', () => {
  const { view } = renderMenu();
  try {
    const menu = canvasMenu();
    expect(menu.getAttribute('aria-label')).toBe(en.contextMenu.canvasLabel);
    expect(en.contextMenu.canvasLabel).not.toBe(en.contextMenu.label);
    expect(Array.from(menu.querySelectorAll('[data-command-id]')).map((button) => button.getAttribute('data-command-id'))).toEqual(['undo', 'redo', 'addShape']);
    expect(menu.querySelectorAll('[role="separator"]')).toHaveLength(1);
    expect(menu.querySelector('[aria-keyshortcuts]')).toBeNull();
    const labels: Record<RibbonCommandId, string> = en.ribbon.commands as Record<RibbonCommandId, string>;
    for (const entry of CANVAS_CONTEXT_ENTRIES) expect(labels[entry.id]).toBeDefined();
  } finally {
    view.unmount();
  }
});

test('each entry runs its command and closes the menu', () => {
  for (const id of ['undo', 'redo', 'addShape'] as const) {
    const { view, calls, closed } = renderMenu();
    try {
      const item = document.querySelector(`[role="menu"] [data-command-id="${id}"]`) as HTMLElement;
      expect(item).not.toBeNull();
      fireEvent.click(item);
      if (id === 'undo') expect(calls.undos).toBe(1);
      if (id === 'redo') expect(calls.redos).toBe(1);
      if (id === 'addShape') expect(calls.added[0]?.[0]).toBe('page');
      expect(document.querySelector('[role="menu"]')).toBeNull();
      expect(closed[0]).toBe('focus');
    } finally {
      view.unmount();
    }
  }
});

test('focus starts on the first enabled entry when history is unavailable', () => {
  const { view } = renderMenu({ canUndo: false, canRedo: false });
  try {
    expect(document.activeElement?.getAttribute('data-command-id')).toBe('addShape');
    expect((document.querySelector('[data-command-id="undo"]') as HTMLButtonElement).disabled).toBe(true);
    expect((document.querySelector('[data-command-id="redo"]') as HTMLButtonElement).disabled).toBe(true);
  } finally {
    view.unmount();
  }
});

test('escape closes the menu and returns focus', () => {
  const { view, closed, focusTarget } = renderMenu({ withFocusTarget: true });
  try {
    expect(canvasMenu()).not.toBeNull();
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
    expect(canvasMenu()).not.toBeNull();
    fireEvent.mouseDown(document.body);
    expect(document.querySelector('[role="menu"]')).toBeNull();
    expect(closed).toEqual(['close']);
    expect(calls).toEqual({ undos: 0, redos: 0, added: [] });
  } finally {
    view.unmount();
  }
});

test('the menu flips inside the viewport near an edge', () => {
  const originalRect = HTMLElement.prototype.getBoundingClientRect;
  HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
    if (this.getAttribute?.('role') === 'menu') return { x: 0, y: 0, width: 220, height: 340, top: 0, left: 0, right: 220, bottom: 340, toJSON: () => ({}) } as DOMRect;
    return originalRect.call(this);
  };
  const { view } = renderMenu({ position: { top: 1000, left: 740 } });
  try {
    const menu = canvasMenu() as HTMLElement;
    expect(menu.style.left).toBe(`${Math.min(740, window.innerWidth - 220 - 4)}px`);
    expect(menu.style.top).toBe(`${Math.min(1000, window.innerHeight - 340 - 4)}px`);
  } finally {
    HTMLElement.prototype.getBoundingClientRect = originalRect;
    view.unmount();
  }
});
