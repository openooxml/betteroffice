import { expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import type { DiagramHandle } from '@betteroffice/vsdx';
import { createT, en } from '@betteroffice/vsdx-i18n';
import { Ribbon } from './Ribbon';
import { RibbonCommandsProvider, createRibbonCommands } from './commands';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const { cleanup, fireEvent, render } = await import('@testing-library/react');

function stubDiagram(ids: string[] = []) {
  return {
    snapshot: () => ({ pages: [{ id: 'page', sourcePartPath: 'page', name: 'Page', shapes: ids.map((id) => ({ id, sourceId: 1, name: id, children: [], cells: [] })) }] }),
    canUndo: () => true,
    canRedo: () => true,
  } as unknown as DiagramHandle;
}

function renderRibbon(diagram: DiagramHandle, selection: { pageId: string; shapeId: string; hit: { kind: 'shape'; shapeId: string } } | null) {
  cleanup();
  return render(<RibbonCommandsProvider handle={diagram} snapshot={diagram.snapshot()} pageId="page" selection={selection} onMutation={() => {}} onError={() => {}} onDownload={() => {}}><Ribbon t={createT(en)} /></RibbonCommandsProvider>);
}

test('renders all tabs and supports click and roving arrow-key selection', () => {
  cleanup();
  const diagram = { snapshot: () => ({ pages: [{ id: 'page', sourcePartPath: 'page', name: 'Page', shapes: [] }] }), canUndo: () => false, canRedo: () => false } as unknown as DiagramHandle;
  const view = render(<RibbonCommandsProvider handle={diagram} snapshot={diagram.snapshot()} pageId="page" selection={null} onMutation={() => {}} onError={() => {}} onDownload={() => {}}><Ribbon t={createT(en)} /></RibbonCommandsProvider>);
  const home = view.getByRole('tab', { name: 'Home' }); const insert = view.getByRole('tab', { name: 'Insert' });
  expect(view.getAllByRole('tab')).toHaveLength(7); expect(home.getAttribute('aria-selected')).toBe('true'); expect(home.tabIndex).toBe(0); expect(insert.tabIndex).toBe(-1);
  fireEvent.click(insert); expect(insert.getAttribute('aria-selected')).toBe('true'); expect(insert.tabIndex).toBe(0);
  fireEvent.keyDown(insert, { key: 'ArrowRight' }); const design = view.getByRole('tab', { name: 'Design' }); expect(design.getAttribute('aria-selected')).toBe('true'); expect(document.activeElement).toBe(design);
});

test('home surface is one flat row with no group-label text nodes', () => {
  const view = renderRibbon(stubDiagram(), null);
  const panel = view.getByTestId('vsdx-ribbon-home-panel');
  expect(panel.style.height).toBe('45px');
  expect(panel.style.display).toBe('flex');
  expect(view.queryByText(en.ribbon.groups.clipboard)).toBeNull();
  expect(view.queryByText(en.ribbon.groups.font)).toBeNull();
  expect(view.queryByText(en.ribbon.groups.paragraph)).toBeNull();
  expect(view.queryByText(en.ribbon.groups.history)).toBeNull();
  expect(view.queryByText(en.ribbon.groups.arrange)).toBeNull();
  for (const group of panel.querySelectorAll('[role="group"]')) expect(group.textContent?.trim() ?? '').toBe('');
  expect(panel.querySelectorAll('[role="separator"]').length).toBeGreaterThan(0);
  view.unmount();
});

test('only the active tab is selected and arrow keys move it', () => {
  const view = renderRibbon(stubDiagram(), null);
  const tabs = view.getAllByRole('tab');
  expect(tabs.filter((tab) => tab.getAttribute('aria-selected') === 'true')).toHaveLength(1);
  expect(view.getByRole('tab', { name: 'Home' }).getAttribute('aria-selected')).toBe('true');
  for (const tab of tabs.filter((tab) => tab.textContent !== 'Home')) expect(tab.getAttribute('aria-selected')).toBe('false');
  const home = view.getByRole('tab', { name: 'Home' });
  fireEvent.keyDown(home, { key: 'ArrowLeft' });
  expect(view.getByRole('tab', { name: 'File' }).getAttribute('aria-selected')).toBe('true');
  expect(document.activeElement).toBe(view.getByRole('tab', { name: 'File' }));
  fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'End' });
  expect(view.getByRole('tab', { name: 'Help' }).getAttribute('aria-selected')).toBe('true');
  fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Home' });
  expect(view.getByRole('tab', { name: 'File' }).getAttribute('aria-selected')).toBe('true');
  view.unmount();
});

test('every rendered command maps to a command id from commands.ts', () => {
  const diagram = stubDiagram(['one']);
  const selection = { pageId: 'page', shapeId: 'one', hit: { kind: 'shape' as const, shapeId: 'one' } };
  const valid = new Set(Object.keys(createRibbonCommands(diagram, null, 'page', () => {}, () => {}, () => {})));
  const view = renderRibbon(diagram, selection);
  for (const tab of ['File', 'Home', 'Insert', 'Design', 'Review', 'View', 'Help']) fireEvent.click(view.getByRole('tab', { name: tab }));
  fireEvent.click(view.getByRole('tab', { name: 'Home' }));
  for (const toggle of view.container.querySelectorAll('[data-split-toggle]')) fireEvent.click(toggle);
  const rendered = view.container.querySelectorAll('[data-command-id]');
  expect(rendered.length).toBeGreaterThan(0);
  for (const node of rendered) expect(valid.has(node.getAttribute('data-command-id') ?? '')).toBe(true);
  for (const node of view.container.querySelectorAll('button[aria-label], input[aria-label]')) {
    if (node.hasAttribute('data-split-toggle')) continue;
    const role = node.parentElement?.getAttribute('role');
    if (role === 'tab' || node.getAttribute('role') === 'tab') continue;
    expect(node.hasAttribute('data-command-id')).toBe(true);
  }
  view.unmount();
});

test('disabled commands keep their labels and stay out of the tab order', () => {
  const view = renderRibbon(stubDiagram(), null);
  const panel = view.getByTestId('vsdx-ribbon-home-panel');
  const disabled = [...panel.querySelectorAll('button[disabled], input[disabled]')];
  expect(disabled.length).toBeGreaterThan(0);
  for (const node of disabled) {
    expect(node.getAttribute('aria-label') ?? '').not.toBe('');
    (node as HTMLElement).focus();
    expect(document.activeElement).not.toBe(node);
  }
  view.unmount();
});

test('tabs without commands render an honest empty state', () => {
  const view = renderRibbon(stubDiagram(), null);
  for (const name of ['Design', 'Review', 'View', 'Help']) {
    fireEvent.click(view.getByRole('tab', { name }));
    expect(view.getByText(en.ribbon.empty)).not.toBeNull();
  }
  fireEvent.click(view.getByRole('tab', { name: 'Insert' }));
  expect(view.queryByText(en.ribbon.empty)).toBeNull();
  view.unmount();
});

function cell(name: string, value: string) {
  return { locator: { sheet: { page: 1 }, shapeId: 1, section: null, row: null, cellName: name }, name, formula: value, value };
}

function richDiagram(shapes: Array<{ id: string; cells?: Array<ReturnType<typeof cell>> }>, calls: { reorder: unknown[][]; mutation: number }) {
  return {
    snapshot: () => ({ pages: [{ id: 'page', sourcePartPath: 'page', name: 'Page', shapes: shapes.map((shape) => ({ id: shape.id, sourceId: 1, name: shape.id, children: [], cells: shape.cells ?? [] })) }] }),
    canUndo: () => true,
    canRedo: () => true,
    reorderShape: (...args: unknown[]) => { calls.reorder.push(args); return {}; },
  } as unknown as DiagramHandle;
}

function selectionFor(shapeId: string) {
  return { pageId: 'page', shapeId, hit: { kind: 'shape' as const, shapeId } };
}

function openArrange(view: ReturnType<typeof render>, toggleId: string) {
  const toggle = view.container.querySelector(`[data-split-toggle="${toggleId}"]`) as HTMLElement;
  expect(toggle.getAttribute('aria-haspopup')).toBe('menu');
  fireEvent.click(toggle);
  expect(toggle.getAttribute('aria-expanded')).toBe('true');
  return toggle;
}

test('opening a split menu moves focus to the first enabled item', () => {
  const view = renderRibbon(stubDiagram(['one', 'two']), selectionFor('one'));
  const toggle = openArrange(view, 'bringToFront');
  const menu = view.getByRole('menu');
  expect(menu).not.toBeNull();
  const first = view.getByRole('menuitem', { name: en.ribbon.commands.bringToFront });
  expect(document.activeElement).toBe(first);
  expect(toggle.getAttribute('aria-expanded')).toBe('true');
  view.unmount();
});

test('opening a split menu focuses the checked item when one is checked', () => {
  const calls = { reorder: [] as unknown[][], mutation: 0 };
  const diagram = richDiagram([{ id: 'one', cells: [cell('FlipX', '1'), cell('FlipY', '0')] }], calls);
  cleanup();
  const view = render(<RibbonCommandsProvider handle={diagram} snapshot={diagram.snapshot()} pageId="page" selection={selectionFor('one')} onMutation={() => {}} onError={() => {}} onDownload={() => {}}><Ribbon t={createT(en)} /></RibbonCommandsProvider>);
  const toggle = view.container.querySelector('[data-split-toggle="rotateRight"]') as HTMLElement;
  fireEvent.click(toggle);
  const checked = view.getByRole('menuitemcheckbox', { name: en.ribbon.commands.flipHorizontal });
  expect(checked.getAttribute('aria-checked')).toBe('true');
  expect(document.activeElement).toBe(checked);
  view.unmount();
});

test('arrow keys cycle and wrap while home and end jump', () => {
  const view = renderRibbon(stubDiagram(['one', 'two']), selectionFor('one'));
  openArrange(view, 'bringToFront');
  const first = view.getByRole('menuitem', { name: en.ribbon.commands.bringToFront });
  const second = view.getByRole('menuitem', { name: en.ribbon.commands.bringForward });
  expect(document.activeElement).toBe(first);
  fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
  expect(document.activeElement).toBe(second);
  fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowDown' });
  expect(document.activeElement).toBe(first);
  fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'ArrowUp' });
  expect(document.activeElement).toBe(second);
  fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Home' });
  expect(document.activeElement).toBe(first);
  fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'End' });
  expect(document.activeElement).toBe(second);
  view.unmount();
});

test('escape closes the menu and returns focus to the trigger', () => {
  const view = renderRibbon(stubDiagram(['one', 'two']), selectionFor('one'));
  const toggle = openArrange(view, 'bringToFront');
  expect(view.queryByRole('menu')).not.toBeNull();
  fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Escape' });
  expect(view.queryByRole('menu')).toBeNull();
  expect(document.activeElement).toBe(toggle);
  expect(toggle.getAttribute('aria-expanded')).toBe('false');
  view.unmount();
});

test('activating an item runs the command and returns focus to the trigger', () => {
  const calls = { reorder: [] as unknown[][], mutation: 0 };
  const diagram = richDiagram([{ id: 'one' }, { id: 'two' }], calls);
  cleanup();
  const view = render(<RibbonCommandsProvider handle={diagram} snapshot={diagram.snapshot()} pageId="page" selection={selectionFor('one')} onMutation={() => { calls.mutation += 1; }} onError={() => {}} onDownload={() => {}}><Ribbon t={createT(en)} /></RibbonCommandsProvider>);
  const toggle = view.container.querySelector('[data-split-toggle="bringToFront"]') as HTMLElement;
  fireEvent.click(toggle);
  const item = view.getByRole('menuitem', { name: en.ribbon.commands.bringToFront });
  fireEvent.click(item);
  expect(calls.reorder.length).toBe(1);
  expect(calls.mutation).toBe(1);
  expect(view.queryByRole('menu')).toBeNull();
  expect(document.activeElement).toBe(toggle);
  view.unmount();
});

test('tab closes the menu instead of trapping focus', () => {
  const view = renderRibbon(stubDiagram(['one', 'two']), selectionFor('one'));
  openArrange(view, 'bringToFront');
  expect(view.queryByRole('menu')).not.toBeNull();
  fireEvent.keyDown(document.activeElement as HTMLElement, { key: 'Tab' });
  expect(view.queryByRole('menu')).toBeNull();
  expect(view.container.querySelector('[role="menu"]')).toBeNull();
  view.unmount();
});
