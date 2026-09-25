import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { useState } from 'react';
import { generateThemeTintShadeMatrix } from '@betteroffice/docx/utils';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import {
  DocxCommandProvider,
  EditorToolbar,
  ToolbarButton,
  ToolbarCommand,
  ToolbarCommandButton,
  ToolbarGroup,
  ToolbarOverflow,
} from '../../index';
import { createDocxCommandController } from '../../commands/createDocxCommandStore';
import { testBinding } from '../../commands/testing';

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');

/** Queries the current document; `screen` stays bound to the first registered one. */
function screen() {
  return within(document.body);
}

let railWidth = 1000;
const originalRect = HTMLElement.prototype.getBoundingClientRect;

function rect(width: number): DOMRect {
  return {
    width,
    height: 28,
    top: 0,
    left: 0,
    right: width,
    bottom: 28,
    x: 0,
    y: 0,
    toJSON() {},
  } as DOMRect;
}

function visibleText(element: Element): string {
  let text = '';
  for (const node of Array.from(element.childNodes)) {
    if (node.nodeType === Node.TEXT_NODE) text += node.textContent ?? '';
    else if (node instanceof Element && !node.hasAttribute('hidden')) text += visibleText(node);
  }
  return text;
}

beforeAll(() => {
  HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
    if (this.getAttribute('role') === 'toolbar') return rect(railWidth);
    if (this.parentElement?.hasAttribute('data-toolbar-items')) {
      const controls = this.querySelectorAll('button').length;
      return rect(Math.max(24, controls * 28 + visibleText(this).length * 7));
    }
    if (this.querySelector(':scope > [data-testid="toolbar-more"]')) return rect(28);
    return originalRect.call(this);
  };
});

afterEach(cleanup);
afterAll(async () => {
  HTMLElement.prototype.getBoundingClientRect = originalRect;
  if (ownsDom) await GlobalRegistrator.unregister();
});

function resize(width: number) {
  railWidth = width;
  act(() => {
    window.dispatchEvent(new Event('resize'));
  });
}

function mount(hostLabel = 'Share', overrides = {}) {
  railWidth = 1000;
  const harness = testBinding({ canUndo: false, ...overrides });
  const controller = createDocxCommandController();
  controller.attach(harness.binding);
  let shared = 0;
  const view = render(
    <DocxCommandProvider commands={controller.store}>
      <EditorToolbar>
        <EditorToolbar.Toolbar>
          <ToolbarGroup label="Formatting">
            <ToolbarCommandButton id="bold" />
            <ToolbarCommandButton id="italic" />
          </ToolbarGroup>
          <ToolbarGroup label="History">
            <ToolbarCommandButton id="undo" />
            <ToolbarCommandButton id="redo" />
          </ToolbarGroup>
          <ToolbarGroup label="Styles">
            <ToolbarCommand id="alignment" />
          </ToolbarGroup>
          <ToolbarButton title={hostLabel} onClick={() => (shared += 1)}>
            {hostLabel}
          </ToolbarButton>
          <ToolbarOverflow label="Custom widget" onSelect={() => (shared += 10)}>
            <span>Custom widget</span>
          </ToolbarOverflow>
          <span data-testid="unrepresented">Plain</span>
        </EditorToolbar.Toolbar>
      </EditorToolbar>
    </DocxCommandProvider>
  );
  return { harness, view, shared: () => shared };
}

function more(): HTMLElement {
  return screen().getByTestId('toolbar-more');
}

function openWithKeyboard(key = 'ArrowDown') {
  const trigger = more();
  act(() => trigger.focus());
  fireEvent.keyDown(trigger, { key });
  return screen().getByRole('menu', { name: 'More actions' });
}

function menuItems(menu: HTMLElement): HTMLElement[] {
  return Array.from(
    menu.querySelectorAll<HTMLElement>(
      '[role="menuitem"], [role="menuitemcheckbox"], [role="menuitemradio"]'
    )
  ).filter((item) => item.closest('[role="menu"]') === menu);
}

describe('toolbar overflow', () => {
  test('keeps everything in the row when it fits', () => {
    mount();
    expect(screen().queryByTestId('toolbar-more')).toBeNull();
  });

  test('moves whole trailing groups into More and keeps every action reachable', () => {
    mount();
    resize(120);
    const toolbar = screen().getByRole('toolbar');
    const group = (label: string) =>
      toolbar.querySelector<HTMLElement>(`[role="group"][aria-label="${label}"]`)!;
    expect(within(toolbar).getByRole('group', { name: 'Formatting' })).toBe(group('Formatting'));
    const history = group('History');
    expect(history.getAttribute('aria-hidden')).toBe('true');
    expect((history as HTMLElement & { inert: boolean }).inert).toBe(true);
    expect(screen().getByTestId('unrepresented').getAttribute('aria-hidden')).toBeNull();

    const menu = openWithKeyboard();
    const labels = menuItems(menu).map((item) => item.dataset.label);
    expect(labels).toEqual(['Undo', 'Redo', 'Alignment', 'Share', 'Custom widget']);
    expect(within(menu).getByRole('group', { name: 'History' })).toBeDefined();

    resize(10);
    const everything = openWithKeyboard();
    expect(menuItems(everything).map((item) => item.dataset.label)).toEqual([
      'Bold',
      'Italic',
      'Undo',
      'Redo',
      'Alignment',
      'Share',
      'Custom widget',
    ]);
  });

  test('long labels need more room', () => {
    mount('Share with the whole team');
    resize(340);
    const menu = openWithKeyboard();
    expect(menuItems(menu).map((item) => item.dataset.label)).toContain(
      'Share with the whole team'
    );
    cleanup();
    mount('Share');
    resize(340);
    expect(screen().queryByTestId('toolbar-more')).toBeNull();
  });

  test('supports menu keyboard navigation, typeahead and Escape', () => {
    mount();
    resize(10);
    const menu = openWithKeyboard();
    const items = menuItems(menu);
    expect(document.activeElement).toBe(items[0]);
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowDown' });
    expect(document.activeElement).toBe(items[1]);
    fireEvent.keyDown(document.activeElement!, { key: 'End' });
    expect(document.activeElement).toBe(items[items.length - 1]);
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowDown' });
    expect(document.activeElement).toBe(items[0]);
    fireEvent.keyDown(document.activeElement!, { key: 'Home' });
    expect(document.activeElement).toBe(items[0]);
    fireEvent.keyDown(document.activeElement!, { key: 'r' });
    expect(document.activeElement?.getAttribute('data-label')).toBe('Redo');
    fireEvent.keyDown(document.activeElement!, { key: 'Escape' });
    expect(screen().queryByRole('menu')).toBeNull();
    expect(document.activeElement).toBe(more());

    const reopened = openWithKeyboard('ArrowUp');
    const last = menuItems(reopened);
    expect(document.activeElement).toBe(last[last.length - 1]);
  });

  test('explains disabled entries and runs enabled ones', async () => {
    const { harness, shared } = mount();
    resize(10);
    const menu = openWithKeyboard();
    const undo = menuItems(menu).find((item) => item.dataset.label === 'Undo')!;
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    const description = document.getElementById(undo.getAttribute('aria-describedby')!);
    expect(description?.textContent).toBe('commands.reasons.nothingToUndo');
    act(() => undo.focus());
    fireEvent.keyDown(undo, { key: 'Enter' });
    expect(screen().getByRole('menu')).toBeDefined();

    const bold = menuItems(menu).find((item) => item.dataset.label === 'Bold')!;
    expect(bold.getAttribute('role')).toBe('menuitemcheckbox');
    expect(bold.getAttribute('aria-checked')).toBe('false');
    act(() => bold.focus());
    fireEvent.keyDown(bold, { key: 'Enter' });
    await act(async () => {});
    expect(screen().queryByRole('menu')).toBeNull();
    expect(document.activeElement).toBe(more());
    expect(harness.calls.map((call) => call.id)).toEqual(['bold']);

    const again = openWithKeyboard();
    const share = menuItems(again).find((item) => item.dataset.label === 'Share')!;
    fireEvent.click(share);
    expect(shared()).toBe(1);
  });

  test('opens and closes submenus from the keyboard', async () => {
    const { harness } = mount();
    resize(10);
    const menu = openWithKeyboard();
    const alignment = menuItems(menu).find((item) => item.dataset.label === 'Alignment')!;
    expect(alignment.getAttribute('aria-haspopup')).toBe('menu');
    act(() => alignment.focus());
    fireEvent.keyDown(alignment, { key: 'ArrowRight' });
    const submenu = screen().getByRole('menu', { name: 'Alignment' });
    const options = menuItems(submenu);
    expect(options.map((item) => item.getAttribute('role'))).toEqual([
      'menuitemradio',
      'menuitemradio',
      'menuitemradio',
      'menuitemradio',
    ]);
    expect(options[0].getAttribute('aria-checked')).toBe('true');
    expect(document.activeElement).toBe(options[0]);
    fireEvent.keyDown(options[0], { key: 'ArrowLeft' });
    expect(screen().queryByRole('menu', { name: 'Alignment' })).toBeNull();
    expect(document.activeElement).toBe(alignment);

    fireEvent.keyDown(alignment, { key: 'Enter' });
    const reopened = menuItems(screen().getByRole('menu', { name: 'Alignment' }));
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowDown' });
    fireEvent.keyDown(document.activeElement!, { key: 'Enter' });
    await act(async () => {});
    expect(reopened.length).toBe(4);
    expect(harness.calls).toEqual([
      { id: 'alignment', args: { value: 'center' }, ordered: true },
    ]);
  });

  test('moves focus to More when a resize hides the focused control', () => {
    mount();
    const share = screen().getByRole('button', { name: 'Share' });
    act(() => share.focus());
    resize(120);
    expect(document.activeElement).toBe(more());
  });

  test('host entries publish their changes to an open menu and run their latest action', () => {
    const selected: string[] = [];
    let update!: (next: { label: string; disabled: boolean }) => void;
    function Host() {
      const [state, setState] = useState({ label: 'Share', disabled: false });
      update = setState;
      return (
        <ToolbarOverflow
          label={state.label}
          disabled={state.disabled}
          description="Sharing is off"
          onSelect={() => selected.push(state.label)}
        >
          <span>{state.label}</span>
        </ToolbarOverflow>
      );
    }
    const harness = testBinding();
    const controller = createDocxCommandController();
    controller.attach(harness.binding);
    render(
      <DocxCommandProvider commands={controller.store}>
        <EditorToolbar>
          <EditorToolbar.Toolbar>
            <ToolbarCommandButton id="bold" />
            <Host />
          </EditorToolbar.Toolbar>
        </EditorToolbar>
      </DocxCommandProvider>
    );
    resize(10);
    const menu = openWithKeyboard();
    const entry = () => menuItems(menu).find((item) => item.dataset.label?.startsWith('Share'))!;
    expect(entry().dataset.label).toBe('Share');

    act(() => update({ label: 'Share link', disabled: true }));
    expect(entry().dataset.label).toBe('Share link');
    expect(entry().getAttribute('aria-disabled')).toBe('true');
    expect(document.getElementById(entry().getAttribute('aria-describedby')!)?.textContent).toBe(
      'Sharing is off'
    );
    act(() => entry().focus());
    fireEvent.keyDown(entry(), { key: 'Enter' });
    expect(selected).toEqual([]);

    act(() => update({ label: 'Share link', disabled: false }));
    expect(entry().hasAttribute('aria-disabled')).toBe(false);
    fireEvent.click(entry());
    expect(selected).toEqual(['Share link']);
  });

  test('bounds the menu to the viewport and wraps long labels', () => {
    const width = window.innerWidth;
    Object.defineProperty(window, 'innerWidth', { value: 240, configurable: true });
    try {
      mount('Share this document with everyone in the organisation and their guests');
      resize(10);
      const menu = openWithKeyboard();
      expect(parseFloat(menu.style.maxWidth)).toBeLessThanOrEqual(240 - 16);
      expect(parseFloat(menu.style.minWidth)).toBeLessThanOrEqual(240 - 16);
      const long = menuItems(menu).find((item) => item.dataset.label?.startsWith('Share'))!;
      expect(long.style.whiteSpace).not.toBe('nowrap');
    } finally {
      Object.defineProperty(window, 'innerWidth', { value: width, configurable: true });
    }
  });
});

interface MenuNode {
  label: string;
  item: HTMLElement;
  children?: MenuNode[];
}

function submenuOf(menu: HTMLElement, label: string): HTMLElement {
  const item = menuItems(menu).find((candidate) => candidate.dataset.label === label);
  if (!item) throw new Error(`No menu entry "${label}"`);
  act(() => item.focus());
  fireEvent.keyDown(item, { key: 'ArrowRight' });
  const submenu = item.parentElement?.querySelector<HTMLElement>(':scope > [role="menu"]');
  if (!submenu) throw new Error(`"${label}" opened no submenu`);
  return submenu;
}

/** Every entry of a menu, opening each submenu in turn. */
function menuTree(menu: HTMLElement): MenuNode[] {
  return menuItems(menu).map((item) => {
    const label = item.dataset.label!;
    if (item.getAttribute('aria-haspopup') !== 'menu') return { label, item };
    const submenu = submenuOf(menu, label);
    const children = menuTree(submenu);
    fireEvent.keyDown(menuItems(submenu)[0], { key: 'ArrowLeft' });
    return { label, item, children };
  });
}

function find(nodes: MenuNode[], path: string[]): MenuNode {
  const [head, ...rest] = path;
  const node = nodes.find((candidate) => candidate.label === head);
  if (!node) throw new Error(`No entry "${head}" among ${nodes.map((n) => n.label).join(', ')}`);
  return rest.length === 0 ? node : find(node.children ?? [], rest);
}

function labels(node: MenuNode): string[] {
  return (node.children ?? []).map((child) => child.label);
}

/** Opens More, walks `path` through the submenus and activates the last entry. */
function choose(path: string[]) {
  let menu = openWithKeyboard();
  for (const label of path.slice(0, -1)) menu = submenuOf(menu, label);
  const item = menuItems(menu).find((candidate) => candidate.dataset.label === path.at(-1));
  if (!item) throw new Error(`No entry "${path.at(-1)}"`);
  act(() => item.focus());
  fireEvent.keyDown(item, { key: 'Enter' });
}

describe('overflow coverage', () => {
  function mountDefault(children?: React.ReactNode) {
    const harness = testBinding();
    const controller = createDocxCommandController();
    controller.attach(harness.binding);
    render(
      <DocxCommandProvider commands={controller.store}>
        <EditorToolbar>
          <EditorToolbar.Toolbar>{children}</EditorToolbar.Toolbar>
        </EditorToolbar>
      </DocxCommandProvider>
    );
    resize(10);
    return harness;
  }

  test('every control of the default toolbar is reachable with all its choices', () => {
    mountDefault();
    const tree = menuTree(openWithKeyboard());
    fireEvent.keyDown(document.activeElement!, { key: 'Escape' });

    for (const label of [
      'Undo',
      'Redo',
      'Bold',
      'Italic',
      'Underline',
      'Strikethrough',
      'Insert link',
      'Superscript',
      'Subscript',
      'Bullet List',
      'Numbered List',
      'Decrease Indent',
      'Increase Indent',
      'Image properties',
      'Clear formatting',
      'Toggle comments sidebar',
    ]) {
      expect(find(tree, [label]).children).toBeUndefined();
    }
    expect(labels(find(tree, ['Paragraph style']))).toEqual(['Normal text', 'Heading 1']);
    expect(labels(find(tree, ['Font']))).toEqual(['Arial', 'Georgia']);
    expect(labels(find(tree, ['Font size'])).at(-1)).toBe('Custom size…');
    expect(labels(find(tree, ['Zoom level']))).toEqual(['50%', '75%', '100%', '125%', '150%', '200%']);
    expect(labels(find(tree, ['Alignment']))).toHaveLength(4);
    expect(labels(find(tree, ['Line spacing']))).toHaveLength(4);
    expect(labels(find(tree, ['Text wrapping']))).toHaveLength(6);
    expect(labels(find(tree, ['Transform']))).toHaveLength(4);
    expect(labels(find(tree, ['Editing mode']))).toHaveLength(3);

    const matrix = generateThemeTintShadeMatrix(null);
    for (const path of [['Font Color'], ['Text Highlight Color'], ['Table', 'Border Color'], ['Table', 'Cell Fill Color']]) {
      const colors = find(tree, path);
      const [clear, theme, ...rest] = labels(colors);
      expect(['Automatic', 'No Color']).toContain(clear);
      expect(theme).toBe('Theme Colors');
      expect(rest).toHaveLength(11);
      expect(rest.at(-1)).toBe('Custom color…');
      const columns = find(colors.children!, ['Theme Colors']).children!;
      expect(columns).toHaveLength(10);
      expect(columns.flatMap((column) => column.children!.map((cell) => cell.label))).toEqual(
        matrix[0].flatMap((_, column) => matrix.map((row) => row[column].label))
      );
    }

    expect(labels(find(tree, ['Table']))).toEqual([
      'Insert row above',
      'Insert row below',
      'Insert column left',
      'Insert column right',
      'Merge cells',
      'Split cell',
      'Select entire table',
      'Delete row',
      'Delete column',
      'Delete table',
      'Borders',
      'Border Color',
      'Border width',
      'Cell Fill Color',
      'Vertical alignment',
      'Table alignment',
      'Toggle header row',
      'Distribute columns evenly',
      'Auto-fit to contents',
      'Toggle no-wrap',
      'Table properties...',
    ]);
    expect(labels(find(tree, ['Table', 'Borders']))).toEqual([
      'All borders',
      'Outside borders',
      'Inside borders',
      'Top border',
      'Bottom border',
      'Left border',
      'Right border',
      'No borders',
    ]);
    expect(labels(find(tree, ['Table', 'Border width']))).toEqual([
      '0.5 pt',
      '1 pt',
      '1.5 pt',
      '2 pt',
      '3 pt',
    ]);
    expect(labels(find(tree, ['Table', 'Vertical alignment']))).toEqual(['Top', 'Middle', 'Bottom']);
    expect(labels(find(tree, ['Table', 'Table alignment']))).toEqual([
      'Align table left',
      'Align table center',
      'Align table right',
    ]);
  });

  test('overflow choices run the same commands as the controls they replace', async () => {
    const harness = mountDefault();
    const matrix = generateThemeTintShadeMatrix(null);
    const accent = matrix[1][4];
    const runs: [string[], unknown][] = [
      [['Table', 'Borders', 'Top border'], 'borderTop'],
      [['Table', 'Border width', '1.5 pt'], { type: 'borderWidth', size: 12 }],
      [['Table', 'Cell Fill Color', 'Red'], { type: 'cellFillColor', color: 'FF0000' }],
      [['Table', 'Border Color', 'Automatic'], { type: 'borderColor', color: '000000' }],
      [
        ['Table', 'Table alignment', 'Align table center'],
        { type: 'tableProperties', props: { justification: 'center' } },
      ],
      [
        ['Font Color', 'Theme Colors', matrix[0][4].label, accent.label],
        { color: { themeColor: accent.themeSlot, themeTint: accent.tint } },
      ],
      [['Text Highlight Color', 'Yellow'], { color: 'FFFF00' }],
      [['Font size', '14'], { points: 14 }],
    ];
    for (const [path, args] of runs) {
      choose(path);
      await act(async () => {});
      expect(harness.calls.at(-1)?.args).toEqual(args);
    }
  });

  test('custom values are asked for in an accessible dialog', async () => {
    const harness = mountDefault();
    choose(['Font Color', 'Custom color…']);
    const dialog = screen().getByRole('dialog', { name: 'Font Color' });
    const input = within(dialog).getByLabelText('Hex color, such as FF0000') as HTMLInputElement;
    expect(document.activeElement).toBe(input);
    const apply = within(dialog).getByRole('button', { name: 'Apply' });
    fireEvent.change(input, { target: { value: 'zz' } });
    expect(apply.hasAttribute('disabled')).toBe(true);
    fireEvent.change(input, { target: { value: 'abcdef' } });
    fireEvent.submit(dialog);
    await act(async () => {});
    expect(harness.calls.at(-1)).toEqual({
      id: 'textColor',
      args: { color: { rgb: 'ABCDEF' } },
      ordered: true,
    });
    expect(screen().queryByRole('dialog')).toBeNull();
    expect(document.activeElement).toBe(more());

    choose(['Font size', 'Custom size…']);
    const size = screen().getByRole('dialog', { name: 'Font size' });
    fireEvent.change(within(size).getByLabelText('Size in points'), { target: { value: '13.5' } });
    fireEvent.submit(size);
    await act(async () => {});
    expect(harness.calls.at(-1)?.args).toEqual({ points: 13.5 });

    choose(['Font size', 'Custom size…']);
    const cancelled = screen().getByRole('dialog', { name: 'Font size' });
    fireEvent.keyDown(within(cancelled).getByLabelText('Size in points'), { key: 'Escape' });
    expect(screen().queryByRole('dialog')).toBeNull();
    expect(document.activeElement).toBe(more());
  });

  test('custom values apply only to the document and selection the prompt opened for', async () => {
    const harness = mountDefault();
    const submit = async (title: string, label: string, value: string) => {
      const dialog = screen().getByRole('dialog', { name: title });
      fireEvent.change(within(dialog).getByLabelText(label), { target: { value } });
      fireEvent.submit(dialog);
      await act(async () => {});
    };
    const ids = () => harness.calls.map((call) => call.id);

    choose(['Font Color', 'Custom color…']);
    harness.state.document = {};
    await submit('Font Color', 'Hex color, such as FF0000', '00FF00');
    choose(['Table', 'Cell Fill Color', 'Custom color…']);
    harness.state.document = {};
    await submit('Cell Fill Color', 'Hex color, such as FF0000', '00FF00');
    expect(ids()).toEqual([]);

    choose(['Font size', 'Custom size…']);
    harness.state.targetChanged = true;
    await submit('Font size', 'Size in points', '13');
    expect(ids()).toEqual([]);

    harness.state.targetChanged = false;
    choose(['Font size', 'Custom size…']);
    await submit('Font size', 'Size in points', '13');
    expect(harness.calls).toEqual([{ id: 'fontSize', args: { points: 13 }, ordered: true }]);
  });

  test('host-composed table insertion offers every size of the grid', async () => {
    const harness = mountDefault(<ToolbarCommand id="insertTable" />);
    const tree = menuTree(openWithKeyboard());
    fireEvent.keyDown(document.activeElement!, { key: 'Escape' });
    const rows = find(tree, ['Table']).children!;
    expect(rows.map((row) => row.label)).toEqual([1, 2, 3, 4, 5, 6].map((n) => `Rows: ${n}`));
    expect(rows.flatMap((row) => row.children!.map((cell) => cell.label))).toHaveLength(36);
    choose(['Table', 'Rows: 3', '4 x 3 Table']);
    await act(async () => {});
    expect(harness.calls.at(-1)).toEqual({
      id: 'insertTable',
      args: { rows: 3, columns: 4 },
      ordered: true,
    });
  });
});

