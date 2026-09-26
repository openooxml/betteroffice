import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { createT, en } from '@betteroffice/xlsx-i18n';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import {
  EditorToolbar,
  Toolbar,
  ToolbarButton,
  ToolbarCommand,
  ToolbarCommandButton,
  ToolbarGroup,
  ToolbarOverflow,
  XlsxCommandProvider,
} from '../../index';
import { createXlsxCommandController } from '../../commands/createXlsxCommandStore';
import { testBinding } from '../../commands/testing';

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');

function screen() {
  return within(document.body);
}

let railWidth = 4000;
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
    if (this.querySelector(':scope > [data-testid="xlsx-toolbar-more"]')) return rect(28);
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

function mount(overrides = {}) {
  railWidth = 4000;
  const harness = testBinding({ translate: createT(en), canUndo: false, ...overrides });
  const controller = createXlsxCommandController();
  controller.attach(harness.binding);
  let shared = 0;
  const view = render(
    <XlsxCommandProvider commands={controller.store}>
      <EditorToolbar mode="commands">
        <EditorToolbar.Toolbar>
          <ToolbarGroup label="Formatting">
            <ToolbarCommandButton id="bold" />
            <ToolbarCommandButton id="italic" />
          </ToolbarGroup>
          <ToolbarGroup label="History">
            <ToolbarCommandButton id="undo" />
            <ToolbarCommandButton id="redo" />
          </ToolbarGroup>
          <ToolbarGroup label="Layout">
            <ToolbarCommand id="horizontalAlignment" />
            <ToolbarCommand id="fontSize" />
            <ToolbarCommand id="textColor" />
          </ToolbarGroup>
          <ToolbarButton title="Share" onClick={() => (shared += 1)}>
            Share
          </ToolbarButton>
          <ToolbarOverflow label="Custom widget" onSelect={() => (shared += 10)}>
            <span>Custom widget</span>
          </ToolbarOverflow>
          <span data-testid="unrepresented">Plain</span>
        </EditorToolbar.Toolbar>
      </EditorToolbar>
    </XlsxCommandProvider>
  );
  return { harness, controller, view, shared: () => shared };
}

function more(): HTMLElement {
  return screen().getByTestId('xlsx-toolbar-more');
}

function openWithKeyboard(key = 'ArrowDown') {
  const trigger = more();
  act(() => trigger.focus());
  fireEvent.keyDown(trigger, { key });
  return screen().getByRole('menu', { name: 'More toolbar items' });
}

function menuItems(menu: HTMLElement): HTMLElement[] {
  return Array.from(
    menu.querySelectorAll<HTMLElement>(
      '[role="menuitem"], [role="menuitemcheckbox"], [role="menuitemradio"]'
    )
  ).filter((item) => item.closest('[role="menu"]') === menu);
}

function labels(menu: HTMLElement): (string | undefined)[] {
  return menuItems(menu).map((item) => item.dataset.label);
}

describe('xlsx toolbar overflow', () => {
  test('keeps everything in the row when it fits', () => {
    mount();
    expect(screen().queryByTestId('xlsx-toolbar-more')).toBeNull();
  });

  test('moves whole trailing groups into More and keeps every action reachable', () => {
    mount();
    resize(200);
    const toolbar = screen().getByRole('toolbar');
    const group = (label: string) =>
      toolbar.querySelector<HTMLElement>(`[role="group"][aria-label="${label}"]`)!;
    expect(group('Formatting').getAttribute('aria-hidden')).toBeNull();
    expect(group('Layout').getAttribute('aria-hidden')).toBe('true');
    expect((group('Layout') as HTMLElement & { inert: boolean }).inert).toBe(true);
    expect(screen().getByTestId('unrepresented').getAttribute('aria-hidden')).toBeNull();

    const menu = openWithKeyboard();
    expect(labels(menu)).toEqual([
      'Horizontal alignment',
      'Font size',
      'Text color',
      'Share',
      'Custom widget',
    ]);
    within(menu).getByRole('group', { name: 'Layout' });

    resize(10);
    expect(labels(openWithKeyboard())).toEqual([
      'Bold',
      'Italic',
      'Undo',
      'Redo',
      'Horizontal alignment',
      'Font size',
      'Text color',
      'Share',
      'Custom widget',
    ]);
  });

  test('keeps a group with a control that has no menu entry in the row', () => {
    const shared: string[] = [];
    railWidth = 4000;
    const controller = createXlsxCommandController();
    controller.attach(testBinding({ translate: createT(en) }).binding);
    render(
      <XlsxCommandProvider commands={controller.store}>
        <EditorToolbar mode="commands">
          <EditorToolbar.Toolbar>
            <ToolbarGroup label="Sharing">
              <ToolbarCommandButton id="bold" />
              <button type="button" onClick={() => shared.push('share')}>
                Share
              </button>
            </ToolbarGroup>
            <ToolbarGroup label="History">
              <ToolbarCommandButton id="undo" />
              <ToolbarCommandButton id="redo" />
            </ToolbarGroup>
          </EditorToolbar.Toolbar>
        </EditorToolbar>
      </XlsxCommandProvider>
    );
    resize(10);
    const toolbar = screen().getByRole('toolbar');
    const group = (label: string) =>
      toolbar.querySelector<HTMLElement>(`[role="group"][aria-label="${label}"]`)!;
    expect(group('Sharing').getAttribute('aria-hidden')).toBeNull();
    expect((group('Sharing') as HTMLElement & { inert?: boolean }).inert).toBeFalsy();
    fireEvent.click(within(toolbar).getByRole('button', { name: 'Share' }));
    expect(shared).toEqual(['share']);
    expect(group('History').getAttribute('aria-hidden')).toBe('true');
    expect(labels(openWithKeyboard())).toEqual(['Undo', 'Redo']);
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
    fireEvent.keyDown(document.activeElement!, { key: 'Home' });
    fireEvent.keyDown(document.activeElement!, { key: 'r' });
    expect(document.activeElement?.getAttribute('data-label')).toBe('Redo');
    fireEvent.keyDown(document.activeElement!, { key: 'Escape' });
    expect(screen().queryByRole('menu')).toBeNull();
    expect(document.activeElement).toBe(more());
  });

  test('explains disabled entries and runs enabled ones', async () => {
    const { harness, shared } = mount();
    resize(10);
    const menu = openWithKeyboard();
    const undo = menuItems(menu).find((item) => item.dataset.label === 'Undo')!;
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    expect(document.getElementById(undo.getAttribute('aria-describedby')!)?.textContent).toBe(
      'Nothing to undo.'
    );
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

    const share = menuItems(openWithKeyboard()).find((item) => item.dataset.label === 'Share')!;
    fireEvent.click(share);
    expect(shared()).toBe(1);
  });

  test('presents choices as radio submenus', async () => {
    const { harness } = mount();
    resize(10);
    const menu = openWithKeyboard();
    const alignment = menuItems(menu).find(
      (item) => item.dataset.label === 'Horizontal alignment'
    )!;
    expect(alignment.getAttribute('aria-haspopup')).toBe('menu');
    act(() => alignment.focus());
    fireEvent.keyDown(alignment, { key: 'ArrowRight' });
    const options = menuItems(screen().getByRole('menu', { name: 'Horizontal alignment' }));
    expect(options.map((item) => [item.getAttribute('role'), item.getAttribute('aria-checked')])).toEqual([
      ['menuitemradio', 'true'],
      ['menuitemradio', 'false'],
      ['menuitemradio', 'false'],
    ]);
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowDown' });
    fireEvent.keyDown(document.activeElement!, { key: 'Enter' });
    await act(async () => {});
    expect(harness.calls).toEqual([
      { id: 'horizontalAlignment', args: { value: 'center' }, ordered: true },
    ]);
  });

  test('asks for a custom font size in a dialog bound to the selection it opened with', async () => {
    const { harness } = mount();
    resize(10);
    const size = menuItems(openWithKeyboard()).find((item) => item.dataset.label === 'Font size')!;
    act(() => size.focus());
    fireEvent.keyDown(size, { key: 'ArrowRight' });
    const submenu = screen().getByRole('menu', { name: 'Font size' });
    const custom = menuItems(submenu).find((item) => item.dataset.label === 'Custom size…')!;
    act(() => custom.focus());
    fireEvent.keyDown(custom, { key: 'Enter' });
    const dialog = screen().getByRole('dialog', { name: 'Font size' });
    const input = within(dialog).getByLabelText('Font size in points') as HTMLInputElement;
    expect(document.activeElement).toBe(input);
    expect(input.value).toBe('11');
    fireEvent.change(input, { target: { value: '500' } });
    expect((within(dialog).getByRole('button', { name: 'Apply' }) as HTMLButtonElement).disabled).toBe(true);
    fireEvent.change(input, { target: { value: '13' } });
    fireEvent.submit(dialog);
    await act(async () => {});
    expect(screen().queryByRole('dialog')).toBeNull();
    expect(document.activeElement).toBe(more());
    expect(harness.calls).toEqual([{ id: 'fontSize', args: { points: 13 }, ordered: true }]);
  });

  test('moves focus to More when a resize hides the focused control', () => {
    mount();
    const share = screen().getByRole('button', { name: 'Share' });
    act(() => share.focus());
    resize(120);
    expect(document.activeElement).toBe(more());
  });

  test('keeps legacy appended content reachable in a dialog, still mounted', () => {
    railWidth = 4000;
    let clicks = 0;
    render(
      <Toolbar onFormat={() => {}}>
        <button type="button" onClick={() => (clicks += 1)}>
          Legacy tool
        </button>
      </Toolbar>
    );
    const tool = screen().getByRole('button', { name: 'Legacy tool' });
    resize(10);
    const entry = menuItems(openWithKeyboard()).find((item) => item.dataset.label === 'More tools')!;
    act(() => entry.focus());
    fireEvent.keyDown(entry, { key: 'Enter' });
    const dialog = screen().getByRole('dialog', { name: 'More tools' });
    const moved = within(dialog).getByRole('button', { name: 'Legacy tool' });
    expect(moved).toBe(tool);
    expect(document.activeElement).toBe(tool);
    fireEvent.click(tool);
    expect(clicks).toBe(1);
    fireEvent.keyDown(tool, { key: 'Escape' });
    expect(screen().queryByRole('dialog')).toBeNull();
    expect(document.activeElement).toBe(more());
  });
});
