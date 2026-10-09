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

  test.each([
    ['converges when hidden groups measure narrower', false],
    ['converges when groups measure narrower while More takes row space', true],
  ] as const)('%s', (_name, rowDependent) => {
    const previousRect = HTMLElement.prototype.getBoundingClientRect;
    railWidth = 100;
    HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
      if (rowDependent && this.hasAttribute('data-toolbar-items')) {
        const trigger = this.parentElement?.querySelector<HTMLElement>(
          '[data-testid="xlsx-toolbar-more"]'
        )?.parentElement?.parentElement;
        const moreWidth =
          trigger && trigger.style.position !== 'absolute'
            ? trigger.getBoundingClientRect().width
            : 0;
        return rect(railWidth - moreWidth);
      }
      if (this.parentElement?.hasAttribute('data-toolbar-items')) {
        if (rowDependent) return rect(this.parentElement.getBoundingClientRect().width * 0.6);
        return rect(this.style.position === 'absolute' ? 20 : 100);
      }
      if (this.querySelector(':scope > span > [data-testid="xlsx-toolbar-more"]')) return rect(28);
      return previousRect.call(this);
    };
    try {
      const controller = createXlsxCommandController();
      controller.attach(testBinding({ translate: createT(en) }).binding);
      const tree = (
        <XlsxCommandProvider commands={controller.store}>
          <EditorToolbar mode="commands">
            <EditorToolbar.Toolbar style={{ padding: 0 }}>
              <ToolbarGroup label="A">
                <ToolbarCommandButton id="bold" />
              </ToolbarGroup>
              <ToolbarGroup label="B">
                <ToolbarCommandButton id="italic" />
              </ToolbarGroup>
            </EditorToolbar.Toolbar>
          </EditorToolbar>
        </XlsxCommandProvider>
      );
      expect(() => render(tree)).not.toThrow();
      const groups = document.querySelectorAll('[data-toolbar-items] > [role="group"]');
      const expectedHidden = [rowDependent ? null : 'true', 'true'];
      expect(Array.from(groups, (group) => group.getAttribute('aria-hidden'))).toEqual(expectedHidden);
      const trigger = screen().getByRole('button', { name: 'More toolbar items' });
      expect(trigger).toBe(more());
      expect(trigger.parentElement?.parentElement?.style.position).toBe('');
      expect(() => resize(100)).not.toThrow();
      expect(Array.from(groups, (group) => group.getAttribute('aria-hidden'))).toEqual(expectedHidden);
      expect(screen().getByRole('button', { name: 'More toolbar items' })).toBe(trigger);
      expect(trigger.parentElement?.parentElement?.style.position).toBe('');
      expect(labels(openWithKeyboard())).toEqual(rowDependent ? ['Italic'] : ['Bold', 'Italic']);
    } finally {
      cleanup();
      HTMLElement.prototype.getBoundingClientRect = previousRect;
    }
  });

  test('preserves scrolling on an unchanged measurement with focus outside the row', () => {
    railWidth = 10;
    const controller = createXlsxCommandController();
    controller.attach(testBinding({ translate: createT(en) }).binding);
    const tree = () => (
      <>
        <button type="button">Outside</button>
        <XlsxCommandProvider commands={controller.store}>
          <EditorToolbar mode="commands">
            <EditorToolbar.Toolbar style={{ padding: 0 }}>
              <span data-testid="unrepresented" style={{ flexShrink: 0 }}>
                Plain
              </span>
              <ToolbarGroup label="Formatting">
                <ToolbarCommandButton id="bold" />
              </ToolbarGroup>
            </EditorToolbar.Toolbar>
          </EditorToolbar>
        </XlsxCommandProvider>
      </>
    );
    const view = render(tree());
    const row = screen().getByRole('toolbar').querySelector<HTMLElement>('[data-toolbar-items]')!;
    const units = Array.from(row.children) as HTMLElement[];
    const trigger = more();
    const wrapper = trigger.parentElement!.parentElement!;
    const moreWidth = 28 + (parseFloat(getComputedStyle(wrapper).marginLeft) || 0);
    const unitWidth = (unit: HTMLElement) => (unit.dataset.testid === 'unrepresented' ? 100 : 20);
    let scrollLeft = 0;
    Object.defineProperties(row, {
      clientWidth: {
        get: () => railWidth - (wrapper.style.position === 'absolute' ? 0 : moreWidth),
      },
      scrollWidth: {
        get: () =>
          Math.max(
            row.clientWidth,
            units.reduce(
              (sum, unit) => sum + (unit.style.position === 'absolute' ? 0 : unitWidth(unit)),
              0
            )
          ),
      },
      scrollLeft: {
        get: () => scrollLeft,
        set: (value: number) => {
          scrollLeft = Math.max(0, Math.min(value, row.scrollWidth - row.clientWidth));
        },
      },
    });
    const mockRect = (element: HTMLElement, width: () => number) => {
      element.getBoundingClientRect = () => {
        row.scrollLeft = row.scrollLeft;
        return rect(width());
      };
    };
    mockRect(row, () => row.clientWidth);
    mockRect(wrapper, () => 28);
    units.forEach((unit) => mockRect(unit, () => unitWidth(unit)));
    resize(100);
    expect(row.style.overflowX).toBe('auto');
    expect(units[0].getAttribute('aria-hidden')).toBeNull();
    expect(units[1].getAttribute('aria-hidden')).toBe('true');
    act(() => screen().getByRole('button', { name: 'Outside' }).focus());
    expect(row.contains(document.activeElement)).toBe(false);
    expect(row.scrollWidth - row.clientWidth).toBe(28);
    row.scrollLeft = 28;
    expect(row.scrollLeft).toBe(28);
    view.rerender(tree());
    expect(more()).toBe(trigger);
    expect(row.scrollLeft).toBe(28);
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
    const row = toolbar.querySelector<HTMLElement>('[data-toolbar-items]')!;
    expect(row.style.overflowX).not.toBe('auto');

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

  test('keeps a group with a control that has no menu entry reachable in the row', () => {
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
    const share = within(toolbar).getByRole('button', { name: 'Share' });
    fireEvent.click(share);
    expect(shared).toEqual(['share']);
    const row = toolbar.querySelector<HTMLElement>('[data-toolbar-items]')!;
    expect(row.style.overflowX).toBe('auto');
    row.getBoundingClientRect = () => rect(10);
    share.getBoundingClientRect = () =>
      ({ ...rect(28), left: 30 - row.scrollLeft, right: 58 - row.scrollLeft }) as DOMRect;
    act(() => share.focus());
    expect(document.activeElement).toBe(share);
    expect(row.scrollLeft).toBe(30);
    row.scrollLeft = 0;
    resize(9);
    expect(row.scrollLeft).toBe(30);
    expect(group('History').getAttribute('aria-hidden')).toBe('true');
    expect(labels(openWithKeyboard())).toEqual(['Undo', 'Redo']);
    resize(4000);
    expect(row.style.overflowX).not.toBe('auto');
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
