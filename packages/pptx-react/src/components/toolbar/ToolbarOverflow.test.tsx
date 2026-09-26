import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import {
  EditorToolbar,
  PptxCommandProvider,
  Toolbar,
  ToolbarButton,
  ToolbarCommand,
  ToolbarCommandButton,
  ToolbarDropdown,
  ToolbarGroup,
  ToolbarMenuItem,
  ToolbarOverflow,
} from '../../index';
import { createPptxCommandController } from '../../commands/createPptxCommandStore';
import { testBinding } from '../../commands/testing';
import { LocaleProvider } from '../../i18n';

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');

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

beforeAll(() => {
  HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
    if (this.getAttribute('role') === 'toolbar') return rect(railWidth);
    if (this.parentElement?.hasAttribute('data-toolbar-items')) {
      return rect(Math.max(24, this.querySelectorAll('button, input, select').length * 30));
    }
    if (this.querySelector(':scope > span > [data-testid="pptx-toolbar-more"]')) return rect(28);
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

async function settle() {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

function mount(overrides: Parameters<typeof testBinding>[0] = {}) {
  railWidth = 1000;
  const harness = testBinding(overrides);
  const controller = createPptxCommandController();
  controller.attach(harness.binding);
  let shared = 0;
  let widget = 0;
  render(
    <PptxCommandProvider commands={controller.store}>
      <EditorToolbar mode="commands">
        <EditorToolbar.Toolbar>
          <ToolbarGroup label="Text">
            <ToolbarCommandButton id="bold" />
            <ToolbarCommandButton id="italic" />
          </ToolbarGroup>
          <ToolbarGroup label="History">
            <ToolbarCommandButton id="undo" />
            <ToolbarCommandButton id="redo" />
          </ToolbarGroup>
          <ToolbarGroup label="Paragraph">
            <ToolbarCommand id="alignment" />
          </ToolbarGroup>
          <ToolbarGroup label="Size">
            <ToolbarCommand id="fontSize" />
          </ToolbarGroup>
          <ToolbarCommand id="zOrder" />
          <ToolbarButton title="Share" onClick={() => (shared += 1)}>
            Share
          </ToolbarButton>
          <ToolbarOverflow label="Custom widget" onSelect={() => (widget += 1)}>
            <span>Custom</span>
          </ToolbarOverflow>
          <ToolbarDropdown title="Host menu" trigger="Menu">
            {(close) => (
              <ToolbarMenuItem label="Host item" onClick={() => (shared += 10)} close={close} />
            )}
          </ToolbarDropdown>
          <span data-testid="unrepresented">Plain</span>
        </EditorToolbar.Toolbar>
      </EditorToolbar>
    </PptxCommandProvider>
  );
  return { harness, shared: () => shared, widget: () => widget };
}

function more(): HTMLElement {
  return screen().getByTestId('pptx-toolbar-more');
}

function openMenu(key = 'ArrowDown'): HTMLElement {
  const trigger = more();
  act(() => trigger.focus());
  fireEvent.keyDown(trigger, { key });
  return screen().getByRole('menu', { name: 'More' });
}

function items(menu: HTMLElement): HTMLElement[] {
  return Array.from(
    menu.querySelectorAll<HTMLElement>(
      '[role="menuitem"], [role="menuitemcheckbox"], [role="menuitemradio"]'
    )
  ).filter((item) => item.closest('[role="menu"]') === menu);
}

function key(target: Element, value: string) {
  fireEvent.keyDown(target, { key: value });
}

describe('PPTX toolbar overflow', () => {
  test('keeps every control in the row while it fits', () => {
    mount();
    expect(screen().queryByTestId('pptx-toolbar-more')).toBeNull();
  });

  test('moves trailing units into an accessible More menu, in host order', () => {
    mount();
    resize(200);
    expect(screen().getByTestId('pptx-bold').closest('[aria-hidden="true"]')).toBeNull();
    expect(screen().getByTestId('pptx-align-left').closest('[aria-hidden="true"]')).not.toBeNull();
    expect(screen().getByTestId('unrepresented').closest('[aria-hidden="true"]')).toBeNull();
    const menu = openMenu();
    expect(items(menu).map((item) => item.dataset.label)).toEqual([
      'Alignment',
      'Font size',
      'Arrange',
      'Share',
      'Custom widget',
      'Host menu',
    ]);
    expect(document.activeElement).toBe(items(menu)[0]);
    expect(menu.querySelector('[role="group"]')?.getAttribute('aria-labelledby')).toBeTruthy();
  });

  test('keeps a group with a control that has no menu entry in the row', () => {
    const shared: string[] = [];
    railWidth = 1000;
    const controller = createPptxCommandController();
    controller.attach(testBinding().binding);
    render(
      <PptxCommandProvider commands={controller.store}>
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
      </PptxCommandProvider>
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
    expect(items(openMenu()).map((item) => item.dataset.label)).toEqual(['Undo', 'Redo']);
  });

  test('navigates with arrows, Home, End and typeahead; Enter runs and returns focus', async () => {
    const { shared } = mount();
    resize(200);
    const menu = openMenu();
    key(document.activeElement!, 'End');
    expect((document.activeElement as HTMLElement).dataset.label).toBe('Host menu');
    key(document.activeElement!, 'Home');
    expect((document.activeElement as HTMLElement).dataset.label).toBe('Alignment');
    key(document.activeElement!, 'ArrowDown');
    expect((document.activeElement as HTMLElement).dataset.label).toBe('Font size');
    key(document.activeElement!, 's');
    expect((document.activeElement as HTMLElement).dataset.label).toBe('Share');
    key(document.activeElement!, 'Enter');
    expect(screen().queryByRole('menu', { name: 'More' })).toBeNull();
    expect(document.activeElement).toBe(more());
    expect(shared()).toBe(1);
    expect(menu.isConnected).toBe(false);
  });

  test('presents selectors as radio submenus bound to the commands', async () => {
    const { harness } = mount();
    resize(200);
    openMenu();
    key(document.activeElement!, 'ArrowRight');
    const submenu = screen().getByRole('menu', { name: 'Alignment' });
    const radios = items(submenu);
    expect(
      radios.map((item) => [
        item.getAttribute('role'),
        item.dataset.label,
        item.getAttribute('aria-checked'),
      ])
    ).toEqual([
      ['menuitemradio', 'Align Left', 'true'],
      ['menuitemradio', 'Center', 'false'],
      ['menuitemradio', 'Align Right', 'false'],
      ['menuitemradio', 'Justify', 'false'],
    ]);
    expect(document.activeElement).toBe(radios[0]);
    key(document.activeElement!, 'ArrowDown');
    key(document.activeElement!, 'Enter');
    await settle();
    expect(harness.calls).toEqual([{ id: 'alignment', args: { value: 'ctr' } }]);
    expect(document.activeElement).toBe(more());
  });

  test('Escape closes a submenu, then the menu, restoring focus each time', () => {
    mount();
    resize(200);
    openMenu();
    const alignment = document.activeElement as HTMLElement;
    key(alignment, 'ArrowRight');
    key(document.activeElement!, 'Escape');
    expect(document.activeElement).toBe(alignment);
    key(alignment, 'Escape');
    expect(screen().queryByRole('menu', { name: 'More' })).toBeNull();
    expect(document.activeElement).toBe(more());
  });

  test('describes disabled entries and mixed toggles', () => {
    mount({
      canUndo: false,
      text: {
        kind: 'range',
        bold: 'mixed',
        italic: false,
        underline: false,
        fontFamily: null,
        fontSize: null,
        color: null,
        alignment: 'l',
      },
    });
    resize(40);
    const menu = openMenu();
    const bold = items(menu).find((item) => item.dataset.label === 'Bold')!;
    expect(bold.getAttribute('role')).toBe('menuitemcheckbox');
    expect(bold.getAttribute('aria-checked')).toBe('mixed');
    const undo = items(menu).find((item) => item.dataset.label === 'Undo')!;
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    expect(document.getElementById(undo.getAttribute('aria-describedby')!)?.textContent).toBe(
      'There is nothing to undo.'
    );
  });

  test('moves focus to More when a resize hides the focused control', () => {
    mount();
    const bold = screen().getByTestId('pptx-bold');
    act(() => bold.focus());
    resize(40);
    expect(document.activeElement).toBe(more());
  });

  test('opens a hidden host dropdown from the menu and returns focus to More', async () => {
    const { shared } = mount();
    resize(200);
    openMenu('ArrowUp');
    expect((document.activeElement as HTMLElement).dataset.label).toBe('Host menu');
    key(document.activeElement!, 'Enter');
    await settle();
    const popup = screen().getByRole('menu', { name: 'Host menu' });
    expect(document.activeElement).toBe(within(popup).getByRole('menuitem', { name: 'Host item' }));
    key(document.activeElement!, 'Escape');
    expect(document.activeElement).toBe(more());
    expect(shared()).toBe(0);
  });

  test('asks for a custom font size in an accessible dialog', async () => {
    const { harness } = mount();
    resize(200);
    openMenu();
    key(document.activeElement!, 'ArrowDown');
    key(document.activeElement!, 'ArrowRight');
    const sizes = screen().getByRole('menu', { name: 'Font size' });
    const custom = items(sizes).find((item) => item.dataset.label === 'Custom size…')!;
    act(() => custom.focus());
    key(custom, 'Enter');
    const dialog = screen().getByRole('dialog', { name: 'Font size' });
    const input = within(dialog).getByLabelText('Size in points') as HTMLInputElement;
    expect(document.activeElement).toBe(input);
    fireEvent.change(input, { target: { value: '30' } });
    fireEvent.submit(dialog);
    await settle();
    expect(harness.calls).toEqual([{ id: 'fontSize', args: { points: 30 } }]);
    expect(document.activeElement).toBe(more());
  });

  test('keeps legacy host content reachable through a popup', () => {
    railWidth = 1000;
    render(
      <LocaleProvider>
        <Toolbar>
          <input aria-label="Host field" />
        </Toolbar>
      </LocaleProvider>
    );
    resize(40);
    const menu = openMenu('ArrowUp');
    expect((document.activeElement as HTMLElement).dataset.label).toBe('More controls');
    key(document.activeElement!, 'Enter');
    const popup = screen().getByRole('dialog', { name: 'More controls' });
    expect(document.activeElement).toBe(within(popup).getByLabelText('Host field'));
    key(document.activeElement!, 'Escape');
    expect(screen().queryByRole('dialog', { name: 'More controls' })).toBeNull();
    expect(document.activeElement).toBe(more());
    expect(menu.isConnected).toBe(false);
  });
});
