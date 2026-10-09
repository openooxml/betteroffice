import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import { MenuDropdown } from './MenuDropdown';
import { TableGridInline } from './TableGridInline';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

const { act, cleanup, fireEvent, render, within } = await import('@testing-library/react');

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function mount() {
  let activated = 0;
  const view = render(
    <>
      <button>Before</button>
      <MenuDropdown
        label="Insert"
        items={[
          { label: 'Image', onClick: () => activated++ },
          { label: 'Unavailable', disabled: true, onClick: () => activated++ },
          { type: 'separator' },
          {
            label: 'Table',
            submenuRole: 'dialog',
            submenuContent: () => <TableGridInline onInsert={() => activated++} />,
          },
          {
            label: 'Break',
            submenuContent: (close) => (
              <>
                <button role="menuitem" tabIndex={-1} aria-disabled="true">
                  Unavailable break
                </button>
                <button
                  role="menuitem"
                  tabIndex={-1}
                  onClick={() => {
                    activated++;
                    close();
                  }}
                >
                  Page break
                </button>
                <button role="menuitem" tabIndex={-1}>
                  Section break
                </button>
              </>
            ),
          },
        ]}
      />
      <button>After</button>
    </>
  );
  const trigger = view.getByRole('menuitem', { name: 'Insert' });
  fireEvent.keyDown(trigger, { key: 'ArrowDown' });
  return { ...view, trigger, activated: () => activated };
}

describe('MenuDropdown keyboard interaction', () => {
  test('opens with keyboard focus and navigates disabled items and separators', () => {
    const view = mount();
    const menu = view.getByRole('menu', { name: 'Insert' });
    const image = within(menu).getByRole('menuitem', { name: 'Image' });
    const unavailable = within(menu).getByRole('menuitem', { name: 'Unavailable' });
    const table = within(menu).getByRole('menuitem', { name: 'Table' });
    expect(document.activeElement === image).toBe(true);
    fireEvent.keyDown(image, { key: 'ArrowDown' });
    expect(document.activeElement === unavailable).toBe(true);
    fireEvent.click(unavailable);
    expect(view.activated()).toBe(0);
    expect(view.queryByRole('menu', { name: 'Insert' })).not.toBeNull();
    fireEvent.keyDown(unavailable, { key: 'ArrowDown' });
    expect(document.activeElement === table).toBe(true);
    fireEvent.keyDown(table, { key: 'End' });
    const breakItem = within(menu).getByRole('menuitem', { name: 'Break' });
    expect(document.activeElement === breakItem).toBe(true);
    fireEvent.keyDown(breakItem, { key: 'ArrowDown' });
    expect(document.activeElement === image).toBe(true);
    fireEvent.keyDown(image, { key: 'ArrowUp' });
    expect(document.activeElement === breakItem).toBe(true);
    fireEvent.keyDown(breakItem, { key: 'Home' });
    expect(document.activeElement === image).toBe(true);
  });

  test.each(['ArrowRight', 'Enter', ' '])(
    'opens a submenu with %s, including an already hovered submenu',
    (key) => {
      const view = mount();
      const item = view.getByRole('menuitem', { name: 'Break' });
      fireEvent.mouseEnter(item.parentElement!);
      expect(view.getByRole('menu', { name: 'Break' })).not.toBeNull();
      fireEvent.keyDown(item, { key });
      expect(
        document.activeElement === view.getByRole('menuitem', { name: 'Unavailable break' })
      ).toBe(true);
    }
  );

  test('shares arrow navigation with nested menus and returns focus one level on Escape', () => {
    const view = mount();
    const item = view.getByRole('menuitem', { name: 'Break' });
    fireEvent.keyDown(item, { key: 'ArrowRight' });
    const first = view.getByRole('menuitem', { name: 'Unavailable break' });
    const second = view.getByRole('menuitem', { name: 'Page break' });
    fireEvent.keyDown(first, { key: 'ArrowDown' });
    expect(document.activeElement === second).toBe(true);
    fireEvent.keyDown(second, { key: 'Escape' });
    expect(view.queryByRole('menu', { name: 'Break' })).toBeNull();
    expect(view.queryByRole('menu', { name: 'Insert' })).not.toBeNull();
    expect(document.activeElement === item).toBe(true);
    fireEvent.keyDown(item, { key: 'Escape' });
    expect(view.queryByRole('menu', { name: 'Insert' })).toBeNull();
    expect(document.activeElement === view.trigger).toBe(true);
  });

  test('returns to the parent item with ArrowLeft', () => {
    const view = mount();
    const item = view.getByRole('menuitem', { name: 'Break' });
    fireEvent.keyDown(item, { key: 'Enter' });
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowLeft' });
    expect(view.queryByRole('menu', { name: 'Break' })).toBeNull();
    expect(document.activeElement === item).toBe(true);
  });

  test.each([false, true])('Tab exits the whole menu with shiftKey=%s', (shiftKey) => {
    const view = mount();
    const item = view.getByRole('menuitem', { name: 'Break' });
    fireEvent.keyDown(item, { key: 'Enter' });
    expect(
      Array.from(
        view.getByRole('menu', { name: 'Insert' }).querySelectorAll('[role="menuitem"]')
      ).every((item) => (item as HTMLElement).tabIndex === -1)
    ).toBe(true);
    const event = new KeyboardEvent('keydown', {
      key: 'Tab',
      shiftKey,
      bubbles: true,
      cancelable: true,
    });
    fireEvent(document.activeElement!, event);
    expect(event.defaultPrevented).toBe(false);
    expect(view.queryByRole('menu', { name: 'Insert' })).toBeNull();
    expect(document.activeElement === view.trigger).toBe(true);
  });

  test('closes a submenu when navigating to another parent item', () => {
    const view = mount();
    const item = view.getByRole('menuitem', { name: 'Table' });
    act(() => item.focus());
    fireEvent.mouseEnter(item.parentElement!);
    fireEvent.keyDown(item, { key: 'ArrowDown' });
    expect(view.queryByRole('grid')).toBeNull();
    expect(document.activeElement === view.getByRole('menuitem', { name: 'Break' })).toBe(true);
  });

  test('keeps a keyboard focused submenu mounted when the pointer leaves', () => {
    const view = mount();
    const item = view.getByRole('menuitem', { name: 'Table' });
    fireEvent.keyDown(item, { key: 'Enter' });
    const cell = view.getAllByRole('gridcell')[0];
    expect(document.activeElement === cell).toBe(true);
    fireEvent.mouseLeave(item.parentElement!);
    expect(cell.isConnected).toBe(true);
    expect(document.activeElement === cell).toBe(true);
  });

  test('leaving another item with the pointer keeps the focused submenu mounted', () => {
    const view = mount();
    const table = view.getByRole('menuitem', { name: 'Table' });
    fireEvent.keyDown(table, { key: 'Enter' });
    const cell = view.getAllByRole('gridcell')[0];
    fireEvent.mouseLeave(view.getByRole('menuitem', { name: 'Image' }).parentElement!);
    expect(cell.isConnected).toBe(true);
    expect(document.activeElement === cell).toBe(true);
  });

  test('switching submenus with the pointer returns focus to the new parent item', () => {
    const view = mount();
    fireEvent.keyDown(view.getByRole('menuitem', { name: 'Table' }), { key: 'Enter' });
    const breakItem = view.getByRole('menuitem', { name: 'Break' });
    fireEvent.mouseEnter(breakItem.parentElement!);
    expect(view.queryByRole('grid')).toBeNull();
    expect(view.getByRole('menu', { name: 'Break' })).not.toBeNull();
    expect(document.activeElement === breakItem).toBe(true);
  });

  test('returns focus after activation but preserves focus set by a command', () => {
    const view = mount();
    fireEvent.click(document.activeElement!);
    expect(view.activated()).toBe(1);
    expect(document.activeElement === view.trigger).toBe(true);
    cleanup();
    const restored = render(
      <>
        <button>Document</button>
        <MenuDropdown
          label="Insert"
          items={[
            {
              label: 'Command',
              onClick: () => document.querySelector<HTMLButtonElement>('button')!.focus(),
            },
          ]}
        />
      </>
    );
    fireEvent.keyDown(restored.getByRole('menuitem', { name: 'Insert' }), { key: 'ArrowDown' });
    fireEvent.click(document.activeElement!);
    expect(document.activeElement === restored.getByRole('button', { name: 'Document' })).toBe(
      true
    );
  });

  test('moves horizontally between menubar triggers instead of opening their menus', () => {
    const view = render(
      <div role="menubar">
        <MenuDropdown label="File" items={[{ label: 'Open' }]} />
        <MenuDropdown label="Insert" items={[{ label: 'Image' }]} />
      </div>
    );
    const first = view.getByRole('menuitem', { name: 'File' });
    const last = view.getByRole('menuitem', { name: 'Insert' });
    act(() => first.focus());
    fireEvent.keyDown(first, { key: 'ArrowRight' });
    expect(document.activeElement === last).toBe(true);
    expect(view.queryByRole('menu')).toBeNull();
    fireEvent.keyDown(last, { key: 'ArrowRight' });
    expect(document.activeElement === first).toBe(true);
    fireEvent.keyDown(first, { key: 'ArrowLeft' });
    expect(document.activeElement === last).toBe(true);
  });

  test('moves from a popup to the adjacent menubar menu and closes the old one', () => {
    const view = render(
      <div role="menubar">
        <MenuDropdown label="File" items={[{ label: 'Open' }]} />
        <MenuDropdown label="Insert" items={[{ label: 'Image' }]} />
      </div>
    );
    fireEvent.keyDown(view.getByRole('menuitem', { name: 'File' }), { key: 'ArrowDown' });
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowRight' });
    expect(view.queryByRole('menu', { name: 'File' })).toBeNull();
    expect(view.getByRole('menu', { name: 'Insert' })).not.toBeNull();
    expect(document.activeElement === view.getByRole('menuitem', { name: 'Insert' })).toBe(true);
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowDown' });
    expect(document.activeElement === view.getByRole('menuitem', { name: 'Image' })).toBe(true);
  });

  test('does not close a popup when horizontal navigation wraps to its only menubar trigger', () => {
    const view = render(
      <div role="menubar">
        <MenuDropdown label="Insert" items={[{ label: 'Image' }]} />
      </div>
    );
    const trigger = view.getByRole('menuitem', { name: 'Insert' });
    fireEvent.keyDown(trigger, { key: 'ArrowDown' });
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowRight' });
    expect(view.queryByRole('menu', { name: 'Insert' })).not.toBeNull();
    expect(document.activeElement === trigger).toBe(true);
  });

  test('does not open or activate a disabled submenu', () => {
    const view = render(
      <MenuDropdown
        label="Insert"
        items={[
          {
            label: 'Table',
            disabled: true,
            submenuContent: () => <TableGridInline onInsert={() => {}} />,
          },
        ]}
      />
    );
    fireEvent.keyDown(view.getByRole('menuitem', { name: 'Insert' }), { key: 'ArrowDown' });
    const item = view.getByRole('menuitem', { name: 'Table' });
    expect(document.activeElement === item).toBe(true);
    expect(item.getAttribute('aria-disabled')).toBe('true');
    fireEvent.mouseEnter(item.parentElement!);
    fireEvent.keyDown(item, { key: 'Enter' });
    fireEvent.click(item);
    expect(view.queryByRole('grid')).toBeNull();
  });

  test('keeps disabled menubar items focusable without allowing them to open', () => {
    const view = render(
      <div role="menubar">
        <MenuDropdown label="File" items={[{ label: 'Open' }]} />
        <MenuDropdown label="Insert" disabled items={[{ label: 'Image' }]} />
      </div>
    );
    fireEvent.keyDown(view.getByRole('menuitem', { name: 'File' }), { key: 'ArrowRight' });
    const disabled = view.getByRole('menuitem', { name: 'Insert' });
    expect(document.activeElement === disabled).toBe(true);
    fireEvent.keyDown(disabled, { key: 'Enter' });
    fireEvent.click(disabled);
    expect(view.queryByRole('menu')).toBeNull();
    fireEvent.keyDown(disabled, { key: 'ArrowLeft' });
    expect(document.activeElement === view.getByRole('menuitem', { name: 'File' })).toBe(true);
  });
});
