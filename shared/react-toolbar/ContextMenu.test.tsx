import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { act, useRef } from 'react';
import type { OverflowMenuEntry } from './MenuList';
import { loadMenuList } from './menuListLoader';

await loadMenuList();

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
const { createRoot } = await import('react-dom/client');
const { ContextMenu, commandMenuItem, isContextMenuKey, menuSections, useContextMenu } = await import('./ContextMenu');

let unmount: (() => void) | null = null;
afterEach(() => {
  act(() => unmount?.());
  unmount = null;
  document.body.innerHTML = '';
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function mount(node: React.ReactNode) {
  const host = document.createElement('div');
  document.body.append(host);
  const root = createRoot(host);
  act(() => root.render(node));
  unmount = () => root.unmount();
  return host;
}

function key(target: Element, init: KeyboardEventInit) {
  act(() => {
    target.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }));
  });
}

const item = (id: string, onSelect = () => {}, disabled = false): OverflowMenuEntry =>
  ({ kind: 'item', id, label: id, disabled, description: disabled ? `${id} is unavailable` : undefined, onSelect });

function focusOrigin() {
  const origin = document.createElement('button');
  origin.textContent = 'origin';
  document.body.append(origin);
  origin.focus();
  return origin;
}

test('a pointer-opened menu is a labelled menu that takes focus', () => {
  focusOrigin();
  mount(<ContextMenu label="Cell" point={{ x: 40, y: 50 }} entries={[item('Cut'), item('Paste', () => {}, true)]} onClose={() => {}} />);
  const menu = document.querySelector('[role="menu"]')!;
  expect(menu.getAttribute('aria-label')).toBe('Cell');
  expect(document.activeElement).toBe(menu);
  const items = [...menu.querySelectorAll('[role="menuitem"]')];
  expect(items.map((node) => node.textContent)).toEqual(['Cut', 'Paste']);
  expect(items[1].getAttribute('aria-disabled')).toBe('true');
  const described = document.getElementById(items[1].getAttribute('aria-describedby')!);
  expect(described?.textContent).toBe('Paste is unavailable');
});

test('Escape outside the menu closes it without moving focus', () => {
  focusOrigin();
  const closes: string[] = [];
  mount(<ContextMenu label="Cell" point={{ x: 0, y: 0 }} entries={[item('Cut')]} onClose={(reason) => closes.push(reason)} />);
  const elsewhere = document.createElement('input');
  document.body.append(elsewhere);
  elsewhere.focus();
  key(elsewhere, { key: 'Escape' });
  expect(closes).toEqual(['escape']);
  expect(document.activeElement).toBe(elsewhere);
});

test('a keyboard-opened menu focuses its first item and arrows move between items', () => {
  focusOrigin();
  mount(<ContextMenu label="Cell" modality="keyboard" point={{ x: 0, y: 0 }} entries={[item('Cut'), item('Copy')]} onClose={() => {}} />);
  expect(document.activeElement?.textContent).toBe('Cut');
  key(document.activeElement!, { key: 'ArrowDown' });
  expect(document.activeElement?.textContent).toBe('Copy');
  key(document.activeElement!, { key: 'Home' });
  expect(document.activeElement?.textContent).toBe('Cut');
});

test('Enter runs the item, closes the menu and returns focus first', () => {
  const origin = focusOrigin();
  const events: string[] = [];
  mount(<ContextMenu label="Cell" modality="keyboard" point={{ x: 0, y: 0 }}
    entries={[item('Copy', () => events.push(`run:${document.activeElement === origin}`))]}
    onClose={(reason) => events.push(`close:${reason}`)} />);
  key(document.activeElement!, { key: 'Enter' });
  expect(events).toEqual(['close:select', 'run:true']);
});

test('a disabled item does not run', () => {
  focusOrigin();
  let ran = 0;
  const closes: string[] = [];
  mount(<ContextMenu label="Cell" modality="keyboard" point={{ x: 0, y: 0 }}
    entries={[item('Paste', () => { ran += 1; }, true)]} onClose={(reason) => closes.push(reason)} />);
  key(document.activeElement!, { key: 'Enter' });
  expect(ran).toBe(0);
  expect(closes).toEqual([]);
});

test('Escape closes and returns focus; a press outside closes without moving it', () => {
  const origin = focusOrigin();
  const closes: string[] = [];
  mount(<ContextMenu label="Cell" point={{ x: 0, y: 0 }} entries={[item('Cut')]} onClose={(reason) => closes.push(reason)} />);
  key(document.querySelector('[role="menu"]')!, { key: 'Escape' });
  expect(closes).toEqual(['escape']);
  expect(document.activeElement).toBe(origin);
  act(() => unmount?.());
  const other = document.createElement('input');
  document.body.append(other);
  mount(<ContextMenu label="Cell" point={{ x: 0, y: 0 }} entries={[item('Cut')]} onClose={(reason) => closes.push(reason)} />);
  act(() => { other.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, button: 0 })); });
  expect(closes).toEqual(['escape', 'outside']);
  act(() => {
    document.querySelector('[role="menu"]')?.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, button: 0 }));
  });
  expect(closes).toEqual(['escape', 'outside']);
});

test('ArrowRight opens a submenu whose items run, ArrowLeft closes it', () => {
  focusOrigin();
  const ran: string[] = [];
  mount(<ContextMenu label="Cell" modality="keyboard" point={{ x: 0, y: 0 }} onClose={() => {}} entries={[{
    kind: 'submenu', id: 'insert', label: 'Insert',
    entries: [item('Rows above', () => ran.push('above')), item('Rows below', () => ran.push('below'))],
  }]} />);
  const parent = document.activeElement!;
  expect(parent.getAttribute('aria-haspopup')).toBe('menu');
  key(parent, { key: 'ArrowRight' });
  expect(parent.getAttribute('aria-expanded')).toBe('true');
  expect(document.querySelector('[role="menu"][aria-label="Insert"]')).not.toBeNull();
  expect(document.activeElement?.textContent).toBe('Rows above');
  key(document.activeElement!, { key: 'ArrowLeft' });
  expect(document.querySelector('[role="menu"][aria-label="Insert"]')).toBeNull();
  expect(document.activeElement).toBe(parent);
  key(parent, { key: 'ArrowRight' });
  key(document.activeElement!, { key: 'ArrowDown' });
  key(document.activeElement!, { key: 'Enter' });
  expect(ran).toEqual(['below']);
});

function Surface(props: { visible?: boolean; resolve: Parameters<typeof useContextMenu>[0]['resolve'] }) {
  const surface = useRef<HTMLDivElement>(null);
  const menu = useContextMenu({ surface, resolve: props.resolve, keyboardPoint: () => ({ x: 7, y: 9 }) });
  return <>{props.visible !== false && <div ref={surface} tabIndex={0} data-testid="surface" />}{menu.element}</>;
}

test('attaches to a surface mounted after the hook', () => {
  const resolve = () => ({ label: 'Surface menu', entries: () => [item('Cut')] });
  const host = document.createElement('div');
  document.body.append(host);
  const root = createRoot(host);
  unmount = () => root.unmount();
  act(() => root.render(<Surface visible={false} resolve={resolve} />));
  act(() => root.render(<Surface resolve={resolve} />));
  const surface = host.querySelector('[data-testid="surface"]')!;
  const click = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, button: 2 });
  act(() => { surface.dispatchEvent(click); });
  expect(click.defaultPrevented).toBe(true);
  expect(document.querySelector('[role="menu"][aria-label="Surface menu"]')).not.toBeNull();
});

test('right-click, Shift+F10 and the Menu key open the menu on the surface', () => {
  const requests: string[] = [];
  const host = mount(<Surface resolve={(request) => {
    requests.push(`${request.modality}@${request.x},${request.y}`);
    return { label: 'Surface menu', entries: () => [item('Cut')] };
  }} />);
  const surface = host.querySelector<HTMLElement>('[data-testid="surface"]')!;
  surface.focus();
  const click = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, button: 2, clientX: 30, clientY: 40 });
  act(() => { surface.dispatchEvent(click); });
  expect(click.defaultPrevented).toBe(true);
  expect(document.querySelector('[role="menu"][aria-label="Surface menu"]')).not.toBeNull();
  key(document.querySelector('[role="menu"]')!, { key: 'Escape' });
  expect(document.querySelector('[role="menu"]')).toBeNull();
  const shiftF10 = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'F10', shiftKey: true });
  act(() => { surface.dispatchEvent(shiftF10); });
  expect(shiftF10.defaultPrevented).toBe(true);
  expect(document.activeElement?.textContent).toBe('Cut');
  key(document.activeElement!, { key: 'Escape' });
  key(surface, { key: 'ContextMenu' });
  expect(requests).toEqual(['pointer@30,40', 'keyboard@7,9', 'keyboard@7,9']);
});

test('a surface without a menu for the request keeps the browser menu', () => {
  const host = mount(<Surface resolve={() => null} />);
  const surface = host.querySelector('[data-testid="surface"]')!;
  const click = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, button: 2 });
  act(() => { surface.dispatchEvent(click); });
  expect(click.defaultPrevented).toBe(false);
  expect(document.querySelector('[role="menu"]')).toBeNull();
});

test('menu helpers', () => {
  expect(isContextMenuKey({ key: 'F10', shiftKey: true, ctrlKey: false, metaKey: false, altKey: false })).toBe(true);
  expect(isContextMenuKey({ key: 'F10', shiftKey: true, ctrlKey: true, metaKey: false, altKey: false })).toBe(false);
  expect(menuSections([[], [item('a')], [], [item('b')]]).map((entry) => entry.kind)).toEqual(['item', 'separator', 'item']);
  const entry = commandMenuItem({ id: 'paste', label: 'Paste', shortcut: 'Ctrl+V',
    state: { enabled: false, disabledReason: { message: 'Read only' } }, run: () => {} });
  expect(entry).toMatchObject({ kind: 'item', disabled: true, description: 'Read only', shortcut: 'Ctrl+V' });
});

test('returnFocus replaces the opener as the focus target on close', () => {
  focusOrigin();
  const editor = document.createElement('textarea');
  document.body.append(editor);
  mount(<ContextMenu label="Cell" modality="keyboard" point={{ x: 0, y: 0 }} entries={[item('Cut')]}
    returnFocus={() => editor.focus()} onClose={() => {}} />);
  key(document.activeElement!, { key: 'Escape' });
  expect(document.activeElement).toBe(editor);
});
