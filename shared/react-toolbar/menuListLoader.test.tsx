import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, test } from 'bun:test';
import { useLayoutEffect } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { ContextMenu } from './ContextMenu';
import { OverflowMenu } from './OverflowMenu';
import { loadedMenuList, loadMenuList, menuListLoader, preloadMenuList, resetMenuList } from './menuListLoader';

const { act, cleanup, fireEvent, render, screen } = await import('@testing-library/react');

const original = menuListLoader.load;
const entries = [{ kind: 'item' as const, id: 'one', label: 'One', onSelect() {} }];
const settle = () => act(async () => { for (let index = 0; index < 20; index += 1) await new Promise((resolve) => setTimeout(resolve, 0)); });

beforeEach(resetMenuList);
afterEach(() => { cleanup(); menuListLoader.load = original; });
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

test('an idle preload makes the next overflow menu open in the same turn', async () => {
  render(<OverflowMenu label="More" trigger="More" entries={entries} testId="more" />);
  await settle();
  expect(loadedMenuList()).toBeDefined();
  fireEvent.click(screen.getByTestId('more'));
  expect(screen.getByRole('menu', { name: 'More' })).toBeDefined();
});

test('an overflow menu opened before the list arrives shows once it has loaded', async () => {
  let finish!: () => void;
  const held = new Promise<void>((resolve) => { finish = resolve; });
  menuListLoader.load = () => held.then(original);
  render(<OverflowMenu label="More" trigger="More" entries={entries} testId="more" />);
  fireEvent.click(screen.getByTestId('more'));
  expect(screen.queryByRole('menu', { name: 'More' })).toBeNull();
  finish();
  await settle();
  expect(screen.getByRole('menu', { name: 'More' })).toBeDefined();
});

test('a failed load keeps the overflow menu closed, refocuses the trigger and retries on the next open', async () => {
  let fail = true;
  menuListLoader.load = () => fail ? Promise.reject(new Error('offline')) : original();
  render(<OverflowMenu label="More" trigger="More" entries={entries} testId="more" />);
  await settle();
  const trigger = screen.getByTestId('more');
  fireEvent.click(trigger);
  await settle();
  expect(screen.queryByRole('menu', { name: 'More' })).toBeNull();
  expect(trigger.getAttribute('aria-expanded')).toBe('false');
  expect(document.activeElement).toBe(trigger);
  fail = false;
  fireEvent.click(trigger);
  await settle();
  expect(screen.getByRole('menu', { name: 'More' })).toBeDefined();
});

test('a failed load closes a context menu and returns focus, and the next one retries', async () => {
  let fail = true;
  menuListLoader.load = () => fail ? Promise.reject(new Error('offline')) : original();
  const reasons: string[] = [];
  let focused = 0;
  const menu = (key: number) => <ContextMenu key={key} label="Actions" entries={entries} point={{ x: 1, y: 1 }}
    returnFocus={() => { focused += 1; }} onClose={(reason) => { reasons.push(reason); }} />;
  const view = render(menu(1));
  await settle();
  expect(screen.queryByRole('menu', { name: 'Actions' })).toBeNull();
  expect(reasons).toEqual(['escape']);
  expect(focused).toBe(1);
  fail = false;
  view.rerender(menu(2));
  await settle();
  expect(screen.getByRole('menu', { name: 'Actions' })).toBeDefined();
});

test.each(['Escape', 'Tab'])('%s closes an overflow menu still waiting for its list, which then never appears', async (key) => {
  let finish!: () => void;
  const held = new Promise<void>((resolve) => { finish = resolve; });
  menuListLoader.load = () => held.then(original);
  render(<OverflowMenu label="More" trigger="More" entries={entries} testId="more" />);
  const trigger = screen.getByTestId('more');
  fireEvent.keyDown(trigger, { key: 'Enter' });
  await settle();
  fireEvent.keyDown(trigger, { key });
  await settle();
  expect(trigger.getAttribute('aria-expanded')).toBe('false');
  finish();
  await settle();
  expect(screen.queryByRole('menu', { name: 'More' })).toBeNull();
});

test('Tab closes a context menu still waiting for its list', async () => {
  let finish!: () => void;
  const held = new Promise<void>((resolve) => { finish = resolve; });
  menuListLoader.load = () => held.then(original);
  const reasons: string[] = [];
  render(<ContextMenu label="Actions" entries={entries} point={{ x: 1, y: 1 }} onClose={(reason) => { reasons.push(reason); }} />);
  fireEvent.keyDown(document.body, { key: 'Tab' });
  expect(reasons).toEqual(['tab']);
  finish();
  await settle();
});

test('dismissal keys reach a pending menu even when the focused editor stops their propagation', async () => {
  let finish!: () => void;
  const held = new Promise<void>((resolve) => { finish = resolve; });
  menuListLoader.load = () => held.then(original);
  const reasons: string[] = [];
  render(<div>
    <input aria-label="Cell" onKeyDown={(event) => event.stopPropagation()} />
    <OverflowMenu label="More" trigger="More" entries={entries} testId="more" />
    <ContextMenu label="Actions" entries={entries} point={{ x: 1, y: 1 }} onClose={(reason) => { reasons.push(reason); }} />
  </div>);
  const trigger = screen.getByTestId('more');
  fireEvent.click(trigger);
  await settle();
  expect(trigger.getAttribute('aria-expanded')).toBe('true');
  fireEvent.keyDown(screen.getByLabelText('Cell'), { key: 'Escape' });
  await settle();
  expect(trigger.getAttribute('aria-expanded')).toBe('false');
  expect(reasons).toEqual(['escape']);
  finish();
  await settle();
  expect(screen.queryByRole('menu')).toBeNull();
});

test('a list that loads after a menu renders but before its effect still shows the menu', async () => {
  const component = await original();
  menuListLoader.load = () => ({
    then: (loaded: (value: typeof component) => unknown) => Promise.resolve(loaded(component)),
  }) as unknown as Promise<typeof component>;
  function LoadDuringCommit() {
    useLayoutEffect(() => { void loadMenuList(); }, []);
    return null;
  }
  render(<div>
    <ContextMenu label="Actions" entries={entries} point={{ x: 1, y: 1 }} onClose={() => {}} />
    <LoadDuringCommit />
  </div>);
  await settle();
  expect(screen.getByRole('menu', { name: 'Actions' })).toBeDefined();
});

test('while a menu waits for its list, other keys never reach the page, and they do again once it has loaded', async () => {
  for (const kind of ['overflow', 'context'] as const) {
    resetMenuList();
    let finish!: () => void;
    const held = new Promise<void>((resolve) => { finish = resolve; });
    menuListLoader.load = () => held.then(original);
    const reached: string[] = [];
    const view = render(<div>
      <input aria-label="Stage" onKeyDown={(event) => { reached.push(event.key); }} />
      {kind === 'overflow' ? <OverflowMenu label="More" trigger="More" entries={entries} testId="more" />
        : <ContextMenu label="Actions" entries={entries} point={{ x: 1, y: 1 }} onClose={() => {}} />}
    </div>);
    if (kind === 'overflow') fireEvent.click(screen.getByTestId('more'));
    await settle();
    const stage = screen.getByLabelText('Stage');
    const blocked = fireEvent.keyDown(stage, { key: 'Delete', cancelable: true });
    expect([kind, blocked, reached]).toEqual([kind, false, []]);
    finish();
    await settle();
    expect(screen.getByRole('menu')).toBeDefined();
    fireEvent.keyDown(stage, { key: 'Delete' });
    expect([kind, reached]).toEqual([kind, ['Delete']]);
    view.unmount();
  }
});

test('a cancelled idle preload loads nothing', async () => {
  let loads = 0;
  menuListLoader.load = () => { loads += 1; return original(); };
  preloadMenuList()();
  await settle();
  expect(loads).toBe(0);
});
