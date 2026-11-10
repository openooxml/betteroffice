import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import { createSelectionSettle } from './selectionSettle';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

function clocked(run: (h: ReturnType<typeof harness>) => void) {
  const timeout = globalThis.setTimeout;
  const cancel = globalThis.clearTimeout;
  let now = 0;
  let serial = 0;
  const timers = new Map<number, { at: number; run(): void }>();
  globalThis.setTimeout = ((run: () => void, delay = 0) => {
    const id = ++serial;
    timers.set(id, { at: now + delay, run });
    return id;
  }) as unknown as typeof setTimeout;
  globalThis.clearTimeout = ((id: number) => { timers.delete(id); }) as unknown as typeof clearTimeout;
  const h = harness((ms) => {
    now += ms;
    for (const [id, timer] of [...timers]) if (timer.at <= now) { timers.delete(id); timer.run(); }
  }, () => timers.size);
  try { run(h); }
  finally {
    h.machine.dispose();
    h.root.remove();
    globalThis.setTimeout = timeout;
    globalThis.clearTimeout = cancel;
  }
}

function harness(tick: (ms: number) => void, pending: () => number) {
  const root = document.createElement('textarea');
  document.body.append(root);
  const sources: string[] = [];
  let unsettled = 0;
  let resettled = 0;
  let ready = true;
  const machine = createSelectionSettle({ root,
    canSettle: () => ready,
    onUnsettle: () => { unsettled++; },
    onSettle: (source) => sources.push(source),
    onResettle: () => { resettled++; },
  });
  const down = (pointerType = 'mouse') => root.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, pointerType }));
  const up = () => document.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
  return { root, machine, tick, pending, sources, down, up, ready: (value: boolean) => { ready = value; }, unsettled: () => unsettled, resettled: () => resettled };
}

test('drag settles exactly once 350 ms after release', () => clocked((h) => {
  h.down();
  h.machine.notifySelectionChange('1:2');
  h.tick(500);
  h.machine.notifySelectionChange('1:9');
  expect(h.sources).toEqual([]);
  h.up();
  h.tick(349);
  expect(h.sources).toEqual([]);
  h.tick(1);
  h.tick(500);
  expect(h.sources).toEqual(['pointer']);
  expect(h.unsettled()).toBeGreaterThan(0);
}));

for (const count of [2, 3]) test(`${count} clicks settle only after the last click`, () => clocked((h) => {
  for (let click = 0; click < count; click++) {
    h.down();
    h.machine.notifySelectionChange(`word:${click}`);
    h.up();
    h.tick(100);
    expect(h.sources).toEqual([]);
  }
  h.tick(249);
  expect(h.sources).toEqual([]);
  h.tick(1);
  expect(h.sources).toEqual(['pointer']);
}));

test('three shift arrows settle 350 ms after the last selection change', () => clocked((h) => {
  for (let index = 1; index <= 3; index++) {
    h.root.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowRight', shiftKey: true }));
    h.machine.notifySelectionChange(`0:${index}`);
    h.tick(100);
  }
  h.tick(249);
  expect(h.sources).toEqual([]);
  h.tick(1);
  expect(h.sources).toEqual(['keyboard']);
}));

test('touch release retains its source through a delayed selection push', () => clocked((h) => {
  h.down('touch');
  h.up();
  h.machine.notifySelectionChange('touch');
  h.tick(350);
  expect(h.sources).toEqual(['touch']);
}));

test('a selection bump while waiting keeps the last pointer action as its source', () => clocked((h) => {
  h.ready(false);
  h.down();
  h.machine.notifySelectionChange('pointer');
  h.up();
  h.tick(350);
  h.machine.notifySelectionChange('published-pointer');
  h.ready(true);
  h.tick(349);
  expect(h.sources).toEqual([]);
  h.tick(1);
  expect(h.sources).toEqual(['pointer']);
}));

test('the last local action in the window supplies the settle source', () => clocked((h) => {
  h.down();
  h.machine.notifySelectionChange('pointer');
  h.up();
  h.tick(200);
  h.root.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowRight' }));
  h.machine.notifySelectionChange('keyboard');
  h.tick(350);
  expect(h.sources).toEqual(['keyboard']);
  h.root.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowRight' }));
  h.machine.notifySelectionChange('keyboard-again');
  h.tick(200);
  h.down('touch');
  h.machine.notifySelectionChange('touch');
  h.up();
  h.tick(350);
  expect(h.sources).toEqual(['keyboard', 'touch']);
  h.machine.notifySelectionChange('programmatic');
  h.tick(350);
  expect(h.sources).toEqual(['keyboard', 'touch', 'keyboard']);
}));

test('settling the same selection again restores the state without another event', () => clocked((h) => {
  h.machine.notifySelectionChange('same');
  h.tick(350);
  h.down();
  h.up();
  h.tick(350);
  expect(h.sources).toEqual(['keyboard']);
  expect(h.resettled()).toBe(1);
}));

test('T23: a typing key restarts the quiet window', () => clocked((h) => {
  h.machine.notifySelectionChange('before');
  h.tick(300);
  h.root.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'x' }));
  h.machine.notifySelectionChange('after');
  h.tick(50);
  expect(h.sources).toEqual([]);
  h.tick(299);
  expect(h.sources).toEqual([]);
  h.tick(1);
  expect(h.sources).toEqual(['keyboard']);
}));

test('D1: programmatic selection changes have the keyboard source and no API flag', () => clocked((h) => {
  const args: Parameters<typeof h.machine.notifySelectionChange> = ['programmatic'];
  type HasApiFlag = Parameters<typeof h.machine.notifySelectionChange> extends [string] ? false : true;
  const hasApiFlag: HasApiFlag = false;
  expect(hasApiFlag).toBe(false);
  expect(h.machine.notifySelectionChange.length).toBe(1);
  Reflect.apply(h.machine.notifySelectionChange, h.machine, [...args, true]);
  h.tick(350);
  expect(h.sources).toEqual(['keyboard']);
}));

test('T6: focus change neither cancels nor restarts a pending settle', () => clocked((h) => {
  h.machine.notifySelectionChange('focus');
  h.tick(200);
  h.root.dispatchEvent(new FocusEvent('focusout', { bubbles: true }));
  h.root.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
  h.tick(149);
  expect(h.sources).toEqual([]);
  h.tick(1);
  expect(h.sources).toEqual(['keyboard']);
}));

test('T23: continuous typing longer than the window settles once after the last key', () => clocked((h) => {
  for (let index = 0; index < 8; index++) {
    h.root.dispatchEvent(new KeyboardEvent('keydown', { key: 'x', bubbles: true }));
    h.machine.notifySelectionChange(`typed:${index}`);
    h.tick(100);
    expect(h.sources).toEqual([]);
  }
  h.tick(249);
  expect(h.sources).toEqual([]);
  h.tick(1);
  expect(h.sources).toEqual(['keyboard']);
}));

for (const end of ['outside-window', 'blur', 'lostpointercapture', 'touch-scroll'] as const) {
  test(`T24: Pressed exits after ${end}`, () => clocked((h) => {
    h.down(end === 'touch-scroll' ? 'touch' : 'mouse');
    h.machine.notifySelectionChange(end);
    h.tick(700);
    expect(h.sources).toEqual([]);
    if (end === 'outside-window') window.dispatchEvent(new PointerEvent('pointerup'));
    if (end === 'blur') window.dispatchEvent(new Event('blur'));
    if (end === 'lostpointercapture') h.root.dispatchEvent(new PointerEvent('lostpointercapture', { bubbles: true }));
    if (end === 'touch-scroll') document.dispatchEvent(new PointerEvent('pointercancel', { bubbles: true, pointerType: 'touch' }));
    h.tick(349);
    expect(h.sources).toEqual([]);
    h.tick(1);
    expect(h.sources).toEqual([end === 'touch-scroll' ? 'touch' : 'pointer']);
    h.tick(700);
    expect(h.sources).toHaveLength(1);
  }));
}

test('pointer cancellation releases the gesture and dispose clears its timer', () => clocked((h) => {
  h.down('touch');
  h.machine.notifySelectionChange('cancelled');
  document.dispatchEvent(new PointerEvent('pointercancel', { bubbles: true }));
  h.tick(350);
  expect(h.sources).toEqual(['touch']);
  h.machine.notifySelectionChange('disposed');
  expect(h.pending()).toBe(1);
  h.machine.dispose();
  expect(h.pending()).toBe(0);
  h.tick(350);
  h.down();
  h.up();
  h.machine.notifySelectionChange('ignored');
  h.tick(350);
  expect(h.sources).toEqual(['touch']);
}));


test('a slow selection mapping settles once when ready and keeps its pointer source', () => clocked((h) => {
  h.ready(false);
  h.down();
  h.machine.notifySelectionChange('acknowledged-selection');
  h.up();
  h.tick(700);
  expect(h.sources).toEqual([]);
  h.machine.notifySelectionChange('acknowledged-selection');
  expect(h.sources).toEqual([]);
  h.ready(true);
  h.machine.notifySelectionChange('acknowledged-selection');
  expect(h.sources).toEqual(['pointer']);
  h.machine.notifySelectionChange('acknowledged-selection');
  h.tick(700);
  expect(h.sources).toEqual(['pointer']);
}));
