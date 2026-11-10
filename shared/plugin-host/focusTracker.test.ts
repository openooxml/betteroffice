import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, spyOn, test } from 'bun:test';
import { trackFocus } from './focusTracker';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const disposers: (() => void)[] = [];
afterEach(() => {
  for (const dispose of disposers.splice(0)) dispose();
  document.body.replaceChildren();
});
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});
const settle = () => new Promise((done) => setTimeout(done, 0));

function setup() {
  const editor = document.createElement('div');
  const floating = document.createElement('div');
  const input = document.createElement('input');
  const second = document.createElement('input');
  const outside = document.createElement('button');
  editor.append(input);
  floating.append(second);
  document.body.append(editor, floating, outside);
  const roots: Element[] = [editor];
  const changes: boolean[] = [];
  const dispose = trackFocus(() => roots, (focused) => changes.push(focused));
  disposers.push(dispose);
  changes.length = 0;
  return { input, second, outside, roots, floating, changes, dispose };
}

test('moving between inside elements and dynamically added roots never blurs', async () => {
  const { input, second, roots, floating, changes } = setup();
  roots.push(floating);
  input.focus();
  input.dispatchEvent(new FocusEvent('focusout', { bubbles: true, relatedTarget: null }));
  second.focus();
  await settle();
  expect(changes).toEqual([true]);
});

test('moving from the document to a panel blurs', async () => {
  const { input, outside, changes } = setup();
  input.focus();
  outside.focus();
  await settle();
  expect(changes).toEqual([true, false]);
});

test('window blur clears focus even when activeElement remains inside', async () => {
  const { input, changes } = setup();
  input.focus();
  input.dispatchEvent(new FocusEvent('focusout', { bubbles: true, relatedTarget: null }));
  window.dispatchEvent(new Event('blur'));
  await settle();
  expect(changes).toEqual([true, false]);
});

test('removing a focused element clears focus', async () => {
  const { input, changes } = setup();
  input.focus();
  input.remove();
  await settle();
  expect(changes).toEqual([true, false]);
});

test('disposing cancels listeners and a queued focus check', async () => {
  const { input, outside, changes, dispose } = setup();
  input.focus();
  input.dispatchEvent(new FocusEvent('focusout', { bubbles: true, relatedTarget: null }));
  dispose();
  outside.focus();
  await settle();
  window.dispatchEvent(new Event('blur'));
  input.focus();
  expect(changes).toEqual([true]);
});

test('observes only focus roots while focused and disconnects on blur', async () => {
  const observe = spyOn(MutationObserver.prototype, 'observe');
  const disconnect = spyOn(MutationObserver.prototype, 'disconnect');
  try {
    const { input, outside, roots } = setup();
    expect(observe).not.toHaveBeenCalled();
    input.focus();
    expect(observe.mock.calls).toEqual([[roots[0], { childList: true, subtree: true }]]);
    const connected = disconnect.mock.calls.length;
    outside.focus();
    await settle();
    expect(disconnect.mock.calls.length).toBeGreaterThan(connected);
    expect(observe.mock.calls).toHaveLength(1);
    input.focus();
    expect(observe.mock.calls).toHaveLength(2);
    const focused = disconnect.mock.calls.length;
    window.dispatchEvent(new Event('blur'));
    expect(disconnect.mock.calls.length).toBeGreaterThan(focused);
  } finally {
    observe.mockRestore();
    disconnect.mockRestore();
  }
});
