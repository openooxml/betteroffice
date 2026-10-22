import { faultControl, open, replyControl, story, type } from './peerlessTestHarness';
import { afterEach, beforeEach, expect, mock, test } from 'bun:test';
import type { DeckSnapshot, PresentationCommandAck } from '@betteroffice/pptx';

const { act, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
let idleCallbacks: Map<number, IdleRequestCallback>;

beforeEach(() => {
  idleCallbacks = new Map();
  let nextId = 0;
  for (const [name, value] of [
    ['requestIdleCallback', (callback: IdleRequestCallback) => {
      const id = ++nextId;
      idleCallbacks.set(id, callback);
      return id;
    }],
    ['cancelIdleCallback', (id: number) => { idleCallbacks.delete(id); }],
  ] as const) {
    const original = Object.getOwnPropertyDescriptor(window, name);
    Object.defineProperty(window, name, { configurable: true, value });
    restorers.push(() => {
      if (original) Object.defineProperty(window, name, original);
      else Reflect.deleteProperty(window, name);
    });
  }
});

afterEach(() => { for (const restore of restorers.splice(0).reverse()) restore(); });

function runIdle(): void {
  const callbacks = [...idleCallbacks.values()];
  idleCallbacks.clear();
  for (const callback of callbacks) callback({ didTimeout: false, timeRemaining: () => 50 });
}

async function commit() {
  const changed = mock((_snapshot: DeckSnapshot) => {});
  const error = mock(() => {});
  const value = await open({ onChange: changed, onError: error });
  const target = await story(value.api);
  await act(async () => {
    await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 });
  });
  let committed: PresentationCommandAck | undefined;
  const off = value.client.subscribe((ack) => {
    if (ack.outcomes.some((outcome) => outcome.committed)) committed = ack;
  });
  try {
    await act(async () => {
      type(value.view.getByLabelText('Slide text'), 'Accepted');
      await value.api.flushPendingInput();
    });
  } finally { off(); }
  expect(committed).toBeDefined();
  expect(changed).not.toHaveBeenCalled();
  return { ...value, changed, error, committed: committed! };
}

test('failure delivers the pending committed change before idle runs exactly once', async () => {
  const value = await commit();
  expect(value.committed.change).toBeDefined();
  act(() => { faultControl.crash(); });
  expect(value.changed).toHaveBeenCalledTimes(1);
  expect(value.changed.mock.calls[0]![0]).toBe(value.committed.change!.snapshot);
  expect(value.error).toHaveBeenCalledTimes(1);
  await act(async () => { runIdle(); faultControl.crash(); });
  expect(value.changed).toHaveBeenCalledTimes(1);
});

test('failure while the idle drain waits on the fence delivers the committed change exactly once', async () => {
  const value = await commit();
  expect(value.committed.change).toBeDefined();
  await value.api.whenHydrated();
  replyControl.hold();
  act(() => { type(value.view.getByLabelText('Slide text'), 'Held'); });
  await waitFor(() => expect(replyControl.count()).toBe(1));
  await act(async () => { runIdle(); });
  expect(value.changed).not.toHaveBeenCalled();
  act(() => { faultControl.crash(); });
  expect(value.changed).toHaveBeenCalledTimes(1);
  expect(value.changed.mock.calls[0]![0]).toBe(value.committed.change!.snapshot);
  expect(value.error).toHaveBeenCalledTimes(1);
  await act(async () => { replyControl.release(); runIdle(); });
  expect(value.changed).toHaveBeenCalledTimes(1);
});

test('idle delivers the committed change exactly once without failure', async () => {
  const value = await commit();
  expect(value.committed.change).toBeDefined();
  await value.api.whenHydrated();
  await waitFor(async () => {
    await act(async () => { runIdle(); });
    expect(value.changed).toHaveBeenCalledTimes(1);
  });
  expect(value.changed.mock.calls[0]![0]).toBe(value.committed.change!.snapshot);
  await act(async () => { runIdle(); });
  expect(value.changed).toHaveBeenCalledTimes(1);
  expect(value.error).not.toHaveBeenCalled();
});

