import { expect, mock, test } from 'bun:test';
import { open, story, type, replyControl } from './peerlessTestHarness';
import { type AsyncCollaborationReplica } from '@betteroffice/pptx';
import { PptxWorkerUnsupportedError } from '@betteroffice/pptx/worker';

const { act, fireEvent, waitFor } = await import('@testing-library/react');

function composition(input: HTMLElement, kind: string, text = ''): void {
  const event = new Event(kind, { bubbles: true });
  Object.defineProperty(event, 'data', { value: text });
  input.dispatchEvent(event);
}

async function collaborators() {
  let replica!: AsyncCollaborationReplica;
  const local = await open({ collaboration: { clientId: 2001, onReplica: (value) => { if (value) replica = value; } } });
  const remote = await open({ collaboration: { clientId: 2002 } });
  const localStory = await story(local.api);
  const remoteStory = await story(remote.api);
  await local.api.selectText({ slide: 1, shapeId: localStory.shapeId, storyId: localStory.storyId, start: 2, end: 2 });
  await remote.api.selectText({ slide: 1, shapeId: remoteStory.shapeId, storyId: remoteStory.storyId, start: 2, end: 2 });
  const vector = await replica.encodeStateVector();
  await act(async () => {
    type(remote.view.getByRole('textbox', { name: 'Slide text' }), 'R');
    await remote.api.flushPendingInput();
  });
  const update = await (await remote.api.handleAsync()).encodeStateAsUpdate(vector);
  return { local, replica, update };
}

test('remote merge preserves IME text and permits the next composition', async () => {
  const { local, replica, update } = await collaborators();
  const input = local.view.getByRole('textbox', { name: 'Slide text' });
  await act(async () => {
    composition(input, 'compositionstart');
    composition(input, 'compositionupdate', '漢');
    await local.api.flushPendingInput();
    await replica.applyUpdate(update);
    composition(input, 'compositionend', '漢');
    await local.api.flushPendingInput();
  });
  expect((await story(local.api)).text).toContain('漢');
  await act(async () => {
    composition(input, 'compositionstart');
    composition(input, 'compositionupdate', '字');
    composition(input, 'compositionend', '字');
    await local.api.flushPendingInput();
  });
  expect((await story(local.api)).text).toContain('漢字');
});

test('local typing chains after an unacknowledged remote merge', async () => {
  const { local, replica, update } = await collaborators();
  replyControl.hold();
  const merging = replica.applyUpdate(update);
  try {
    await waitFor(() => expect(replyControl.count()).toBe(1));
    act(() => { type(local.view.getByRole('textbox', { name: 'Slide text' }), 'NEXT'); });
  } finally { replyControl.release(); }
  await act(async () => { await merging; await local.api.flushPendingInput(); });
  const text = (await story(local.api)).text;
  expect(text).toContain('R');
  expect(text).toContain('NEXT');
});

test('copy writes selected text and cut deletes it through a worker command', async () => {
  const { api, view, client } = await open();
  const selected = await story(api);
  await api.selectText({ slide: 1, shapeId: selected.shapeId, storyId: selected.storyId, start: 0, end: 5 });
  const original = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
  const writeText = mock(async (_text: string) => {});
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText } });
  try {
    const input = view.getByRole('textbox', { name: 'Slide text' });
    act(() => { fireEvent.keyDown(input, { key: 'c', ctrlKey: true }); });
    await waitFor(() => expect(writeText).toHaveBeenCalledWith('alpha'));
    expect((await story(api)).text).toBe(selected.text);
    const throughSeq = client.state.throughSeq;
    act(() => { fireEvent.keyDown(input, { key: 'x', metaKey: true }); });
    await waitFor(() => expect(writeText).toHaveBeenCalledTimes(2));
    await waitFor(() => expect(client.state.throughSeq).toBeGreaterThan(throughSeq));
    expect((await story(api)).text).toBe(selected.text.slice(5));
    await act(async () => { await api.commands.execute('undo', null); });
    expect((await story(api)).text).toBe(selected.text);
  } finally {
    if (original) Object.defineProperty(navigator, 'clipboard', original);
    else Reflect.deleteProperty(navigator, 'clipboard');
  }
});

test('legacy sync members throw typed refusals naming their async replacements', async () => {
  const { api } = await open();
  const legacy = api as unknown as { handle: unknown };
  for (const [operation, alternative, invoke] of [
    ['handle', 'handleAsync()', () => legacy.handle],
  ] as const) {
    expect(invoke).toThrow(PptxWorkerUnsupportedError);
    try { invoke(); } catch (error) {
      expect(error).toMatchObject({ code: 'sync-api-unavailable', operation });
      expect((error as Error).message).toContain(alternative);
    }
  }
});

test('cut waits for an unresolved selection before reading and deleting text', async () => {
  const { api, view, client } = await open();
  const selected = await story(api);
  const original = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
  const writeText = mock(async (_text: string) => {});
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText } });
  replyControl.hold();
  const selecting = api.selectText({ slide: 1, shapeId: selected.shapeId, storyId: selected.storyId, start: 0, end: 5 });
  try {
    await waitFor(() => expect(replyControl.count()).toBe(1));
    act(() => { fireEvent.keyDown(view.getByRole('textbox', { name: 'Slide text' }), { key: 'x', ctrlKey: true }); });
    expect(writeText).not.toHaveBeenCalled();
    const throughSeq = client.state.throughSeq;
    replyControl.release();
    await act(async () => { await selecting; });
    await waitFor(() => expect(writeText).toHaveBeenCalledWith('alpha'));
    await waitFor(() => expect(client.state.throughSeq).toBeGreaterThan(throughSeq + 1));
    expect((await story(api)).text).toBe(selected.text.slice(5));
  } finally {
    replyControl.release();
    if (original) Object.defineProperty(navigator, 'clipboard', original);
    else Reflect.deleteProperty(navigator, 'clipboard');
  }
});
