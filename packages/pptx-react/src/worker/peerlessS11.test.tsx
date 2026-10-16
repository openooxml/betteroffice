import { expect, mock, test } from 'bun:test';
import type { AsyncCollaborationReplica, PptxPresencePeer, PptxPresenceCursor } from '@betteroffice/pptx';
import { act, waitFor } from '@testing-library/react';
import { open, story, type, replyControl } from './peerlessTestHarness';

test('two_client_sync', async () => {
  let leftReplica!: AsyncCollaborationReplica;
  let rightReplica!: AsyncCollaborationReplica;
  const left = await open({ collaboration: { clientId: 1001, onReplica: (value) => { if (value) leftReplica = value; } } });
  const right = await open({ collaboration: { clientId: 1002, onReplica: (value) => { if (value) rightReplica = value; } } });
  await waitFor(() => { expect(leftReplica).toBeTruthy(); expect(rightReplica).toBeTruthy(); });
  const leftVector = await leftReplica.encodeStateVector();
  const rightVector = await rightReplica.encodeStateVector();
  expect(leftReplica.encodeStateVector()).toBeInstanceOf(Promise);
  const leftStory = await story(left.api);
  const rightStory = await story(right.api);
  await left.api.selectText({ slide: 1, shapeId: leftStory.shapeId, storyId: leftStory.storyId, start: 2, end: 2 });
  await right.api.selectText({ slide: 1, shapeId: rightStory.shapeId, storyId: rightStory.storyId, start: 2, end: 2 });
  const updates: string[] = [];
  const off = leftReplica.onUpdate((_bytes: Uint8Array, origin: string) => { updates.push(origin); });
  await act(async () => {
    type(left.view.getByRole('textbox', { name: 'Slide text' }), 'L');
    type(right.view.getByRole('textbox', { name: 'Slide text' }), 'R');
    await Promise.all([left.api.flushPendingInput(), right.api.flushPendingInput()]);
    const leftUpdate = await leftReplica.encodeStateAsUpdate(rightVector);
    const rightUpdate = await rightReplica.encodeStateAsUpdate(leftVector);
    await Promise.all([leftReplica.applyUpdate(rightUpdate), rightReplica.applyUpdate(leftUpdate)]);
  });
  expect((await story(left.api)).text).toBe((await story(right.api)).text);
  expect((await story(left.api)).text).toContain('L');
  expect((await story(left.api)).text).toContain('R');
  expect(updates.filter((value) => value === 'local')).toHaveLength(1);
  expect(updates.filter((value) => value === 'remote')).toHaveLength(1);
  off();
  await act(async () => { await left.api.commands!.execute('undo', null); });
  expect((await story(left.api)).text).not.toContain('L');
  expect((await story(left.api)).text).toContain('R');
});

test('remote_undo_and_readonly_remote', async () => {
  let leftReplica!: AsyncCollaborationReplica;
  let rightReplica!: AsyncCollaborationReplica;
  const left = await open({ collaboration: { clientId: 1003, onReplica: (value) => { if (value) leftReplica = value; } } });
  const right = await open({ readOnly: true, collaboration: { clientId: 1004, onReplica: (value) => { if (value) rightReplica = value; } } });
  await waitFor(() => { expect(leftReplica).toBeTruthy(); expect(rightReplica).toBeTruthy(); });
  const selected = await story(left.api);
  await left.api.selectText({ slide: 1, shapeId: selected.shapeId, storyId: selected.storyId, start: 2, end: 2 });
  const vector = await rightReplica.encodeStateVector();
  await act(async () => {
    type(left.view.getByRole('textbox', { name: 'Slide text' }), 'LOCAL');
    await left.api.flushPendingInput();
    await rightReplica.applyUpdate(await leftReplica.encodeStateAsUpdate(vector));
  });
  expect((await story(right.api)).text).toContain('LOCAL');
  const readonlyVector = await leftReplica.encodeStateVector();
  const update = await rightReplica.encodeStateAsUpdate(readonlyVector);
  await leftReplica.applyUpdate(update);
  await act(async () => { await left.api.commands!.execute('undo', null); });
  expect((await story(left.api)).text).not.toContain('LOCAL');
  expect(left.view.queryByRole('alert')).toBeNull();
});

test('presence_retirement', async () => {
  const replicas: unknown[] = [];
  const cursor = mock((_value: PptxPresenceCursor | null) => {});
  const listeners = new Set<(value: readonly PptxPresencePeer[]) => void>();
  const presence = { peers: [], setCursor: cursor, onPresence: (listener: (value: readonly PptxPresencePeer[]) => void) => {
    listeners.add(listener); return () => { listeners.delete(listener); };
  } };
  const value = await open({ collaboration: { clientId: 1005, presence,
    onReplica: (replica) => { replicas.push(replica); } } });
  await waitFor(() => expect(replicas.some(Boolean)).toBe(true));
  const peer = { state: { clientId: 1006, clock: 1, user: { name: 'Remote', color: '#123456' },
    cursor: { slideId: value.client.state.activeSlideId!, shapeId: value.box.shapeId! } }, lastSeen: 1, cursorMovedAt: 1 };
  act(() => { for (const listener of listeners) listener([peer]); });
  expect(value.view.getByLabelText('Collaborators').textContent).toContain('Remote');
  expect(value.view.getByTestId('pptx-remote-shape')).toBeTruthy();
  value.view.unmount();
  expect(replicas[replicas.length - 1]).toBeNull();
  expect(cursor.mock.calls[cursor.mock.calls.length - 1][0]).toBeNull();
  expect(listeners.size).toBe(0);
  await expect((replicas.find(Boolean) as AsyncCollaborationReplica).encodeStateVector()).rejects.toThrow();
});

test('pending_remote_rebases_selection_without_local_echo_loop', async () => {
  let leftReplica!: AsyncCollaborationReplica;
  let rightReplica!: AsyncCollaborationReplica;
  const left = await open({ collaboration: { clientId: 1101, onReplica: (value) => { if (value) leftReplica = value; } } });
  const right = await open({ collaboration: { clientId: 1102, onReplica: (value) => { if (value) rightReplica = value; } } });
  const leftStory = await story(left.api);
  const rightStory = await story(right.api);
  await left.api.selectText({ slide: 1, shapeId: leftStory.shapeId, storyId: leftStory.storyId, start: 2, end: 2 });
  await right.api.selectText({ slide: 1, shapeId: rightStory.shapeId, storyId: rightStory.storyId, start: 0, end: 0 });
  const vector = await leftReplica.encodeStateVector();
  await act(async () => { type(right.view.getByRole('textbox', { name: 'Slide text' }), 'REMOTE'); await right.api.flushPendingInput(); });
  const update = await rightReplica.encodeStateAsUpdate(vector);
  replyControl.hold();
  let remote!: Promise<unknown>;
  await act(async () => {
    type(left.view.getByRole('textbox', { name: 'Slide text' }), 'L');
    remote = leftReplica.applyUpdate(update);
    type(left.view.getByRole('textbox', { name: 'Slide text' }), 'NEXT');
  });
  await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
  await act(async () => { replyControl.release(); await remote; await left.api.flushPendingInput(); });
  expect((await story(left.api)).text).toBe('REMOTEalLNEXTpha beta\nx\nalpha beta');
  expect(left.client.state.selection).toMatchObject({ anchor: 13, focus: 13 });
});
