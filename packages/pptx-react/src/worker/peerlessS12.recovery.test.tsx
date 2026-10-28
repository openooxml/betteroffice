import { faultControl, open, story, type } from './peerlessTestHarness';
import { expect, mock, spyOn, test } from 'bun:test';
import type { DeckSnapshot } from '@betteroffice/pptx';
import type { PresentationCommandSaveResult } from '@betteroffice/pptx/worker';
import * as wasm from '../../../pptx/src/wasm/loader';
import { peerlessRecoveryFactory } from './peerlessRecoveryJournal';
const { act, waitFor } = await import('@testing-library/react');

test('recovery retains typing after an edited automatic checkpoint', async () => {
  const error = mock(() => {});
  const value = await open({ onError: error });
  const target = await story(value.api);
  await act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 }); });
  const checkpoint = spyOn(value.client, 'checkpoint');
  try {
    const input = value.view.getByLabelText('Slide text');
    await act(async () => { type(input, 'FIRST'); await value.api.flushPendingInput(); });
    await waitFor(() => expect(checkpoint).toHaveBeenCalledTimes(1), { timeout: 3000 });
    await act(async () => { await checkpoint.mock.results[0]!.value; });
    await act(async () => {
      type(input, 'SECOND');
      await value.api.flushPendingInput();
      faultControl.crash();
    });
    const acceptedSeq = value.client.state.throughSeq;
    await waitFor(() => expect(error).toHaveBeenCalledTimes(1));
    for (let attempt = 0; attempt < 2; attempt++) {
      const recovered = await value.api.recoverySave();
      const saved = wasm.openPresentation(recovered.bytes);
      try {
        const content = saved.readContent();
        expect(content.ok).toBe(true);
        if (!content.ok) throw new Error(content.failure.message);
        expect(content.stories[0]!.text).toBe(`FIRSTSECOND${target.text}`);
        expect(recovered.throughSeq).toBe(acceptedSeq);
      } finally { saved.dispose(); }
    }
    expect(checkpoint).toHaveBeenCalledTimes(1);
  } finally { checkpoint.mockRestore(); }
});

test('recovery retains acknowledged typing before the first checkpoint', async () => {
  const error = mock(() => {});
  const value = await open({ onError: error });
  const target = await story(value.api);
  await act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 }); });
  const checkpoint = spyOn(value.client, 'checkpoint');
  try {
    const input = value.view.getByLabelText('Slide text');
    await act(async () => { type(input, 'Accepted'); await value.api.flushPendingInput(); });
    const acceptedSeq = value.client.state.throughSeq;
    expect(checkpoint).not.toHaveBeenCalled();
    faultControl.unknown();
    await act(async () => { type(input, 'Uncertain'); await value.api.flushPendingInput().catch(() => {}); });
    await waitFor(() => expect(error).toHaveBeenCalledTimes(1));
    const recovered = await value.api.recoverySave();
    const saved = wasm.openPresentation(recovered.bytes);
    try {
      const content = saved.readContent();
      expect(content.ok).toBe(true);
      if (!content.ok) throw new Error(content.failure.message);
      expect(content.stories[0]!.text).toBe(`Accepted${target.text}`);
      expect(recovered.throughSeq).toBe(acceptedSeq);
    } finally { saved.dispose(); }
    expect(checkpoint).not.toHaveBeenCalled();
  } finally { checkpoint.mockRestore(); }
});

test('unknown_outcome', async () => {
  const error = mock(() => {});
  const changed = mock((_snapshot: DeckSnapshot) => {});
  const value = await open({ onError: error, onChange: changed });
  const target = await story(value.api);
  await act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 }); });
  const input = value.view.getByLabelText('Slide text');
  await act(async () => { type(input, 'Covered'); await value.api.flushPendingInput(); });
  await waitFor(() => expect(changed).toHaveBeenCalled());
  const checkpoint = await value.api.recoverySave();
  faultControl.unknown();
  await act(async () => { type(input, 'Uncertain'); await value.api.flushPendingInput().catch(() => {}); });
  await waitFor(() => expect(error).toHaveBeenCalledTimes(1));
  const recovered = await value.api.recoverySave();
  expect(recovered.bytes).toEqual(checkpoint.bytes);
  expect(value.view.getByRole('button', { name: 'Download recovered presentation' })).toBeDefined();
  expect(value.stage.getAttribute('aria-readonly')).toBe('true');
  await expect(value.api.saveAsync()).rejects.toThrow();
  value.view.unmount();
  await expect(value.api.recoverySave()).rejects.toThrow();
});

test('open and three automatic checkpoints encode one recovery update', async () => {
  const original = wasm.openPresentation;
  const encoded = mock(() => {});
  const updates: Uint8Array[] = [];
  const opened = spyOn(wasm, 'openPresentation').mockImplementation((...args) => {
    const handle = original(...args);
    const encode = handle.encodeStateAsUpdate.bind(handle);
    handle.encodeStateAsUpdate = (...args) => { encoded(); const update = encode(...args); updates.push(update); return update; };
    return handle;
  });
  try {
    const value = await open();
    const target = await story(value.api);
    await act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 }); });
    const checkpoint = spyOn(value.client, 'checkpoint');
    try {
      for (let index = 0; index < 3; index++) {
        await act(async () => { type(value.view.getByLabelText('Slide text'), String(index)); await value.api.flushPendingInput(); });
        await waitFor(() => expect(checkpoint).toHaveBeenCalledTimes(index + 1), { timeout: 3000 });
        await act(async () => {
          const result = (await checkpoint.mock.results[index]!.value) as PresentationCommandSaveResult | undefined;
          expect(result).toBeDefined();
          expect('recoveryUpdate' in result!).toBe(false);
        });
      }
      expect(encoded).toHaveBeenCalledTimes(1);
      expect(updates[0]!.buffer.byteLength).toBe(0);
    } finally { checkpoint.mockRestore(); }
  } finally { opened.mockRestore(); }
});

for (const rebase of [false, true]) test(`recovery replays four edits across checkpoints with rebase ${rebase}`, async () => {
  const create = peerlessRecoveryFactory.create;
  const factory = spyOn(peerlessRecoveryFactory, 'create').mockImplementation((bytes) => create(bytes, rebase ? 0 : undefined));
  const error = mock(() => {});
  try {
    const value = await open({ onError: error });
    const target = await story(value.api);
    await act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 }); });
    const checkpoint = spyOn(value.client, 'checkpoint');
    try {
      let checkpoints = 0;
      for (let index = 1; index <= 4; index++) {
        await act(async () => { type(value.view.getByLabelText('Slide text'), String(index)); await value.api.flushPendingInput(); });
        if (index === 1 || index === 3) {
          const call = checkpoints++;
          await waitFor(() => expect(checkpoint).toHaveBeenCalledTimes(checkpoints), { timeout: 3000 });
          await act(async () => {
            const result = (await checkpoint.mock.results[call]!.value) as PresentationCommandSaveResult | undefined;
            expect(Boolean(result?.recoveryUpdate)).toBe(rebase);
            expect(checkpoint.mock.calls[call]?.[0]).toEqual({ rebase });
          });
        }
      }
      const throughSeq = value.client.state.throughSeq;
      await act(async () => { faultControl.crash(); });
      await waitFor(() => expect(error).toHaveBeenCalledTimes(1));
      const recovered = await value.api.recoverySave();
      const saved = wasm.openPresentation(recovered.bytes);
      try {
        const content = saved.readContent();
        expect(content.ok).toBe(true);
        if (!content.ok) throw new Error(content.failure.message);
        expect(content.stories[0]!.text).toBe(`1234${target.text}`);
        expect(recovered.throughSeq).toBe(throughSeq);
      } finally { saved.dispose(); }
    } finally { checkpoint.mockRestore(); }
  } finally { factory.mockRestore(); }
});

test('hook retains the transferred open recovery buffer by identity', async () => {
  const create = peerlessRecoveryFactory.create;
  let recovery: ReturnType<typeof create> | undefined;
  const factory = spyOn(peerlessRecoveryFactory, 'create').mockImplementation((...args) => {
    recovery = create(...args);
    return recovery;
  });
  try {
    const value = await open();
    expect(value.recoveryUpdate).toBeInstanceOf(ArrayBuffer);
    expect(recovery?.baseline?.update.buffer).toBe(value.recoveryUpdate);
    expect('recoveryUpdate' in value.client.state).toBe(false);
    expect(value.client.takeRecoveryUpdate?.()).toBeUndefined();
  } finally { factory.mockRestore(); }
});
