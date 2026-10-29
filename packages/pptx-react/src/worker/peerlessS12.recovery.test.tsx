import { deck, faultControl, open, story, type } from './peerlessTestHarness';
import { afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
import type { DeckSnapshot } from '@betteroffice/pptx';
import type { PresentationCommandSaveResult } from '@betteroffice/pptx/worker';
import * as wasm from '../../../pptx/src/wasm/loader';
import { peerlessSessionOpener } from './usePeerlessPresentation';
import { peerlessRecoveryFactory } from './peerlessRecoveryJournal';
import { peerlessCheckpointTimer } from './usePeerlessPresentation';
const { act, waitFor } = await import('@testing-library/react');

const last = <T,>(items: readonly T[]): T | undefined => items[items.length - 1];
let dueCheckpoints: (() => void)[] = [];
let checkpointTimer: ReturnType<typeof spyOn> | undefined;
beforeEach(() => {
  dueCheckpoints = [];
  checkpointTimer = spyOn(peerlessCheckpointTimer, 'start').mockImplementation((callback) => {
    dueCheckpoints.push(callback);
    return () => { dueCheckpoints = dueCheckpoints.filter((due) => due !== callback); };
  });
});
afterEach(() => { checkpointTimer?.mockRestore(); });

async function fireCheckpointTimer(): Promise<void> {
  await waitFor(() => expect(dueCheckpoints).toHaveLength(1), { timeout: 3000 });
  await act(async () => { dueCheckpoints.shift()!(); });
}

test('recovery retains typing after an edited automatic checkpoint', async () => {
  const error = mock(() => {});
  const value = await open({ onError: error });
  const target = await story(value.api);
  await act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 }); });
  const checkpoint = spyOn(value.client, 'checkpoint');
  try {
    const input = value.view.getByLabelText('Slide text');
    await act(async () => { type(input, 'FIRST'); await value.api.flushPendingInput(); });
    await fireCheckpointTimer();
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

test('recovery after a crash on a deck whose state exceeds the update cap', async () => {
  const updateCap = 64 * 1024 * 1024;
  let bytes: Uint8Array;
  {
    const zip = await JSZip.loadAsync(await deck());
    const media = new Uint8Array(48 * 1024 * 1024 + 64 * 1024);
    let random = 0x12345678;
    for (let offset = 0; offset < media.length; offset++) {
      random ^= random << 13;
      random ^= random >>> 17;
      random ^= random << 5;
      media[offset] = random & 0xff;
    }
    media.set(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a7N8AAAAASUVORK5CYII=', 'base64'));
    zip.file('ppt/media/large.png', media);
    zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>',
      '<Default Extension="png" ContentType="image/png"/></Types>'));
    const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    zip.file('ppt/slides/_rels/slide2.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${rel}/image" Target="../media/large.png"/></Relationships>`);
    zip.file('ppt/slides/slide2.xml', (await zip.file('ppt/slides/slide2.xml')!.async('string')).replace('</p:spTree>',
      '<p:pic><p:nvPicPr><p:cNvPr id="3" name="Large picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="7000000" y="100000"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree>'));
    bytes = await zip.generateAsync({ type: 'uint8array', compression: 'STORE' });
  }
  const error = mock(() => {});
  const opened = spyOn(wasm, 'openPresentation');
  let value: Awaited<ReturnType<typeof open>>;
  try {
    value = await open({ file: bytes, onError: error });
    const handle = opened.mock.results[0]!.value as wasm.PresentationHandle;
    expect(handle.encodeStateAsUpdate().byteLength).toBeGreaterThan(updateCap);
  } finally { opened.mockRestore(); }
  const target = await story(value.api);
  const checkpoint = spyOn(value.client, 'checkpoint');
  try {
    await act(async () => {
      await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 });
      type(value.view.getByLabelText('Slide text'), 'Accepted');
      await value.api.flushPendingInput();
      faultControl.crash();
    });
    const acceptedSeq = value.client.state.throughSeq;
    await waitFor(() => expect(error).toHaveBeenCalledTimes(1), { timeout: 15000 });
    expect(checkpoint).not.toHaveBeenCalled();
    const recovered = await value.api.recoverySave();
    const saved = wasm.openPresentation(recovered.bytes);
    try {
      const content = saved.readContent();
      expect(content.ok).toBe(true);
      if (!content.ok) throw new Error(content.failure.message);
      expect(content.stories[0]!.text).toBe(`Accepted${target.text}`);
      expect(recovered.throughSeq).toBe(acceptedSeq);
    } finally { saved.dispose(); }
  } finally { checkpoint.mockRestore(); }
}, 60_000);

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

test('source open and three ordinary checkpoints encode no recovery update', async () => {
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
        await fireCheckpointTimer();
        await waitFor(() => expect(checkpoint).toHaveBeenCalledTimes(index + 1), { timeout: 3000 });
        await act(async () => {
          const result = (await checkpoint.mock.results[index]!.value) as PresentationCommandSaveResult | undefined;
          expect(result).toBeDefined();
          expect('recoveryUpdate' in result!).toBe(false);
        });
      }
      expect(encoded).not.toHaveBeenCalled();
      expect(updates).toHaveLength(0);
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
          await fireCheckpointTimer();
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

test('hook retains the transferred initial room recovery buffer by identity', async () => {
  const create = peerlessRecoveryFactory.create;
  let recovery: ReturnType<typeof create> | undefined;
  const factory = spyOn(peerlessRecoveryFactory, 'create').mockImplementation((...args) => {
    recovery = create(...args);
    return recovery;
  });
  try {
    const value = await open({ collaboration: { clientId: 9001 } });
    expect(value.client.state.recoveryBaseline).toBe('update');
    expect(value.recoveryUpdate).toBeInstanceOf(ArrayBuffer);
    expect(recovery?.baseline?.kind === 'update' ? recovery.baseline.update.buffer : undefined).toBe(value.recoveryUpdate);
    expect('recoveryUpdate' in value.client.state).toBe(false);
    expect(value.client.takeRecoveryUpdate?.()).toBeUndefined();
  } finally { factory.mockRestore(); }
});

async function equalParts(actual: Uint8Array, expected: Uint8Array) {
  const [left, right] = await Promise.all([JSZip.loadAsync(actual), JSZip.loadAsync(expected)]);
  const paths = (zip: JSZip) => Object.keys(zip.files).filter((path) => !zip.files[path]!.dir).sort();
  expect(paths(left)).toEqual(paths(right));
  for (const path of paths(left)) expect(await left.file(path)!.async('uint8array')).toEqual(await right.file(path)!.async('uint8array'));
}

for (const rebase of [false, true]) test(`source recovery preserves insert, delete, move and slide add with middle rebase ${rebase}`, async () => {
  const create = peerlessRecoveryFactory.create;
  let recovery: ReturnType<typeof create> | undefined;
  const factory = spyOn(peerlessRecoveryFactory, 'create').mockImplementation((...args) => {
    recovery = create(...args);
    return recovery;
  });
  const error = mock(() => {});
  try {
    const value = await open({ onError: error });
    expect(value.client.state.recoveryBaseline).toBe('source');
    expect(value.recoveryUpdate).toBeUndefined();
    const target = await story(value.api);
    const at = (offset: number) => ({ source: 'confirmed' as const, selection: { slideId: target.slideId,
      shapeId: target.shapeId, storyId: target.storyId, anchor: offset, focus: offset, affinity: 'downstream' as const } });
    const input = { inputTime: 0, undoGroup: 'recovery-test' };
    await act(async () => {
      const ack = await value.client.dispatch({ kind: 'insertText', payload: { ...input, at: at(0), text: 'typed' } });
      expect(last(ack.outcomes)?.committed).toBe(true);
      await value.api.recoverySave();
    });
    expect(recovery?.baseline?.kind).toBe('source');
    await act(async () => {
      const ack = await value.client.dispatch({ kind: 'deleteBackward', payload: { ...input, unit: 'codePoint', at: at(5) } });
      expect(last(ack.outcomes)?.committed).toBe(true);
    });
    const bound = spyOn(recovery!, 'needsRebase').mockReturnValueOnce(rebase);
    try { await act(async () => { await value.api.recoverySave(); }); }
    finally { bound.mockRestore(); }
    expect(recovery?.baseline?.kind).toBe(rebase ? 'update' : 'source');
    await act(async () => {
      const selected = await value.client.dispatch({ kind: 'selectObject', payload: { ...input,
        frame: value.client.state.frame!, shapeId: target.shapeId } });
      const object = selected.selection;
      if (object?.kind !== 'object' || !object.rect) throw new Error('Expected a movable object');
      const ack = await value.client.dispatch({ kind: 'moveShape', payload: { ...input,
        frame: value.client.state.frame!, target: object, x: object.rect.x + 10000, y: object.rect.y + 20000 } });
      expect(last(ack.outcomes)?.committed).toBe(true);
      await value.api.recoverySave();
    });
    await act(async () => {
      const ack = await value.client.dispatch({ kind: 'invoke', payload: { ...input, method: 'insertSlide', args: [2] } });
      expect(last(ack.outcomes)?.committed).toBe(true);
    });
    const reference = await value.client.save();
    const throughSeq = value.client.state.throughSeq;
    expect(recovery!.hasUncheckpointedUpdates()).toBe(true);
    await act(async () => { faultControl.crash(); });
    await waitFor(() => expect(error).toHaveBeenCalledTimes(1));
    const recovered = await value.api.recoverySave();
    expect(recovered.throughSeq).toBe(throughSeq);
    await equalParts(recovered.bytes, reference);
  } finally { factory.mockRestore(); }
});

for (const initialUpdate of [false, true]) test(`recovery preserves a remote insertion before a dependent local delete with update baseline ${initialUpdate}`, async () => {
  const error = mock(() => {});
  const bytes = await deck();
  await wasm.initWasm(await readFile(new URL('../../../pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  const remote = wasm.openPresentation(bytes, { clientId: 9003 });
  try {
    if (initialUpdate) {
      const content = remote.readContent();
      if (!content.ok) throw new Error(content.failure.message);
      remote.insertText(content.stories[0]!.storyId, 0, 'ROOM', {});
    }
    const value = await open({ file: bytes, onError: error, ...(initialUpdate ? {
      collaboration: { clientId: 9004, initialUpdate: remote.encodeStateAsUpdate() },
    } : {}) });
    expect(value.client.state.recoveryBaseline).toBe(initialUpdate ? 'update' : 'source');
    await remoteThenLocal(value, await story(value.api), remote, error);
  } finally { remote.dispose(); }
});

async function remoteThenLocal(value: Awaited<ReturnType<typeof open>>, target: Awaited<ReturnType<typeof story>>,
  remote: wasm.PresentationHandle, error: () => void) {
  const updates: Uint8Array[] = [];
  const off = remote.onUpdate((update) => updates.push(update.slice()));
  try { remote.insertText(target.storyId, 0, 'REMOTE', {}); } finally { off(); }
  expect(updates.length).toBeGreaterThan(0);
  await act(async () => {
    const ack = await value.client.dispatch({ kind: 'applyRemoteUpdate', payload: { bytes: updates.map((update) => update.slice().buffer),
      inputTime: 0, undoGroup: 'remote' } });
    expect(ack.updates?.some((update) => update.origin === 'remote')).toBe(true);
    await value.api.recoverySave();
    const local = await value.client.dispatch({ kind: 'deleteBackward', payload: { inputTime: 1, undoGroup: 'local', unit: 'codePoint',
      at: { source: 'confirmed', selection: { slideId: target.slideId, shapeId: target.shapeId, storyId: target.storyId,
        anchor: 6, focus: 6, affinity: 'downstream' } } } });
    expect(last(local.outcomes)?.committed).toBe(true);
  });
  const reference = await value.client.save();
  const throughSeq = value.client.state.throughSeq;
  await act(async () => { faultControl.crash(); });
  await waitFor(() => expect(error).toHaveBeenCalledTimes(1));
  const opener = spyOn(peerlessSessionOpener, 'open');
  opener.mockClear();
  try {
    const recovered = await value.api.recoverySave();
    expect(recovered.throughSeq).toBe(throughSeq);
    expect(opener).toHaveBeenCalledTimes(1);
    expect(Boolean(opener.mock.calls[0]?.[1]?.initialUpdate)).toBe(value.client.state.recoveryBaseline === 'update');
    await equalParts(recovered.bytes, reference);
  } finally { opener.mockRestore(); }
}

for (const fault of ['gap', 'engine'] as const) test(`source recovery rejects ${fault}`, async () => {
  const create = peerlessRecoveryFactory.create;
  let recovery: ReturnType<typeof create> | undefined;
  const factory = spyOn(peerlessRecoveryFactory, 'create').mockImplementation((...args) => {
    recovery = create(...args);
    return recovery;
  });
  const error = mock(() => {});
  try {
    const value = await open({ onError: error });
    const target = await story(value.api);
    await act(async () => {
      await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 });
      type(value.view.getByLabelText('Slide text'), 'accepted');
      await value.api.flushPendingInput();
      faultControl.crash();
    });
    await waitFor(() => expect(error).toHaveBeenCalledTimes(1));
    if (fault === 'gap') recovery!.journal.shift();
    else {
      const baseline = recovery!.baseline;
      if (baseline?.kind !== 'source') throw new Error('Expected source baseline');
      baseline.engine = 'different-engine';
    }
    await expect(value.api.recoverySave()).rejects.toThrow('Recovery baseline is unavailable');
  } finally { factory.mockRestore(); }
});
