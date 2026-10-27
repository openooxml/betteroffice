import { afterEach, expect, test } from 'bun:test';
import { MAX_UPDATE_BYTES, createResidentDocxCommandHooks } from './residentCommandActions';
import { createResidentCollaboration, readResidentPresence } from './residentCollaboration';
import { createYrsSession } from './index';
import { applyFrameDelta, decodeFrameDelta } from '../layout/render/frameDelta';
import { MAX_COLLABORATION_FRAME_BYTES } from '../collaboration/protocol';
import type { DocxCommandEnvelope, DocxCommandHooksEnvironment, DocxCommandInitRequest } from './residentCommandProtocol';
import { TIMEOUT, closeCommandEngines, command, context, openCommandEngine, selection, texts } from './residentCommandTestSupport';

afterEach(closeCommandEngines);

const collaborating: DocxCommandHooksEnvironment = {
  collaboration: true,
  readSelectionLayout: () => { throw new Error('Selection layout is unavailable'); },
};

async function replica(checkpoint?: Uint8Array, clientId = 9101) {
  const arm = await openCommandEngine(['Seed'], { checkpoint, clientId });
  const buffer = createResidentCollaboration(arm.engine);
  const init = { session: 'collaboration-worker', incarnation: 1, epoch: 1, readOnly: false,
    selection: arm.state.selection, fonts: [], layoutInput: '{}', layoutWithRegions: true,
    layoutRevision: 1, extras: '{}' } satisfies DocxCommandInitRequest;
  const hooks = buffer.wrap(createResidentDocxCommandHooks(arm.engine, init, collaborating));
  const run = async (kind: string, payload: unknown) => {
    const intent = command(arm, kind as never, payload as never);
    const ctx = context(arm, intent);
    const prepared = await hooks.preflight(intent, ctx);
    const result = prepared.status === 'ready' ? await hooks.apply(prepared.prepared, ctx) : prepared;
    if (result.status === 'applied') arm.version++;
    return { ...result, selection: hooks.selection(), editState: hooks.editState() };
  };
  return { arm, buffer, hooks, run, init };
}

test('initial state is loaded once and the replica uses its own client', async () => {
  const a = await replica();
  const b = await replica(a.arm.engine.encodeState(), 9102);
  expect(texts(b.arm)).toEqual(['Seed']);
  expect(b.arm.native.client_id()).toBe(9102);
  expect(b.buffer.drain()).toBeUndefined();
  await b.run('insertText', { text: 'B' });
  const peer = await createYrsSession({ clientId: 9103 });
  try {
    peer.loadState(b.arm.engine.encodeState());
    expect(peer.paragraphs('body').map(p => p.text)).toEqual(['BSeed']);
    expect(b.arm.engine.encodeStateVector()).not.toEqual(a.arm.engine.encodeStateVector());
  } finally { peer.destroy(); a.buffer.dispose(); b.buffer.dispose(); }
}, TIMEOUT);

test('local updates drain in order, remote merges do not echo and duplicates are noop', async () => {
  const a = await replica();
  const b = await replica(a.arm.engine.encodeState(), 9102);
  try {
    await b.run('insertText', { text: 'B' });
    const pushed = b.buffer.drain()!;
    expect(pushed.updates.length).toBeGreaterThan(0);
    expect(await a.run('applyRemoteUpdate', { update: pushed.updates[0]!.slice(0) })).toMatchObject({ status: 'applied' });
    expect(a.buffer.drain()?.updates ?? []).toEqual([]);
    expect(await a.run('applyRemoteUpdate', { update: pushed.updates[0]!.slice(0) })).toMatchObject({ status: 'noop' });
    expect(texts(a.arm)).toEqual(texts(b.arm));
  } finally { a.buffer.dispose(); b.buffer.dispose(); }
}, TIMEOUT);

test('remote edits stay outside local undo and rebase the confirmed caret', async () => {
  const a = await replica();
  const b = await replica(a.arm.engine.encodeState(), 9102);
  try {
    const id = a.arm.engine.paragraphs('body')[0]!.paraId;
    await a.run('select', { selection: selection(id, 2) });
    await a.run('insertText', { text: 'A' });
    await b.run('insertText', { text: 'B' });
    const update = b.buffer.drain()!.updates[0]!;
    const merged = await a.run('applyRemoteUpdate', { update });
    expect(merged.selection?.focus.offset).toBe(4);
    await a.run('undo', {});
    expect(texts(a.arm)).toEqual(['BSeed']);
    await a.run('redo', {});
    expect(texts(a.arm)[0]).toContain('A');
    expect(texts(a.arm)[0]).toContain('B');
  } finally { a.buffer.dispose(); b.buffer.dispose(); }
}, TIMEOUT);

test('read-only merges bypass basis admission and invalid updates leave the document unchanged', async () => {
  const a = await replica();
  const b = await replica(a.arm.engine.encodeState(), 9102);
  try {
    await a.run('setReadOnly', { readOnly: true });
    await b.run('insertText', { text: 'B' });
    const update = b.buffer.drain()!.updates[0]!;
    const stale = command(a.arm, 'applyRemoteUpdate', { update }, { basis: { epoch: 1, version: 999 } });
    expect((await a.hooks.preflight(stale, context(a.arm, stale))).status).toBe('ready');
    expect((await a.run('applyRemoteUpdate', { update })).status).toBe('applied');
    const before = a.arm.engine.encodeStateVector();
    expect(await a.run('applyRemoteUpdate', { update: new Uint8Array([255]).buffer })).toMatchObject({ status: 'refused', reason: 'invalid-update' });
    expect(await a.run('applyRemoteUpdate', { update: new ArrayBuffer(0) })).toMatchObject({ status: 'refused', reason: 'invalid-payload' });
    expect(a.arm.engine.encodeStateVector()).toEqual(before);
  } finally { a.buffer.dispose(); b.buffer.dispose(); }
}, TIMEOUT);

test('worker sticky selections and main-thread encoded selections resolve both ways', async () => {
  const a = await replica();
  const peer = await createYrsSession({ clientId: 9103 });
  try {
    peer.loadState(a.arm.engine.encodeState());
    const id = a.arm.engine.paragraphs('body')[0]!.paraId;
    await a.run('select', { selection: selection(id, 2) });
    const confirmed = a.hooks.selection()!;
    expect(peer.resolveStickyPosition({ story: 'body', encoded: confirmed.focus.sticky! })?.offset).toBe(2);
    peer.setSelection({ story: 'body', paraId: id, offset: 3 });
    expect(a.arm.engine.resolveStickyPosition({ story: 'body', encoded: peer.encodeSelection()!.head })?.offset).toBe(3);
  } finally { peer.destroy(); a.buffer.dispose(); }
}, TIMEOUT);

test('a session without collaboration refuses remote merges as unsupported', async () => {
  const a = await replica();
  try {
    const plain = createResidentDocxCommandHooks(a.arm.engine, { session: 'plain', incarnation: 1, epoch: 1, readOnly: false,
      selection: a.arm.state.selection, fonts: [], layoutInput: '{}', layoutWithRegions: true, layoutRevision: 1, extras: '{}' });
    const intent = command(a.arm, 'applyRemoteUpdate', { update: new Uint8Array([1]).buffer });
    expect(await plain.preflight(intent, context(a.arm, intent))).toMatchObject({ status: 'refused', reason: 'unsupported-command' });
    expect(texts(a.arm)).toEqual(['Seed']);
  } finally { a.buffer.dispose(); }
}, TIMEOUT);

test('remote merges are admitted from the command policy in a collaborating session', async () => {
  const a = await replica();
  try {
    const hooks = createResidentDocxCommandHooks(a.arm.engine, { ...a.init, readOnly: true }, collaborating);
    const admit = async (payload: unknown, extra: Partial<DocxCommandEnvelope> = {}) => {
      const intent = command(a.arm, 'applyRemoteUpdate', payload as { update: ArrayBuffer }, extra);
      return hooks.preflight(intent, context(a.arm, intent));
    };
    expect(await admit({ update: new Uint8Array([1]).buffer })).toMatchObject({ status: 'ready' });
    expect(await admit({ update: new Uint8Array([1]).buffer }, { basis: { epoch: 2, version: 999 } })).toMatchObject({ status: 'ready' });
    expect(await admit({ update: new ArrayBuffer(MAX_COLLABORATION_FRAME_BYTES + 1) })).toMatchObject({ status: 'ready' });
    for (const payload of [{}, { update: new ArrayBuffer(0) }, { update: new ArrayBuffer(MAX_UPDATE_BYTES + 1) },
      { update: new Uint8Array([1]) }, { update: new Uint8Array([1]).buffer, selection: null }]) {
      expect(await admit(payload)).toMatchObject({ status: 'refused', reason: 'invalid-payload' });
    }
    expect(texts(a.arm)).toEqual(['Seed']);
  } finally { a.buffer.dispose(); }
}, TIMEOUT);

test('presence skips carets and rects on pages the retained frame does not hold', async () => {
  const a = await replica();
  try {
    const frame = applyFrameDelta(null, decodeFrameDelta(a.arm.engine.buildDisplayListFrame('{}', 0)));
    const paraId = a.arm.engine.paragraphs('body')[0]!.paraId;
    const sticky = (offset: number) => a.arm.engine.encodeStickyPosition({ story: 'body', paraId, offset }).encoded;
    const peers = [{ clientId: 9103, clock: 1, cursor: { story: 'body', anchor: sticky(1), head: sticky(3) } }];
    const held = await readResidentPresence(a.arm.engine, frame, peers, [0, 1], 1);
    expect(held.peers[0]!.caret).toMatchObject({ pageIndex: 0 });
    expect(held.peers[0]!.rects.length).toBeGreaterThan(0);
    const missing = await readResidentPresence(a.arm.engine, { ...frame, pages: [] }, peers, [0, 1], 1);
    expect(missing.peers[0]).toMatchObject({ status: 'ok', caret: null, rects: [] });
  } finally { a.buffer.dispose(); }
}, TIMEOUT);

test('remote merges mark their stories for the next save projection', async () => {
  const a = await replica();
  const b = await replica(a.arm.engine.encodeState(), 9102);
  try {
    const marked: string[] = [];
    const mark = a.arm.engine.markProjectionStories;
    a.arm.engine.markProjectionStories = (stories) => { marked.push(...stories); mark(stories); };
    await b.run('insertText', { text: 'B' });
    expect(await a.run('applyRemoteUpdate', { update: b.buffer.drain()!.updates[0]! })).toMatchObject({ status: 'applied' });
    expect(marked).toContain('body');
  } finally { a.buffer.dispose(); b.buffer.dispose(); }
}, TIMEOUT);
