import { afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import type { CollaborationUpdateOrigin } from '../collaboration/types';
import { createYrsSession, type YrsSession } from './index';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
import { wrapSession } from './yrsSessionFacade';
import * as editWasm from './wasm/index';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))));

let nextClientId = 6200;
const clientId = (): number => ++nextClientId;

const owned: Array<{ destroy(): void }> = [];
afterEach(() => {
  for (const session of owned.splice(0)) {
    try { session.destroy(); } catch {}
  }
});
function own<T extends { destroy(): void }>(session: T): T {
  owned.push(session);
  return session;
}

type Seen = Array<{ update: Uint8Array; origin: CollaborationUpdateOrigin }>;
const record = (seen: Seen) => (update: Uint8Array, origin: CollaborationUpdateOrigin): void => {
  seen.push({ update: update.slice(), origin });
};
const bytes = (seen: Seen): number[][] => seen.map((event) => [...event.update]);

async function stateFor(text: string): Promise<Uint8Array> {
  const seed = await createYrsSession({ clientId: clientId() });
  try {
    seed.loadStories([{ storyId: 'body', paragraphs: [{ text }, { text: 'second paragraph' }] }]);
    return seed.encodeState();
  } finally {
    seed.destroy();
  }
}

async function facadeFromState(state: Uint8Array): Promise<YrsSession> {
  const session = own(await createYrsSession({ clientId: clientId() }));
  session.loadState(state);
  return session;
}

async function residentFromState(state: Uint8Array): Promise<ResidentEngineSession> {
  const session = own(await createResidentEngineSession(undefined, clientId()));
  session.loadState(state);
  return session;
}

function bodyText(session: { paragraphs(story: string): Array<{ text: string }> }): string {
  return session.paragraphs('body').map((paragraph) => paragraph.text).join('\n');
}

function firstRange(session: { paragraphs(story: string): Array<{ paraId: string }> }) {
  const paraId = session.paragraphs('body')[0]!.paraId;
  return { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 1 } };
}

test('r1_resident_listener_reads_state_during_delivery', async () => {
  const engine = await residentFromState(await stateFor('alpha'));
  const errors: unknown[] = [];
  const vectors: Uint8Array[] = [];
  const updates: Uint8Array[] = [];
  engine.onUpdate(() => {
    try {
      vectors.push(engine.encodeStateVector());
      updates.push(engine.encodeStateAsUpdate());
    } catch (error) {
      errors.push(error);
    }
  });
  engine.applyRawOps('body', [{ op: 'insert', index: 0, text: 'X ' }]);
  expect(errors).toEqual([]);
  expect(vectors.length).toBeGreaterThan(0);
  expect(updates.length).toBeGreaterThan(0);
  expect(vectors.at(-1)).toEqual(engine.encodeStateVector());
});

test('r2_resident_listener_write_delivers_fifo_exactly_once_to_every_listener', async () => {
  const state = await stateFor('alpha');
  const engine = await residentFromState(state);
  const first: Seen = [];
  const second: Seen = [];
  let wrote = false;
  engine.onUpdate((update, origin) => {
    first.push({ update: update.slice(), origin });
    if (!wrote) {
      wrote = true;
      engine.applyRawOps('body', [{ op: 'insert', index: 0, text: 'B ' }]);
    }
  });
  engine.onUpdate(record(second));
  engine.applyRawOps('body', [{ op: 'insert', index: 0, text: 'A ' }]);
  expect(first).toHaveLength(2);
  expect(second).toHaveLength(2);
  expect(bytes(second)).toEqual(bytes(first));
  expect(first.map((event) => event.origin)).toEqual(['local', 'local']);
  const peer = await facadeFromState(state);
  for (const event of second) peer.applyUpdate(event.update);
  expect(bodyText(peer)).toBe(bodyText(engine));
  expect(peer.encodeStateVector()).toEqual(engine.encodeStateVector());
});

test('r3_resident_and_facade_over_one_raw_session_both_receive_every_edit', async () => {
  const state = await stateFor('alpha');
  const nativeCreate = editWasm.createEditSession;
  const raws: Array<ReturnType<typeof nativeCreate>> = [];
  const create = spyOn(editWasm, 'createEditSession').mockImplementation((...args) => {
    const raw = nativeCreate(...args);
    raws.push(raw);
    return raw;
  });
  try {
    const engineClientId = clientId();
    const engine = own(await createResidentEngineSession(undefined, engineClientId));
    engine.loadState(state);
    const facade = own(wrapSession(raws[0]!, engineClientId));
    const residentSeen: Seen = [];
    const facadeSeen: Seen = [];
    engine.onUpdate(record(residentSeen));
    facade.onUpdate(record(facadeSeen));
    const paraId = facade.paragraphs('body')[0]!.paraId;
    facade.insertText({ story: 'body', paraId, offset: 0 }, 'F ');
    expect(residentSeen).toHaveLength(1);
    expect(facadeSeen).toHaveLength(1);
    engine.applyRawOps('body', [{ op: 'insert', index: 0, text: 'R ' }]);
    expect(residentSeen).toHaveLength(2);
    expect(facadeSeen).toHaveLength(2);
    expect(bytes(facadeSeen)).toEqual(bytes(residentSeen));
    expect(residentSeen.map((event) => event.origin)).toEqual(['local', 'local']);
  } finally {
    create.mockRestore();
  }
});

test('r4_resident_throwing_listener_does_not_block_other_listeners_or_the_edit', async () => {
  const engine = await residentFromState(await stateFor('alpha'));
  const seen: Seen = [];
  engine.onUpdate(() => { throw new Error('listener failure'); });
  engine.onUpdate(record(seen));
  expect(() => engine.applyRawOps('body', [{ op: 'insert', index: 0, text: 'X ' }])).not.toThrow();
  expect(seen).toHaveLength(1);
  expect(seen[0]!.origin).toBe('local');
});

test('g1_facade_listeners_run_before_mutating_calls_return', async () => {
  const facade = await facadeFromState(await stateFor('alpha beta'));
  const seen: Seen = [];
  facade.onUpdate(record(seen));
  const paraId = facade.paragraphs('body')[0]!.paraId;
  facade.insertText({ story: 'body', paraId, offset: 0 }, 'X ');
  expect(seen).toHaveLength(1);
  expect(seen[0]!.origin).toBe('local');
  expect(facade.undo()).toBe(true);
  expect(seen).toHaveLength(2);
  expect(seen[1]!.origin).toBe('local');
  expect(facade.redo()).toBe(true);
  expect(seen).toHaveLength(3);
  expect(seen[2]!.origin).toBe('local');
  expect(bodyText(facade)).toContain('X alpha beta');
});

async function expectRemoteThenCleanupAndChunks(
  makeReceiver: (state: Uint8Array) => Promise<YrsSession | ResidentEngineSession>
): Promise<void> {
  const state = await stateFor('a');
  const receiver = await makeReceiver(state);
  const peer = await facadeFromState(state);
  const seen: Seen = [];
  receiver.onUpdate(record(seen));
  const peerSeen: Seen = [];
  peer.onUpdate(record(peerSeen));
  peer.formatRange(firstRange(peer), { bold: true });
  peer.formatRange(firstRange(peer), { bold: null });
  receiver.formatRange(firstRange(receiver), { bold: true });
  seen.length = 0;
  const peerUpdates = peerSeen.filter((event) => event.origin === 'local').map((event) => event.update);
  expect(peerUpdates.length).toBeGreaterThan(0);
  let cleanupSeen = false;
  for (const update of peerUpdates) {
    const start = seen.length;
    receiver.applyUpdate(update);
    const delivered = seen.slice(start);
    if (delivered.length === 0) continue;
    expect(delivered[0]!.origin).toBe('remote');
    const firstLocal = delivered.findIndex((event) => event.origin === 'local');
    if (firstLocal >= 0) {
      cleanupSeen = true;
      expect(firstLocal).toBeGreaterThan(0);
      expect(delivered.slice(firstLocal).every((event) => event.origin === 'local')).toBe(true);
    }
  }
  expect(cleanupSeen).toBe(true);

  const chunkSource = await facadeFromState(state);
  const chunkPeerSeen: Seen = [];
  chunkSource.onUpdate(record(chunkPeerSeen));
  const chunkParaId = chunkSource.paragraphs('body')[0]!.paraId;
  chunkSource.insertText({ story: 'body', paraId: chunkParaId, offset: 0 }, '1');
  chunkSource.insertText({ story: 'body', paraId: chunkParaId, offset: 0 }, '2');
  chunkSource.insertText({ story: 'body', paraId: chunkParaId, offset: 0 }, '3');
  const chunks = chunkPeerSeen.filter((event) => event.origin === 'local').map((event) => event.update);
  expect(chunks.length).toBeGreaterThanOrEqual(3);
  const chunkReceiver = await makeReceiver(state);
  const chunkSeen: Seen = [];
  chunkReceiver.onUpdate(record(chunkSeen));
  chunkReceiver.applyUpdateChunks(chunks);
  expect(chunkSeen.length).toBeGreaterThanOrEqual(3);
  expect(chunkSeen.every((event) => event.origin === 'remote')).toBe(true);
  expect(bytes(chunkSeen)).toEqual(chunks.map((chunk) => [...chunk]));
  const third = await facadeFromState(state);
  for (const event of chunkSeen) third.applyUpdate(event.update);
  expect(third.encodeStateVector()).toEqual(chunkReceiver.encodeStateVector());
}

test('g2_facade_apply_update_remote_then_cleanup_and_chunks_in_order', async () => {
  await expectRemoteThenCleanupAndChunks(facadeFromState);
});

test('g2_resident_apply_update_remote_then_cleanup_and_chunks_in_order', async () => {
  await expectRemoteThenCleanupAndChunks(residentFromState);
});

test('g3_facade_listener_write_during_delivery_is_fifo_exactly_once', async () => {
  const facade = await facadeFromState(await stateFor('alpha'));
  const first: Seen = [];
  const second: Seen = [];
  let wrote = false;
  facade.onUpdate((update, origin) => {
    first.push({ update: update.slice(), origin });
    if (!wrote) {
      wrote = true;
      const paraId = facade.paragraphs('body')[0]!.paraId;
      facade.insertText({ story: 'body', paraId, offset: 0 }, 'B ');
    }
  });
  facade.onUpdate(record(second));
  const paraId = facade.paragraphs('body')[0]!.paraId;
  facade.insertText({ story: 'body', paraId, offset: 0 }, 'A ');
  expect(first).toHaveLength(2);
  expect(second).toHaveLength(2);
  expect(bytes(second)).toEqual(bytes(first));
  expect(first.map((event) => event.origin)).toEqual(['local', 'local']);
  expect(bodyText(facade)).toContain('B A alpha');
});
