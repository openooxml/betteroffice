import { beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { openPresentation, type PresentationHandle } from '../index';
import { initWasm } from '../wasm/loader';
import { CollaborationProvider } from './provider';
import type { CollaborationTransportEvent } from './types';

const root = resolve(import.meta.dir, '../../../..');
let fixture: Uint8Array;

beforeAll(async () => {
  const [wasm, pptx] = await Promise.all([
    readFile(resolve(root, 'packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
  ]);
  await initWasm(wasm);
  fixture = pptx;
});

function replicas(count: number, text: string) {
  const source = openPresentation(fixture, { clientId: 4200 });
  const slideId = source.snapshot().slides[0].id;
  const { shapeId } = source.addTextBox(slideId, {
    name: 'Cleanup',
    rect: { x: 500_000, y: 500_000, width: 3_000_000, height: 700_000 },
    text,
    style: {},
  });
  const storyId = source.snapshot().slides[0].shapes.find((shape) => shape.id === shapeId)!
    .textStories[0].id;
  const seed = source.encodeStateAsUpdate();
  source.dispose();
  const handles = Array.from({ length: count }, (_, index) =>
    openPresentation(fixture, { clientId: 4201 + index, initialUpdate: seed })
  );
  return { handles, storyId };
}

type DeleteSet = Array<[client: number, ranges: Array<[start: number, end: number]>]>;

/** Decodes a v1 update that must hold no structs into its delete set, ranges merged. */
function deleteOnly(update: Uint8Array): DeleteSet {
  let at = 0;
  const read = () => {
    let value = 0;
    for (let scale = 1; ; scale *= 128) {
      if (at >= update.length) throw new Error('truncated update');
      const byte = update[at++];
      value += (byte & 0x7f) * scale;
      if (byte < 0x80) return value;
    }
  };
  expect(read()).toBe(0);
  const clients = new Map<number, Array<[number, number]>>();
  for (let count = read(); count > 0; count -= 1) {
    const client = read();
    const ranges = clients.get(client) ?? [];
    for (let length = read(); length > 0; length -= 1) {
      const start = read();
      ranges.push([start, start + read()]);
    }
    clients.set(client, ranges);
  }
  expect(at).toBe(update.length);
  return [...clients].sort(([a], [b]) => a - b).map(([client, ranges]) => {
    const merged: Array<[number, number]> = [];
    for (const [start, end] of ranges.sort(([a], [b]) => a - b)) {
      const last = merged[merged.length - 1];
      if (last && start <= last[1]) last[1] = Math.max(last[1], end);
      else merged.push([start, end]);
    }
    return [client, merged];
  });
}

/** A replica's whole delete set: a diff against its own state vector holds nothing else. */
function deleteSet(handle: PresentationHandle): DeleteSet {
  return deleteOnly(handle.encodeStateAsUpdate(handle.encodeStateVector()));
}

function clocks(set: DeleteSet): string[] {
  return set.flatMap(([client, ranges]) => ranges.flatMap(([start, end]) =>
    Array.from({ length: end - start }, (_, offset) => `${client}:${start + offset}`)));
}

/** The 'local' updates one edit or one remote apply published, and the deletions they carry. */
type Sent = { updates: Uint8Array[]; deleted: Set<string> };

function outboxes(handles: PresentationHandle[]) {
  const local = handles.map(() => [] as Uint8Array[]);
  const stops = handles.map((handle, index) => handle.onUpdate((update, origin) => {
    if (origin === 'local') local[index].push(update.slice());
  }));
  let published = 0;
  const edit = (index: number, change: () => void): Sent => {
    const before = new Set(clocks(deleteSet(handles[index])));
    change();
    const deleted = clocks(deleteSet(handles[index])).filter((id) => !before.has(id));
    return { updates: local[index].splice(0), deleted: new Set(deleted) };
  };
  /** What `to` publishes while applying `sent` must be delete-only; `exact` also requires it
   * to be exactly the deletions `to` made beyond what `sent` carried. */
  const deliver = (sent: Sent, to: number, exact: boolean): Sent => {
    const before = new Set(clocks(deleteSet(handles[to])));
    for (const update of sent.updates) handles[to].applyUpdate(update);
    const updates = local[to].splice(0);
    const deleted = updates.flatMap((update) => clocks(deleteOnly(update)));
    if (exact) {
      const expected = clocks(deleteSet(handles[to]))
        .filter((id) => !before.has(id) && !sent.deleted.has(id));
      expect(deleted.sort()).toEqual(expected.sort());
    }
    published += updates.length;
    return { updates, deleted: new Set(deleted) };
  };
  return {
    edit,
    deliver,
    published: () => published,
    quiet: () => local.every((updates) => updates.length === 0),
    stop: () => stops.forEach((stop) => stop()),
  };
}

function text(handle: PresentationHandle, storyId: string): string {
  return handle.story(storyId).paragraphs[0].runs.map((run) => run.text).join('');
}

/** The seeded edit stream: a replica index and an edit to run on it. */
function randomEdits(handles: PresentationHandle[], storyId: string) {
  let seed = 0x5eed;
  const random = () => {
    seed = (seed * 1_103_515_245 + 12_345) % 2_147_483_648;
    return seed / 2_147_483_648;
  };
  const pick = (limit: number) => Math.floor(random() * limit);
  return () => {
    const index = pick(handles.length);
    const handle = handles[index];
    const length = text(handle, storyId).length;
    const roll = random();
    if (roll < 0.15 || length < 3) {
      const at = pick(length + 1);
      return { index, change: () => handle.insertText(storyId, at, 'xy') };
    }
    if (roll < 0.25) {
      const start = pick(length - 1);
      return { index, change: () => handle.deleteText(storyId, start, start + 1) };
    }
    const start = pick(length - 1);
    const end = start + 1 + pick(length - start);
    const key = random() < 0.6 ? 'bold' : 'italic';
    const value = random() < 0.5;
    return { index, change: () => handle.formatText(storyId, start, end, { [key]: value }) };
  };
}

function expectConverged(handles: PresentationHandle[], storyId: string) {
  const [first, ...rest] = handles;
  for (const handle of rest) {
    expect([...handle.encodeStateVector()]).toEqual([...first.encodeStateVector()]);
    expect(handle.story(storyId)).toEqual(first.story(storyId));
    expect(deleteSet(handle)).toEqual(deleteSet(first));
  }
}

test('seeded concurrent deck text edits converge through incremental updates alone', () => {
  const { handles, storyId } = replicas(2, 'abcdefghij');
  const net = outboxes(handles);
  const inboxes = handles.map(() => [] as Sent[]);
  const sent: Uint8Array[] = [];
  const broadcast = (message: Sent, from: number) => {
    sent.push(...message.updates);
    if (message.updates.length === 0) return;
    inboxes.forEach((inbox, to) => { if (to !== from) inbox.push(message); });
  };
  const settle = (): number => {
    let rounds = 0;
    while (inboxes.some((inbox) => inbox.length > 0)) {
      rounds += 1;
      if (rounds > 8) throw new Error('replicas kept publishing updates');
      inboxes.forEach((inbox, to) => {
        for (const message of inbox.splice(0)) broadcast(net.deliver(message, to, true), to);
      });
    }
    return rounds;
  };
  const next = randomEdits(handles, storyId);
  try {
    for (let edit = 0; edit < 300; edit += 1) {
      const { index, change } = next();
      broadcast(net.edit(index, change), index);
      if (edit % 3 === 2) {
        expect(settle()).toBeLessThanOrEqual(4);
        expectConverged(handles, storyId);
      }
    }
    expect(net.published()).toBeGreaterThan(0);
    for (const update of sent) {
      for (const handle of handles) handle.applyUpdate(update);
    }
    expect(net.quiet()).toBe(true);
    expectConverged(handles, storyId);
  } finally {
    net.stop();
    for (const handle of handles) handle.dispose();
  }
}, 30_000);

test('a cleanup made by a relaying deck replica reaches the replica behind it', () => {
  const { handles, storyId } = replicas(3, 'abcdefghij');
  const net = outboxes(handles);
  const neighbours = [[1], [0, 2], [1]];
  const queue: Array<{ message: Sent; to: number; from: number }> = [];
  const publish = (message: Sent, from: number, except = -1) => {
    if (message.updates.length === 0) return;
    for (const to of neighbours[from]) if (to !== except) queue.push({ message, to, from });
  };
  const settle = () => {
    for (let rounds = 0; queue.length > 0; rounds += 1) {
      if (rounds > 12) throw new Error('replicas kept publishing updates');
      for (const { message, to, from } of queue.splice(0)) {
        publish(net.deliver(message, to, false), to);
        if (to === 1) publish(message, 1, from);
      }
    }
  };
  const next = randomEdits(handles, storyId);
  try {
    for (let edit = 0; edit < 300; edit += 1) {
      const { index, change } = next();
      publish(net.edit(index, change), index);
      if (edit % 3 === 2) {
        settle();
        expectConverged(handles, storyId);
      }
    }
  } finally {
    net.stop();
    for (const handle of handles) handle.dispose();
  }
}, 30_000);

test('a deck provider pair converges through seeded concurrent edits', () => {
  const { handles, storyId } = replicas(2, 'abcdefghij');
  const queues: Uint8Array[][] = [[], []];
  const listeners: Array<((event: CollaborationTransportEvent) => void) | undefined> = [];
  const providers = handles.map((handle, side) => new CollaborationProvider(handle, {
    connect() {},
    disconnect() {},
    send(data) { queues[side].push(data.slice()); return true; },
    onEvent(listener) {
      listeners[side] = listener;
      return () => { listeners[side] = undefined; };
    },
  }));
  const deliver = () => {
    for (let rounds = 0; queues.some((queue) => queue.length > 0); rounds += 1) {
      if (rounds > 8) throw new Error('providers kept publishing frames');
      for (const side of [0, 1]) {
        for (const data of queues[side].splice(0)) listeners[1 - side]?.({ type: 'message', data });
      }
    }
  };
  const next = randomEdits(handles, storyId);
  try {
    for (const provider of providers) provider.connect();
    for (const listener of listeners) listener?.({ type: 'open' });
    deliver();
    expect(providers.map((provider) => provider.synced)).toEqual([true, true]);
    for (let edit = 0; edit < 150; edit += 1) {
      next().change();
      if (edit % 3 === 2) {
        deliver();
        expectConverged(handles, storyId);
      }
    }
  } finally {
    for (const provider of providers) provider.destroy();
    for (const handle of handles) handle.dispose();
  }
}, 30_000);

test('deletions held until their structs arrive are not republished', () => {
  const { handles, storyId } = replicas(2, 'ab');
  const [a, b] = handles;
  const net = outboxes(handles);
  try {
    const inserted = net.edit(0, () => a.insertText(storyId, 1, 'xy'));
    const dependent = net.edit(0, () => a.insertText(storyId, 3, 'z'));
    const removal = net.edit(0, () => a.deleteText(storyId, 1, 2));
    for (const message of [removal, dependent, inserted]) {
      expect(net.deliver(message, 1, false).updates).toEqual([]);
    }
    expect(text(b, storyId)).toBe('ayzb');
    expectConverged(handles, storyId);
  } finally {
    net.stop();
    for (const handle of handles) handle.dispose();
  }
});
