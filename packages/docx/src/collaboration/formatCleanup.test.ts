import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createEditSession, preloadEditWasm, type EditSession } from '../wasm/edit';
import { createYrsSession, type YrsInlineFormatDelta, type YrsSession } from '../yrs';
import { CollaborationProvider } from './provider';
import type { CollaborationTransportEvent } from './types';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))));

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
      const last = merged.at(-1);
      if (last && start <= last[1]) last[1] = Math.max(last[1], end);
      else merged.push([start, end]);
    }
    return [client, merged];
  });
}

/** A replica's whole delete set: a diff against its own state vector holds nothing else. */
function deleteSet(session: YrsSession): DeleteSet {
  return deleteOnly(session.encodeStateAsUpdate(session.encodeStateVector()));
}

function clocks(set: DeleteSet): string[] {
  return set.flatMap(([client, ranges]) => ranges.flatMap(([start, end]) =>
    Array.from({ length: end - start }, (_, offset) => `${client}:${start + offset}`)));
}

async function replicas(count: number, text: string): Promise<YrsSession[]> {
  const seed = await createYrsSession({ clientId: 1001 });
  seed.loadStories([{ storyId: 'body', paragraphs: [{ text }] }]);
  const state = seed.encodeState();
  seed.destroy();
  const sessions: YrsSession[] = [];
  for (let index = 0; index < count; index += 1) {
    const session = await createYrsSession({ clientId: 1002 + index });
    session.loadState(state);
    sessions.push(session);
  }
  return sessions;
}

function paragraph(session: YrsSession) {
  return session.paragraphs('body')[0];
}

function format(session: YrsSession, start: number, end: number, delta: YrsInlineFormatDelta) {
  const { paraId } = paragraph(session);
  session.formatRange(
    { story: 'body', start: { paraId, offset: start }, end: { paraId, offset: end } },
    delta
  );
}

/** Every replica forwards only its 'local' updates to every other replica, as the provider does. */
function mesh(sessions: YrsSession[]) {
  const inboxes = sessions.map(() => [] as Uint8Array[]);
  const sent: Uint8Array[] = [];
  const stops = sessions.map((session, from) => session.onUpdate((update, origin) => {
    if (origin !== 'local') return;
    sent.push(update.slice());
    inboxes.forEach((inbox, to) => { if (to !== from) inbox.push(update.slice()); });
  }));
  const settle = (): number => {
    let rounds = 0;
    while (inboxes.some((inbox) => inbox.length > 0)) {
      rounds += 1;
      if (rounds > 8) throw new Error('replicas kept publishing updates');
      sessions.forEach((session, to) => {
        for (const update of inboxes[to].splice(0)) session.applyUpdate(update);
      });
    }
    return rounds;
  };
  return { settle, sent, stop: () => stops.forEach((stop) => stop()) };
}

function expectConverged(sessions: YrsSession[]) {
  const [first, ...rest] = sessions;
  for (const session of rest) {
    expect(session.encodeStateVector()).toEqual(first.encodeStateVector());
    expect(paragraph(session).text).toBe(paragraph(first).text);
    expect(session.storySegments('body')).toEqual(first.storySegments('body'));
    expect(session.storyChecksum('body')).toBe(first.storyChecksum('body'));
    expect(deleteSet(session)).toEqual(deleteSet(first));
  }
}

test('two replicas converge when replaying a peer update cleans up formatting', async () => {
  const sessions = await replicas(2, 'a');
  const [a, b] = sessions;
  const network = mesh(sessions);
  try {
    format(b, 0, 1, { bold: true });
    format(b, 0, 1, { bold: null });
    format(a, 0, 1, { bold: true });
    network.settle();
    expectConverged(sessions);
  } finally {
    network.stop();
    for (const session of sessions) session.destroy();
  }
});

test('a cleanup made by a relaying replica reaches the replica behind it', async () => {
  const sessions = await replicas(3, 'ab');
  const [a, b, c] = sessions;
  const toA: Uint8Array[] = [];
  const toC: Uint8Array[] = [];
  const fromA: Uint8Array[] = [];
  const fromC: Uint8Array[] = [];
  const stops = [
    a.onUpdate((update, origin) => { if (origin === 'local') fromA.push(update.slice()); }),
    c.onUpdate((update, origin) => { if (origin === 'local') fromC.push(update.slice()); }),
    b.onUpdate((update, origin) => {
      if (origin !== 'local') return;
      toA.push(update.slice());
      toC.push(update.slice());
    }),
  ];
  const settle = () => {
    for (let rounds = 0; fromA.length + fromC.length + toA.length + toC.length > 0; rounds += 1) {
      if (rounds > 8) throw new Error('replicas kept publishing updates');
      for (const update of fromA.splice(0)) { b.applyUpdate(update); toC.push(update); }
      for (const update of fromC.splice(0)) { b.applyUpdate(update); toA.push(update); }
      for (const update of toA.splice(0)) a.applyUpdate(update);
      for (const update of toC.splice(0)) c.applyUpdate(update);
    }
  };
  try {
    format(b, 0, 1, { bold: true });
    format(b, 0, 1, { bold: null });
    format(a, 0, 1, { bold: true });
    format(c, 0, 2, { italic: true });
    settle();
    expectConverged(sessions);
  } finally {
    for (const stop of stops) stop();
    for (const session of sessions) session.destroy();
  }
});

test('published cleanups go quiet and replaying them changes nothing', async () => {
  const sessions = await replicas(2, 'abc');
  const [a, b] = sessions;
  const network = mesh(sessions);
  try {
    format(b, 0, 2, { bold: true });
    format(b, 1, 3, { bold: null });
    format(a, 0, 3, { bold: true });
    format(a, 1, 2, { italic: true });
    expect(network.settle()).toBeLessThanOrEqual(3);
    expectConverged(sessions);
    const published = network.sent.length;
    const checksum = a.storyChecksum('body');
    for (const update of network.sent.slice()) {
      a.applyUpdate(update);
      b.applyUpdate(update);
    }
    expect(network.settle()).toBe(0);
    expect(network.sent.length).toBe(published);
    expect(a.storyChecksum('body')).toBe(checksum);
    expectConverged(sessions);
  } finally {
    network.stop();
    for (const session of sessions) session.destroy();
  }
});

test('seeded concurrent edits converge through incremental updates alone', async () => {
  const sessions = await replicas(2, 'abcdefghij');
  const network = mesh(sessions);
  let seed = 0x5eed;
  const random = () => {
    seed = (seed * 1_103_515_245 + 12_345) % 2_147_483_648;
    return seed / 2_147_483_648;
  };
  const pick = (limit: number) => Math.floor(random() * limit);
  try {
    for (let edit = 0; edit < 600; edit += 1) {
      const session = sessions[pick(sessions.length)];
      const length = paragraph(session).text.length;
      const roll = random();
      if (roll < 0.15 || length < 3) {
        session.insertText({ story: 'body', paraId: paragraph(session).paraId, offset: pick(length + 1) }, 'xy');
      } else if (roll < 0.25) {
        const start = pick(length - 1);
        const { paraId } = paragraph(session);
        session.deleteRange({
          story: 'body',
          start: { paraId, offset: start },
          end: { paraId, offset: start + 1 },
        });
      } else {
        const start = pick(length - 1);
        const end = start + 1 + pick(length - start);
        const key = random() < 0.6 ? 'bold' : 'italic';
        format(session, start, end, { [key]: random() < 0.5 ? true : null });
      }
      if (edit % 3 === 2) {
        network.settle();
        expectConverged(sessions);
      }
    }
  } finally {
    network.stop();
    for (const session of sessions) session.destroy();
  }
});

test('a provider pair converges after a formatting cleanup', async () => {
  const sessions = await replicas(2, 'a');
  const [a, b] = sessions;
  const queues: Uint8Array[][] = [[], []];
  const listeners: Array<((event: CollaborationTransportEvent) => void) | undefined> = [];
  const providers = sessions.map((session, side) => new CollaborationProvider(session, {
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
  try {
    for (const provider of providers) provider.connect();
    for (const listener of listeners) listener?.({ type: 'open' });
    deliver();
    expect(providers.map((provider) => provider.synced)).toEqual([true, true]);
    format(b, 0, 1, { bold: true });
    format(b, 0, 1, { bold: null });
    format(a, 0, 1, { bold: true });
    deliver();
    expectConverged(sessions);
  } finally {
    for (const provider of providers) provider.destroy();
    for (const session of sessions) session.destroy();
  }
});

test('a remote apply publishes exactly the deletions it made beyond the update', async () => {
  const sessions = await replicas(2, 'a');
  const [a, b] = sessions;
  const local = sessions.map(() => [] as Uint8Array[]);
  const stops = sessions.map((session, index) => session.onUpdate((update, origin) => {
    if (origin === 'local') local[index].push(update.slice());
  }));
  type Sent = { updates: Uint8Array[]; deleted: Set<string> };
  const edit = (index: number, change: () => void): Sent => {
    const before = new Set(clocks(deleteSet(sessions[index])));
    change();
    const deleted = clocks(deleteSet(sessions[index])).filter((id) => !before.has(id));
    return { updates: local[index].splice(0), deleted: new Set(deleted) };
  };
  const deliver = (sent: Sent, to: number): Sent => {
    const before = new Set(clocks(deleteSet(sessions[to])));
    for (const update of sent.updates) sessions[to].applyUpdate(update);
    const updates = local[to].splice(0);
    const published = updates.flatMap((update) => clocks(deleteOnly(update)));
    const expected = clocks(deleteSet(sessions[to]))
      .filter((id) => !before.has(id) && !sent.deleted.has(id));
    expect(published.sort()).toEqual(expected.sort());
    return { updates, deleted: new Set(published) };
  };
  try {
    let queue: Array<[Sent, number]> = [
      [edit(1, () => format(b, 0, 1, { bold: true })), 0],
      [edit(1, () => format(b, 0, 1, { bold: null })), 0],
      [edit(0, () => format(a, 0, 1, { bold: true })), 1],
    ];
    let published = 0;
    for (let rounds = 0; queue.length > 0; rounds += 1) {
      if (rounds > 8) throw new Error('replicas kept publishing updates');
      queue = queue.flatMap(([sent, to]): Array<[Sent, number]> => {
        const reply = deliver(sent, to);
        published += reply.updates.length;
        return reply.updates.length > 0 ? [[reply, 1 - to]] : [];
      });
    }
    expect(published).toBeGreaterThan(0);
    expectConverged(sessions);
  } finally {
    for (const stop of stops) stop();
    for (const session of sessions) session.destroy();
  }
});

test('the polling observer queues the cleanup as a local delete-only event', () => {
  const [a, b] = [createEditSession(1101), createEditSession(1102)];
  try {
    const receipt = JSON.parse(a.load_json(JSON.stringify([
      { storyId: 'body', paragraphs: [{ text: 'a' }] },
    ]))) as { body: string[] };
    const [paraId] = receipt.body;
    b.load(a.encode_state());
    a.start_update_event_observation();
    b.start_update_event_observation();
    const bold = (session: EditSession, value: true | null) =>
      session.format_range('body', paraId, 0, paraId, 1, JSON.stringify({ bold: value }));
    const outboxes = new Map<EditSession, Uint8Array[]>([[a, []], [b, []]]);
    const drain = (session: EditSession): Uint8Array[] => {
      const local: Uint8Array[] = [];
      for (let event = session.drain_update_event(); event.length > 0;) {
        if (event[0] === 0) local.push(event.slice(1));
        event = session.drain_update_event();
      }
      return local;
    };
    bold(b, true);
    bold(b, null);
    bold(a, true);
    for (const session of [a, b]) outboxes.get(session)!.push(...drain(session));
    let published = 0;
    for (let rounds = 0; [...outboxes.values()].some((outbox) => outbox.length > 0); rounds += 1) {
      if (rounds > 8) throw new Error('replicas kept publishing updates');
      for (const [from, to] of [[a, b], [b, a]] as Array<[EditSession, EditSession]>) {
        for (const update of outboxes.get(from)!.splice(0)) {
          to.apply_update(update);
          const cleanup = drain(to);
          for (const deletion of cleanup) deleteOnly(deletion);
          published += cleanup.length;
          outboxes.get(to)!.push(...cleanup);
        }
      }
    }
    expect(published).toBeGreaterThan(0);
    expect([...b.encode_state_vector()]).toEqual([...a.encode_state_vector()]);
    expect(deleteOnly(b.encode_diff(b.encode_state_vector())))
      .toEqual(deleteOnly(a.encode_diff(a.encode_state_vector())));
    expect(b.story_checksum('body')).toBe(a.story_checksum('body'));
  } finally {
    a.free();
    b.free();
  }
});

test('deletions held until their structs arrive are not republished', async () => {
  const sessions = await replicas(2, 'ab');
  const [a, b] = sessions;
  const sent: Uint8Array[][] = [];
  let published = 0;
  const stops = [
    a.onUpdate((update, origin) => { if (origin === 'local') sent.at(-1)?.push(update.slice()); }),
    b.onUpdate((_update, origin) => { if (origin === 'local') published += 1; }),
  ];
  const edit = (change: () => void) => {
    sent.push([]);
    change();
  };
  try {
    const { paraId } = paragraph(a);
    edit(() => a.insertText({ story: 'body', paraId, offset: 1 }, 'xy'));
    edit(() => a.insertText({ story: 'body', paraId, offset: 3 }, 'z'));
    edit(() => a.deleteRange({
      story: 'body',
      start: { paraId, offset: 1 },
      end: { paraId, offset: 2 },
    }));
    const [inserted, dependent, removal] = sent;
    for (const update of [...removal, ...dependent, ...inserted]) b.applyUpdate(update);
    expect(paragraph(b).text).toBe('ayzb');
    expect(published).toBe(0);
    expectConverged(sessions);
  } finally {
    for (const stop of stops) stop();
    for (const session of sessions) session.destroy();
  }
});

test('a held worker deletion is published once its struct arrives, a held peer one is not', async () => {
  const sessions = await replicas(3, 'ab');
  const [a, b, worker] = sessions;
  const sent: Uint8Array[][] = [];
  const published: Uint8Array[] = [];
  const stops = [
    ...[a, worker].map((session) => session.onUpdate((update, origin) => {
      if (origin === 'local') sent.at(-1)?.push(update.slice());
    })),
    b.onUpdate((update, origin) => { if (origin === 'local') published.push(update.slice()); }),
  ];
  const edit = (change: () => void) => {
    sent.push([]);
    change();
  };
  try {
    const { paraId } = paragraph(a);
    edit(() => a.insertText({ story: 'body', paraId, offset: 1 }, 'xyz'));
    for (const update of sent[0]) worker.applyUpdate(update);
    edit(() => a.deleteRange({
      story: 'body',
      start: { paraId, offset: 3 },
      end: { paraId, offset: 4 },
    }));
    edit(() => worker.deleteRange({
      story: 'body',
      start: { paraId, offset: 1 },
      end: { paraId, offset: 2 },
    }));
    const [inserted, peerRemoval, workerRemoval] = sent;
    for (const update of peerRemoval) b.applyUpdate(update);
    for (const update of workerRemoval) b.applyLocalUpdate(update);
    for (const update of inserted) b.applyUpdate(update);
    expect(paragraph(b).text).toBe('ayb');
    expect(published.map((update) => clocks(deleteOnly(update)).length)).toEqual([1]);
    for (const update of published) a.applyUpdate(update);
    expectConverged([a, b]);
  } finally {
    for (const stop of stops) stop();
    for (const session of sessions) session.destroy();
  }
});
