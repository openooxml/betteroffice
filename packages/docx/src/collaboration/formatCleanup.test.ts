import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsInlineFormatDelta, type YrsSession } from '../yrs';
import { CollaborationProvider } from './provider';
import type { CollaborationTransportEvent } from './types';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))));

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
