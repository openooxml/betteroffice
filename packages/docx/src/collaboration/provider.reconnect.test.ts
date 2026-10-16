import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from '../yrs';
import { createAsyncCollaborationReplica } from './asyncReplica';
import { CollaborationProvider } from './provider';
import { CHUNKED_SYNC_CLIENT, decodeMessages, encodeVarUint, stateVectorDominates } from './protocol';
import type { CollaborationError, CollaborationTransport, CollaborationTransportEvent } from './types';

type Kind = 'sync' | 'async';

let seed: Uint8Array;

beforeAll(async () => {
  preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  const source = await createYrsSession({ clientId: 300 });
  source.loadStories([{ storyId: 'body', paragraphs: [{ text: 'alpha beta' }, { text: 'second paragraph' }] }]);
  seed = source.encodeState();
  source.destroy();
});

function later<T>(operation: () => T): Promise<T> {
  return new Promise((done, fail) => setTimeout(() => {
    try { done(operation()); } catch (cause) { fail(cause); }
  }, 0));
}

async function openPeer(kind: Kind, clientId: number) {
  const session: YrsSession = await createYrsSession({ clientId });
  session.loadState(seed);
  const paragraph = (index: number) => session.paragraphs('body')[index].paraId;
  const replica = kind === 'sync' ? session : createAsyncCollaborationReplica({
    clientId,
    encodeStateVector: () => later(() => session.encodeStateVector()),
    encodeStateAsUpdate: (vector) => later(() => session.encodeStateAsUpdate(vector)),
    applyUpdate: (update) => later(() => session.applyUpdate(update)),
    onUpdate: (listener) => session.onUpdate(listener),
  });
  return {
    session,
    connect: (transport: CollaborationTransport) => new CollaborationProvider(replica, transport),
    edit: (index: number, text: string) => { session.insertText({ story: 'body', paraId: paragraph(index), offset: 0 }, text); },
    erase: (index: number, start: number, end: number) => {
      const paraId = paragraph(index);
      session.deleteRange({ story: 'body', start: { paraId, offset: start }, end: { paraId, offset: end } });
    },
    text: () => session.paragraphs('body').map((entry) => entry.text),
  };
}

function relay(size: number, vectorOf: (index: number) => Promise<Uint8Array>) {
  const listeners = Array.from({ length: size }, () => new Set<(event: CollaborationTransportEvent) => void>());
  const online = new Array<boolean>(size).fill(false);
  const sent: { from: number; data: Uint8Array }[] = [];
  const missing = new Array<number>(size).fill(0);
  let queue: { from: number; to: number; data: Uint8Array }[] = [];
  const setOnline = (index: number, open: boolean) => {
    online[index] = open;
    if (!open) queue = queue.filter((frame) => frame.from !== index && frame.to !== index);
    for (const listener of [...listeners[index]]) listener({ type: open ? 'open' : 'close' });
  };
  const reopen = (...indices: number[]) => {
    for (const index of indices) online[index] = true;
    for (const index of indices) for (const listener of [...listeners[index]]) listener({ type: 'open' });
  };
  const transports = listeners.map((set, from): CollaborationTransport => ({
    connect: () => setOnline(from, true),
    disconnect: () => setOnline(from, false),
    send(data) {
      sent.push({ from, data: data.slice() });
      for (let to = 0; to < size; to += 1) {
        if (to !== from && online[from] && online[to]) queue.push({ from, to, data: data.slice() });
      }
      return true;
    },
    onEvent(listener) {
      set.add(listener);
      return () => { set.delete(listener); };
    },
  }));
  const flush = async ({ to = Array.from({ length: size }, (_, index) => index), twice = false } = {}) => {
    for (let idle = 0; idle < 64;) {
      await new Promise((resolve) => setTimeout(resolve, 0));
      const at = queue.findIndex((frame) => to.includes(frame.to));
      if (at < 0) { idle += 1; continue; }
      idle = 0;
      const [frame] = queue.splice(at, 1);
      const requests = decodeMessages(frame.data)
        .flatMap((message) => message.type === 'sync-step-1' ? [message.stateVector] : []);
      const copies = twice && requests.length ? 2 : 1;
      for (let copy = 0; copy < copies; copy += 1) {
        if (requests.length) {
          const vector = await vectorOf(frame.to);
          missing[frame.to] += requests.filter((request) => !stateVectorDominates(vector, request)).length;
        }
        for (const listener of [...listeners[frame.to]]) listener({ type: 'message', data: frame.data });
      }
    }
  };
  const count = (from: number, type: string, since: number) => sent.slice(since)
    .filter((frame) => frame.from === from)
    .flatMap((frame) => decodeMessages(frame.data))
    .filter((message) => message.type === type).length;
  const vectors = (from: number, since: number) => sent.slice(since)
    .filter((frame) => frame.from === from)
    .flatMap((frame) => decodeMessages(frame.data))
    .flatMap((message) => message.type === 'sync-step-1' ? [clocks(message.stateVector)] : []);
  return { transports, setOnline, reopen, flush, sent, missing, count, vectors };
}

function clocks(vector: Uint8Array): Map<bigint, bigint> {
  const values: bigint[] = [];
  for (let index = 0; index < vector.length;) {
    let value = 0n;
    for (let shift = 0n; ; shift += 7n) {
      const byte = vector[index++];
      value |= BigInt(byte & 0x7f) << shift;
      if (byte < 0x80) break;
    }
    values.push(value);
  }
  return new Map(Array.from({ length: Number(values[0]) }, (_, entry) => [values[2 * entry + 1], values[2 * entry + 2]]));
}

function phantomClockTo(client: number, from: number, to: number): Uint8Array {
  return Uint8Array.from([1, 1, ...encodeVarUint(client), ...encodeVarUint(from), 0, ...encodeVarUint(to - from), 0]);
}

async function network(kind: Kind, size: number) {
  const peers = await Promise.all(Array.from({ length: size }, (_, index) => openPeer(kind, 201 + index)));
  const net = relay(size, async (index) => peers[index].session.encodeStateVector());
  const providers = peers.map((peer, index) => peer.connect(net.transports[index]));
  const errors: CollaborationError[] = [];
  for (const provider of providers) provider.onError((error) => { errors.push(error); });
  for (const provider of providers) provider.connect();
  await net.flush();
  return {
    peers, net, errors,
    close() {
      for (const provider of providers) provider.destroy();
      for (const peer of peers) peer.session.destroy();
    },
  };
}

for (const kind of ['sync', 'async'] as const) {
  describe(`${kind} replicas reconnecting over a peer transport`, () => {
    it('publish offline edits to a peer that saw no open', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        net.setOnline(0, false);
        left.edit(0, 'offline ');
        net.setOnline(0, true);
        await net.flush();
        expect(right.text()).toEqual(left.text());
        expect(right.text()[0]).toBe('offline alpha beta');
      } finally { close(); }
    });

    it('converge three replicas on a relay after one reconnects', async () => {
      const { peers, net, close } = await network(kind, 3);
      try {
        net.setOnline(2, false);
        peers[2].edit(0, 'offline ');
        peers[0].edit(1, 'online ');
        await net.flush();
        const mark = net.sent.length;
        net.setOnline(2, true);
        await net.flush();
        expect(peers[1].text()).toEqual(peers[0].text());
        expect(peers[2].text()).toEqual(peers[0].text());
        expect(peers[0].text()).toEqual(['offline alpha beta', 'online second paragraph']);
        expect([0, 1, 2].map((index) => net.count(index, 'sync-step-1', mark))).toEqual([1, 1, 1]);
      } finally { close(); }
    });

    it('request missing state once and go quiet', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        let mark = net.sent.length;
        await net.flush();
        expect(net.sent.length).toBe(mark);

        net.setOnline(0, false);
        net.setOnline(0, true);
        await net.flush();
        expect([net.count(0, 'sync-step-1', mark), net.count(1, 'sync-step-1', mark)]).toEqual([1, 1]);

        mark = net.sent.length;
        net.setOnline(0, false);
        left.edit(0, 'offline ');
        net.setOnline(0, true);
        await net.flush();
        expect([net.count(0, 'sync-step-1', mark), net.count(1, 'sync-step-1', mark)]).toEqual([1, 1]);
        expect([net.count(0, 'sync-step-2', mark), net.count(1, 'sync-step-2', mark)]).toEqual([1, 1]);
        expect(right.text()).toEqual(left.text());

        mark = net.sent.length;
        await net.flush();
        expect(net.sent.length).toBe(mark);
      } finally { close(); }
    });

    it('spend one extra exchange per side when both reopen diverged, then go quiet', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        net.setOnline(0, false);
        net.setOnline(1, false);
        left.edit(0, 'left ');
        right.edit(1, 'right ');
        await net.flush();
        const mark = net.sent.length;
        net.reopen(0, 1);
        await net.flush();
        expect(right.text()).toEqual(left.text());
        expect(left.text()).toEqual(['left alpha beta', 'right second paragraph']);
        expect([0, 1].map((index) => [net.count(index, 'sync-step-1', mark), net.count(index, 'sync-step-2', mark)]))
          .toEqual([[2, 2], [2, 2]]);
        const settled = net.sent.length;
        await net.flush();
        expect(net.sent.length).toBe(settled);
      } finally { close(); }
    });

    it('publish offline deletions to a peer that saw no open', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        net.setOnline(0, false);
        const vector = left.session.encodeStateVector();
        left.erase(0, 0, 6);
        expect(left.session.encodeStateVector()).toEqual(vector);
        net.setOnline(0, true);
        await net.flush();
        expect(right.text()).toEqual(left.text());
        expect(right.text()[0]).toBe('beta');
      } finally { close(); }
    });

    it('publish an offline deletion at the native clock limit', async () => {
      const { peers: [left, right], net, errors, close } = await network(kind, 2);
      try {
        const own = left.session.clientId;
        const from = Number(clocks(left.session.encodeStateVector()).get(BigInt(own)) ?? 0n);
        const filler = phantomClockTo(own, from, 0xffff_ffff);
        for (const peer of [left, right]) peer.session.applyUpdate(filler.slice());
        await net.flush();
        expect(clocks(right.session.encodeStateVector()).get(BigInt(own))).toBe(0xffff_ffffn);
        net.setOnline(0, false);
        left.erase(0, 0, 6);
        const mark = net.sent.length;
        net.setOnline(0, true);
        await net.flush();
        expect(net.vectors(0, mark).map((vector) => vector.get(BigInt(own)))).toEqual([0xffff_ffffn]);
        expect(errors).toEqual([]);
        expect(right.text()).toEqual(left.text());
        expect(right.text()[0]).toBe('beta');
      } finally { close(); }
    });

    it('publish offline deletions to every replica on a relay', async () => {
      const { peers, net, close } = await network(kind, 3);
      try {
        peers[2].edit(1, 'cleared ');
        await net.flush();
        expect(peers[0].text()[1]).toBe('cleared second paragraph');
        net.setOnline(2, false);
        peers[2].erase(1, 0, 8);
        peers[2].erase(0, 0, 6);
        net.setOnline(2, true);
        await net.flush();
        expect(peers[1].text()).toEqual(peers[0].text());
        expect(peers[2].text()).toEqual(peers[0].text());
        expect(peers[0].text()).toEqual(['beta', 'second paragraph']);
      } finally { close(); }
    });

    it('send at most one reverse request per newer state received, then go quiet', async () => {
      const { peers, net, close } = await network(kind, 3);
      try {
        net.setOnline(2, false);
        peers[2].edit(0, 'offline ');
        peers[0].edit(1, 'online ');
        await net.flush();
        const mark = net.sent.length;
        net.missing.fill(0);
        net.setOnline(2, true);
        await net.flush({ twice: true });
        expect(peers[1].text()).toEqual(peers[0].text());
        expect(peers[2].text()).toEqual(peers[0].text());
        expect(peers[0].text()).toEqual(['offline alpha beta', 'online second paragraph']);
        const reverse = [0, 1, 2].map((index) => net.count(index, 'sync-step-1', mark) - (index === 2 ? 1 : 0));
        expect(reverse.slice(0, 2).every((requests) => requests > 0)).toBe(true);
        expect(reverse.map((requests, index) => Math.max(0, requests - net.missing[index]))).toEqual([0, 0, 0]);
        const settled = net.sent.length;
        await net.flush({ twice: true });
        expect(net.sent.length).toBe(settled);
      } finally { close(); }
    });

    it('converge with a relay peer that leaves mid-handshake and rejoins unchanged', async () => {
      const { peers, net, close } = await network(kind, 3);
      try {
        net.setOnline(2, false);
        peers[2].edit(0, 'offline ');
        net.setOnline(2, true);
        await net.flush({ to: [0, 1] });
        net.setOnline(2, false);
        net.setOnline(2, true);
        await net.flush();
        expect(peers[1].text()).toEqual(peers[0].text());
        expect(peers[2].text()).toEqual(peers[0].text());
        expect(peers[0].text()[0]).toBe('offline alpha beta');
      } finally { close(); }
    });
  });
}

describe('a replica holding its initial sync', () => {
  const sorted = (vector: Map<bigint, bigint>) => [...vector].sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));

  it('opens with the chunked sync marker next to the local clock ahead, or the real clock at the native limit', async () => {
    const { peers: [left, right], net, errors, close } = await network('sync', 2);
    try {
      const own = BigInt(left.session.clientId);
      const marker: [bigint, bigint] = [BigInt(CHUNKED_SYNC_CLIENT), 0n];
      const reopen = async () => {
        net.setOnline(0, false);
        const before = clocks(left.session.encodeStateVector());
        const mark = net.sent.length;
        net.setOnline(0, true);
        await net.flush();
        return { before, opened: net.vectors(0, mark)[0], updates: net.count(0, 'update', mark) };
      };

      const ahead = await reopen();
      expect(sorted(ahead.opened))
        .toEqual(sorted(new Map([...ahead.before, [own, (ahead.before.get(own) ?? 0n) + 1n], marker])));
      expect(ahead.updates).toBe(0);

      const filler = phantomClockTo(Number(own), Number(ahead.before.get(own) ?? 0n), 0xffff_ffff);
      for (const peer of [left, right]) peer.session.applyUpdate(filler.slice());
      await net.flush();
      const limit = await reopen();
      expect(limit.before.get(own)).toBe(0xffff_ffffn);
      expect(sorted(limit.opened)).toEqual(sorted(new Map([...limit.before, marker])));
      expect(limit.updates).toBe(1);
      expect(errors).toEqual([]);
      expect(right.text()).toEqual(left.text());
    } finally { close(); }
  });

  it.todo('first joiner in a two-peer relay leaves the chunked hold (deferred: no peer answers its open, and its request waits for that answer)', async () => {
    const { peers: [first, second], net, errors, close } = await network('sync', 2);
    try {
      second.edit(0, 'live ');
      await net.flush();
      expect(errors).toEqual([]);
      expect(first.text()).toEqual(second.text());
    } finally { close(); }
  });
});
