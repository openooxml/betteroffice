import { describe, expect, it } from 'bun:test';
import { openWorkbook } from '../wasm/loader';
import { createTestWorkbookSession, loadWorkbookSessionFixtures } from '../session/testHelpers';
import { asyncReplicaFor } from './asyncReplica';
import { CollaborationProvider } from './provider';
import {
  decodeMessages,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeVarUint,
  stateVectorDominates,
  withoutDocumentFingerprint,
} from './protocol';
import type { CollaborationReplica, CollaborationTransport, CollaborationTransportEvent } from './types';

type Kind = 'sync' | 'async';

interface Peer {
  replica: CollaborationReplica;
  edit(col: number, input: string): Promise<void>;
  undo(): Promise<void>;
  read(from?: number): Promise<string[]>;
  vector(): Promise<Uint8Array>;
  dispose(): unknown;
}

const fixtures = loadWorkbookSessionFixtures();

async function openPeer(kind: Kind, clientId: number): Promise<Peer> {
  const { fixture } = await fixtures;
  if (kind === 'sync') {
    const handle = openWorkbook(fixture, { collaborative: true, clientId });
    return {
      replica: handle,
      edit: async (col, input) => { expect(handle.editCell(0, 2, col + 1, input).applied).toBe(true); },
      undo: async () => { expect(handle.undo().applied).toBe(true); },
      read: async (from = 0) => [1, 2].map((col) => handle.cell(0, 2, from + col).input),
      vector: async () => handle.encodeStateVector(),
      dispose: () => handle.dispose(),
    };
  }
  const session = await createTestWorkbookSession(fixture, undefined, { collaborative: true, clientId });
  const applied = async (ticket: { ack: Promise<{ status: string }> }) => {
    expect((await ticket.ack).status).toBe('applied');
  };
  return {
    replica: session.replica!,
    edit: (col, input) => applied(session.editSession!.submit('editCell', { payload: { args: [0, 2, col + 1, input] } })),
    undo: () => applied(session.editSession!.submit('undo', { payload: { args: [] } })),
    read: async (from = 0) => {
      const range = `${String.fromCharCode(66 + from)}3:${String.fromCharCode(67 + from)}3`;
      return (await session.call.cellInputs(0, range)).cells[0].map((cell) => cell.input);
    },
    vector: () => asyncReplicaFor(session.replica!)!.encodeStateVector(),
    dispose: () => session.dispose(),
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
  return { transports, setOnline, reopen, flush, sent, missing, count };
}

function connectLegacyPeer(replica: CollaborationReplica, transport: CollaborationTransport, errors: string[]) {
  const target = asyncReplicaFor(replica) ?? replica;
  let tail = Promise.resolve();
  transport.onEvent((event) => {
    tail = tail.then(async () => {
      if (event.type === 'open') transport.send(encodeSyncStep1(await target.encodeStateVector()));
      if (event.type !== 'message') return;
      for (const message of decodeMessages(event.data)) {
        if (message.type === 'sync-step-1') {
          const vector = withoutDocumentFingerprint(message.stateVector).slice();
          transport.send(encodeSyncStep2(await target.encodeStateAsUpdate(vector)));
        } else if (message.type === 'sync-step-2' || message.type === 'update') {
          await target.applyUpdate(withoutDocumentFingerprint(message.update).slice());
        }
      }
    }).catch((cause) => { errors.push(`legacy: ${String(cause)}`); });
  });
  transport.connect();
}

function clients(vector: Uint8Array): bigint[] {
  const values: bigint[] = [];
  const bytes = withoutDocumentFingerprint(vector);
  for (let index = 0; index < bytes.length;) {
    let value = 0n;
    for (let shift = 0n; ; shift += 7n) {
      const byte = bytes[index++];
      value |= BigInt(byte & 0x7f) << shift;
      if (byte < 0x80) break;
    }
    values.push(value);
  }
  return values.filter((_, index) => index % 2 === 1);
}

function phantomClients(count: number): Uint8Array {
  const bytes = [...encodeVarUint(count)];
  for (let client = 1; client <= count; client += 1) bytes.push(1, ...encodeVarUint(1_000_000 + client), 0, 0, 1);
  bytes.push(0);
  return Uint8Array.from(bytes);
}

async function network(kind: Kind, size: number, legacy?: number) {
  const peers = await Promise.all(Array.from({ length: size }, (_, index) => openPeer(kind, 101 + index)));
  const net = relay(size, (index) => peers[index].vector());
  const errors: string[] = [];
  const providers = peers.flatMap((peer, index) => index === legacy
    ? [] : [new CollaborationProvider(peer.replica, net.transports[index])]);
  for (const provider of providers) {
    provider.onError((error) => errors.push(`${error.code}: ${error.message}`));
    provider.connect();
  }
  if (legacy !== undefined) connectLegacyPeer(peers[legacy].replica, net.transports[legacy], errors);
  await net.flush();
  return {
    peers, net, providers, errors,
    async close() {
      for (const provider of providers) provider.destroy();
      await Promise.all(peers.map((peer) => peer.dispose()));
    },
  };
}

for (const kind of ['sync', 'async'] as const) {
  describe(`${kind} replicas reconnecting over a peer transport`, () => {
    it('publish offline edits to a peer that saw no open', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        net.setOnline(0, false);
        await left.edit(0, 'offline');
        net.setOnline(0, true);
        await net.flush();
        expect(await right.read()).toEqual(await left.read());
        expect((await right.read())[0]).toBe('offline');
      } finally { await close(); }
    });

    it('keep edits from both sides of a split', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        net.setOnline(0, false);
        await left.edit(0, 'left offline');
        await right.edit(1, 'right online');
        await net.flush();
        net.setOnline(0, true);
        await net.flush();
        const cells = await Promise.all([left.read(), right.read()]);
        expect(cells[1]).toEqual(cells[0]);
        expect(cells[0].slice(0, 2)).toEqual(['left offline', 'right online']);
      } finally { await close(); }
    });

    it('converge three replicas on a relay after one reconnects', async () => {
      const { peers, net, close } = await network(kind, 3);
      try {
        net.setOnline(2, false);
        await peers[2].edit(0, 'offline');
        await peers[0].edit(1, 'online');
        await net.flush();
        const mark = net.sent.length;
        net.setOnline(2, true);
        await net.flush();
        const cells = await Promise.all(peers.map((peer) => peer.read()));
        expect(cells[1]).toEqual(cells[0]);
        expect(cells[2]).toEqual(cells[0]);
        expect(cells[0].slice(0, 2)).toEqual(['offline', 'online']);
        expect([0, 1, 2].map((index) => net.count(index, 'sync-step-1', mark))).toEqual([1, 1, 1]);
      } finally { await close(); }
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
        await left.edit(0, 'offline');
        net.setOnline(0, true);
        await net.flush();
        expect([net.count(0, 'sync-step-1', mark), net.count(1, 'sync-step-1', mark)]).toEqual([1, 1]);
        expect([net.count(0, 'sync-step-2', mark), net.count(1, 'sync-step-2', mark)]).toEqual([1, 1]);
        expect(await right.read()).toEqual(await left.read());

        mark = net.sent.length;
        await net.flush();
        expect(net.sent.length).toBe(mark);
      } finally { await close(); }
    });

    it('spend one extra exchange per side when both reopen diverged, then go quiet', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        net.setOnline(0, false);
        net.setOnline(1, false);
        await left.edit(0, 'left offline');
        await right.edit(1, 'right offline');
        await net.flush();
        const mark = net.sent.length;
        net.reopen(0, 1);
        await net.flush();
        const cells = await Promise.all([left.read(), right.read()]);
        expect(cells[1]).toEqual(cells[0]);
        expect(cells[0]).toEqual(['left offline', 'right offline']);
        expect([0, 1].map((index) => [net.count(index, 'sync-step-1', mark), net.count(index, 'sync-step-2', mark)]))
          .toEqual([[2, 2], [2, 2]]);
        const settled = net.sent.length;
        await net.flush();
        expect(net.sent.length).toBe(settled);
      } finally { await close(); }
    });


    it('publish offline deletions to a peer that saw no open', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        await left.edit(4, 'cleared');
        await left.edit(5, 'undone');
        await net.flush();
        expect(await right.read(4)).toEqual(['cleared', 'undone']);
        net.setOnline(0, false);
        const vector = await left.vector();
        await left.undo();
        expect(await left.vector()).toEqual(vector);
        await left.edit(4, '');
        expect(await left.vector()).toEqual(vector);
        net.setOnline(0, true);
        await net.flush();
        const cells = await Promise.all([left.read(4), right.read(4)]);
        expect(cells[1]).toEqual(cells[0]);
        expect(cells[0].filter(Boolean)).toEqual([]);
      } finally { await close(); }
    });

    it('publish offline deletions to every replica on a relay', async () => {
      const { peers, net, close } = await network(kind, 3);
      try {
        await peers[2].edit(4, 'cleared');
        await peers[2].edit(5, 'undone');
        await net.flush();
        net.setOnline(2, false);
        await peers[2].undo();
        await peers[2].edit(4, '');
        net.setOnline(2, true);
        await net.flush();
        const cells = await Promise.all(peers.map((peer) => peer.read(4)));
        expect(cells[1]).toEqual(cells[0]);
        expect(cells[2]).toEqual(cells[0]);
        expect(cells[0].filter(Boolean)).toEqual([]);
      } finally { await close(); }
    });

    it('publish an offline deletion at the native state vector entry limit', async () => {
      const { peers: [left, right], net, errors, close } = await network(kind, 2);
      try {
        await right.edit(4, 'cleared');
        await net.flush();
        const filler = phantomClients(65_536 - clients(await left.vector()).length);
        for (const peer of [left, right]) await peer.replica.applyUpdate(filler.slice());
        const vector = await left.vector();
        expect(clients(vector)).toHaveLength(65_536);
        expect(clients(vector)).not.toContain(101n);
        net.setOnline(0, false);
        await left.edit(4, '');
        expect(await left.vector()).toEqual(vector);
        net.setOnline(0, true);
        await net.flush();
        expect(errors).toEqual([]);
        expect(await right.read(4)).toEqual(await left.read(4));
        expect((await right.read(4))[0]).toBe('');
      } finally { await close(); }
    });

    it('send at most one reverse request per newer state received, then go quiet', async () => {
      const { peers, net, close } = await network(kind, 3);
      try {
        net.setOnline(2, false);
        await peers[2].edit(0, 'offline');
        await peers[0].edit(1, 'online');
        await net.flush();
        const mark = net.sent.length;
        net.missing.fill(0);
        net.setOnline(2, true);
        await net.flush({ twice: true });
        const cells = await Promise.all(peers.map((peer) => peer.read()));
        expect(cells[1]).toEqual(cells[0]);
        expect(cells[2]).toEqual(cells[0]);
        expect(cells[0]).toEqual(['offline', 'online']);
        const reverse = [0, 1, 2].map((index) => net.count(index, 'sync-step-1', mark) - (index === 2 ? 1 : 0));
        expect(reverse.slice(0, 2).every((requests) => requests > 0)).toBe(true);
        expect(reverse.map((requests, index) => Math.max(0, requests - net.missing[index]))).toEqual([0, 0, 0]);
        const settled = net.sent.length;
        await net.flush({ twice: true });
        expect(net.sent.length).toBe(settled);
      } finally { await close(); }
    });

    it('converge after an offline overwrite of existing cells without dropping the link', async () => {
      const { peers: [left, right], net, providers, errors, close } = await network(kind, 2);
      try {
        await left.edit(0, 'before');
        await right.edit(1, 'theirs');
        await net.flush();
        net.setOnline(0, false);
        await left.edit(0, 'after');
        await left.edit(1, 'mine');
        net.setOnline(0, true);
        await net.flush();
        expect(errors).toEqual([]);
        expect(providers.map((provider) => [provider.status, provider.synced]))
          .toEqual([['connected', true], ['connected', true]]);
        expect(await right.read()).toEqual(['after', 'mine']);
        expect(await left.read()).toEqual(['after', 'mine']);
      } finally { await close(); }
    });

    it('spend exactly one extra exchange per open, then go quiet', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        await left.edit(4, 'cleared');
        await net.flush();
        const opens: number[][] = [];
        for (const change of [async () => {}, () => left.edit(4, ''), () => left.edit(0, 'offline')]) {
          const mark = net.sent.length;
          net.setOnline(0, false);
          await change();
          net.setOnline(0, true);
          await net.flush();
          opens.push([0, 1].flatMap((index) => ['sync-step-1', 'sync-step-2'].map((type) => net.count(index, type, mark))));
          const settled = net.sent.length;
          await net.flush();
          expect(net.sent.length).toBe(settled);
        }
        expect(opens).toEqual([[1, 1, 1, 1], [1, 1, 1, 1], [1, 1, 1, 1]]);
        expect(await right.read(4)).toEqual(await left.read(4));
        expect(await right.read()).toEqual(await left.read());
      } finally { await close(); }
    });

    it('stay valid with a peer that answers without requesting state back', async () => {
      const { peers: [left, right], net, providers: [provider], errors, close } = await network(kind, 2, 1);
      try {
        net.setOnline(0, false);
        net.setOnline(1, false);
        await left.edit(0, 'ours offline');
        await right.edit(1, 'theirs offline');
        const mark = net.sent.length;
        net.reopen(0, 1);
        await net.flush();
        expect(errors).toEqual([]);
        expect([provider.status, provider.synced]).toEqual(['connected', true]);
        expect(net.count(1, 'sync-step-1', mark)).toBe(1);
        expect(net.count(0, 'sync-step-2', mark)).toBeGreaterThan(0);
        const cells = await Promise.all([left.read(), right.read()]);
        expect(cells[1]).toEqual(cells[0]);
        expect(cells[0]).toEqual(['ours offline', 'theirs offline']);
      } finally { await close(); }
    });

    it('converge with a relay peer that leaves mid-handshake and rejoins unchanged', async () => {
      const { peers, net, close } = await network(kind, 3);
      try {
        net.setOnline(2, false);
        await peers[2].edit(0, 'offline');
        net.setOnline(2, true);
        await net.flush({ to: [0, 1] });
        net.setOnline(2, false);
        net.setOnline(2, true);
        await net.flush();
        const cells = await Promise.all(peers.map((peer) => peer.read()));
        expect(cells[1]).toEqual(cells[0]);
        expect(cells[2]).toEqual(cells[0]);
        expect(cells[0][0]).toBe('offline');
      } finally { await close(); }
    });
  });
}
