import { describe, expect, it } from 'bun:test';
import { openWorkbook } from '../wasm/loader';
import { createTestWorkbookSession, loadWorkbookSessionFixtures } from '../session/testHelpers';
import { asyncReplicaFor } from './asyncReplica';
import { CollaborationProvider } from './provider';
import { decodeMessages } from './protocol';
import type { CollaborationReplica, CollaborationTransport, CollaborationTransportEvent } from './types';

type Kind = 'sync' | 'async';

interface Peer {
  replica: CollaborationReplica;
  edit(col: number, input: string): Promise<void>;
  undo(): Promise<void>;
  read(): Promise<string[]>;
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
      read: async () => [1, 2].map((col) => handle.cell(0, 2, col).input),
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
    read: async () => (await session.call.cellInputs(0, 'B3:C3')).cells[0].map((cell) => cell.input),
    vector: () => asyncReplicaFor(session.replica!)!.encodeStateVector(),
    dispose: () => session.dispose(),
  };
}

function relay(size: number) {
  const listeners = Array.from({ length: size }, () => new Set<(event: CollaborationTransportEvent) => void>());
  const online = new Array<boolean>(size).fill(false);
  const sent: { from: number; data: Uint8Array }[] = [];
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
  const flush = async () => {
    for (let idle = 0; idle < 64;) {
      await new Promise((resolve) => setTimeout(resolve, 0));
      const frame = queue.shift();
      if (!frame) { idle += 1; continue; }
      idle = 0;
      for (const listener of [...listeners[frame.to]]) listener({ type: 'message', data: frame.data });
    }
  };
  const carries = (data: Uint8Array, type: string) => decodeMessages(data).some((message) => message.type === type);
  const holdBack = async (held: string, doubled: string) => {
    for (let idle = 0, moved = 0; idle < 64 && moved < 32;) {
      await new Promise((resolve) => setTimeout(resolve, 0));
      const at = queue.findIndex((frame) => !carries(frame.data, held));
      if (at < 0) { idle += 1; continue; }
      idle = 0;
      moved += 1;
      const [frame] = queue.splice(at, 1);
      const copies = carries(frame.data, doubled) ? 2 : 1;
      for (let copy = 0; copy < copies; copy += 1) {
        for (const listener of [...listeners[frame.to]]) listener({ type: 'message', data: frame.data });
      }
    }
  };
  const count = (from: number, type: string, since: number) => sent.slice(since)
    .filter((frame) => frame.from === from)
    .flatMap((frame) => decodeMessages(frame.data))
    .filter((message) => message.type === type).length;
  return { transports, setOnline, reopen, flush, holdBack, sent, count };
}

async function network(kind: Kind, size: number) {
  const peers = await Promise.all(Array.from({ length: size }, (_, index) => openPeer(kind, 101 + index)));
  const net = relay(size);
  const providers = peers.map((peer, index) => new CollaborationProvider(peer.replica, net.transports[index]));
  for (const provider of providers) provider.connect();
  await net.flush();
  return {
    peers, net,
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
        expect([net.count(0, 'sync-step-1', mark), net.count(1, 'sync-step-1', mark)]).toEqual([1, 0]);

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

    it('request each missing state once while answers are delayed and duplicated', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        net.setOnline(0, false);
        net.setOnline(1, false);
        await left.edit(0, 'left offline');
        await right.edit(1, 'right offline');
        const mark = net.sent.length;
        net.reopen(0, 1);
        await net.holdBack('sync-step-2', 'sync-step-1');
        await net.flush();
        const cells = await Promise.all([left.read(), right.read()]);
        expect(cells[1]).toEqual(cells[0]);
        expect(cells[0]).toEqual(['left offline', 'right offline']);
        expect([0, 1].map((index) => net.count(index, 'sync-step-1', mark))).toEqual([2, 2]);
        const settled = net.sent.length;
        await net.flush();
        expect(net.sent.length).toBe(settled);
      } finally { await close(); }
    });

    it('publish offline deletions to a peer that saw no open', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        await left.edit(0, 'cleared');
        await left.edit(1, 'undone');
        await net.flush();
        expect(await right.read()).toEqual(['cleared', 'undone']);
        net.setOnline(0, false);
        const vector = await left.vector();
        await left.undo();
        expect(await left.vector()).toEqual(vector);
        await left.edit(0, '');
        expect(await left.vector()).toEqual(vector);
        net.setOnline(0, true);
        await net.flush();
        const cells = await Promise.all([left.read(), right.read()]);
        expect(cells[1]).toEqual(cells[0]);
        expect(cells[0].filter(Boolean)).toEqual([]);
      } finally { await close(); }
    });

    it('publish offline deletions to every replica on a relay', async () => {
      const { peers, net, close } = await network(kind, 3);
      try {
        await peers[2].edit(0, 'cleared');
        await peers[2].edit(1, 'undone');
        await net.flush();
        net.setOnline(2, false);
        await peers[2].undo();
        await peers[2].edit(0, '');
        net.setOnline(2, true);
        await net.flush();
        const cells = await Promise.all(peers.map((peer) => peer.read()));
        expect(cells[1]).toEqual(cells[0]);
        expect(cells[2]).toEqual(cells[0]);
        expect(cells[0].filter(Boolean)).toEqual([]);
      } finally { await close(); }
    });
  });
}
