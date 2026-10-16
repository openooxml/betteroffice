import { describe, expect, it } from 'bun:test';
import { openWorkbook } from '../wasm/loader';
import { createTestWorkbookSession, loadWorkbookSessionFixtures } from '../session/testHelpers';
import { CollaborationProvider } from './provider';
import { decodeMessages } from './protocol';
import type { CollaborationReplica, CollaborationTransport, CollaborationTransportEvent } from './types';

type Kind = 'sync' | 'async';

interface Peer {
  replica: CollaborationReplica;
  edit(col: number, input: string): Promise<void>;
  read(): Promise<string[]>;
  dispose(): unknown;
}

const fixtures = loadWorkbookSessionFixtures();

async function openPeer(kind: Kind, clientId: number): Promise<Peer> {
  const { fixture } = await fixtures;
  if (kind === 'sync') {
    const handle = openWorkbook(fixture, { collaborative: true, clientId });
    return {
      replica: handle,
      edit: async (col, input) => { handle.editCell(0, 0, col, input); },
      read: async () => [0, 1, 2].map((col) => handle.cell(0, 0, col).input),
      dispose: () => handle.dispose(),
    };
  }
  const session = await createTestWorkbookSession(fixture, undefined, { collaborative: true, clientId });
  return {
    replica: session.replica!,
    edit: async (col, input) => {
      const ack = await session.editSession!.submit('editCell', { payload: { args: [0, 0, col, input] } }).ack;
      expect(ack.status).toBe('applied');
    },
    read: async () => (await session.call.cellInputs(0, 'A1:C1')).cells[0].map((cell) => cell.input),
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
  const count = (from: number, type: string, since: number) => sent.slice(since)
    .filter((frame) => frame.from === from)
    .flatMap((frame) => decodeMessages(frame.data))
    .filter((message) => message.type === type).length;
  return { transports, setOnline, flush, sent, count };
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
  });
}
