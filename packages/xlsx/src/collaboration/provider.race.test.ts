import { describe, expect, it } from 'bun:test';
import { openWorkbook } from '../wasm/loader';
import { createTestWorkbookSession, loadWorkbookSessionFixtures } from '../session/testHelpers';
import { asyncReplicaFor, createAsyncCollaborationReplica } from './asyncReplica';
import { CollaborationProvider } from './provider';
import type { CollaborationReplica, CollaborationTransport, CollaborationTransportEvent } from './types';

type Kind = 'sync' | 'async';

const fixtures = loadWorkbookSessionFixtures();

async function openPeer(kind: Kind, clientId: number) {
  const { fixture } = await fixtures;
  if (kind === 'sync') {
    const handle = openWorkbook(fixture, { collaborative: true, clientId });
    return {
      replica: handle as CollaborationReplica,
      edit: async (col: number, input: string) => { expect(handle.editCell(0, 2, col + 1, input).applied).toBe(true); },
      undo: async () => { expect(handle.undo().applied).toBe(true); },
      read: async (from = 0) => [1, 2].map((col) => handle.cell(0, 2, from + col).input),
      vector: async () => handle.encodeStateVector(),
      hold: () => { throw new Error('sync replicas read their vector inline'); },
      dispose: () => handle.dispose(),
    };
  }
  const session = await createTestWorkbookSession(fixture, undefined, { collaborative: true, clientId });
  const inner = asyncReplicaFor(session.replica!)!;
  let gate: Promise<void> | undefined;
  const replica = createAsyncCollaborationReplica({
    clientId,
    encodeStateVector: () => {
      const read = inner.encodeStateVector();
      return gate ? gate.then(() => read) : read;
    },
    encodeStateAsUpdate: (vector) => inner.encodeStateAsUpdate(vector),
    applyUpdate: (update) => inner.applyUpdate(update),
    onUpdate: (listener) => inner.onUpdate(listener),
  });
  const applied = async (ticket: { ack: Promise<{ status: string }> }) => {
    expect((await ticket.ack).status).toBe('applied');
  };
  return {
    replica: replica as CollaborationReplica,
    edit: (col: number, input: string) => applied(session.editSession!.submit('editCell', { payload: { args: [0, 2, col + 1, input] } })),
    undo: () => applied(session.editSession!.submit('undo', { payload: { args: [] } })),
    read: async (from = 0) => {
      const range = `${String.fromCharCode(66 + from)}3:${String.fromCharCode(67 + from)}3`;
      return (await session.call.cellInputs(0, range)).cells[0].map((cell) => cell.input);
    },
    vector: () => inner.encodeStateVector(),
    hold: () => {
      let release!: () => void;
      gate = new Promise((done) => { release = done; });
      return () => { gate = undefined; release(); };
    },
    dispose: () => session.dispose(),
  };
}

function relay(size: number) {
  const listeners = Array.from({ length: size }, () => new Set<(event: CollaborationTransportEvent) => void>());
  const online = new Array<boolean>(size).fill(false);
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
  const flush = async ({ to = Array.from({ length: size }, (_, index) => index) } = {}) => {
    for (let idle = 0; idle < 64;) {
      await new Promise((resolve) => setTimeout(resolve, 0));
      const at = queue.findIndex((frame) => to.includes(frame.to));
      if (at < 0) { idle += 1; continue; }
      idle = 0;
      const [frame] = queue.splice(at, 1);
      for (const listener of [...listeners[frame.to]]) listener({ type: 'message', data: frame.data });
    }
  };
  return { transports, setOnline, flush, pending: () => queue.length };
}

type Link = ReturnType<typeof relay>;

async function opened(...links: Link[]) {
  for (let tick = 0; tick < 64 && links.some((link) => !link.pending()); tick += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  expect(links.map((link) => link.pending() > 0)).toEqual(links.map(() => true));
}

async function settle(links: Link[]) {
  for (let round = 0; round < 16 && links.some((link) => link.pending()); round += 1) {
    for (const link of links) await link.flush();
  }
  expect(links.map((link) => link.pending())).toEqual(links.map(() => 0));
}

function attach(replica: CollaborationReplica, transport: CollaborationTransport, errors: string[]) {
  const provider = new CollaborationProvider(replica, transport);
  provider.onError((error) => errors.push(`${error.code}: ${error.message}`));
  provider.connect();
  return provider;
}

for (const kind of ['sync', 'async'] as const) {
  describe(`${kind} replicas racing a reconnect`, () => {
    it('publish an offline deletion when a relay forwards the next insert first', async () => {
      const [a, c, b] = await Promise.all([101, 102, 103].map((clientId) => openPeer(kind, clientId)));
      const [ac, cb, ab] = [relay(2), relay(2), relay(2)];
      const errors: string[] = [];
      const providers = [
        attach(a.replica, ac.transports[0], errors), attach(c.replica, ac.transports[1], errors),
        attach(c.replica, cb.transports[0], errors), attach(b.replica, cb.transports[1], errors),
        attach(a.replica, ab.transports[0], errors), attach(b.replica, ab.transports[1], errors),
      ];
      try {
        await settle([ac, cb, ab]);
        await a.edit(4, 'cleared');
        await a.edit(5, 'undone');
        await settle([ac, cb, ab]);
        expect(await b.read(4)).toEqual(['cleared', 'undone']);
        ac.setOnline(0, false);
        ab.setOnline(0, false);
        const vector = await a.vector();
        await a.undo();
        await a.edit(4, '');
        expect(await a.vector()).toEqual(vector);
        ac.setOnline(0, true);
        ab.setOnline(0, true);
        await opened(ac, ab);
        await a.edit(0, 'online');
        await ac.flush({ to: [1] });
        cb.setOnline(1, false);
        cb.setOnline(1, true);
        await cb.flush();
        await settle([ab, ac, cb]);
        expect(errors).toEqual([]);
        const cells = await Promise.all([a, c, b].map(async (peer) => [...await peer.read(), ...await peer.read(4)]));
        expect(cells[1]).toEqual(cells[0]);
        expect(cells[2]).toEqual(cells[0]);
        expect(cells[0][0]).toBe('online');
        expect(cells[0].slice(2).filter(Boolean)).toEqual([]);
      } finally {
        for (const provider of providers) provider.destroy();
        await Promise.all([a, c, b].map((peer) => peer.dispose()));
      }
    });

    if (kind === 'async') {
      it('publish an offline deletion when the next insert leaves before the open vector', async () => {
        const [a, b] = await Promise.all([101, 102].map((clientId) => openPeer(kind, clientId)));
        const link = relay(2);
        const errors: string[] = [];
        const providers = [attach(a.replica, link.transports[0], errors), attach(b.replica, link.transports[1], errors)];
        try {
          await link.flush();
          await a.edit(4, 'cleared');
          await a.edit(5, 'undone');
          await link.flush();
          expect(await b.read(4)).toEqual(['cleared', 'undone']);
          link.setOnline(0, false);
          const vector = await a.vector();
          await a.undo();
          await a.edit(4, '');
          expect(await a.vector()).toEqual(vector);
          const release = a.hold();
          link.setOnline(0, true);
          await a.edit(0, 'online');
          release();
          await link.flush();
          expect(errors).toEqual([]);
          const cells = await Promise.all([a, b].map(async (peer) => [...await peer.read(), ...await peer.read(4)]));
          expect(cells[1]).toEqual(cells[0]);
          expect(cells[0][0]).toBe('online');
          expect(cells[0].slice(2).filter(Boolean)).toEqual([]);
        } finally {
          for (const provider of providers) provider.destroy();
          await Promise.all([a, b].map((peer) => peer.dispose()));
        }
      });
    }
  });
}
