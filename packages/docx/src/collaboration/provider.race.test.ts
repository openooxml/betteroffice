import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from '../yrs';
import { createAsyncCollaborationReplica } from './asyncReplica';
import { CollaborationProvider } from './provider';
import type { CollaborationTransport, CollaborationTransportEvent } from './types';

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
  let gate: Promise<void> | undefined;
  const replica = kind === 'sync' ? session : createAsyncCollaborationReplica({
    clientId,
    encodeStateVector: () => {
      if (!gate) return later(() => session.encodeStateVector());
      const vector = session.encodeStateVector();
      return gate.then(() => vector);
    },
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
    hold: () => {
      let release!: () => void;
      gate = new Promise((done) => { release = done; });
      return () => { gate = undefined; release(); };
    },
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
  const pending = (from?: number) => queue.filter((frame) => from === undefined || frame.from === from).length;
  return { transports, setOnline, flush, pending };
}

type Link = ReturnType<typeof relay>;

async function opened(...links: Link[]) {
  for (let tick = 0; tick < 64 && links.some((link) => !link.pending(0)); tick += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  expect(links.map((link) => link.pending(0) > 0)).toEqual(links.map(() => true));
}

async function settle(links: Link[]) {
  for (let round = 0; round < 16 && (round === 0 || links.some((link) => link.pending())); round += 1) {
    for (const link of links) await link.flush();
  }
  expect(links.map((link) => link.pending())).toEqual(links.map(() => 0));
}

function attach(peer: Awaited<ReturnType<typeof openPeer>>, transport: CollaborationTransport, errors: string[]) {
  const provider = peer.connect(transport);
  provider.onError((error) => errors.push(`${error.code}: ${error.message}`));
  provider.connect();
  return provider;
}

for (const kind of ['sync', 'async'] as const) {
  describe(`${kind} replicas racing a reconnect`, () => {
    it('publish an offline deletion when a relay forwards the next insert first', async () => {
      const [a, c, b] = await Promise.all([301, 302, 303].map((clientId) => openPeer(kind, clientId)));
      const [ac, cb, ab] = [relay(2), relay(2), relay(2)];
      const errors: string[] = [];
      const providers = [
        attach(a, ac.transports[0], errors), attach(c, ac.transports[1], errors),
        attach(c, cb.transports[0], errors), attach(b, cb.transports[1], errors),
        attach(a, ab.transports[0], errors), attach(b, ab.transports[1], errors),
      ];
      try {
        await settle([ac, cb, ab]);
        ac.setOnline(0, false);
        ab.setOnline(0, false);
        const vector = a.session.encodeStateVector();
        a.erase(0, 0, 2);
        expect(a.session.encodeStateVector()).toEqual(vector);
        ac.setOnline(0, true);
        ab.setOnline(0, true);
        await opened(ac, ab);
        a.edit(0, 'online ');
        await ac.flush({ to: [1] });
        cb.setOnline(1, false);
        cb.setOnline(1, true);
        await cb.flush();
        await settle([ab, ac, cb]);
        expect(errors).toEqual([]);
        expect([a, c, b].map((peer) => peer.text()[0])).toEqual(['online pha beta', 'online pha beta', 'online pha beta']);
        expect(c.text()).toEqual(a.text());
        expect(b.text()).toEqual(a.text());
      } finally {
        for (const provider of providers) provider.destroy();
        for (const peer of [a, c, b]) peer.session.destroy();
      }
    });

    if (kind === 'async') {
      it('publish an offline deletion when the next insert leaves before the open vector', async () => {
        const [a, b] = await Promise.all([301, 302].map((clientId) => openPeer(kind, clientId)));
        const link = relay(2);
        const errors: string[] = [];
        const providers = [attach(a, link.transports[0], errors), attach(b, link.transports[1], errors)];
        try {
          await link.flush();
          link.setOnline(0, false);
          const vector = a.session.encodeStateVector();
          a.erase(0, 0, 2);
          expect(a.session.encodeStateVector()).toEqual(vector);
          const release = a.hold();
          link.setOnline(0, true);
          await new Promise((done) => setTimeout(done, 0));
          a.edit(0, 'online ');
          expect(link.pending(0)).toBe(1);
          release();
          await link.flush();
          expect(errors).toEqual([]);
          expect([a, b].map((peer) => peer.text()[0])).toEqual(['online pha beta', 'online pha beta']);
          expect(b.text()).toEqual(a.text());
        } finally {
          for (const provider of providers) provider.destroy();
          for (const peer of [a, b]) peer.session.destroy();
        }
      });
    }
  });
}
