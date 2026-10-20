import { beforeAll, describe, expect, it } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { initWasm, openPresentation, type PresentationHandle } from '../wasm/loader';
import { CollaborationProvider } from './provider';
import {
  decodeMessages,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeVarUint,
  OPEN_SYNC_CLIENT,
  parseSyncRequest,
  withoutDocumentFingerprint,
} from './protocol';
import type { AsyncCollaborationReplica, CollaborationTransport, CollaborationTransportEvent } from './types';

type Kind = 'sync' | 'async';

const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
let bytes: Uint8Array;
let seed: Uint8Array;

beforeAll(async () => {
  await initWasm(await readFile(resolve(import.meta.dir, '../wasm/generated/pptx_wasm_bg.wasm')));
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="${rel}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="900000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>alpha</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>');
  bytes = await zip.generateAsync({ type: 'uint8array' });
  const source = openPresentation(bytes, { clientId: 300 });
  seed = source.encodeStateAsUpdate();
  source.dispose();
});

function later<T>(operation: () => T): Promise<T> {
  return new Promise((done, fail) => setTimeout(() => {
    try { done(operation()); } catch (cause) { fail(cause); }
  }, 0));
}

function openPeer(kind: Kind, clientId: number) {
  const deck: PresentationHandle = openPresentation(bytes, { clientId, initialUpdate: seed });
  const story = () => deck.snapshot().slides[0].shapes[0].textStories[0].id;
  const replica: AsyncCollaborationReplica = {
    clientId,
    encodeStateVector: () => later(() => deck.encodeStateVector()),
    encodeStateAsUpdate: (vector) => later(() => deck.encodeStateAsUpdate(vector)),
    applyUpdate: (update) => later(() => deck.applyUpdate(update)),
    onUpdate: (listener) => deck.onUpdate(listener),
  };
  return {
    deck,
    connect: (transport: CollaborationTransport) => kind === 'sync'
      ? new CollaborationProvider(deck, transport)
      : new CollaborationProvider(replica, transport),
    edit: (text: string) => { deck.insertText(story(), 0, text); },
    erase: (start: number, end: number) => { deck.deleteText(story(), start, end); },
    text: () => deck.story(story()).paragraphs.flatMap((paragraph) => paragraph.runs.map((run) => run.text)).join(''),
  };
}

function relay(size: number) {
  const listeners = Array.from({ length: size }, () => new Set<(event: CollaborationTransportEvent) => void>());
  const online = new Array<boolean>(size).fill(false);
  const sent: { from: number; data: Uint8Array }[] = [];
  const opens = new Array<number>(size).fill(0);
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
        opens[frame.to] += requests.filter((request) => parseSyncRequest(request).open).length;
        for (const listener of [...listeners[frame.to]]) listener({ type: 'message', data: frame.data });
      }
    }
  };
  const count = (from: number, type: string, since: number) => sent.slice(since)
    .filter((frame) => frame.from === from)
    .flatMap((frame) => decodeMessages(frame.data))
    .filter((message) => message.type === type).length;
  return { transports, setOnline, reopen, flush, sent, opens, count };
}

function connectLegacyPeer(deck: PresentationHandle, transport: CollaborationTransport, errors: string[]) {
  let tail = Promise.resolve();
  transport.onEvent((event) => {
    tail = tail.then(async () => {
      if (event.type === 'open') transport.send(encodeSyncStep1(await later(() => deck.encodeStateVector())));
      if (event.type !== 'message') return;
      for (const message of decodeMessages(event.data)) {
        if (message.type === 'sync-step-1') {
          const vector = withoutDocumentFingerprint(message.stateVector).slice();
          transport.send(encodeSyncStep2(await later(() => deck.encodeStateAsUpdate(vector))));
        } else if (message.type === 'sync-step-2' || message.type === 'update') {
          const update = withoutDocumentFingerprint(message.update).slice();
          await later(() => deck.applyUpdate(update));
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
  const peers = Array.from({ length: size }, (_, index) => openPeer(kind, 201 + index));
  const net = relay(size);
  const errors: string[] = [];
  const providers = peers.flatMap((peer, index) => index === legacy ? [] : [peer.connect(net.transports[index])]);
  for (const provider of providers) {
    provider.onError((error) => errors.push(`${error.code}: ${error.message}`));
    provider.connect();
  }
  if (legacy !== undefined) connectLegacyPeer(peers[legacy].deck, net.transports[legacy], errors);
  await net.flush();
  return {
    peers, net, providers, errors,
    close() {
      for (const provider of providers) provider.destroy();
      for (const peer of peers) peer.deck.dispose();
    },
  };
}

for (const kind of ['sync', 'async'] as const) {
  describe(`${kind} replicas reconnecting over a peer transport`, () => {
    it('publish offline edits to a peer that saw no open', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        net.setOnline(0, false);
        left.edit('offline');
        net.setOnline(0, true);
        await net.flush();
        expect(right.deck.snapshot()).toEqual(left.deck.snapshot());
        expect(right.text()).toBe('offlinealpha');
      } finally { close(); }
    });

    it('keep edits from both sides of a split', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        net.setOnline(0, false);
        left.edit('left ');
        right.edit('right ');
        await net.flush();
        net.setOnline(0, true);
        await net.flush();
        expect(right.deck.snapshot()).toEqual(left.deck.snapshot());
        expect(left.text()).toContain('left ');
        expect(left.text()).toContain('right ');
      } finally { close(); }
    });

    it('converge three replicas on a relay after one reconnects', async () => {
      const { peers, net, close } = await network(kind, 3);
      try {
        net.setOnline(2, false);
        peers[2].edit('offline ');
        peers[0].edit('online ');
        await net.flush();
        const mark = net.sent.length;
        net.setOnline(2, true);
        await net.flush();
        expect(peers[1].deck.snapshot()).toEqual(peers[0].deck.snapshot());
        expect(peers[2].deck.snapshot()).toEqual(peers[0].deck.snapshot());
        expect(peers[0].text()).toContain('offline ');
        expect(peers[0].text()).toContain('online ');
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
        left.edit('offline');
        net.setOnline(0, true);
        await net.flush();
        expect([net.count(0, 'sync-step-1', mark), net.count(1, 'sync-step-1', mark)]).toEqual([1, 1]);
        expect([net.count(0, 'sync-step-2', mark), net.count(1, 'sync-step-2', mark)]).toEqual([1, 1]);
        expect(right.deck.snapshot()).toEqual(left.deck.snapshot());

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
        left.edit('left ');
        right.edit('right ');
        await net.flush();
        const mark = net.sent.length;
        net.reopen(0, 1);
        await net.flush();
        expect(right.deck.snapshot()).toEqual(left.deck.snapshot());
        expect(left.text()).toContain('left ');
        expect(left.text()).toContain('right ');
        expect([0, 1].map((index) => [net.count(index, 'sync-step-1', mark), net.count(index, 'sync-step-2', mark)]))
          .toEqual([[2, 2], [2, 2]]);
        const settled = net.sent.length;
        await net.flush();
        expect(net.sent.length).toBe(settled);
      } finally { close(); }
    });


    it('publish an offline deletion to a peer that saw no open', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        net.setOnline(0, false);
        const vector = left.deck.encodeStateVector();
        left.erase(0, 2);
        expect(left.deck.encodeStateVector()).toEqual(vector);
        net.setOnline(0, true);
        await net.flush();
        expect(right.deck.snapshot()).toEqual(left.deck.snapshot());
        expect(right.text()).toBe('pha');
      } finally { close(); }
    });

    it('publish an offline deletion to every replica on a relay', async () => {
      const { peers, net, close } = await network(kind, 3);
      try {
        net.setOnline(2, false);
        peers[2].erase(0, 2);
        net.setOnline(2, true);
        await net.flush();
        expect(peers[1].deck.snapshot()).toEqual(peers[0].deck.snapshot());
        expect(peers[2].deck.snapshot()).toEqual(peers[0].deck.snapshot());
        expect(peers[0].text()).toBe('pha');
      } finally { close(); }
    });

    it('publish an offline deletion at the native state vector entry limit', async () => {
      const { peers: [left, right], net, errors, close } = await network(kind, 2);
      try {
        const filler = phantomClients(65_536 - clients(left.deck.encodeStateVector()).length);
        for (const peer of [left, right]) peer.deck.applyUpdate(filler.slice());
        const vector = left.deck.encodeStateVector();
        expect(clients(vector)).toHaveLength(65_536);
        expect(clients(vector)).not.toContain(201n);
        net.setOnline(0, false);
        left.erase(0, 2);
        expect(left.deck.encodeStateVector()).toEqual(vector);
        net.setOnline(0, true);
        await net.flush();
        expect(errors).toEqual([]);
        expect(right.deck.snapshot()).toEqual(left.deck.snapshot());
        expect(right.text()).toBe('pha');
      } finally { close(); }
    });

    it('publish an offline deletion when the history already holds the open request client', async () => {
      const { peers: [left, right], net, errors, close } = await network(kind, 2);
      try {
        const history = Uint8Array.from([1, 1, ...encodeVarUint(OPEN_SYNC_CLIENT), 0, 0, 3, 0]);
        for (const peer of [left, right]) peer.deck.applyUpdate(history.slice());
        expect(clients(left.deck.encodeStateVector())).toContain(BigInt(OPEN_SYNC_CLIENT));
        net.setOnline(0, false);
        left.erase(0, 2);
        net.setOnline(0, true);
        await net.flush();
        expect(errors).toEqual([]);
        expect(right.deck.snapshot()).toEqual(left.deck.snapshot());
        expect(right.text()).toBe('pha');
      } finally { close(); }
    });

    it('mark an open request at one entry below the native limit', async () => {
      const { peers: [left, right], net, errors, close } = await network(kind, 2, 1);
      try {
        const filler = phantomClients(65_535 - clients(left.deck.encodeStateVector()).length);
        for (const peer of [left, right]) peer.deck.applyUpdate(filler.slice());
        expect(clients(left.deck.encodeStateVector())).toHaveLength(65_535);
        net.setOnline(0, false);
        right.edit('legacy ');
        const mark = net.sent.length;
        net.setOnline(0, true);
        await net.flush();
        const [request] = net.sent.slice(mark).filter((frame) => frame.from === 0)
          .flatMap((frame) => decodeMessages(frame.data))
          .flatMap((message) => message.type === 'sync-step-1' ? [message.stateVector] : []);
        expect(parseSyncRequest(request).open).toBe(true);
        expect(clients(request)).toHaveLength(65_536);
        expect(errors).toEqual([]);
        expect(left.deck.snapshot()).toEqual(right.deck.snapshot());
        expect(left.text()).toBe('legacy alpha');
      } finally { close(); }
    });

    it('answer every delivered open request with one reverse request, then go quiet', async () => {
      const { peers, net, close } = await network(kind, 3);
      try {
        net.setOnline(2, false);
        peers[2].edit('offline ');
        peers[0].edit('online ');
        await net.flush();
        const mark = net.sent.length;
        net.opens.fill(0);
        net.setOnline(2, true);
        await net.flush({ twice: true });
        expect(peers[1].deck.snapshot()).toEqual(peers[0].deck.snapshot());
        expect(peers[2].deck.snapshot()).toEqual(peers[0].deck.snapshot());
        expect(peers[0].text()).toContain('offline ');
        expect(peers[0].text()).toContain('online ');
        const reverse = [0, 1, 2].map((index) => net.count(index, 'sync-step-1', mark) - (index === 2 ? 1 : 0));
        expect(net.opens).toEqual([2, 2, 0]);
        expect(reverse).toEqual(net.opens);
        const settled = net.sent.length;
        await net.flush({ twice: true });
        expect(net.sent.length).toBe(settled);
      } finally { close(); }
    });

    for (const change of ['move', 'resize'] as const) {
      it(`converge after an offline shape ${change} without dropping the link`, async () => {
        const { peers: [left, right], net, providers, errors, close } = await network(kind, 2);
        try {
          const slide = left.deck.snapshot().slides[0];
          const shape = slide.shapes[0];
          net.setOnline(0, false);
          const { after } = change === 'move'
            ? left.deck.moveShape(slide.id, shape.id, shape.x + 250000, shape.y + 125000)
            : left.deck.resizeShape(slide.id, shape.id, shape.width - 1000000, shape.height + 300000);
          net.setOnline(0, true);
          await net.flush();
          expect(errors).toEqual([]);
          expect(providers.map((provider) => [provider.status, provider.synced]))
            .toEqual([['connected', true], ['connected', true]]);
          expect(right.deck.snapshot()).toEqual(left.deck.snapshot());
          const { x, y, width, height } = right.deck.snapshot().slides[0].shapes[0];
          expect({ x, y, width, height }).toEqual(after);
          expect(after).not.toEqual({ x: shape.x, y: shape.y, width: shape.width, height: shape.height });
        } finally { close(); }
      });
    }

    it('spend exactly one extra exchange per open, then go quiet', async () => {
      const { peers: [left, right], net, close } = await network(kind, 2);
      try {
        const opens: number[][] = [];
        for (const change of [() => {}, () => left.erase(0, 1), () => left.edit('offline ')]) {
          const mark = net.sent.length;
          net.setOnline(0, false);
          change();
          net.setOnline(0, true);
          await net.flush();
          opens.push([0, 1].flatMap((index) => ['sync-step-1', 'sync-step-2'].map((type) => net.count(index, type, mark))));
          const settled = net.sent.length;
          await net.flush();
          expect(net.sent.length).toBe(settled);
        }
        expect(opens).toEqual([[1, 1, 1, 1], [1, 1, 1, 1], [1, 1, 1, 1]]);
        expect(right.deck.snapshot()).toEqual(left.deck.snapshot());
        expect(right.text()).toBe('offline lpha');
      } finally { close(); }
    });

    it('stay valid with a peer that answers without requesting state back', async () => {
      const { peers: [left, right], net, providers: [provider], errors, close } = await network(kind, 2, 1);
      try {
        net.setOnline(0, false);
        net.setOnline(1, false);
        left.edit('ours ');
        right.edit('theirs ');
        const mark = net.sent.length;
        net.reopen(0, 1);
        await net.flush();
        expect(errors).toEqual([]);
        expect([provider.status, provider.synced]).toEqual(['connected', true]);
        expect(net.count(1, 'sync-step-1', mark)).toBe(1);
        expect(net.count(0, 'sync-step-2', mark)).toBeGreaterThan(0);
        expect(right.deck.snapshot()).toEqual(left.deck.snapshot());
        expect(left.text()).toContain('ours ');
        expect(left.text()).toContain('theirs ');
      } finally { close(); }
    });

    it('converge with a relay peer that leaves mid-handshake and rejoins unchanged', async () => {
      const { peers, net, close } = await network(kind, 3);
      try {
        net.setOnline(2, false);
        peers[2].edit('offline ');
        net.setOnline(2, true);
        await net.flush({ to: [0, 1] });
        net.setOnline(2, false);
        net.setOnline(2, true);
        await net.flush();
        expect(peers[1].deck.snapshot()).toEqual(peers[0].deck.snapshot());
        expect(peers[2].deck.snapshot()).toEqual(peers[0].deck.snapshot());
        expect(peers[0].text()).toBe('offline alpha');
      } finally { close(); }
    });
  });
}
