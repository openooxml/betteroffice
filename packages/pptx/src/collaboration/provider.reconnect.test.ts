import { beforeAll, describe, expect, it } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { initWasm, openPresentation, type PresentationHandle } from '../wasm/loader';
import { CollaborationProvider } from './provider';
import { decodeMessages } from './protocol';
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
    text: () => deck.story(story()).paragraphs.flatMap((paragraph) => paragraph.runs.map((run) => run.text)).join(''),
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
  const peers = Array.from({ length: size }, (_, index) => openPeer(kind, 201 + index));
  const net = relay(size);
  const providers = peers.map((peer, index) => peer.connect(net.transports[index]));
  for (const provider of providers) provider.connect();
  await net.flush();
  return {
    peers, net,
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
        expect([net.count(0, 'sync-step-1', mark), net.count(1, 'sync-step-1', mark)]).toEqual([1, 0]);

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
  });
}
