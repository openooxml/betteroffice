import { beforeAll, describe, expect, it } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { initWasm, openPresentation, type PresentationHandle } from '../wasm/loader';
import { CollaborationProvider } from './provider';
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
  let gate: Promise<void> | undefined;
  const replica: AsyncCollaborationReplica = {
    clientId,
    encodeStateVector: () => {
      if (!gate) return later(() => deck.encodeStateVector());
      const vector = deck.encodeStateVector();
      return gate.then(() => vector);
    },
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

function attach(peer: ReturnType<typeof openPeer>, transport: CollaborationTransport, errors: string[]) {
  const provider = peer.connect(transport);
  provider.onError((error) => errors.push(`${error.code}: ${error.message}`));
  provider.connect();
  return provider;
}

for (const kind of ['sync', 'async'] as const) {
  describe(`${kind} replicas racing a reconnect`, () => {
    it('publish an offline deletion when a relay forwards the next insert first', async () => {
      const [a, c, b] = [201, 202, 203].map((clientId) => openPeer(kind, clientId));
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
        const vector = a.deck.encodeStateVector();
        a.erase(0, 2);
        expect(a.deck.encodeStateVector()).toEqual(vector);
        ac.setOnline(0, true);
        ab.setOnline(0, true);
        await opened(ac, ab);
        a.edit('online ');
        await ac.flush({ to: [1] });
        cb.setOnline(1, false);
        cb.setOnline(1, true);
        await cb.flush();
        await settle([ab, ac, cb]);
        expect(errors).toEqual([]);
        expect(c.deck.snapshot()).toEqual(a.deck.snapshot());
        expect(b.deck.snapshot()).toEqual(a.deck.snapshot());
        expect([a, c, b].map((peer) => peer.text())).toEqual(['online pha', 'online pha', 'online pha']);
      } finally {
        for (const provider of providers) provider.destroy();
        for (const peer of [a, c, b]) peer.deck.dispose();
      }
    });

    if (kind === 'async') {
      it('publish an offline deletion when the next insert leaves before the open vector', async () => {
        const [a, b] = [201, 202].map((clientId) => openPeer(kind, clientId));
        const link = relay(2);
        const errors: string[] = [];
        const providers = [attach(a, link.transports[0], errors), attach(b, link.transports[1], errors)];
        try {
          await link.flush();
          link.setOnline(0, false);
          const vector = a.deck.encodeStateVector();
          a.erase(0, 2);
          expect(a.deck.encodeStateVector()).toEqual(vector);
          const release = a.hold();
          link.setOnline(0, true);
          a.edit('online ');
          await new Promise((done) => setTimeout(done, 0));
          release();
          await link.flush();
          expect(errors).toEqual([]);
          expect(b.deck.snapshot()).toEqual(a.deck.snapshot());
          expect([a, b].map((peer) => peer.text())).toEqual(['online pha', 'online pha']);
        } finally {
          for (const provider of providers) provider.destroy();
          for (const peer of [a, b]) peer.deck.dispose();
        }
      });
    }
  });
}
