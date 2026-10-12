import { createRoot } from 'react-dom/client';
import { createRef } from 'react';
import JSZip from 'jszip';
import { DocxEditor, type DocxEditorProps, type DocxEditorRef } from '@betteroffice/docx-react';
import { CollaborationProvider, type AsyncReplicaAdapter, type CollaborationReplica, type CollaborationTransport, type CollaborationTransportEvent } from '@betteroffice/docx/collaboration';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import type { DocxPeerlessCollaborationProbe } from './docx-peerless-collab-probe';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '@betteroffice/docx-react/styles.css';

const workerOnly = (): Partial<DocxEditorProps> => ({ experimentalPeerless: true });
const mixed = new URLSearchParams(location.search).has('mixed');
setGoogleFontsEnabled(false);
const fonts = { resolve: () => async () => (await fetch(fontUrl)).arrayBuffer() };

async function syntheticDocument(text: string) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/></w:rPr><w:t>${text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

const endpoints = [0, 1].map(() => ({ connected: false, held: false, queued: [] as Uint8Array[], listeners: new Set<(event: CollaborationTransportEvent) => void>() }));
const emit = (index: number, event: CollaborationTransportEvent) => { for (const listener of endpoints[index]!.listeners) listener(event); };
function transport(index: number): CollaborationTransport {
  const endpoint = endpoints[index]!;
  return {
    connect() { endpoint.connected = true; emit(index, { type: 'open' }); },
    disconnect() { endpoint.connected = false; endpoint.queued.length = 0; emit(index, { type: 'close' }); },
    send(data) {
      const target = endpoints[1 - index]!;
      if (!target.connected) return true;
      if (target.held) target.queued.push(data.slice());
      else queueMicrotask(() => { if (target.connected) emit(1 - index, { type: 'message', data: data.slice() }); });
      return true;
    },
    onEvent(listener) { endpoint.listeners.add(listener); return () => { endpoint.listeners.delete(listener); }; },
  };
}

const file = await syntheticDocument('Seed');
const errors: string[] = [];
const slots = [0, 1].map(index => ({ ref: createRef<DocxEditorRef>(), file, clientId: 4300 + index,
  initialUpdate: undefined as Uint8Array | undefined, replica: null as CollaborationReplica | null,
  provider: undefined as CollaborationProvider | undefined, generation: 0 }));
const root = createRoot(document.getElementById('root')!);
let count = 1;

const onReplica = slots.map((slot, index) => (replica: CollaborationReplica | null) => {
  slot.provider?.destroy(); slot.provider = undefined;
  slot.replica = replica;
  if (replica) {
    slot.generation++;
    slot.provider = new CollaborationProvider(replica, transport(index), { user: { name: index ? 'Writer B' : 'Writer A' } });
    slot.provider.onError(error => errors.push(error.message));
    slot.provider.connect();
  }
  render();
});
function render() {
  root.render(<div style={{ display: 'flex', height: '100%' }}>{slots.slice(0, count).map((slot, index) =>
    <div key={index} data-editor={index} style={{ width: '50%', height: '100%', minWidth: 0 }}>
      <DocxEditor ref={slot.ref} documentBuffer={slot.file} {...(mixed && index === 1 ? {} : workerOnly())}
        showToolbar={false} showRuler={false} showOutlineButton={false} measurementFontProvider={fonts}
        collaboration={{ clientId: slot.clientId, initialUpdate: slot.initialUpdate, onReplica: onReplica[index], presence: slot.provider }}
        onError={error => errors.push(error.message)} />
    </div>)}</div>);
}
const asyncSide = (index: number) => {
  const replica = slots[index]!.replica;
  if (!replica) throw new Error('Replica is not ready');
  return (replica as Partial<AsyncReplicaAdapter>).async ?? replica;
};
const waitReplica = async (index: number, previous = 0) => {
  for (let attempt = 0; attempt < 600 && slots[index]!.generation <= previous; attempt++) await new Promise(resolve => setTimeout(resolve, 100));
  if (!slots[index]!.replica || slots[index]!.generation <= previous) throw new Error('Replica open timed out');
};
const reopen = async (text?: string) => {
  const generations = slots.map(slot => slot.generation);
  const update = text === undefined ? await asyncSide(0).encodeStateAsUpdate() : undefined;
  if (text !== undefined) {
    const bytes = await syntheticDocument(text);
    for (const slot of slots) slot.file = bytes;
    count = 1;
  }
  for (const slot of slots) { slot.clientId += 100; slot.initialUpdate = update?.slice(); }
  render();
  await waitReplica(0, generations[0]);
  if (text !== undefined) { slots[1]!.initialUpdate = await asyncSide(0).encodeStateAsUpdate(); count = 2; render(); }
  await waitReplica(1, generations[1]);
};
export const probe: DocxPeerlessCollaborationProbe = {
  errors,
  ready: Promise.resolve(),
  async text(index: number) {
    const read = await slots[index]!.ref.current!.readParagraphs({ story: 'body', view: 'accepted' });
    if (!read.ok) throw new Error(read.failure.message);
    return read.paragraphs.map(paragraph => paragraph.text).join('\n');
  },
  painted(index: number) { return Array.from(document.querySelectorAll(`[data-editor="${index}"] .canvas-page-mirror`))
    .map(host => host.textContent ?? '').join(''); },
  generation(index: number) { return slots[index]!.generation; },
  synced(index: number) { return slots[index]!.provider?.synced ?? false; },
  hold(index: number) { endpoints[index]!.held = true; },
  release(index: number) {
    const endpoint = endpoints[index]!; endpoint.held = false;
    for (const data of endpoint.queued.splice(0).reverse()) { emit(index, { type: 'message', data }); emit(index, { type: 'message', data: data.slice() }); }
  },
  drop(index: number) { endpoints[index]!.queued.length = 0; },
  disconnect(index: number) { slots[index]!.provider!.disconnect(); },
  reconnect(index: number) { slots[index]!.provider!.connect(); },
  room: () => reopen(),
  swap: () => reopen('Replacement'),
};
window.__docxPeerlessCollab = probe;
probe.ready = (async () => {
  render(); await waitReplica(0);
  slots[1]!.initialUpdate = await asyncSide(0).encodeStateAsUpdate();
  count = 2; render(); await waitReplica(1);
})();
