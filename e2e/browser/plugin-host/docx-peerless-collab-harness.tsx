import { createRoot } from 'react-dom/client';
import { createRef } from 'react';
import JSZip from 'jszip';
import { DocxEditor, type DocxEditorRef } from '@betteroffice/docx-react';
import { CollaborationProvider, createAsyncCollaborationReplica, type AsyncReplicaAdapter, type CollaborationReplica, type CollaborationTransport, type CollaborationTransportEvent } from '@betteroffice/docx/collaboration';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import type { DocxPeerlessCollaborationProbe } from './docx-peerless-collab-probe';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '@betteroffice/docx-react/styles.css';

const query = new URLSearchParams(location.search);
const mixed = query.has('mixed');
const pinned = !query.has('unpinned');
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
// A held endpoint still answers handshakes; only data frames wait for release.
const syncRequest = (data: Uint8Array) => data[0] === 0 && data[1] === 0;
function transport(index: number): CollaborationTransport {
  const endpoint = endpoints[index]!;
  let joined = false;
  return {
    connect() {
      endpoint.connected = true;
      const join = !joined && endpoints[1 - index]!.connected;
      joined = true;
      if (join) emit(1 - index, { type: 'close' });
      emit(index, { type: 'open' });
      if (join) emit(1 - index, { type: 'open' });
    },
    disconnect() { endpoint.connected = false; endpoint.queued.length = 0; emit(index, { type: 'close' }); },
    send(data) {
      const target = endpoints[1 - index]!;
      if (!target.connected) return true;
      if (target.held && !syncRequest(data)) target.queued.push(data.slice());
      else queueMicrotask(() => { if (target.connected) emit(1 - index, { type: 'message', data: data.slice() }); });
      return true;
    },
    onEvent(listener) { endpoint.listeners.add(listener); return () => { endpoint.listeners.delete(listener); }; },
  };
}

const file = await syntheticDocument('Seed');
const errors: string[] = [];
const checkpoints: Uint8Array[] = [];
const slots = [0, 1].map(index => ({ ref: createRef<DocxEditorRef>(), file, clientId: 4300 + index,
  initialUpdate: undefined as Uint8Array | undefined, replica: null as CollaborationReplica | null,
  provider: undefined as CollaborationProvider | undefined, generation: 0 }));
const root = createRoot(document.getElementById('root')!);
let count = 1;
const mergeHooks: Array<(() => void) | undefined> = [];
// The provider merges through this replica, so a hook sees each worker merge right after it is dispatched.
const observed = (index: number, replica: CollaborationReplica) => {
  const hosted = (replica as Partial<AsyncReplicaAdapter>).async;
  const merged = <T,>(merging: T) => { const hook = mergeHooks[index]; mergeHooks[index] = undefined; hook?.(); return merging; };
  return hosted ? createAsyncCollaborationReplica({ ...hosted,
    applyUpdate: update => merged(hosted.applyUpdate(update)),
    ...(hosted.applyUpdateChunks && { applyUpdateChunks: (updates: readonly Uint8Array[]) => merged(hosted.applyUpdateChunks!(updates)) }) }) : replica;
};

const onReplica = slots.map((slot, index) => (replica: CollaborationReplica | null) => {
  slot.provider?.destroy(); slot.provider = undefined;
  slot.replica = replica;
  if (replica) {
    slot.generation++;
    slot.provider = new CollaborationProvider(observed(index, replica), transport(index), { user: { name: index ? 'Writer B' : 'Writer A' } });
    slot.provider.onError(error => errors.push(error.message));
    slot.provider.connect();
  }
  render();
});
function render() {
  root.render(<div style={{ display: 'flex', height: '100%' }}>{slots.slice(0, count).map((slot, index) =>
    <div key={index} data-editor={index} style={{ width: '50%', height: '100%', minWidth: 0 }}>
      <DocxEditor ref={slot.ref} documentBuffer={slot.file} {...(mixed && index === 1 ? { legacySyncApi: true } : {})}
        showToolbar={false} showRuler={false} showOutlineButton={false} measurementFontProvider={fonts}
        collaboration={{ clientId: pinned ? slot.clientId : undefined, initialUpdate: slot.initialUpdate, onReplica: onReplica[index], presence: slot.provider }}
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
const stateVectorEntries = (bytes: Uint8Array) => {
  let at = 0;
  const next = () => { let value = 0; for (let scale = 1; ; scale *= 128) { const byte = bytes[at++]!; value += (byte & 127) * scale; if (byte < 128) return value; } };
  return Array.from({ length: next() }, () => [next(), next()] as [number, number]).sort((a, b) => a[0] - b[0]);
};
const proposalResult = <T extends { ok: boolean }>(result: T) => {
  if (!result.ok) throw new Error(JSON.stringify(result));
  return result as Extract<T, { ok: true }>;
};
const reopen = async (source?: string | Uint8Array) => {
  const generations = slots.map(slot => slot.generation);
  const update = source === undefined ? await asyncSide(0).encodeStateAsUpdate() : undefined;
  if (source !== undefined) {
    const bytes = typeof source === 'string' ? await syntheticDocument(source) : source;
    for (const slot of slots) slot.file = bytes;
    count = 1;
  }
  for (const slot of slots) { slot.clientId += 100; slot.initialUpdate = update?.slice(); }
  render();
  await waitReplica(0, generations[0]);
  if (source !== undefined) { slots[1]!.initialUpdate = await asyncSide(0).encodeStateAsUpdate(); count = 2; render(); }
  await waitReplica(1, generations[1]);
};
export const probe: DocxPeerlessCollaborationProbe = {
  errors,
  ready: Promise.resolve(),
  load: bytes => reopen(new Uint8Array(bytes)),
  async text(index: number, view: 'accepted' | 'original' = 'accepted') {
    for (let attempt = 0; attempt < 20; attempt++) {
      try {
        const read = await slots[index]!.ref.current!.readParagraphs({ story: 'body', view });
        if (!read.ok) throw new Error(read.failure.message);
        return read.paragraphs.map(paragraph => paragraph.text).join('\n');
      } catch (error) {
        if (!(error instanceof Error) || !error.message.includes('superseded')) throw error;
        await new Promise(resolve => setTimeout(resolve, 10));
      }
    }
    throw new Error('Document text did not settle');
  },
  painted(index: number) { return Array.from(document.querySelectorAll(`[data-editor="${index}"] .canvas-page-mirror`))
    .map(host => host.textContent ?? '').join(''); },
  generation(index: number) { return slots[index]!.generation; },
  clientId(index: number) { return slots[index]!.replica?.clientId ?? null; },
  async stateVector(index: number) { return stateVectorEntries(await asyncSide(index).encodeStateVector()); },
  synced(index: number) { return slots[index]!.provider?.synced ?? false; },
  hold(index: number) { endpoints[index]!.held = true; },
  release(index: number) {
    const endpoint = endpoints[index]!; endpoint.held = false;
    for (const data of endpoint.queued.splice(0).reverse()) { emit(index, { type: 'message', data }); emit(index, { type: 'message', data: data.slice() }); }
  },
  drop(index: number) { endpoints[index]!.queued.length = 0; },
  heldUpdates(index: number) { return endpoints[index]!.queued.filter(data => data[0] === 0 && data[1] === 2).length; },
  disconnect(index: number) { slots[index]!.provider!.disconnect(); },
  dropMidMerge(index: number) {
    const dropped = new Promise<void>(resolve => { mergeHooks[index] = () => { slots[index]!.provider!.disconnect(); resolve(); }; });
    probe.release(index);
    return dropped;
  },
  reconnect(index: number) { slots[index]!.provider!.connect(); },
  restart() {
    const open = [0, 1].filter(index => endpoints[index]!.connected);
    for (const index of open) emit(index, { type: 'close' });
    for (const index of open) emit(index, { type: 'open' });
  },
  async checkpoint(index: number) { checkpoints[index] = await asyncSide(index).encodeStateAsUpdate(); },
  async rejoin(index: number) {
    const slot = slots[index]!; const generation = slot.generation;
    slot.initialUpdate = checkpoints[index]!.slice(); render();
    await waitReplica(index, generation);
  },
  async propose(index: number, text: string) {
    const editor = slots[index]!.ref.current!;
    const paragraph = (await editor.getParagraphIdentities()).paragraphs.find(identity => identity.session?.story === 'body')!.session!;
    const { version } = await editor.resolveParagraphAnchors([paragraph]);
    proposalResult(await editor.proposeChanges({ expectVersion: version, proposals: [{ id: 'proposal', paragraph, op: 'insertText', at: 'end', text,
      suggest: { author: 'Writer A', date: '2026-10-01T00:00:00Z' } }] }));
  },
  async settle(index: number, state: 'accepted' | 'rejected') {
    const editor = slots[index]!.ref.current!;
    const current = await editor.getProposals();
    const ids = current.proposals.map(proposal => proposal.id);
    const decided = proposalResult(await editor.setProposalStates({ expectVersion: current.version,
      expectPreviewVersion: current.previewVersion, changes: ids.map(id => ({ id, state })) }));
    proposalResult(await editor.withdrawProposals({ expectVersion: decided.snapshot.version, ids }));
  },
  room: () => reopen(),
  swap: () => reopen('Replacement'),
};
window.__docxPeerlessCollab = probe;
probe.ready = (async () => {
  render(); await waitReplica(0);
  slots[1]!.initialUpdate = await asyncSide(0).encodeStateAsUpdate();
  count = 2; render(); await waitReplica(1);
})();
