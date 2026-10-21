import { createRoot } from 'react-dom/client';
import { createRef } from 'react';
import JSZip from 'jszip';
import { DocxEditor, type DocxEditorRef } from '@betteroffice/docx-react';
import { CollaborationProvider, createAsyncCollaborationReplica, type AsyncReplicaAdapter, type CollaborationReplica,
  type CollaborationTransport, type CollaborationTransportEvent } from '@betteroffice/docx/collaboration';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import { classifyFrame, RetainedUpdateLog } from '../../../apps/relay/src/retention';
import { MAX_COLLABORATION_FRAME_BYTES } from '../../../shared/collaboration-limits';
import type { DocxRelayJoinProbe, DocxRelayJoinResult } from './docx-relay-join-probe';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '@betteroffice/docx-react/styles.css';

setGoogleFontsEnabled(false);
const fonts = { resolve: () => async () => (await fetch(fontUrl)).arrayBuffer() };

async function syntheticDocument(text: string) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/></w:rPr><w:t>${text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

/** The relay's room in-page: the same retained log, replayed frame by frame to a joiner, one task per frame. */
const log = new RetainedUpdateLog(512, MAX_COLLABORATION_FRAME_BYTES);
const sockets = new Set<(event: CollaborationTransportEvent) => void>();
const deliver = (socket: (event: CollaborationTransportEvent) => void, data: Uint8Array) =>
  setTimeout(() => { if (sockets.has(socket)) socket({ type: 'message', data }); }, 0);
let replayed = 0;
function transport(): CollaborationTransport {
  const listeners = new Set<(event: CollaborationTransportEvent) => void>();
  const socket = (event: CollaborationTransportEvent) => { for (const listener of listeners) listener(event); };
  return {
    connect() {
      sockets.add(socket);
      socket({ type: 'open' });
      const frames = log.snapshot();
      replayed = frames.length;
      for (const frame of frames) deliver(socket, frame);
      deliver(socket, log.syncRequest());
    },
    disconnect() { sockets.delete(socket); socket({ type: 'close' }); },
    send(data) {
      const bytes = data.slice();
      const { kind } = classifyFrame(bytes);
      for (const response of log.responses(bytes)) deliver(socket, response);
      if (kind === 'document') log.retain(bytes);
      for (const peer of sockets) if (peer !== socket) deliver(peer, bytes.slice());
      return true;
    },
    onEvent(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
  };
}

const file = await syntheticDocument('Seed');
const errors: string[] = [];
let merges = 0;
let joiner: { legacy: boolean } | null = null;
let seed: Uint8Array | undefined;
const refs = [createRef<DocxEditorRef>(), createRef<DocxEditorRef>()];
const replicas: (CollaborationReplica | null)[] = [null, null];
const providers: (CollaborationProvider | undefined)[] = [];
const counted = (replica: CollaborationReplica): CollaborationReplica => {
  const hosted = (replica as Partial<AsyncReplicaAdapter>).async;
  if (hosted) {
    return createAsyncCollaborationReplica({ ...hosted,
      applyUpdate(update) { merges++; return hosted.applyUpdate(update); },
      ...(hosted.applyUpdateChunks && { applyUpdateChunks(updates: readonly Uint8Array[]) { merges++; return hosted.applyUpdateChunks!(updates); } }) });
  }
  return Object.assign(Object.create(replica) as CollaborationReplica, {
    applyUpdate(update: Uint8Array) { merges++; return replica.applyUpdate(update); },
    ...(replica.applyUpdateChunks && { applyUpdateChunks(updates: readonly Uint8Array[]) { merges++; return replica.applyUpdateChunks!(updates); } }),
  });
};
const onReplica = [0, 1].map(index => (replica: CollaborationReplica | null) => {
  providers[index]?.destroy(); providers[index] = undefined;
  replicas[index] = replica;
  if (!replica) return;
  const provider = new CollaborationProvider(index ? counted(replica) : replica, transport(), { user: { name: index ? 'Joiner' : 'Writer' } });
  provider.onError(error => errors.push(error.message));
  providers[index] = provider;
  provider.connect();
});
const root = createRoot(document.getElementById('root')!);
function render() {
  const slots = joiner ? [0, 1] : [0];
  root.render(<div style={{ display: 'flex', height: '100%' }}>{slots.map(index =>
    <div key={index} data-editor={index} style={{ width: '50%', height: '100%', minWidth: 0 }}>
      <DocxEditor ref={refs[index]} documentBuffer={file} {...(index === 1 && joiner?.legacy ? { legacySyncApi: true } : {})}
        showToolbar={false} showRuler={false} showOutlineButton={false} measurementFontProvider={fonts}
        collaboration={{ clientId: 5100 + index, initialUpdate: index ? seed : undefined, onReplica: onReplica[index], presence: providers[index] }}
        onError={error => errors.push(error.message)} />
    </div>)}</div>);
}
const painted = (index: number) => Array.from(document.querySelectorAll(`[data-editor="${index}"] .canvas-page-mirror`))
  .map(host => host.textContent ?? '').join('');
const until = async (check: () => boolean) => {
  for (let attempt = 0; attempt < 600 && !check(); attempt++) await new Promise(resolve => setTimeout(resolve, 100));
  if (!check()) throw new Error('Timed out');
};

export const probe: DocxRelayJoinProbe = {
  errors,
  ready: Promise.resolve(),
  retained: () => log.snapshot().length,
  async text(index) {
    const read = await refs[index]!.current!.readParagraphs({ story: 'body', view: 'accepted' });
    if (!read.ok) throw new Error(read.failure.message);
    return read.paragraphs.map(paragraph => paragraph.text).join('\n');
  },
  async join(legacy, expected) {
    merges = 0;
    joiner = { legacy };
    const start = performance.now();
    let last = painted(1); let paints = 0; let done = 0;
    const sample = () => {
      const now = painted(1);
      if (now !== last) { last = now; paints++; }
      if (now.includes(expected)) { done = performance.now(); return; }
      requestAnimationFrame(sample);
    };
    requestAnimationFrame(sample);
    render();
    await until(() => done > 0);
    const result: DocxRelayJoinResult = { ms: Math.round(done - start), paints, merges, replayed, text: last };
    return result;
  },
};
window.__docxRelayJoin = probe;
probe.ready = (async () => {
  render();
  await until(() => replicas[0] !== null && providers[0]?.synced === true);
  seed = await ((replicas[0] as Partial<AsyncReplicaAdapter>).async ?? replicas[0]!).encodeStateAsUpdate();
})();
