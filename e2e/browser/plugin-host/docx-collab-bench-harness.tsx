import { createRoot } from 'react-dom/client';
import JSZip from 'jszip';
import { DocxEditor } from '@betteroffice/docx-react';
import { CollaborationProvider, type CollaborationReplica, type CollaborationTransport,
  type CollaborationTransportEvent } from '@betteroffice/docx/collaboration';
import { decodeMessages, encodeSyncStep2, encodeUpdate, parseSyncRequest } from '@betteroffice/docx/collaboration/protocol';
import { createYrsSession } from '@betteroffice/docx/yrs';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import type { DocxCollabBenchProbe, DocxCollabBenchResult } from './docx-collab-bench-probe';
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

const file = await syntheticDocument('Seed');
const peer = await createYrsSession({ clientId: 6002 });
peer.openDocx(file, true);
const seed = peer.encodeStateAsUpdate();
const paraId = peer.paragraphs('body')[0]!.paraId;
const peerText = () => peer.paragraphs('body')[0]!.text;
const errors: string[] = [];
let dispatchMs = 0;
let answered = false;

/** The peer's next `count` edits, one Update frame each. */
function frames(count: number): Uint8Array[] {
  const updates: Uint8Array[] = [];
  const off = peer.onUpdate((update) => updates.push(update.slice()));
  for (let index = 0; index < count; index++) {
    peer.insertText({ story: 'body', paraId, offset: peerText().length }, String.fromCharCode(97 + (index % 26)));
  }
  off();
  return updates.map((update) => encodeUpdate(update));
}

const listeners = new Set<(event: CollaborationTransportEvent) => void>();
const emit = (event: CollaborationTransportEvent) => {
  const started = performance.now();
  for (const listener of [...listeners]) listener(event);
  dispatchMs += performance.now() - started;
};
let onSyncRequest: ((stateVector: Uint8Array) => void) | null = null;
const transport: CollaborationTransport = {
  connect() {},
  disconnect() {},
  send(data) {
    if (!answered) {
      for (const message of decodeMessages(data)) {
        if (message.type === 'sync-step-1') onSyncRequest?.(parseSyncRequest(message.stateVector).stateVector);
      }
    }
    return true;
  },
  onEvent(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
};

let provider: CollaborationProvider | undefined;
const onReplica = (replica: CollaborationReplica | null) => {
  provider?.destroy(); provider = undefined;
  if (!replica) return;
  provider = new CollaborationProvider(replica, transport);
  provider.onError((error) => errors.push(error.message));
  provider.connect();
  emit({ type: 'open' });
};

const painted = () => Array.from(document.querySelectorAll('.canvas-page-mirror')).map((host) => host.textContent ?? '').join('');

function measure(expected: string, run: () => void, frameCount: number): Promise<DocxCollabBenchResult> {
  return new Promise((resolve) => {
    dispatchMs = 0;
    let last = painted(); let paints = 0;
    const started = performance.now();
    const sample = () => {
      const now = painted();
      if (now !== last) { last = now; paints++; }
      if (now.includes(expected) && provider?.synced) {
        resolve({ frames: frameCount, ms: Math.round(performance.now() - started), paints, dispatchMs: Math.round(dispatchMs * 10) / 10 });
        return;
      }
      requestAnimationFrame(sample);
    };
    requestAnimationFrame(sample);
    run();
  });
}

const deliver = (queue: Uint8Array[], paceMs: number, then?: () => void) => {
  const next = () => {
    const frame = queue.shift();
    if (!frame) { then?.(); return; }
    emit({ type: 'message', data: frame });
    setTimeout(next, paceMs);
  };
  setTimeout(next, 0);
};

const root = createRoot(document.getElementById('root')!);
export const probe: DocxCollabBenchProbe = {
  errors,
  ready: Promise.resolve(),
  async join(edits) {
    const replay = frames(edits);
    const expected = peerText();
    const request = new Promise<Uint8Array>((resolve) => {
      onSyncRequest = (stateVector) => { onSyncRequest = null; resolve(stateVector); };
    });
    root.render(<div style={{ height: '100%' }}>
      <DocxEditor documentBuffer={file} showToolbar={false} showRuler={false} showOutlineButton={false}
        measurementFontProvider={fonts} collaboration={{ clientId: 6001, initialUpdate: seed, onReplica }}
        onError={(error) => errors.push(error.message)} />
    </div>);
    const stateVector = await request;
    return measure(expected, () => deliver([...replay], 0, () => {
      answered = true;
      emit({ type: 'message', data: encodeSyncStep2(peer.encodeStateAsUpdate(stateVector)) });
    }), replay.length);
  },
  type(edits, paceMs) {
    const queue = frames(edits);
    return measure(peerText(), () => deliver(queue, paceMs), queue.length);
  },
};
window.__docxCollabBench = probe;
