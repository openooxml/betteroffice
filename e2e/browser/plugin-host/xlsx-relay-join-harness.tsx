import { createRoot } from 'react-dom/client';
import JSZip from 'jszip';
import { XlsxEditor } from '@betteroffice/xlsx-react';
import { CollaborationProvider, createAsyncCollaborationReplica, type AsyncReplicaAdapter, type CollaborationReplica,
  type CollaborationTransport, type CollaborationTransportEvent } from '@betteroffice/xlsx/collaboration';
import { RetainedUpdateLog } from '../../../apps/relay/src/retention';
import { classifyFrame } from '../../../apps/relay/src/retention';
import { MAX_COLLABORATION_FRAME_BYTES } from '../../../shared/collaboration-limits';
import type { XlsxRelayJoinProbe, XlsxRelayJoinResult } from './xlsx-relay-join-probe';

async function workbook() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetFormatPr defaultRowHeight="20"/><cols><col min="1" max="1" width="20" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>seed</t></is></c></row></sheetData></worksheet>');
  return zip.generateAsync({ type: 'uint8array' });
}

const paints = new WeakMap<HTMLCanvasElement, Set<string>>();
const clears = [0, 0];
const nativeClear = CanvasRenderingContext2D.prototype.clearRect;
CanvasRenderingContext2D.prototype.clearRect = function (this: CanvasRenderingContext2D, ...args: Parameters<typeof nativeClear>) {
  nativeClear.apply(this, args);
  const host = this.canvas.closest<HTMLElement>('[data-editor]');
  if (!host || this.canvas.dataset.paintSource !== 'worker' && !this.canvas.closest('[data-legacy]')) return;
  paints.set(this.canvas, new Set());
  clears[Number(host.dataset.editor)]! += 1;
};
const nativeText = CanvasRenderingContext2D.prototype.fillText;
CanvasRenderingContext2D.prototype.fillText = function (this: CanvasRenderingContext2D, ...args: Parameters<typeof nativeText>) {
  nativeText.apply(this, args);
  paints.get(this.canvas)?.add(args[0]);
};
const painted = (index: number) => {
  const texts = new Set<string>();
  for (const canvas of document.querySelectorAll<HTMLCanvasElement>(`[data-editor="${index}"] canvas`)) {
    for (const text of paints.get(canvas) ?? []) texts.add(text);
  }
  return [...texts];
};

const log = new RetainedUpdateLog(512, MAX_COLLABORATION_FRAME_BYTES);
const sockets = new Set<(event: CollaborationTransportEvent) => void>();
const deliver = (socket: (event: CollaborationTransportEvent) => void, data: Uint8Array) =>
  setTimeout(() => { if (sockets.has(socket)) socket({ type: 'message', data }); }, 0);
let replayed = 0;
let lastWriterSend = 0;
function transport(index: number): CollaborationTransport {
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
      if (kind === 'document') {
        log.retain(bytes);
        if (index === 0) lastWriterSend = performance.now();
      }
      for (const peer of sockets) if (peer !== socket) deliver(peer, bytes.slice());
      return true;
    },
    onEvent(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
  };
}

const file = await workbook();
const errors: string[] = [];
let applies = 0;
let joiner: { legacy: boolean } | null = null;
let seed: Uint8Array | undefined;
const replicas: (CollaborationReplica | null)[] = [null, null];
const providers: (CollaborationProvider | undefined)[] = [];
const counted = (replica: CollaborationReplica): CollaborationReplica => {
  const hosted = (replica as Partial<AsyncReplicaAdapter>).async;
  if (hosted) return createAsyncCollaborationReplica({ ...hosted, clientId: hosted.clientId,
    encodeStateVector: () => hosted.encodeStateVector(), encodeStateAsUpdate: (vector) => hosted.encodeStateAsUpdate(vector),
    onUpdate: (listener) => hosted.onUpdate(listener),
    applyUpdate(update) { applies++; return hosted.applyUpdate(update); },
    ...(hosted.applyUpdateChunks && { applyUpdateChunks(updates: readonly Uint8Array[]) { applies++; return hosted.applyUpdateChunks!(updates); } }) });
  return Object.assign(Object.create(replica) as CollaborationReplica, {
    applyUpdate(update: Uint8Array) { applies++; return replica.applyUpdate(update); },
  });
};
const onReplica = [0, 1].map(index => (replica: CollaborationReplica | null) => {
  providers[index]?.destroy(); providers[index] = undefined;
  replicas[index] = replica;
  if (!replica) return;
  const provider = new CollaborationProvider(index ? counted(replica) : replica, transport(index), { user: { name: index ? 'Joiner' : 'Writer' } });
  provider.onError(error => errors.push(error.message));
  providers[index] = provider;
  provider.connect();
});
const root = createRoot(document.getElementById('root')!);
function render() {
  const slots = joiner ? [0, 1] : [0];
  root.render(<div style={{ display: 'flex', height: '100%' }}>{slots.map(index => {
    const collaboration = { clientId: 5200 + index, initialUpdate: index ? seed : undefined, onReplica: onReplica[index] };
    return <div key={index} data-editor={index} {...(index === 1 && joiner?.legacy ? { 'data-legacy': '' } : {})}
      style={{ width: '50%', height: '100%', minWidth: 0 }}>
      {index === 1 && joiner?.legacy
        ? <XlsxEditor legacySyncApi file={file} collaboration={collaboration} onError={error => errors.push(error.message)} />
        : <XlsxEditor file={file} collaboration={collaboration} onError={error => errors.push(error.message)} />}
    </div>;
  })}</div>);
}
const until = async (check: () => boolean) => {
  for (let attempt = 0; attempt < 1200 && !check(); attempt++) await new Promise(resolve => setTimeout(resolve, 50));
  if (!check()) throw new Error(`Timed out; painted ${JSON.stringify(painted(1))}`);
};

let longTaskMs = 0;
new PerformanceObserver((list) => { for (const entry of list.getEntries()) longTaskMs += entry.duration; })
  .observe({ type: 'longtask', buffered: false });

function measure(expected: string[], from: () => number): Promise<XlsxRelayJoinResult> {
  return new Promise((resolve) => {
    const start = performance.now();
    let last = ''; let paints = 0; let previous = start; let maxFrameGapMs = 0; let slowFrames = 0;
    const sample = (now: number) => {
      const gap = now - previous; previous = now;
      maxFrameGapMs = Math.max(maxFrameGapMs, gap);
      if (gap > 50) slowFrames++;
      const texts = painted(1);
      const key = texts.slice().sort().join('\u0000');
      if (key !== last) { last = key; paints++; }
      if (expected.every(value => texts.includes(value))) {
        resolve({ ms: Math.round(performance.now() - from()), paints, clears: clears[1]!, applies, replayed,
          longTaskMs: Math.round(longTaskMs), maxFrameGapMs: Math.round(maxFrameGapMs), slowFrames });
        return;
      }
      if (now - start > 60_000) { errors.push(`watch timed out; painted ${JSON.stringify(texts)}`); resolve({ ms: -1, paints, clears: clears[1]!, applies, replayed, longTaskMs: Math.round(longTaskMs), maxFrameGapMs: Math.round(maxFrameGapMs), slowFrames }); return; }
      requestAnimationFrame(sample);
    };
    requestAnimationFrame(sample);
  });
}

export const probe: XlsxRelayJoinProbe = {
  errors,
  ready: Promise.resolve(),
  retained: () => log.snapshot().length,
  painted,
  mark() { applies = 0; clears[1] = 0; longTaskMs = 0; },
  async join(legacy, expected) {
    probe.mark();
    joiner = { legacy };
    const start = performance.now();
    const result = measure(expected, () => start);
    render();
    return result;
  },
  watch(expected) { return measure(expected, () => lastWriterSend); },
};
window.__xlsxRelayJoin = probe;
probe.ready = (async () => {
  await document.fonts.ready;
  render();
  await until(() => replicas[0] !== null && providers[0]?.synced === true);
  seed = await ((replicas[0] as Partial<AsyncReplicaAdapter>).async ?? replicas[0]!).encodeStateAsUpdate();
})();
