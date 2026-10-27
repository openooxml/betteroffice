import { createRoot } from 'react-dom/client';
import JSZip from 'jszip';
import { PptxEditor, type PptxAsyncEditorApi } from '@betteroffice/pptx-react';
import { CollaborationProvider, attachCollaborationPipe, collaborationPipeFor, type AsyncCollaborationReplica, type CollaborationReplica,
  type CollaborationTransport, type CollaborationTransportEvent } from '@betteroffice/pptx/collaboration';
import { peerlessPainterFactories, peerlessViewStoreFactory } from '../../../packages/pptx-react/src/worker/usePeerlessPresentation';
import { RetainedUpdateLog, classifyFrame } from '../../../apps/relay/src/retention';
import { MAX_COLLABORATION_FRAME_BYTES } from '../../../shared/collaboration-limits';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import type { PptxRelayJoinProbe, PptxRelayJoinResult } from './pptx-relay-join-probe';

async function deck() {
  const zip = new JSZip();
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
    <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  </Types>`);
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${r}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file('ppt/presentation.xml', `<p:presentation xmlns:p="${p}" xmlns:r="${r}">
    <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
    <p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/>
  </p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', rels(`<Relationship Id="rId1" Type="${r}/slide" Target="slides/slide1.xml"/>`));
  zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="${p}" xmlns:a="${a}"><p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Editable text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="6400800" cy="1828800"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:solidFill><a:srgbClr val="172554"/></a:solidFill>
        <a:latin typeface="Liberation Sans"/></a:rPr><a:t>initial</a:t></a:r><a:endParaRPr sz="2400"/></a:p></p:txBody>
    </p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}

const texts = new WeakMap<HTMLCanvasElement, Set<string>>();
const drawn = new WeakMap<HTMLCanvasElement, string[]>();
const commits = [0, 0];
const editorOf = (canvas: HTMLCanvasElement) => Number(canvas.closest<HTMLElement>('[data-editor]')?.dataset.editor ?? -1);
const painter = peerlessPainterFactories.frame;
peerlessPainterFactories.frame = (canvas, options) => painter(canvas, { ...options,
  onCommit(frame, view) {
    options?.onCommit?.(frame, view);
    const index = editorOf(canvas);
    if (index < 0) return;
    commits[index]! += 1;
    texts.set(canvas, new Set(frame.displayList.primitives.flatMap((primitive) => primitive.kind === 'textBox'
      ? primitive.lines.map((line) => line.runs.map((run) => run.text).join('')) : [])));
  },
});
const nativeClear = CanvasRenderingContext2D.prototype.clearRect;
CanvasRenderingContext2D.prototype.clearRect = function (this: CanvasRenderingContext2D, ...args: Parameters<typeof nativeClear>) {
  nativeClear.apply(this, args);
  if (!this.canvas.closest('[data-legacy]')) return;
  texts.set(this.canvas, new Set());
  drawn.set(this.canvas, []);
  commits[editorOf(this.canvas)]! += 1;
};
const nativeText = CanvasRenderingContext2D.prototype.fillText;
CanvasRenderingContext2D.prototype.fillText = function (this: CanvasRenderingContext2D, ...args: Parameters<typeof nativeText>) {
  nativeText.apply(this, args);
  if (!this.canvas.closest('[data-legacy]')) return;
  // The painter may draw a run in pieces, so the run's text is what one paint drew, in order.
  const pieces = drawn.get(this.canvas);
  pieces?.push(args[0]);
  if (pieces) texts.get(this.canvas)?.add(pieces.join(''));
};
const painted = (index: number) => {
  const found = new Set<string>();
  for (const canvas of document.querySelectorAll<HTMLCanvasElement>(`[data-editor="${index}"] canvas`)) {
    for (const text of texts.get(canvas) ?? []) found.add(text);
  }
  return [...found];
};

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
      for (const response of log.responses(bytes)) deliver(socket, response);
      if (classifyFrame(bytes).kind === 'document') log.retain(bytes);
      for (const peer of sockets) if (peer !== socket) deliver(peer, bytes.slice());
      return true;
    },
    onEvent(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
  };
}

type AnyReplica = CollaborationReplica | AsyncCollaborationReplica;
const errors: string[] = [];
let applies = 0;
let pipePosts = 0;
let commandFrames = 0;
let viewPublications = 0;
let joiner: { legacy: boolean } | null = null;
const createViewStore = peerlessViewStoreFactory.create;
peerlessViewStoreFactory.create = (client, painters, options) => {
  const joining = joiner !== null;
  const installed = new WeakSet<object>();
  const store = createViewStore(client, { ...painters, paintFrame(frame, view, media) {
    if (joining && !installed.has(frame.displayList)) { commandFrames++; installed.add(frame.displayList); }
    painters.paintFrame(frame, view, media);
  } }, options);
  if (joining) store.subscribePresentation(() => { viewPublications++; });
  return store;
};
let seed: Uint8Array | undefined;
let file: Uint8Array;
let fonts: { family: string; bytes: Uint8Array }[];
let writer: PptxAsyncEditorApi | null = null;
const replicas: (AnyReplica | null)[] = [null, null];
const providers: (CollaborationProvider | undefined)[] = [];
const counted = (replica: AnyReplica): AnyReplica => {
  const counter: Record<PropertyKey, unknown> = { clientId: replica.clientId };
  for (let owner: object | null = replica; owner && owner !== Object.prototype; owner = Object.getPrototypeOf(owner)) {
    for (const key of Object.getOwnPropertyNames(owner)) {
      const value = (replica as unknown as Record<string, unknown>)[key];
      if (key === 'constructor' || key in counter || typeof value !== 'function') continue;
      counter[key] = (...args: unknown[]) => {
        if (key === 'applyUpdate' || key === 'applyUpdateChunks') applies++;
        return (value as (...args: unknown[]) => unknown).apply(replica, args);
      };
    }
  }
  const channel = collaborationPipeFor(replica);
  if (channel) attachCollaborationPipe(counter as unknown as AsyncCollaborationReplica, { open(receive, lost) {
    const endpoint = channel.open(receive, lost);
    return { post(ops, transfer) { pipePosts++; return endpoint.post(ops, transfer); }, close: () => endpoint.close() };
  } });
  return counter as unknown as AnyReplica;
};
const onReplica = [0, 1].map((index) => (replica: AnyReplica | null) => {
  providers[index]?.destroy(); providers[index] = undefined;
  replicas[index] = replica;
  if (!replica) return;
  const watched = index ? counted(replica) : replica;
  const provider = new CollaborationProvider(watched as AsyncCollaborationReplica, transport(), { user: { name: index ? 'Joiner' : 'Writer' } });
  provider.onError((error) => errors.push(error.message));
  providers[index] = provider;
  provider.connect();
});
const root = createRoot(document.getElementById('root')!);
function render() {
  const slots = joiner ? [0, 1] : [0];
  root.render(<div style={{ display: 'flex', height: '100%' }}>{slots.map((index) => {
    const collaboration = { clientId: 5200 + index, initialUpdate: index ? seed : undefined, onReplica: onReplica[index] };
    const legacy = index === 1 && joiner?.legacy;
    return <div key={index} data-editor={index} {...(legacy ? { 'data-legacy': '' } : {})} style={{ width: '50%', height: '100%', minWidth: 0 }}>
      {legacy
        ? <PptxEditor legacySyncApi file={file} fonts={fonts} clientId={5201}
          collaboration={collaboration as never} onError={(error) => errors.push(error.message)} />
        : <PptxEditor file={file} fonts={fonts} clientId={5200 + index} collaboration={collaboration as never}
          onError={(error) => errors.push(error.message)} onReady={(api) => { if (index === 0) writer = api; }} />}
    </div>;
  })}</div>);
}
const until = async (check: () => boolean, what: string) => {
  for (let attempt = 0; attempt < 1200 && !check(); attempt++) await new Promise((resolve) => setTimeout(resolve, 50));
  if (!check()) throw new Error(`Timed out waiting for ${what}`);
};

let longTaskMs = 0;
new PerformanceObserver((list) => { for (const entry of list.getEntries()) longTaskMs += entry.duration; })
  .observe({ type: 'longtask', buffered: false });

function measure(expected: string): Promise<PptxRelayJoinResult> {
  return new Promise((resolve) => {
    const start = performance.now();
    let last = ''; let paints = 0; let previous = start; let maxFrameGapMs = 0; let slowFrames = 0;
    const result = (ms: number) => ({ ms, paints, commits: commits[1]!, applies, pipePosts, commandFrames, viewPublications, replayed,
      longTaskMs: Math.round(longTaskMs), maxFrameGapMs: Math.round(maxFrameGapMs), slowFrames });
    const sample = (now: number) => {
      const gap = now - previous; previous = now;
      maxFrameGapMs = Math.max(maxFrameGapMs, gap);
      if (gap > 50) slowFrames++;
      const found = painted(1);
      const key = found.slice().sort().join('\u0000');
      if (key !== last) { last = key; paints++; }
      if (found.some((text) => text.includes(expected))) { resolve(result(Math.round(performance.now() - start))); return; }
      if (now - start > 90_000) { errors.push(`join timed out; painted ${JSON.stringify(found)}`); resolve(result(-1)); return; }
      requestAnimationFrame(sample);
    };
    requestAnimationFrame(sample);
  });
}

export const probe: PptxRelayJoinProbe = {
  errors,
  ready: Promise.resolve(),
  retained: () => log.snapshot().length,
  async edit(count) {
    let value = '';
    for (let edit = 0; edit < count; edit++) {
      const content = await writer!.readContent();
      if (!content?.ok) throw new Error('Writer content is unavailable');
      const story = content.stories[0]!;
      value = `edit ${edit}`;
      const result = await writer!.applyEdits({ expectVersion: content.version, steps: [{ op: 'replaceText',
        target: { kind: 'range', slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId, start: 0, end: story.text.length },
        text: value }] });
      if (!result?.ok) throw new Error(`Writer edit ${edit} failed: ${JSON.stringify(result)}`);
    }
    return value;
  },
  async join(legacy, expected) {
    applies = 0; pipePosts = 0; commandFrames = 0; viewPublications = 0; commits[1] = 0; longTaskMs = 0;
    joiner = { legacy };
    const result = measure(expected);
    render();
    return result;
  },
};
window.__pptxRelayJoin = probe;
probe.ready = (async () => {
  const response = await fetch(fontUrl);
  fonts = [{ family: 'Liberation Sans', bytes: new Uint8Array(await response.arrayBuffer()) }];
  file = await deck();
  await document.fonts.ready;
  render();
  await until(() => writer !== null && replicas[0] !== null && providers[0]?.synced === true, 'writer sync');
  seed = await (replicas[0] as AsyncCollaborationReplica).encodeStateAsUpdate();
})();
