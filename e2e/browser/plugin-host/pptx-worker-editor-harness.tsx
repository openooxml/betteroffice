import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { flushSync } from 'react-dom';
import JSZip from 'jszip';
import { PptxEditor, type PptxWorkerEditorApi } from '@betteroffice/pptx-react';
import {
  openPresentation, paintSlide, sizeCanvasForSlide, PptxPeerNotReadyError, PptxWorkerEditorFailedError,
  type PptxWorkerEditorAccess, type PptxWorkerEditorSession,
} from '@betteroffice/pptx';
import { createPresentationEditorSession } from '../../../packages/pptx/src/session/editorSession';
import { workerEditorSessionOpener } from '../../../packages/pptx-react/src/worker/useEditableSessionPresentation';
import { createWorkerTransport } from '../../../shared/office-session/transport';
import { isClientMessage, isHostMessage } from '../../../shared/office-session/protocol';
import { SessionFailure } from '../../../shared/office-session/types';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import { difference, installPaintProbe, pixels, until, type WorkerEditorProbe } from './pptx-worker-editor-probe';

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

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

const paints = installPaintProbe();
const mounted = deferred<void>();
let ready = deferred<void>();
const apis: PptxWorkerEditorApi[] = [];
const owners: PptxWorkerEditorSession[] = [];
const gates: { worker: Worker; hydration: boolean; replay: boolean; queued: (() => void)[]; release(): void }[] = [];
let access: PptxWorkerEditorAccess;
let api: PptxWorkerEditorApi;
let baseline: { canvas: HTMLCanvasElement; bytes: Uint8ClampedArray }[] = [];
const errors: WorkerEditorProbe['errors'] = [];

workerEditorSessionOpener.open = (bytes, options) => {
  const worker = new Worker(new URL('../../../packages/pptx/src/session/pptxSessionWorker.mjs', import.meta.url), { type: 'module' });
  const native = createWorkerTransport(worker);
  const methods = new Map<number, string>();
  const gate = { worker, hydration: owners.length === 0, replay: false, queued: [] as (() => void)[],
    release() { for (const deliver of this.queued.splice(0)) deliver(); } };
  gates.push(gate);
  const owner = createPresentationEditorSession(bytes, options ?? {}, {
    ...native,
    post(message, transfer) {
      if (isClientMessage(message) && message.kind === 'call') methods.set(message.id, message.method);
      native.post(message, transfer);
    },
    listen(listener) {
      return native.listen((message) => {
        const method = isHostMessage(message) && message.kind === 'reply' ? methods.get(message.id) : undefined;
        if ((gate.hydration && method === 'attachPeer') || (gate.replay && method === 'replay'))
          gate.queued.push(() => listener(message));
        else listener(message);
      });
    },
  });
  owners.push(owner);
  owner.subscribe((state) => { if (state.initialFrame) paints.tag(state.initialFrame); });
  const frame = owner.frame;
  owner.frame = async (slideId) => {
    const result = await frame(slideId);
    paints.tag(result);
    return result;
  };
  return owner;
};

const owner = () => owners.at(-1)!;
const gate = () => gates.at(-1)!;
const main = () => document.querySelector<HTMLCanvasElement>('[data-testid="pptx-slide-canvas"]')!;
const overlay = () => main()?.parentElement?.querySelector<HTMLCanvasElement>('canvas[aria-hidden="true"]') ?? null;
const story = () => {
  const result = access.readContent();
  if (!result.ok || !result.stories[0]) throw new Error('Editable story is missing');
  return result.stories[0];
};
const text = () => story().text;
const textBox = () => {
  const box = access.layoutSlide(0).primitives.find((primitive) => primitive.kind === 'textBox');
  if (!box || box.kind !== 'textBox') throw new Error('Text geometry is missing');
  return box;
};
const caret = (position: number) => {
  const line = textBox().lines.find((line) => line.caretStops.some((stop) => stop.position === position))!;
  return { line, x: line.caretStops.find((stop) => stop.position === position)!.x };
};

async function reopen(bytes: Uint8Array) {
  const opened = openPresentation(bytes, { fonts });
  try {
    const read = opened.readContent();
    if (!read.ok) throw new Error('Reopened content is unavailable');
    return { bytes: bytes.byteLength, text: read.stories[0].text };
  } finally { opened.dispose(); }
}

async function rejection(operation: () => unknown) {
  try { await operation(); return 'accepted'; }
  catch (error) { return error instanceof Error ? error.name : String(error); }
}

const root = createRoot(document.getElementById('worker-editor-root')!);
let file: Uint8Array;
let fonts: { family: string; bytes: Uint8Array }[];
let readOnly = false;
const render = () => root.render(<div data-editor>
  <PptxEditor file={file} fonts={fonts} experimentalWorkerOpen clientId={42} readOnly={readOnly}
    onError={(error) => errors.push({ name: error.name,
      code: error instanceof PptxWorkerEditorFailedError || error instanceof PptxPeerNotReadyError ? error.code : '',
      cause: error instanceof PptxWorkerEditorFailedError && error.cause instanceof SessionFailure ? error.cause.code : '',
      typed: error instanceof PptxWorkerEditorFailedError })}
    onReady={(value) => {
      if (!('handleAsync' in value)) throw new Error('Worker editor session was replaced by a viewer');
      api = value;
      apis.push(value);
      void value.handleAsync().then((peer) => { access = peer; ready.resolve(); });
    }} />
</div>);

const probe: WorkerEditorProbe = {
  mounted: mounted.promise,
  get ready() { return ready.promise; },
  errors,
  state: () => ({ hydrated: owner().hydrated, sequence: owner().state.sequence,
    acknowledged: owner().state.acknowledgedSequence, held: gate().queued.length }),
  refuse: () => {
    try { owner().apply({ method: 'insertText', args: ['missing', 0, 'refused'] }); return 'accepted'; }
    catch (error) { return error instanceof Error ? error.name : String(error); }
  },
  releaseHydration: async () => {
    gate().hydration = false;
    gate().release();
    await owner().whenHydrated();
    await ready.promise;
    await api.whenHydrated();
    await api.commands.execute('zoom', { scale: 1 });
  },
  hold: () => { gate().replay = true; },
  release: async () => {
    gate().replay = false;
    gate().release();
    await api.flushPendingInput();
  },
  text,
  select: (start, end) => {
    const { shapeId, storyId } = story();
    return api.selectText({ slide: 1, shapeId, storyId, start, end });
  },
  point: (position) => {
    const rect = main().getBoundingClientRect();
    const { line, x } = caret(position);
    const scale = rect.width / access.layoutSlide(0).width;
    const point = { x: rect.left + (x + 0.1) * scale, y: rect.top + (line.y + line.height / 2) * scale };
    const hit = api.getPositionAtPoint(point.x, point.y);
    return { ...point, position: hit?.kind === 'text' ? hit.position : null };
  },
  baseline: async () => {
    const canvases = await until(() => {
      const canvases = Array.from(document.querySelectorAll<HTMLCanvasElement>('[data-testid="pptx-slide-canvas"], aside canvas'));
      return canvases.length === 2 && canvases.every((canvas) => paints.paints.get(canvas)?.finished) ? canvases : undefined;
    }, 'baseline paint completion');
    baseline = canvases.map((canvas) => ({ canvas, bytes: pixels(canvas).slice() }));
  },
  unchanged: () => baseline.length === 2 && baseline.every(({ canvas, bytes }) =>
    canvas.isConnected && difference(bytes, pixels(canvas)).differingPixels === 0),
  parity: async () => {
    const targets = await until(() => {
      const canvases = Array.from(document.querySelectorAll<HTMLCanvasElement>('[data-testid="pptx-slide-canvas"], aside canvas'));
      return canvases.length === 2 && canvases.every((canvas) => {
        const paint = paints.paints.get(canvas);
        return paint?.finished && paint.frame?.sequence === owner().state.sequence;
      }) ? canvases : undefined;
    }, 'slide and thumbnail adoption');
    return Promise.all(targets.map(async (canvas) => {
      const paint = paints.paints.get(canvas)!;
      const frame = paint.frame!;
      const expected = document.createElement('canvas');
      const scale = parseFloat(canvas.style.width) / frame.displayList.width;
      sizeCanvasForSlide(expected, frame.displayList, devicePixelRatio, scale);
      if (expected.width !== canvas.width || expected.height !== canvas.height) throw new Error('Backing stores differ');
      await paintSlide(expected.getContext('2d')!, frame.displayList, devicePixelRatio, scale);
      return { thumbnail: !!canvas.closest('aside'), dpr: devicePixelRatio, sequence: frame.sequence,
        ...difference(pixels(expected), pixels(canvas)) };
    }));
  },
  overlay: async (position) => {
    const canvas = await until(() => {
      const canvas = overlay();
      return canvas && pixels(canvas).some((value, index) => index % 4 === 3 && value > 0) ? canvas : undefined;
    }, 'visible peer overlay');
    const expected = document.createElement('canvas');
    expected.width = canvas.width;
    expected.height = canvas.height;
    const ctx = expected.getContext('2d')!;
    const scale = devicePixelRatio * (parseFloat(canvas.style.width) / access.layoutSlide(0).width);
    ctx.setTransform(scale, 0, 0, scale, 0, 0);
    ctx.fillStyle = '#1d4ed8';
    const { line, x } = caret(position);
    ctx.fillRect(x, line.y, 1.5, line.height);
    const actual = pixels(canvas);
    return { differingPixels: difference(pixels(expected), actual).differingPixels,
      pixels: actual.filter((value, index) => index % 4 === 3 && value > 0).length };
  },
  provenance: () => ({ total: paints.history.length,
    thumbnails: paints.history.filter((paint) => paint.thumbnail).length,
    unknown: paints.history.filter((paint) => !paint.worker).length }),
  save: async () => reopen(await api.saveAsync()),
  toggleReadOnly: async (value) => {
    readOnly = value;
    flushSync(render);
    await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    return { sessions: owners.length, ready: apis.length };
  },
  edit: async (value) => {
    const { slideId, shapeId, storyId } = story();
    const position = text().length;
    const result = await api.applyEdits({ expectVersion: await api.version(),
      steps: [{ op: 'insertText', target: { kind: 'range', slideId, shapeId, storyId,
        start: position, end: position }, at: 'end', text: value }] });
    return result.ok ? 'accepted' : result.failure.code;
  },
  proposal: async () => {
    access.propose('browser', 'Review', [{ type: 'replaceText', storyId: story().storyId, start: 0, end: 1, text: '?' }]);
    api.refreshProposals();
    await api.flushPendingInput();
  },
  fail: () => {
    gate().worker.terminate();
    void api.saveAsync().catch(() => {});
  },
  rejected: async (index) => {
    const retained = apis[index];
    return Promise.all([
      () => retained.saveAsync(), () => retained.handleAsync(), () => retained.readContent(),
      () => retained.applyEdits({ expectVersion: 'retired', steps: [] }),
      () => retained.selectText({ slide: 1, shapeId: 's', storyId: 's', start: 0, end: 0 }),
      () => retained.commands.execute('undo', null), () => retained.flushPendingInput(),
    ].map(rejection));
  },
  recover: async (index) => {
    const saved = await apis[index].recoverySave();
    return { ...await reopen(saved.bytes), recovery: saved.recovery };
  },
  replace: async () => {
    ready = deferred<void>();
    file = file.slice();
    render();
    await ready.promise;
    await api.commands.execute('zoom', { scale: 1 });
  },
  unmount: async () => {
    root.unmount();
    await until(() => owner().state.stage === 'disposed' ? true : undefined, 'owner disposal');
  },
};
window.__pptxWorkerEditor = probe;
void (async () => {
  const response = await fetch(fontUrl);
  if (!response.ok) throw new Error(`Font fetch failed: ${response.status}`);
  fonts = [{ family: 'Liberation Sans', bytes: new Uint8Array(await response.arrayBuffer()) }];
  file = await deck();
  await document.fonts.ready;
  render();
  await until(() => gate()?.queued.length ? true : undefined, 'held hydration acknowledgement');
  await probe.baseline();
  mounted.resolve();
})();
