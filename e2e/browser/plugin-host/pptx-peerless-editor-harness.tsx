import { createElement } from 'react';
import { flushSync } from 'react-dom';
import type { Root } from 'react-dom/client';
import JSZip from 'jszip';
import { EditorToolbar, PptxEditor, type PptxAsyncEditorApi, type PptxAsyncEditorProps, type PptxAsyncCollaborationOptions } from '@betteroffice/pptx-react';
import type { AsyncCollaborationReplica, PptxFontFace, PresentationCommandClient } from '@betteroffice/pptx';
import { peerlessPainterFactories, peerlessSessionOpener } from '../../../packages/pptx-react/src/worker/usePeerlessPresentation';
import type { PeerlessConfirmedFrame } from '../../../packages/pptx-react/src/worker/peerlessViewStore';
import type { PeerlessEditorProbe } from './pptx-peerless-editor-probe';

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

interface ReplyGate { held: boolean; queued: (() => void)[] }

function gatedWorker(gate: ReplyGate): Worker {
  const worker = new Worker(new URL('../../../packages/pptx/src/session/pptxSessionWorker.mjs', import.meta.url), { type: 'module' });
  const add = worker.addEventListener.bind(worker);
  const remove = worker.removeEventListener.bind(worker);
  const wrappers = new Map<EventListenerOrEventListenerObject, EventListener>();
  worker.addEventListener = (type: string, listener: EventListenerOrEventListenerObject, options?: boolean | AddEventListenerOptions) => {
    if (type !== 'message' || !listener) { add(type, listener, options); return; }
    const wrapped: EventListener = (event) => {
      const deliver = () => typeof listener === 'function' ? listener.call(worker, event) : listener.handleEvent(event);
      if (gate.held) gate.queued.push(deliver); else deliver();
    };
    wrappers.set(listener, wrapped);
    add(type, wrapped, options);
  };
  worker.removeEventListener = (type: string, listener: EventListenerOrEventListenerObject, options?: boolean | EventListenerOptions) => {
    remove(type, listener && wrappers.get(listener) || listener, options);
    if (listener) wrappers.delete(listener);
  };
  return worker;
}

async function twoSlides(bytes: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const xml = async (path: string) => zip.file(path)!.async('string');
  zip.file('[Content_Types].xml', (await xml('[Content_Types].xml')).replace('</Types>',
    '<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'));
  zip.file('ppt/presentation.xml', (await xml('ppt/presentation.xml')).replace('</p:sldIdLst>',
    '<p:sldId id="257" r:id="rId2"/></p:sldIdLst>'));
  zip.file('ppt/_rels/presentation.xml.rels', (await xml('ppt/_rels/presentation.xml.rels')).replace('</Relationships>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>'));
  zip.file('ppt/slides/slide2.xml', (await xml('ppt/slides/slide1.xml')).replace('>initial<', '>other<'));
  return zip.generateAsync({ type: 'uint8array' });
}

interface EditorSlot {
  file: Uint8Array;
  ready: ReturnType<typeof deferred<PptxAsyncEditorApi>>;
  api: PptxAsyncEditorApi | null;
  client: PresentationCommandClient | null;
  replica: AsyncCollaborationReplica | null;
  collaboration: PptxAsyncCollaborationOptions | undefined;
  saved: Uint8Array | null;
  generation: number;
  gate: ReplyGate;
}

export async function mountPeerlessEditor(root: Root, bytes: Uint8Array, fonts: PptxFontFace[]) {
  const file = await twoSlides(bytes);
  const count = new URLSearchParams(location.search).has('collaboration') ? 2 : 1;
  const errors: string[] = [];
  const saves: number[][] = [];
  let readOnly = false;
  let retainToolbar = false;
  const slots: EditorSlot[] = Array.from({ length: count }, (_, index) => {
    const slot: EditorSlot = { file, ready: deferred<PptxAsyncEditorApi>(), api: null, client: null,
      replica: null, collaboration: undefined, saved: null, generation: 0, gate: { held: false, queued: [] } };
    if (count === 2) slot.collaboration = { clientId: 4200 + index, onReplica: (value) => { slot.replica = value; } };
    return slot;
  });
  const frames = new WeakMap<HTMLCanvasElement, PeerlessConfirmedFrame>();
  const painter = peerlessPainterFactories.frame;
  peerlessPainterFactories.frame = (canvas, options) => painter(canvas, { ...options,
    onCommit(frame, view) {
      options?.onCommit?.(frame, view);
      frames.set(canvas, frame);
    },
  });
  const opener = peerlessSessionOpener.open;
  peerlessSessionOpener.open = (source, options) => {
    const slot = slots[(options?.clientId ?? 4200) - 4200]!;
    const generation = ++slot.generation;
    slot.gate = { held: false, queued: [] };
    const gate = slot.gate;
    const session = opener(source, { ...options, worker: () => gatedWorker(gate) });
    void session.ready.then((client) => { if (generation === slot.generation) slot.client = client; }).catch(() => {});
    return session;
  };
  const render = () => root.render(createElement('div', { style: { display: 'flex', height: '100%' } },
    slots.map((slot, index) => createElement('div', { key: index, 'data-editor': true, 'data-peerless-editor': index,
      style: { width: `${100 / count}%`, height: '100%', minWidth: 0 } },
    createElement<PptxAsyncEditorProps>(PptxEditor, { file: slot.file, fonts, clientId: 4200 + index, collaboration: slot.collaboration,
      readOnly, toolbar: retainToolbar ? createElement(EditorToolbar, { mode: 'commands' }) : undefined,
      onError: (error: Error) => errors.push(error.message), onSave: (saved: Uint8Array) => { slot.saved = saved; saves.push(Array.from(saved)); },
      onReady: (api: PptxAsyncEditorApi) => { slot.api = api; slot.ready.resolve(api); },
    })))));
  const slot = (index = 0) => slots[index]!;
  const api = (index = 0) => slot(index).api!;
  const client = (index = 0) => slot(index).client!;
  const canvas = (index = 0) => document.querySelector<HTMLCanvasElement>(
    `[data-peerless-editor="${index}"] [data-testid="pptx-slide-canvas"]`)!;
  const frame = (index = 0) => {
    const value = frames.get(canvas(index));
    if (!value) throw new Error('Worker frame has not been painted');
    return value;
  };
  const box = (index = 0) => {
    const value = frame(index).displayList.primitives.find((primitive) => primitive.kind === 'textBox');
    if (!value || value.kind !== 'textBox') throw new Error('Worker text geometry is missing');
    return value;
  };
  const content: PeerlessEditorProbe['content'] = async (index = 0) => {
    const result = await api(index).readContent();
    if (!result.ok) throw new Error(result.failure.message);
    return result;
  };
  let remoteMerge: Promise<unknown> | undefined;
  const probe: PeerlessEditorProbe = {
    ready: Promise.all(slots.map((value) => value.ready.promise)).then(() => {}), errors,
    mainThreadEngines: () => globalThis.__pptxMainEngineProbe!.count(),
    generation: (index) => slot(index).generation,
    hydrated: (index) => api(index).hydrated,
    saves,
    setReadOnly(value, toolbar = false) { readOnly = value; retainToolbar = toolbar; flushSync(render); },
    content,
    selection: (index) => client(index).state.selection,
    activeSlide: (index) => client(index).state.slides.findIndex((slide) => slide.id === client(index).state.activeSlideId) + 1,
    paintedTexts: (index) => frame(index).displayList.primitives.flatMap((primitive) => primitive.kind === 'textBox'
      ? primitive.lines.map((line) => line.runs.map((run) => run.text).join('')) : []),
    point(position, index = 0) {
      const line = box(index).lines.find((line) => line.caretStops.some((stop) => stop.position === position));
      const stop = line?.caretStops.find((stop) => stop.position === position);
      if (!line || !stop) throw new Error('Worker caret geometry is missing');
      const bounds = canvas(index).getBoundingClientRect();
      const scale = bounds.width / frame(index).displayList.width;
      return { x: bounds.left + (stop.x + 0.1) * scale, y: bounds.top + (line.y + line.height / 2) * scale };
    },
    shapePoint(index = 0) {
      const bounds = canvas(index).getBoundingClientRect();
      const scale = bounds.width / frame(index).displayList.width;
      const textBox = box(index);
      return { x: bounds.left + (textBox.x + 1) * scale, y: bounds.top + (textBox.y + textBox.h / 2) * scale,
        units: { x: client(index).state.size.width / bounds.width, y: client(index).state.size.height / bounds.height } };
    },
    async select(start, end, index = 0) {
      const story = (await content(index)).stories.find((story) => story.slideId === client(index).state.activeSlideId);
      if (!story) throw new Error('Active slide story is missing');
      return api(index).selectText({ slide: probe.activeSlide(index), shapeId: story.shapeId, storyId: story.storyId, start, end });
    },
    flush: (index) => api(index).flushPendingInput(),
    hold(index) { slot(index).gate.held = true; },
    release(index) {
      const gate = slot(index).gate;
      gate.held = false;
      for (const deliver of gate.queued.splice(0)) deliver();
    },
    held: (index) => slot(index).gate.queued.length,
    async saveAndReopen(index = 0) {
      const value = slot(index);
      value.saved = null;
      const result = await api(index).commands.execute('save', null);
      if (!result.ok || !value.saved) throw new Error('Save returned no bytes');
      value.file = value.saved;
      value.ready = deferred<PptxAsyncEditorApi>();
      value.api = null;
      render();
      await value.ready.promise;
      await api(index).whenHydrated();
      return value.file.byteLength;
    },
    async merge() {
      const [left, right] = slots.map((value) => value.replica!);
      const [leftVector, rightVector] = await Promise.all([left!.encodeStateVector(), right!.encodeStateVector()]);
      const [leftUpdate, rightUpdate] = await Promise.all([
        left!.encodeStateAsUpdate(rightVector), right!.encodeStateAsUpdate(leftVector),
      ]);
      await Promise.all([left!.applyUpdate(rightUpdate), right!.applyUpdate(leftUpdate)]);
      await Promise.all(slots.map((value) => value.api!.flushPendingInput()));
    },
    async queueRemoteMerge() {
      const update = await slot(1).replica!.encodeStateAsUpdate();
      remoteMerge = slot(0).replica!.applyUpdate(update);
      void remoteMerge.catch(() => {});
    },
    async settleRemoteMerge() { await remoteMerge; await api(0).flushPendingInput(); },
  };
  window.__pptxPeerlessEditor = probe;
  render();
  await probe.ready;
}
