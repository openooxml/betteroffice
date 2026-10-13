import './peerlessTestCapabilities';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { Suspense } from 'react';
import type { PeerlessSession, PresentationCommandClient, PresentationCommandEnvelope,
  PresentationCommandEvents, PresentationCommandMethods, TextBoxPrimitive } from '@betteroffice/pptx';
import { decodeCommandDisplayList } from '@betteroffice/pptx';
import { createSessionHost } from '../../../../shared/office-session/host';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { createPeerlessSession } from '../../../pptx/src/session/commandSession';
import { registerPresentationCommandHandlers } from '../../../pptx/src/session/commandHost';
import * as wasm from '../../../pptx/src/wasm/loader';
import { PptxEditor } from '../PptxEditor';
import { peerlessSessionOpener } from './usePeerlessPresentation';
import type { PptxAsyncEditorApi } from './createPeerlessEditorApi';
import type { PptxAsyncEditorProps } from '../PptxEditor';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor, within } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
let client: PresentationCommandClient;
let commands: PresentationCommandEnvelope[];
let owner: PeerlessSession | undefined;
let holding = false;
let unknownNext = false;
let held: unknown[] = [];
let releaseReplies = () => {};
let crashWorker = () => {};
let suspensions = 0;
const suspended = new Promise<void>(() => {});

function Suspend({ active }: { active: boolean }) {
  if (active) { suspensions++; throw suspended; }
  return null;
}

function bounded<T>(promise: Promise<T>): Promise<T> {
  let timer: ReturnType<typeof setTimeout>;
  return Promise.race([promise, new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error('Command timed out')), 5000);
  })]).finally(() => clearTimeout(timer));
}

export async function deck(small = false): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="2400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${['alpha beta', 'x', 'alpha beta'].map((text) => `<a:p><a:r><a:rPr sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>${text}</a:t></a:r></a:p>`).join('')}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  if (small) {
    const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
    zip.file('ppt/slides/slide1.xml', slide.replace('cx="6000000" cy="2400000"', 'cx="40000" cy="40000"'));
  }
  zip.file('ppt/slides/slide2.xml', await zip.file('ppt/slides/slide1.xml')!.async('string'));
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>',
    '<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'));
  zip.file('ppt/presentation.xml', (await zip.file('ppt/presentation.xml')!.async('string')).replace('</p:sldIdLst>',
    '<p:sldId id="257" r:id="rId2"/></p:sldIdLst>'));
  zip.file('ppt/_rels/presentation.xml.rels', (await zip.file('ppt/_rels/presentation.xml.rels')!.async('string')).replace('</Relationships>',
    `<Relationship Id="rId2" Type="${rel}/slide" Target="slides/slide2.xml"/></Relationships>`));
  return zip.generateAsync({ type: 'uint8array' });
}

beforeEach(() => {
  const originalRect = HTMLElement.prototype.getBoundingClientRect;
  const toolbarRect = spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
    const rect = originalRect.call(this);
    return this.getAttribute('role') === 'toolbar' ? { ...rect.toJSON(), width: 4000 } as DOMRect : rect;
  });
  restorers.push(() => toolbarRect.mockRestore());
  commands = [];
  holding = false; unknownNext = false; held = []; suspensions = 0;

  owner = undefined;
  const worker = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
  Object.defineProperty(globalThis, 'Worker', { configurable: true, value: class {} });
  restorers.push(() => {
    if (worker) Object.defineProperty(globalThis, 'Worker', worker);
    else Reflect.deleteProperty(globalThis, 'Worker');
  });
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation((function (this: HTMLCanvasElement) {
    return new Proxy({ canvas: this, measureText: (text: string) => ({ width: text.length * 10 }) }, {
      get(target, key) { return key in target ? Reflect.get(target, key) : () => {}; },
    }) as unknown as CanvasRenderingContext2D;
  }) as unknown as HTMLCanvasElement['getContext']);
  restorers.push(() => context.mockRestore());
  for (const [property, extent] of [['clientWidth', 1000], ['clientHeight', 580]] as const) {
    const original = Object.getOwnPropertyDescriptor(HTMLElement.prototype, property);
    Object.defineProperty(HTMLElement.prototype, property, { configurable: true,
      get(this: HTMLElement) { return extent; } });
    restorers.push(() => {
      if (original) Object.defineProperty(HTMLElement.prototype, property, original);
      else Reflect.deleteProperty(HTMLElement.prototype, property);
    });
  }
  const opener = spyOn(peerlessSessionOpener, 'open').mockImplementation((bytes, options = {}) => {
    const pair = createInProcessPair();
    const registration = registerPresentationCommandHandlers({ initWasm: async () => {
      await wasm.initWasm(await readFile(new URL('../../../pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
    } });
    const command = registration.handlers.command;
    const received = spyOn(registration.handlers, 'command').mockImplementation((context, envelope) => {
      commands.push(envelope);
      return command(context, envelope);
    });
    restorers.push(() => received.mockRestore());
    createSessionHost<PresentationCommandMethods, PresentationCommandEvents, null>(pair.host, {
      handlers: registration.handlers, policies: registration.policies, context: null, onDispose: registration.dispose,
    });
    const commandIds = new Set<number>();
    const endpoint = new class extends EventTarget {
      postMessage(message: unknown, transfer?: Transferable[]) {
        const call = message as { kind: string; method: string; id: number };
        if (call.kind === 'call' && (call.method === 'command' || call.method === 'edit')) commandIds.add(call.id);
        pair.client.post(message, transfer);
      }
      terminate() { pair.client.close(); }
    };
    const deliver = (message: unknown) => endpoint.dispatchEvent(new MessageEvent('message', { data: message }));
    crashWorker = () => { endpoint.dispatchEvent(Object.assign(new Event('error'), { message: 'Worker stopped' })); };
    releaseReplies = () => { holding = false; for (const message of held.splice(0)) (message as () => void)(); };
    const off = pair.client.listen((message) => {
      const reply = message as { kind: string; id: number };
      if (unknownNext && reply.kind === 'reply' && commandIds.has(reply.id)) {
        unknownNext = false;
        deliver({ ...reply, ok: false, error: { name: 'CommandOutcomeUnknown', message: 'Edit outcome is unknown', refusal: { seq: client.state.throughSeq + 1 } } });
        return;
      }
      if (holding && reply.kind === 'reply' && commandIds.has(reply.id)) held.push(() => { deliver(message); });
      else deliver(message);
    });
    restorers.push(off);
    const font = readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
    let actual: PeerlessSession | undefined;
    const value = font.then((font) => createPeerlessSession(bytes, { readOnly: options.readOnly, observeChanges: options.observeChanges,
      activeSlideIndex: options.activeSlideIndex, collaboration: options.collaboration, clientId: options.clientId, initialUpdate: options.initialUpdate, worker: () => endpoint as unknown as Worker,
      fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font) }] })).then((session) => { actual = session; return session; });
    owner = { ready: bounded(value.then((session) => session.ready)).then((value) => { client = value; return value; }),
      get failure() { return actual?.failure; },
      subscribeFailure(listener) {
        let live = true;
        let off: (() => void) | undefined;
        void value.then((session) => { if (live) off = session.subscribeFailure(listener); });
        return () => { live = false; off?.(); };
      },
      dispose: async () => { await (await value).dispose(); registration.dispose(); } };
    return owner;
  });
  restorers.push(() => opener.mockRestore());
});

afterEach(async () => {
  try { releaseReplies(); cleanup(); await owner?.dispose(); }
  finally { for (const restore of restorers.splice(0).reverse()) restore(); }
});
afterAll(async () => {
  if (!ownsDom) return;
  await act(async () => { cleanup(); await new Promise((resolve) => setTimeout(resolve, 0)); });
  await GlobalRegistrator.unregister();
});

export async function open(options: Partial<PptxAsyncEditorProps> = {}) {
  let api!: PptxAsyncEditorApi;
  const ready = mock((value: PptxAsyncEditorApi) => { api = value; });
  const bytes = await bounded(deck());
  const ui = (readOnly = options.readOnly, suspend = false) => <Suspense fallback={<div>Waiting</div>}>
    <PptxEditor fonts={[]} file={bytes} onReady={ready} {...options} readOnly={readOnly} />
    <Suspend active={suspend} />
  </Suspense>;
  const mounted = render(ui());
  const view = { ...mounted, ...within(mounted.container) };
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const canvas = view.getByTestId('pptx-slide-canvas') as HTMLCanvasElement;
  const list = decodeCommandDisplayList(client.state.frame!);
  const box = list.primitives.find((primitive) => primitive.kind === 'textBox') as TextBoxPrimitive;
  const bounds = spyOn(canvas, 'getBoundingClientRect').mockReturnValue({ x: 0, y: 0, left: 0, top: 0,
    right: list.width, bottom: list.height, width: list.width, height: list.height, toJSON: () => ({}) });
  restorers.push(() => bounds.mockRestore());
  const stage = view.getByRole('application');
  Object.defineProperties(stage, { setPointerCapture: { configurable: true, value: mock(() => {}) },
    releasePointerCapture: { configurable: true, value: mock(() => {}) } });
  const point = (x = box.x + 1, y = box.y + box.h / 2) =>
    ({ clientX: x, clientY: y, pointerId: 7, isPrimary: true, button: 0, detail: 1 });
  const select = async () => {
    act(() => { fireEvent.pointerDown(canvas, point()); fireEvent.pointerUp(stage, point()); });
    await waitFor(() => expect(client.state.selection).toMatchObject({ kind: 'object', shapeId: box.shapeId }));
    await waitFor(() => expect(view.getAllByTestId(/pptx-resize-/)).toHaveLength(8));
  };
  return { view, stage, canvas, box, point, select, list, bytes, api, ui, client, commands };
}


export async function story(api: PptxAsyncEditorApi) {
  const result = await api.readContent();
  if (!result.ok) throw new Error(result.failure.message);
  return result.stories[0]!;
}

export function type(input: HTMLElement, text: string): void {
  input.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertText', data: text, bubbles: true, cancelable: true }));
}

export const replyControl = { hold() { holding = true; }, count() { return held.length; }, release() { releaseReplies(); } };

export const faultControl = { unknown() { unknownNext = true; }, crash() { crashWorker(); } };
