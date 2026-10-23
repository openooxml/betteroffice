import './peerlessTestCapabilities';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type { TextBoxPrimitive } from '@betteroffice/pptx';
import type {
  PptxWorkerSession, PresentationCommandClient, PresentationCommandEnvelope, PresentationCommandEvents,
  PresentationCommandMethods,
} from '@betteroffice/pptx/worker';
import { decodeCommandDisplayList } from '@betteroffice/pptx/worker';
import { createSessionHost } from '../../../../shared/office-session/host';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { createSharedPresentationCommandClient } from '../../../pptx/src/session/commandClient';
import { registerPresentationCommandHandlers } from '../../../pptx/src/session/commandHost';
import * as wasm from '../../../pptx/src/wasm/loader';
import { PptxEditor } from '../PptxEditor';
import { peerlessPainterFactories, peerlessSessionOpener } from './usePeerlessPresentation';
import { usePptxCommands, usePptxCommandState } from '../commands/hooks';
import { pptxCommandController } from '../commands/createPptxCommandStore';
import type { PptxAsyncEditorApi } from './createPeerlessEditorApi';
import type { PptxAsyncEditorProps } from '../PptxEditor';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
let client: PresentationCommandClient;
let commands: PresentationCommandEnvelope[];
let owner: PptxWorkerSession | undefined;
let refuseFormatting = false;


function bounded<T>(promise: Promise<T>): Promise<T> {
  let timer: ReturnType<typeof setTimeout>;
  return Promise.race([promise, new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error('Command timed out')), 5000);
  })]).finally(() => clearTimeout(timer));
}

async function deck(small = false): Promise<Uint8Array> {
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
  refuseFormatting = false;

  const createPainter = peerlessPainterFactories.frame;
  const painter = spyOn(peerlessPainterFactories, 'frame').mockImplementation((canvas, options) =>
    createPainter(canvas, { ...options, onCommit(frame, view) {
      options?.onCommit?.(frame, view);

    } }));
  restorers.push(() => painter.mockRestore()); owner = undefined;
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
      return command(context, refuseFormatting && envelope.kind === 'formatText'
        ? { ...envelope, payload: { ...envelope.payload, patch: { unexpected: true } } } as unknown as PresentationCommandEnvelope : envelope);
    });
    restorers.push(() => received.mockRestore());
    createSessionHost<PresentationCommandMethods, PresentationCommandEvents, null>(pair.host, {
      handlers: registration.handlers, policies: registration.policies, context: null, onDispose: registration.dispose,
    });
    const ready = bounded(readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url))
      .then((font) => createSharedPresentationCommandClient(bytes, { session: 'insert-session', incarnation: 'insert-worker',
        readOnly: options.readOnly, activeSlideIndex: options.activeSlideIndex,
        fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }] }, pair.client))
      .then((value) => { client = value; return value; }));
    owner = { ready, failure: undefined, subscribeFailure: () => () => {},
      dispose: async () => { try { await bounded(ready.then((value) => value.dispose(), () => {})); }
        finally { registration.dispose(); pair.client.close(); } } };
    return owner;
  });
  restorers.push(() => opener.mockRestore());
});

afterEach(async () => {
  try { cleanup(); await owner?.dispose(); }
  finally { for (const restore of restorers.splice(0).reverse()) restore(); }
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

async function open(options: Partial<PptxAsyncEditorProps> = {}) {
  let api!: PptxAsyncEditorApi;
  const ready = mock((value: PptxAsyncEditorApi) => { api = value; });
  const bytes = await bounded(deck());
  const view = render(<PptxEditor fonts={[]} file={bytes} onReady={ready} {...options} />);
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
  return { view, stage, canvas, box, point, select, list, bytes, api };
}


test('custom_hooks_share_api_commands', async () => {
  let observed: PptxAsyncEditorApi['commands'] | undefined;
  function Chrome() {
    observed = usePptxCommands();
    const state = usePptxCommandState('undo');
    return <span data-testid="custom-chrome">{String(state.enabled)}</span>;
  }
  const { view, api } = await open({ toolbar: <Chrome /> });
  expect(observed === api.commands).toBe(true);
  expect(view.getByTestId('custom-chrome').textContent).toBe('false');
  await act(async () => { expect(await api.commands.execute('tool', { value: 'textBox' })).toMatchObject({ ok: true }); });
  expect(api.commands.getState('tool').value).toBe('textBox');
});

test('toolbar_slots', async () => {
  const { view, bytes } = await open({ toolbar: <span data-testid="replacement-toolbar">custom</span> });
  expect(view.queryByTestId('pptx-editor-toolbar')).toBeNull();
  expect(view.getByTestId('replacement-toolbar').textContent).toBe('custom');
  view.rerender(<PptxEditor fonts={[]} file={bytes} toolbar={null} />);
  expect(view.queryByTestId('replacement-toolbar')).toBeNull();
  expect(view.queryByTestId('pptx-editor-toolbar')).toBeNull();
  view.rerender(<PptxEditor fonts={[]} file={bytes} toolbar={<span>hidden</span>} showToolbar={false} />);
  expect(view.queryByText('hidden')).toBeNull();
});

test('deferred_origin_refuses_replaced_document', async () => {
  const { view, api } = await open();
  const controller = pptxCommandController(api.commands)!;
  const pending = controller.prepare('fontFamily');
  view.rerender(<PptxEditor fonts={[]} file={await deck(true)} />);
  expect(await pending.execute({ family: 'Liberation Sans' })).toMatchObject({ ok: false });
  expect(commands.filter((command) => command.kind === 'formatText')).toHaveLength(0);
});

test('worker_refusal_becomes_typed_command_failure', async () => {
  const { api, select } = await open();
  await select();
  refuseFormatting = true;
  const result = await api.commands.execute('bold', null);
  expect(result).toMatchObject({ ok: false, failure: { code: 'invalid-arguments' } });
});

test('formatting_follows_pending_first_text_click', async () => {
  const { view, canvas, stage, box, point } = await open();
  const line = box.lines[0]!;
  const stop = line.caretStops.find((stop) => stop.position === 2)!;
  const input = view.getByRole('textbox', { name: 'Slide text' });
  act(() => {
    fireEvent.pointerDown(canvas, point(stop.x, line.y + line.height / 2));
    fireEvent.pointerUp(stage, point(stop.x, line.y + line.height / 2));
    fireEvent.keyDown(input, { key: 'b', ctrlKey: true });
  });
  await bounded(client.flush());
  expect(commands.map((command) => command.kind)).toEqual(['selectAtPoint', 'formatText']);
  expect(client.state.editState.toolbar.storedMarks).toEqual({ bold: true });
});
