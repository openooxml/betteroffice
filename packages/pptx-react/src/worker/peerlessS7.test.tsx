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
import { PEERLESS_FLAG, PptxEditor } from '../PptxEditor';
import { peerlessPainterFactories, peerlessSessionOpener } from './usePeerlessPresentation';
import { peerlessImagePreparation } from './usePeerlessInsertion';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
let client: PresentationCommandClient;
let commands: PresentationCommandEnvelope[];
let owner: PptxWorkerSession | undefined;


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
      return command(context, envelope);
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

async function open(small = false) {
  const ready = mock(() => {});
  const bytes = await bounded(deck(small));
  const view = render(<PptxEditor fonts={[]} file={bytes} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
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
  return { view, stage, canvas, box, point, select, list, bytes };
}

test('textbox_focus', async () => {
  const { view, stage, canvas, point } = await open();
  fireEvent.click(view.getByRole('button', { name: 'Text box' }));
  act(() => { fireEvent.pointerDown(canvas, point(200, 150)); fireEvent.pointerUp(stage, point(400, 220)); });
  await waitFor(() => expect(commands.filter((command) => command.kind === 'insertTextBox')).toHaveLength(1));
  await bounded(client.flush());
  const selected = client.state.selection;
  expect(selected).toMatchObject({ anchor: 0, focus: 0 });
  expect(selected?.kind).not.toBe('object');
  const input = view.getByRole('textbox', { name: 'Slide text' });
  expect(document.activeElement === input).toBe(true);
  await act(async () => { input.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertText', data: 'hello',
    bubbles: true, cancelable: true })); await bounded(client.flush()); });
  const read = (await bounded(client.read())).result;
  if (!read.ok) throw new Error(read.failure.message);
  expect(read.stories.some((story) => story.text === 'hello')).toBe(true);
});

test('preset_shape', async () => {
  const { view, stage, canvas, point } = await open();
  fireEvent.click(view.getByRole('button', { name: 'Shape' }));
  fireEvent.click(view.getByRole('menuitemradio', { name: 'Rectangle' }));
  act(() => { fireEvent.pointerDown(canvas, point(200, 150)); fireEvent.pointerUp(stage, point(350, 260)); });
  await waitFor(() => expect(commands.filter((command) => command.kind === 'insertShape')).toHaveLength(1));
  await bounded(client.flush());
  expect(client.state.selection?.kind).toBe('object');
  expect(view.getAllByTestId(/pptx-resize-/)).toHaveLength(8);
  await bounded(client.dispatch({ kind: 'undo', payload: { inputTime: 1, undoGroup: 'undo' } }));
  expect(client.state.editState.canUndo).toBe(false);
});

for (const mode of ['replacement', 'readOnly']) test(`late_image_${mode}`, async () => {
  const { view, bytes } = await open();
  let resolve!: (value: Awaited<ReturnType<typeof peerlessImagePreparation.prepare>>) => void;
  const prepare = spyOn(peerlessImagePreparation, 'prepare').mockImplementation(() => new Promise((done) => { resolve = done; }));
  restorers.push(() => prepare.mockRestore());
  fireEvent.click(view.getByRole('button', { name: 'Insert image' }));
  const picker = view.getByLabelText('Insert picture');
  fireEvent.change(picker, { target: { files: [new File(['png'], 'test.png', { type: 'image/png' })] } });
  expect(prepare).toHaveBeenCalledTimes(1);
  view.rerender(<PptxEditor fonts={[]} file={mode === 'replacement' ? await deck(true) : bytes}
    {...{ [PEERLESS_FLAG]: true as const }} readOnly={mode === 'readOnly'} />);
  await act(async () => { resolve({ bytes: new Uint8Array([1]).buffer, contentType: 'image/png', width: 10, height: 10 }); });
  expect(commands.filter((command) => command.kind === 'insertPicture')).toHaveLength(0);
});

test('accepted_image_picker_selects_picture', async () => {
  const { view } = await open();
  const bytes = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a7N8AAAAASUVORK5CYII=', 'base64')).buffer;
  const prepare = spyOn(peerlessImagePreparation, 'prepare').mockResolvedValue({ bytes, contentType: 'image/png', width: 1, height: 1 });
  restorers.push(() => prepare.mockRestore());
  fireEvent.click(view.getByRole('button', { name: 'Insert image' }));
  fireEvent.change(view.getByLabelText('Insert picture'), { target: { files: [new File([bytes], 'picture.png', { type: 'image/png' })] } });
  await waitFor(() => expect(commands.filter((command) => command.kind === 'insertPicture')).toHaveLength(1));
  await bounded(client.flush());
  expect(client.state.selection?.kind).toBe('object');
  expect(decodeCommandDisplayList(client.state.frame!).primitives.some((primitive) => primitive.kind === 'image')).toBe(true);
});
