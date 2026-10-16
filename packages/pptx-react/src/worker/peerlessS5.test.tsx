import './peerlessTestCapabilities';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type { PeerlessSession, PresentationCommandClient, PresentationCommandEnvelope,
  PresentationCommandEvents, PresentationCommandMethods, TextBoxPrimitive } from '@betteroffice/pptx';
import { decodeCommandDisplayList, indexGeometry, navigateGeometry } from '@betteroffice/pptx';
import { createSessionHost } from '../../../../shared/office-session/host';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { createSharedPresentationCommandClient } from '../../../pptx/src/session/commandClient';
import { registerPresentationCommandHandlers } from '../../../pptx/src/session/commandHost';
import * as wasm from '../../../pptx/src/wasm/loader';
import { PEERLESS_FLAG, PptxEditor } from '../PptxEditor';
import { peerlessSessionOpener } from './usePeerlessPresentation';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, createEvent, fireEvent, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
let client: PresentationCommandClient;
let commands: PresentationCommandEnvelope[];
let owner: PeerlessSession | undefined;

function bounded<T>(promise: Promise<T>): Promise<T> {
  let timer: ReturnType<typeof setTimeout>;
  return Promise.race([promise, new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error('S5 command timed out')), 5000);
  })]).finally(() => clearTimeout(timer));
}

async function deck(): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="2400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${['alpha beta', 'x', 'alpha beta'].map((text) => `<a:p><a:r><a:rPr sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>${text}</a:t></a:r></a:p>`).join('')}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}

beforeEach(() => {
  commands = []; owner = undefined;
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
      get(this: HTMLElement) { return this.getAttribute('role') === 'application' ? extent : 0; } });
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
      .then((font) => createSharedPresentationCommandClient(bytes, { session: 's5', incarnation: 's5-worker',
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

async function open() {
  const ready = mock(() => {});
  const view = render(<PptxEditor fonts={[]} file={await deck()} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const canvas = view.getByTestId('pptx-slide-canvas') as HTMLCanvasElement;
  const list = decodeCommandDisplayList(client.state.frame!);
  const box = list.primitives.find((primitive) => primitive.kind === 'textBox') as TextBoxPrimitive;
  const bounds = spyOn(canvas, 'getBoundingClientRect').mockReturnValue({ x: 0, y: 0, left: 0, top: 0,
    right: list.width, bottom: list.height, width: list.width, height: list.height, toJSON: () => ({}) });
  restorers.push(() => bounds.mockRestore());
  const stage = view.getByRole('application');
  const capture = mock(() => {});
  const release = mock(() => {});
  Object.defineProperties(stage, { setPointerCapture: { configurable: true, value: capture },
    releasePointerCapture: { configurable: true, value: release } });
  const point = (position: number, lineIndex = 0) => {
    const line = box.lines[lineIndex]!;
    const stop = line.caretStops.find((stop) => stop.position === position)!;
    return { clientX: stop.x, clientY: line.y + line.height / 2, pointerId: 7, isPrimary: true, button: 0, detail: 1 };
  };
  return { view, stage, canvas, box, point, capture, release,
    input: view.getByRole('textbox', { name: 'Slide text' }) };
}

test('shift_drag', async () => {
  const { stage, canvas, point, capture, release } = await open();
  act(() => { fireEvent.pointerDown(canvas, point(1)); fireEvent.pointerUp(stage, point(1)); });
  await waitFor(() => expect(client.state.selection?.focus).toBe(1));
  act(() => { fireEvent.pointerDown(canvas, { ...point(4), shiftKey: true }); });
  await waitFor(() => expect(client.state.selection).toMatchObject({ anchor: 1, focus: 4 }));
  act(() => { fireEvent.pointerMove(stage, { ...point(8), buttons: 1 }); fireEvent.pointerUp(stage, point(8)); });
  await waitFor(() => expect(client.state.selection).toMatchObject({ anchor: 1, focus: 8 }));
  expect(capture).toHaveBeenCalledWith(7);
  expect(release).toHaveBeenCalledWith(7);
});

test('shift_click_pending_first_click', async () => {
  const { stage, canvas, point } = await open();
  expect(client.state.selection).toBeNull();
  act(() => {
    fireEvent.pointerDown(canvas, point(1)); fireEvent.pointerUp(stage, point(1));
    fireEvent.pointerDown(canvas, { ...point(4), shiftKey: true }); fireEvent.pointerUp(stage, point(4));
  });
  await waitFor(() => expect(client.state.selection).toMatchObject({ anchor: 1, focus: 4 }));
  expect(commands[1]).toMatchObject({ kind: 'selectAtPoint', afterSeq: commands[0]!.seq,
    payload: { extend: true, at: { source: 'afterSeq', seq: commands[0]!.seq } } });
});

test('drag_moves_coalesce_and_deliver_pointerup', async () => {
  const { stage, canvas, point, box, release } = await open();
  act(() => {
    fireEvent.pointerDown(canvas, point(1));
    for (const position of [4, 5, 6, 6, 7, 8, 8]) {
      fireEvent.pointerMove(stage, { ...point(position), buttons: 1 });
    }
    fireEvent.pointerUp(stage, point(9));
  });
  await waitFor(() => expect(client.state.selection).toMatchObject({ anchor: 1, focus: 9 }));
  expect(commands).toHaveLength(3);
  expect(commands.map((command) => command.kind === 'selectAtPoint' ? command.payload.x : null))
    .toEqual([1, 4, 9].map((position) => box.lines[0]!.caretStops.find((stop) => stop.position === position)!.x));
  expect(release).toHaveBeenCalledWith(7);
  act(() => { fireEvent.pointerDown(canvas, point(1)); });
  await waitFor(() => expect(client.state.selection?.focus).toBe(1));
  const beforeMoves = commands.length;
  act(() => {
    for (const position of [4, 5, 6, 7]) fireEvent.pointerMove(stage, { ...point(position), buttons: 1 });
  });
  await waitFor(() => expect(client.state.selection?.focus).toBe(7));
  expect(commands).toHaveLength(beforeMoves + 2);
  const count = commands.length;
  act(() => {
    fireEvent.pointerMove(stage, { ...point(7), buttons: 1 });
    fireEvent.pointerUp(stage, point(7));
  });
  expect(commands).toHaveLength(count);
});

test('queued_drag_precedes_typing', async () => {
  const { stage, canvas, point, input } = await open();
  act(() => { fireEvent.pointerDown(canvas, point(1)); });
  await waitFor(() => expect(client.state.selection?.focus).toBe(1));
  const start = commands.length;
  act(() => {
    fireEvent.pointerMove(stage, { ...point(4), buttons: 1 });
    fireEvent.pointerMove(stage, { ...point(7), buttons: 1 });
    input.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertText', data: 'Z',
      bubbles: true, cancelable: true }));
  });
  await waitFor(() => expect(commands).toHaveLength(start + 3));
  expect(commands.slice(start).map((command) => command.kind)).toEqual(['selectAtPoint', 'selectAtPoint', 'insertText']);
  const read = (await client.read()).result;
  expect(read.ok && read.stories[0]!.text).toBe('aZeta\nx\nalpha beta');
});

test('expanded_focus_wrapped_affinity', async () => {
  const original = wasm.presentationDisplayListJson;
  const wrapped = spyOn(wasm, 'presentationDisplayListJson').mockImplementation((...args) => {
    const list = JSON.parse(original(...args));
    const box = list.primitives.find((primitive: TextBoxPrimitive) => primitive.kind === 'textBox') as TextBoxPrimitive;
    const first = box.lines[0]!;
    box.lines = [{ ...first, end: 5, caretStops: first.caretStops.filter((stop) => stop.position <= 5) },
      { ...first, start: 5, y: first.y + first.height,
        caretStops: first.caretStops.filter((stop) => stop.position >= 5) },
      ...box.lines.slice(1).map((line) => ({ ...line, y: line.y + first.height }))];
    return JSON.stringify(list);
  });
  restorers.push(() => wrapped.mockRestore());
  const { stage, canvas, input, point } = await open();
  for (const [detail, focus, home] of [[2, 5, 0], [3, 10, 5]] as const) {
    act(() => { fireEvent.pointerDown(canvas, { ...point(2), detail }); fireEvent.pointerUp(stage, point(2)); });
    await waitFor(() => expect(client.state.selection).toMatchObject({ anchor: 0, focus, affinity: 'upstream' }));
    act(() => { fireEvent.keyDown(input, { key: 'Home' }); });
    await waitFor(() => expect(client.state.selection?.focus).toBe(home));
  }
  act(() => {
    fireEvent.pointerDown(canvas, { ...point(8, 1), detail: 2 });
    fireEvent.pointerMove(stage, { ...point(2), buttons: 1 });
  });
  await waitFor(() => expect(client.state.selection).toMatchObject({ anchor: 10, focus: 0, affinity: 'downstream' }));
  act(() => { fireEvent.pointerMove(stage, { ...point(8, 1), buttons: 1 }); fireEvent.pointerUp(stage, point(8, 1)); });
  await waitFor(() => expect(client.state.selection).toMatchObject({ anchor: 6, focus: 10, affinity: 'upstream' }));
});

test('multi_click', async () => {
  const { canvas, stage, point } = await open();
  for (let count = 1; count <= 3; count += 1) {
    const event = createEvent.pointerDown(canvas, { ...point(2), detail: 0 });
    Object.defineProperty(event, 'timeStamp', { value: count * 100 });
    act(() => { fireEvent(canvas, event);
      fireEvent.pointerUp(stage, point(2)); });
    await waitFor(() => expect(client.state.throughSeq).toBeGreaterThanOrEqual(count));
    if (count === 2) expect(client.state.selection).toMatchObject({ anchor: 0, focus: 5 });
    if (count === 3) expect(client.state.selection).toMatchObject({ anchor: 0, focus: 10 });
  }
  act(() => { fireEvent.pointerDown(canvas, { ...point(2), detail: 2 });
    fireEvent.pointerMove(stage, { ...point(8), buttons: 1 }); fireEvent.pointerUp(stage, point(8)); });
  await waitFor(() => expect(client.state.selection).toMatchObject({ anchor: 0, focus: 10 }));
  act(() => { fireEvent.pointerDown(canvas, { ...point(8), detail: 2 });
    fireEvent.pointerMove(stage, { ...point(2), buttons: 1 }); });
  await waitFor(() => expect(client.state.selection).toMatchObject({ anchor: 10, focus: 0 }));
  act(() => { fireEvent.pointerMove(stage, { ...point(8), buttons: 1 }); fireEvent.pointerUp(stage, point(8)); });
  await waitFor(() => expect(client.state.selection).toMatchObject({ anchor: 6, focus: 10 }));
  act(() => { fireEvent.pointerDown(canvas, { ...point(2), detail: 3 });
    fireEvent.pointerMove(stage, { ...point(21, 2), buttons: 1 }); });
  await waitFor(() => expect(client.state.selection).toMatchObject({ anchor: 0, focus: 23 }));
  act(() => { fireEvent.pointerMove(stage, { ...point(2), buttons: 1 }); fireEvent.pointerUp(stage, point(2)); });
  await waitFor(() => expect(client.state.selection).toMatchObject({ anchor: 0, focus: 10 }));
});

test('modifier_edges', async () => {
  const { canvas, stage, input, point } = await open();
  const platform = Object.getOwnPropertyDescriptor(navigator, 'platform');
  restorers.push(() => {
    if (platform) Object.defineProperty(navigator, 'platform', platform);
    else Reflect.deleteProperty(navigator, 'platform');
  });
  for (const [name, word] of [['MacIntel', { altKey: true }], ['Win32', { ctrlKey: true }]] as const) {
    Object.defineProperty(navigator, 'platform', { configurable: true, value: name });
    act(() => { fireEvent.pointerDown(canvas, point(2)); fireEvent.pointerUp(stage, point(2)); });
    await waitFor(() => expect(client.state.selection?.focus).toBe(2));
    act(() => { fireEvent.keyDown(input, { key: 'ArrowRight', ...word, shiftKey: true }); });
    await waitFor(() => expect(client.state.selection).toMatchObject({ anchor: 2, focus: 5 }));
    act(() => { fireEvent.keyDown(input, { key: 'End', ctrlKey: true }); });
    await waitFor(() => expect(client.state.selection?.focus).toBe(23));
    act(() => { fireEvent.keyDown(input, { key: 'Home' }); });
    await waitFor(() => expect(client.state.selection?.focus).toBe(13));
    act(() => { fireEvent.keyDown(input, { key: 'End' }); });
    await waitFor(() => expect(client.state.selection?.focus).toBe(23));
    act(() => { fireEvent.keyDown(input, { key: 'Home', metaKey: true }); });
    await waitFor(() => expect(client.state.selection?.focus).toBe(0));
    act(() => { fireEvent.keyDown(input, { key: 'ArrowRight', metaKey: true }); });
    await waitFor(() => expect(client.state.selection?.focus).toBe(10));
    act(() => { fireEvent.keyDown(input, { key: 'ArrowLeft', metaKey: true }); });
    await waitFor(() => expect(client.state.selection?.focus).toBe(0));
    act(() => { fireEvent.pointerDown(canvas, point(2)); fireEvent.pointerUp(stage, point(2)); });
    await waitFor(() => expect(client.state.selection?.focus).toBe(2));
    act(() => { fireEvent.keyDown(input, { key: 'ArrowRight',
      ...(name === 'MacIntel' ? { ctrlKey: true } : { altKey: true }) }); });
    await waitFor(() => expect(client.state.selection?.focus).toBe(3));
  }
});

test('vertical_goal', async () => {
  const { canvas, stage, input, point, box } = await open();
  act(() => { fireEvent.pointerDown(canvas, point(8)); fireEvent.pointerUp(stage, point(8)); });
  await waitFor(() => expect(client.state.selection?.focus).toBe(8));
  act(() => { fireEvent.keyDown(input, { key: 'ArrowDown' }); });
  await waitFor(() => expect(client.state.selection?.focus).toBe(12));
  act(() => { fireEvent.keyDown(input, { key: 'ArrowDown' }); });
  await waitFor(() => expect(client.state.selection?.focus).toBe(21));
  const index = indexGeometry({ contractVersion: 1, width: 960, height: 540,
    background: { kind: 'solid', color: '#ffffff' }, primitives: [{ ...box,
      lines: [box.lines[0]!, { ...box.lines[0]!, start: 10, end: 20,
        caretStops: box.lines[0]!.caretStops.map((stop) => ({ ...stop, position: stop.position + 10 })) }] }] });
  const selection = { ...client.state.selection!, anchor: 10, focus: 10, affinity: 'downstream' as const };
  expect(navigateGeometry(index, selection, 'lineEnd', false)?.affinity).toBe('upstream');
});

test('pending_unicode', async () => {
  const { canvas, stage, input, point } = await open();
  act(() => { fireEvent.pointerDown(canvas, point(0)); fireEvent.pointerUp(stage, point(0)); });
  await waitFor(() => expect(client.state.selection?.focus).toBe(0));
  const start = commands.length;
  const positions = new Map<number, number | undefined>();
  restorers.push(client.subscribe((ack) => { positions.set(ack.throughSeq, ack.selection?.focus); }));
  act(() => {
    input.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertText', data: 'e\u0301😀',
      bubbles: true, cancelable: true }));
    fireEvent.keyDown(input, { key: 'ArrowLeft' });
    fireEvent.keyDown(input, { key: 'ArrowLeft', ctrlKey: true, altKey: true });
  });
  await waitFor(() => expect(commands).toHaveLength(start + 3));
  const [typed, left, word] = commands.slice(start);
  expect(left).toMatchObject({ kind: 'moveSelection', afterSeq: typed!.seq,
    payload: { at: { source: 'afterSeq', seq: typed!.seq } } });
  expect(word).toMatchObject({ kind: 'moveSelection', afterSeq: left!.seq, payload: { motion: 'wordLeft' } });
  await waitFor(() => expect(client.state.throughSeq).toBe(word!.seq));
  expect(client.state.outcomes.every((outcome) => outcome.status !== 'refused')).toBe(true);
  expect(positions.get(typed!.seq)).toBe(4);
  expect(positions.get(left!.seq)).toBe(2);
  expect(client.state.selection?.focus).toBe(0);
  act(() => { fireEvent.keyDown(input, { key: 'ArrowRight' }); });
  await waitFor(() => expect(client.state.selection?.focus).toBe(2));
  act(() => { fireEvent.keyDown(input, { key: 'ArrowRight' }); });
  await waitFor(() => expect(client.state.selection?.focus).toBe(4));
});
