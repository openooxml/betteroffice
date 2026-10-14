import './peerlessTestCapabilities';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type { PeerlessSession, PresentationCommandClient, PresentationCommandEnvelope,
  PresentationCommandEvents, PresentationCommandMethods, TextBoxPrimitive } from '@betteroffice/pptx';
import { decodeCommandDisplayList } from '@betteroffice/pptx';
import { createSessionHost } from '../../../../shared/office-session/host';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { createSharedPresentationCommandClient } from '../../../pptx/src/session/commandClient';
import { registerPresentationCommandHandlers } from '../../../pptx/src/session/commandHost';
import * as wasm from '../../../pptx/src/wasm/loader';
import { PEERLESS_FLAG, PptxEditor } from '../PptxEditor';
import { peerlessPainterFactories, peerlessSessionOpener } from './usePeerlessPresentation';
import * as i18n from '../i18n';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, createEvent, fireEvent, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
let client: PresentationCommandClient;
let commands: PresentationCommandEnvelope[];
let owner: PeerlessSession | undefined;
let paintedThroughSeq = 0;

function bounded<T>(promise: Promise<T>): Promise<T> {
  let timer: ReturnType<typeof setTimeout>;
  return Promise.race([promise, new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error('S6 command timed out')), 5000);
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
  commands = [];
  paintedThroughSeq = 0;
  const createPainter = peerlessPainterFactories.frame;
  const painter = spyOn(peerlessPainterFactories, 'frame').mockImplementation((canvas, options) =>
    createPainter(canvas, { ...options, onCommit(frame, view) {
      options?.onCommit?.(frame, view);
      paintedThroughSeq = frame.frame.coversThroughSeq;
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
      .then((font) => createSharedPresentationCommandClient(bytes, { session: 's6', incarnation: 's6-worker',
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
  const view = render(<PptxEditor fonts={[]} file={await bounded(deck(small))} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
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
  return { view, stage, canvas, box, point, select, list };
}

test('shape_border', async () => {
  const { view, stage, canvas, box, point, select } = await open();
  await select();
  const line = box.lines[0]!;
  const stop = line.caretStops.find((stop) => stop.position === 2)!;
  act(() => { fireEvent.pointerDown(canvas, point(stop.x, line.y + line.height / 2));
    fireEvent.pointerUp(stage, point(stop.x, line.y + line.height / 2)); });
  await waitFor(() => expect(client.state.selection).toMatchObject({ storyId: box.storyId, focus: 2 }));
  expect(client.state.selection?.kind).not.toBe('object');
  await waitFor(() => expect(view.queryByTestId('pptx-resize-nw')).toBeNull());
});

test('drag_release', async () => {
  const { view, stage, canvas, point, select, list } = await open();
  await select();
  const version = client.state.version;
  const before = client.state.selection;
  act(() => { fireEvent.pointerDown(canvas, point());
    fireEvent.pointerMove(stage, { ...point(point().clientX + 10), buttons: 1 }); });
  expect(client.state.version).toBe(version);
  expect(commands.filter((command) => command.kind === 'moveShape' || command.kind === 'setShapeRect')).toHaveLength(0);
  act(() => { fireEvent.pointerUp(stage, point(point().clientX + 30, point().clientY + 20)); });
  await waitFor(() => expect(commands.filter((command) => command.kind === 'moveShape')).toHaveLength(1));
  await bounded(client.flush());
  const selected = client.state.selection;
  expect(selected).toMatchObject({ kind: 'object', rect: {
    x: 100000 + Math.round(30 * 9144000 / list.width),
    y: 100000 + Math.round(20 * 5143500 / list.height),
  } });
  const moved = selected;
  await bounded(client.dispatch({ kind: 'undo', payload: { inputTime: 1, undoGroup: 'undo' } }));
  const restored = decodeCommandDisplayList(client.state.frame!).primitives.find((primitive) => primitive.kind === 'textBox')!;
  expect(restored.x).toBeCloseTo(100000 * list.width / 9144000, 3);
  expect(restored.y).toBeCloseTo(100000 * list.height / 5143500, 3);
  expect(client.state.editState.canUndo).toBe(false);
  await bounded(client.dispatch({ kind: 'redo', payload: { inputTime: 2, undoGroup: 'redo' } }));
  expect(client.state.version).not.toBe(version);
  expect(before).not.toEqual(moved);
  await waitFor(() => expect(paintedThroughSeq).toBe(client.state.throughSeq));
  act(() => { fireEvent.pointerDown(canvas, point(point().clientX + 30, point().clientY + 20));
    fireEvent.pointerUp(stage, point(point().clientX + 30, point().clientY + 20)); });
  const handle = await waitFor(() => view.getByTestId('pptx-resize-e'));
  act(() => { fireEvent.pointerDown(handle, point());
    fireEvent.pointerMove(stage, { ...point(point().clientX + 10), buttons: 1 });
    fireEvent.pointerUp(stage, point(point().clientX + 30)); });
  await waitFor(() => expect(commands.filter((command) => command.kind === 'setShapeRect')).toHaveLength(1));
  await bounded(client.flush());
  expect(client.state.selection).toMatchObject({ kind: 'object', rect: {
    x: 100000 + Math.round(30 * 9144000 / list.width),
    width: 6000000 + Math.round(30 * 9144000 / list.width), height: 2400000,
  } });
});

test('resize_minimum', async () => {
  const { view, stage, point, select, box } = await open();
  await select();
  const handle = view.getByTestId('pptx-resize-nw');
  act(() => { fireEvent.pointerDown(handle, point(box.x, box.y));
    fireEvent.pointerMove(stage, { ...point(box.x + box.w + 100, box.y + box.h + 100), buttons: 1 });
    fireEvent.pointerUp(stage, point(box.x + box.w + 200, box.y + box.h + 200)); });
  await waitFor(() => expect(commands.filter((command) => command.kind === 'setShapeRect')).toHaveLength(1));
  await bounded(client.flush());
  expect(client.state.selection).toMatchObject({ kind: 'object',
    rect: { x: 6100000 - 76200, y: 2500000 - 76200, width: 76200, height: 76200 } });
  await bounded(client.dispatch({ kind: 'undo', payload: { inputTime: 1, undoGroup: 'undo' } }));
  expect(client.state.editState.canUndo).toBe(false);
});

test('cancel_gesture', async () => {
  const { stage, canvas, point, select, view } = await open();
  await select();
  const version = client.state.version;
  for (const target of [canvas, view.getByTestId('pptx-resize-se')]) {
    act(() => { fireEvent.pointerDown(target, point());
      fireEvent.pointerMove(stage, { ...point(point().clientX + 40), buttons: 1 });
      fireEvent.keyDown(stage, { key: 'Escape' });
      fireEvent.pointerUp(stage, point(point().clientX + 50)); });
    act(() => { fireEvent.pointerDown(target, point());
      fireEvent.pointerMove(stage, { ...point(point().clientX + 40), buttons: 1 });
      fireEvent.pointerCancel(stage, point());
      fireEvent.pointerUp(stage, point(point().clientX + 50)); });
  }
  await bounded(client.flush());
  expect(client.state.version).toBe(version);
  expect(commands.filter((command) => command.kind === 'moveShape' || command.kind === 'setShapeRect')).toHaveLength(0);
});

test('stale_target', async () => {
  const { select } = await open();
  await select();
  const selected = client.state.selection!;
  const frame = client.state.frame!;
  const version = client.state.version;
  const ack = await bounded(client.dispatch({ kind: 'setShapeRect', payload: {
    inputTime: 1, undoGroup: 'stale', frame,
    target: { slideId: selected.slideId, shapeId: 'removed-target', token: 'replaced' },
    rect: { x: 0, y: 0, width: 76200, height: 76200 },
  } }));
  expect(ack.outcomes[0]).toMatchObject({ status: 'refused', reason: 'bad-target', committed: false });
  expect(client.state.version).toBe(version);
  expect(client.state.selection).toEqual(selected);
});

test('small_shape_handle_click_and_jitter_send_no_transform', async () => {
  const { view, stage, point, select } = await open(true);
  await select();
  const version = client.state.version;
  const before = client.state.selection;
  for (const offsets of [[0], [2], [10, 0]]) {
    await act(async () => {
      fireEvent.pointerDown(view.getByTestId('pptx-resize-se'), point());
      for (const offset of offsets) fireEvent.pointerMove(stage, point(point().clientX + offset));
      fireEvent.pointerUp(stage, point(point().clientX + offsets[offsets.length - 1]!));
    });
  }
  await bounded(client.flush());
  expect(commands.filter((command) => command.kind === 'setShapeRect')).toHaveLength(0);
  expect(client.state.version).toBe(version);
  expect(client.state.selection).toEqual(before);
  expect(client.state.editState.canUndo).toBe(false);
});

test('blur_and_lost_capture_cancel_shape_gestures', async () => {
  const { view, stage, canvas, point, select } = await open();
  await select();
  for (const target of [canvas, view.getByTestId('pptx-resize-se')]) {
    for (const deferred of [false]) {
      await act(async () => {
        fireEvent.pointerDown(target, point());
        fireEvent.pointerMove(stage, point(point().clientX + 40));
        if (deferred) fireEvent.pointerUp(stage, point(point().clientX + 50));
        fireEvent.blur(window);
        fireEvent.pointerUp(stage, point(point().clientX + 50));
      });
    }
    await act(async () => {
      fireEvent.pointerDown(target, point());
      fireEvent.pointerMove(stage, point(point().clientX + 40));
      fireEvent.lostPointerCapture(stage, point());
      fireEvent.pointerUp(stage, point(point().clientX + 50));
    });
  }
  await bounded(client.flush());
  expect(commands.filter((command) => command.kind === 'moveShape' || command.kind === 'setShapeRect')).toHaveLength(0);
  expect(client.state.editState.canUndo).toBe(false);
});

test('peerless_interaction_layers_disable_touch_panning', async () => {
  const { stage, canvas } = await open();
  expect(canvas.style.touchAction).toBe('none');
  expect(stage.style.touchAction).toBe('none');
});

test('drag_burst_does_not_render_slide_strip', async () => {
  let stripRenders = 0;
  const original = i18n.useTranslation;
  const translation = spyOn(i18n, 'useTranslation').mockImplementation(() => {
    const value = original();
    return { ...value, t: (...args: Parameters<typeof value.t>) => {
      if (args[0] === 'slides.panelLabel') stripRenders += 1;
      return value.t(...args);
    } };
  });
  restorers.push(() => translation.mockRestore());
  const { view, stage, point, select } = await open();
  await select();
  await act(async () => { fireEvent.pointerDown(view.getByTestId('pptx-resize-se'), point()); });
  const before = stripRenders;
  const outline = view.getByTestId('pptx-shape-selection');
  const transform = outline.style.transform;
  for (let offset = 10; offset <= 60; offset += 10) {
    act(() => { fireEvent.pointerMove(stage, point(point().clientX + offset)); });
  }
  expect(outline.style.transform).not.toBe(transform);
  expect(stripRenders).toBe(before);
  act(() => { fireEvent.pointerCancel(stage, point()); });
});

test('resize_handles_have_accessible_names_and_roles', async () => {
  const { view, select } = await open();
  await select();
  for (const handle of view.getAllByTestId(/pptx-resize-/)) {
    expect(handle.getAttribute('role')).toBe('img');
    expect(handle.getAttribute('aria-label')).toMatch(/^Resize /);
  }
  expect(new Set(view.getAllByTestId(/pptx-resize-/).map((handle) => handle.getAttribute('aria-label'))).size).toBe(8);
});

test('repeated_shape_click_uses_eight_pixel_drag_threshold', async () => {
  const { stage, canvas, point, select } = await open();
  await select();
  const pointer = (kind: 'pointerDown' | 'pointerMove' | 'pointerUp', target: HTMLElement,
    offset: number, time: number) => {
    const event = createEvent[kind](target, point(point().clientX, point().clientY + offset));
    Object.defineProperty(event, 'timeStamp', { value: time });
    fireEvent(target, event);
  };
  await act(async () => {
    pointer('pointerDown', canvas, 0, 10000);
    pointer('pointerUp', stage, 0, 10001);
  });
  await bounded(client.flush());
  await act(async () => {
    pointer('pointerDown', canvas, 0, 10100);
    pointer('pointerMove', stage, 6, 10101);
    pointer('pointerUp', stage, 6, 10102);
  });
  await bounded(client.flush());
  expect(commands.filter((command) => command.kind === 'moveShape')).toHaveLength(0);
  await act(async () => {
    pointer('pointerDown', canvas, 6, 10200);
    pointer('pointerUp', stage, 14, 10201);
  });
  await waitFor(() => expect(commands.filter((command) => command.kind === 'moveShape')).toHaveLength(1));
  await bounded(client.flush());
});

test('released_shape_drag_precedes_immediate_text_click_and_typing', async () => {
  const { stage, canvas, box, point, list, view } = await open();
  const stop = box.lines[0]!.caretStops.find((stop) => stop.position === 2)!;
  const textPoint = point(stop.x, box.lines[0]!.y + box.lines[0]!.height / 2);
  const input = view.getByRole('textbox', { name: 'Slide text' });
  act(() => {
    fireEvent.pointerDown(canvas, point());
    fireEvent.pointerMove(stage, { ...point(point().clientX + 30), buttons: 1 });
    fireEvent.pointerUp(stage, point(point().clientX + 30));
    fireEvent.pointerDown(canvas, textPoint);
    fireEvent.pointerUp(stage, textPoint);
    input.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertText', data: '!',
      bubbles: true, cancelable: true }));
  });
  await bounded(client.flush());
  expect(commands.map((command) => command.kind)).toEqual(['selectObject', 'moveShape', 'selectAtPoint', 'insertText']);
  const saved = (await client.read()).result;
  expect(saved.ok).toBe(true);
  if (!saved.ok) throw new Error(saved.failure.message);
  expect(saved.stories[0]!.text).toBe('al!pha beta\nx\nalpha beta');
  expect(saved.slides[0]!.shapes[0]!.x).toBe(100000 + Math.round(30 * 9144000 / list.width));
});
