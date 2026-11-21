import { expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import { collaborationPipeFor } from '@betteroffice/pptx/collaboration';
import { decodeCommandDisplayList, presentationSharedInputs } from '@betteroffice/pptx/worker';
import type { AsyncCollaborationReplica, ShapeRect, TextBoxPrimitive } from '@betteroffice/pptx';
import { deck, open, replyControl, story, type } from './peerlessTestHarness';
import { createPeerlessViewStore, peerlessInputReference } from './peerlessViewStore';
import { peerlessPainterFactories } from './usePeerlessPresentation';

const { act, fireEvent, waitFor } = await import('@testing-library/react');

function outline(view: Awaited<ReturnType<typeof open>>['view'], resizeHandles = true) {
  const elements = view.getAllByTestId('pptx-text-edit-outline');
  expect(elements).toHaveLength(1);
  expect(view.queryAllByTestId(/pptx-resize-/)).toHaveLength(resizeHandles ? 8 : 0);
  const element = elements[0]!;
  expect(element.style.border).toBe('2px dashed #6b7280');
  expect(element.style.boxShadow).toBe('0 0 0 1px rgba(255, 255, 255, 0.9)');
  return element;
}

function bounds(element: HTMLElement, rect: ShapeRect, scale: number) {
  for (const [property, value] of [['left', rect.x], ['top', rect.y],
    ['width', rect.width], ['height', rect.height]] as const) {
    expect(parseFloat(element.style[property])).toBeCloseTo(value * scale, 3);
  }
}

async function editExisting(h: Awaited<ReturnType<typeof open>>) {
  const line = h.box.lines[0]!;
  const at = h.point(line.caretStops[2]!.x, line.y + line.height / 2);
  act(() => { fireEvent.pointerDown(h.canvas, at); fireEvent.pointerUp(h.stage, at); });
  await waitFor(() => expect(h.client.state.selection).toMatchObject({ shapeId: h.box.shapeId, anchor: 2, focus: 2 }));
  await waitFor(() => outline(h.view));
}

async function groupedDeck(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  const path = 'ppt/slides/slide1.xml';
  const xml = await zip.file(path)!.async('string');
  const shape = xml.match(/<p:sp>.*?<\/p:sp>/)![0];
  const child = shape.replace('id="2" name="Text"', 'id="4" name="Child"');
  const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="5143500"/><a:chOff x="0" y="0"/><a:chExt cx="9144000" cy="5143500"/></a:xfrm></p:grpSpPr>${child}</p:grpSp>`;
  zip.file(path, xml.replace(shape, group));
  return zip.generateAsync({ type: 'uint8array' });
}

function viewStore(client: Awaited<ReturnType<typeof open>>['client']) {
  return createPeerlessViewStore(client, { paintFrame() {}, paintEcho() {}, restoreLine() {} });
}

function resizeBehindNotes(h: Awaited<ReturnType<typeof open>>, store: ReturnType<typeof viewStore>, signal?: AbortSignal) {
  const { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation } = h.client.state.frame!;
  const frame = { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation };
  const at = { source: 'afterSeq' as const, seq: h.client.state.throughSeq };
  const notes = store.dispatch({ kind: 'invoke', payload: {
    method: 'setSlideNotes', args: [slideId, 'Updated notes'], inputTime: 1, undoGroup: 'notes',
  } });
  const resize = store.dispatch({ kind: 'setShapeRect', payload: {
    at, frame, shapeId: h.box.shapeId, handle: 'e', delta: { x: 30, y: 0 }, inputTime: 2, undoGroup: 'resize',
  } }, { signal, reportRefusal: false, waitForPredecessor: true });
  return { notes, resize };
}

test('resize released before object selection commits once in order', async () => {
  const h = await open();
  await editExisting(h);
  const edge = h.point(h.box.x + h.box.w, h.box.y + h.box.h / 2);
  const since = h.commands.length;
  replyControl.hold();
  try {
    act(() => {
      fireEvent.pointerDown(h.view.getByTestId('pptx-resize-e'), edge);
      fireEvent.pointerMove(h.stage, { ...edge, clientX: edge.clientX + 10, buttons: 1 });
      fireEvent.pointerUp(h.stage, { ...edge, clientX: edge.clientX + 30 });
    });
    await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
    expect(h.commands.slice(since).map((command) => command.kind)).toEqual(['selectObject']);
    await act(async () => { replyControl.releaseNext(); });
    await waitFor(() => expect(h.commands.slice(since).map((command) => command.kind))
      .toEqual(['selectObject', 'setShapeRect']));
    expect(h.commands.filter((command) => command.kind === 'setShapeRect')).toHaveLength(1);
  } finally {
    await act(async () => { replyControl.release(); });
  }
  await h.client.flush();
  expect(h.commands.filter((command) => command.kind === 'setShapeRect')).toHaveLength(1);
});

test('pending released resize remains cancellable before object selection replies', async () => {
  const h = await open();
  await editExisting(h);
  const edge = h.point(h.box.x + h.box.w, h.box.y + h.box.h / 2);
  replyControl.hold();
  try {
    act(() => {
      fireEvent.pointerDown(h.view.getByTestId('pptx-resize-e'), edge);
      fireEvent.pointerMove(h.stage, { ...edge, clientX: edge.clientX + 10, buttons: 1 });
      fireEvent.pointerUp(h.stage, { ...edge, clientX: edge.clientX + 30 });
      fireEvent.keyDown(h.stage, { key: 'Escape' });
    });
    await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
  } finally {
    await act(async () => { replyControl.release(); });
  }
  await h.client.flush();
  expect(h.commands.filter((command) => command.kind === 'setShapeRect')).toHaveLength(0);
});

test('text input after cancelling a released resize survives the pending selection', async () => {
  const h = await open();
  await editExisting(h);
  const line = h.box.lines[0]!;
  const textAt = h.point(line.caretStops[2]!.x, line.y + line.height / 2);
  const edge = h.point(h.box.x + h.box.w, h.box.y + h.box.h / 2);
  const since = h.commands.length;
  replyControl.hold();
  try {
    act(() => {
      fireEvent.pointerDown(h.view.getByTestId('pptx-resize-e'), edge);
      fireEvent.pointerMove(h.stage, { ...edge, clientX: edge.clientX + 10, buttons: 1 });
      fireEvent.pointerUp(h.stage, { ...edge, clientX: edge.clientX + 30 });
      fireEvent.keyDown(h.stage, { key: 'Escape' });
      fireEvent.pointerDown(h.canvas, textAt);
      fireEvent.pointerUp(h.stage, textAt);
      type(h.view.getByRole('textbox', { name: 'Slide text' }), '!');
    });
    await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
    expect(h.commands.slice(since).map((command) => command.kind))
      .toEqual(['selectObject', 'selectAtPoint', 'insertText']);
  } finally {
    await act(async () => { replyControl.release(); });
  }
  await h.client.flush();
  expect(h.commands.slice(since).map((command) => command.kind))
    .toEqual(['selectObject', 'selectAtPoint', 'insertText']);
  expect((await story(h.api)).text).toBe('al!pha beta\nx\nalpha beta');
});

test('cancelling a released resize preserves queued clicks and typing', async () => {
  const h = await open();
  await editExisting(h);
  const line = h.box.lines[0]!;
  const textAt = h.point(line.caretStops[2]!.x, line.y + line.height / 2);
  const edge = h.point(h.box.x + h.box.w, h.box.y + h.box.h / 2);
  const input = h.view.getByRole('textbox', { name: 'Slide text' });
  const since = h.commands.length;
  replyControl.hold();
  try {
    act(() => {
      fireEvent.pointerDown(h.view.getByTestId('pptx-resize-e'), edge);
      fireEvent.pointerMove(h.stage, { ...edge, clientX: edge.clientX + 10, buttons: 1 });
      fireEvent.pointerUp(h.stage, { ...edge, clientX: edge.clientX + 30 });
      fireEvent.pointerDown(h.canvas, textAt);
      fireEvent.pointerUp(h.stage, textAt);
      type(input, '!');
      fireEvent.keyDown(h.stage, { key: 'Escape' });
      type(input, '?');
    });
    await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
  } finally {
    await act(async () => { replyControl.release(); });
  }
  expect((await story(h.api)).text).toBe('al!?pha beta\nx\nalpha beta');
  expect(h.commands.slice(since).filter((command) => command.kind === 'setShapeRect')).toHaveLength(0);
});

test('aborting a resize in the view queue keeps it off the worker', async () => {
  const h = await open();
  await h.select();
  const store = viewStore(h.client);
  const controller = new AbortController();
  const since = h.commands.length;
  replyControl.hold();
  try {
    const { notes, resize } = resizeBehindNotes(h, store, controller.signal);
    expect(store.getSnapshot().waitingInputCount).toBe(1);
    controller.abort();
    await act(async () => { replyControl.release(); await notes; });
    expect((await resize).outcomes[0]).toMatchObject({ status: 'cancelled', committed: false });
    expect(h.commands.slice(since).map((command) => command.kind)).toEqual(['invoke']);
  } finally {
    await act(async () => { replyControl.release(); });
    store.dispose();
  }
});

test('shared input can submit without a cancellation method', async () => {
  const h = await open();
  const shared = presentationSharedInputs.get(h.client)!;
  const cancel = Object.getOwnPropertyDescriptor(shared, 'cancel')!;
  Reflect.deleteProperty(shared, 'cancel');
  const store = viewStore(h.client);
  try {
    let status: string | undefined;
    await act(async () => {
      status = await Promise.resolve().then(() => store.dispatch({ kind: 'setChangeObserver', payload: {
        enabled: false, inputTime: 1, undoGroup: 'observer',
      } }, { signal: AbortSignal.abort() })).then((ack) => ack.outcomes[0]?.status, () => undefined);
      await h.client.flush();
    });
    expect(status).toBe('noop');
  } finally {
    store.dispose();
    Object.defineProperty(shared, 'cancel', cancel);
  }
});

test('queued relative resize uses the updated selection predecessor', async () => {
  const h = await open();
  await h.select();
  const store = viewStore(h.client);
  const since = h.commands.length;
  replyControl.hold();
  try {
    const { notes, resize } = resizeBehindNotes(h, store);
    expect(store.getSnapshot().waitingInputCount).toBe(1);
    await waitFor(() => expect(replyControl.count()).toBe(1));
    await act(async () => { replyControl.releaseNext(); await notes; });
    const { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation } = h.client.state.frame!;
    const line = h.box.lines[0]!;
    const click = store.dispatch({ kind: 'selectAtPoint', payload: {
      frame: { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation },
      x: line.caretStops[2]!.x, y: line.y + line.height / 2, extend: false, inputTime: 3, undoGroup: 'click',
    } });
    const typed = store.dispatch({ kind: 'insertText', payload: {
      at: peerlessInputReference(store)!, text: '!', inputTime: 4, undoGroup: 'typing',
    } });
    await act(async () => { replyControl.release(); await Promise.all([resize, click, typed]); });
    const [predecessor, command] = h.commands.slice(since);
    expect(command).toMatchObject({ kind: 'setShapeRect', payload: { at: { source: 'afterSeq', seq: predecessor?.seq } } });
    expect((await resize).outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect((await story(h.api)).text).toBe('al!pha beta\nx\nalpha beta');
  } finally {
    await act(async () => { replyControl.release(); });
    store.dispose();
  }
});

test('a new shape gesture preserves an earlier released resize', async () => {
  const h = await open();
  await editExisting(h);
  const edge = h.point(h.box.x + h.box.w, h.box.y + h.box.h / 2);
  const since = h.commands.length;
  const outcomes: Array<{ seq: number; status: string; committed?: boolean }> = [];
  const unsubscribe = h.client.subscribe((ack) => { outcomes.push(...ack.outcomes); });
  replyControl.hold();
  try {
    act(() => {
      const handle = h.view.getByTestId('pptx-resize-e');
      fireEvent.pointerDown(handle, edge);
      fireEvent.pointerMove(h.stage, { ...edge, clientX: edge.clientX + 10, buttons: 1 });
      fireEvent.pointerUp(h.stage, { ...edge, clientX: edge.clientX + 30 });
      fireEvent.pointerDown(handle, edge);
      fireEvent.pointerMove(h.stage, { ...edge, clientX: edge.clientX + 5, buttons: 1 });
    });
    await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
    await act(async () => { replyControl.releaseNext(); });
    await waitFor(() => expect(h.commands.slice(since).some((command) => command.kind === 'setShapeRect')).toBe(true));
  } finally {
    await act(async () => { replyControl.release(); });
    act(() => { fireEvent.pointerCancel(h.stage, edge); });
    unsubscribe();
  }
  await h.client.flush();
  const resize = h.commands.slice(since).filter((command) => command.kind === 'setShapeRect');
  expect(resize).toHaveLength(1);
  expect(outcomes.filter((outcome) => outcome.seq === resize[0]!.seq))
    .toEqual([expect.objectContaining({ status: 'applied', committed: true })]);
  const saved = (await h.client.read()).result;
  expect(saved.ok).toBe(true);
  if (!saved.ok) throw new Error(saved.failure.message);
  expect(saved.slides[0]!.shapes[0]!.width)
    .toBe(6000000 + Math.round(30 * 9144000 / h.list.width));
});

test('a second handle drag resizes the shape and preserves subsequent typing', async () => {
  const h = await open();
  await editExisting(h);
  const line = h.box.lines[0]!;
  const textAt = h.point(line.caretStops[2]!.x, line.y + line.height / 2);
  const edge = h.point(h.box.x + h.box.w, h.box.y + h.box.h / 2);
  const since = h.commands.length;
  const outcomes: Array<{ seq: number; status: string }> = [];
  const unsubscribe = h.client.subscribe((ack) => { outcomes.push(...ack.outcomes); });
  replyControl.hold();
  try {
    act(() => {
      const handle = h.view.getByTestId('pptx-resize-e');
      fireEvent.pointerDown(handle, edge);
      fireEvent.pointerMove(h.stage, { ...edge, clientX: edge.clientX + 10, buttons: 1 });
      fireEvent.pointerUp(h.stage, { ...edge, clientX: edge.clientX + 30 });
      fireEvent.pointerDown(handle, edge);
      fireEvent.pointerMove(h.stage, { ...edge, clientX: edge.clientX + 10, buttons: 1 });
      fireEvent.pointerUp(h.stage, { ...edge, clientX: edge.clientX + 20 });
      fireEvent.pointerDown(h.canvas, textAt);
      fireEvent.pointerUp(h.stage, textAt);
      type(h.view.getByRole('textbox', { name: 'Slide text' }), '!');
    });
    await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
    await act(async () => { replyControl.release(); });
    const saved = await story(h.api);
    const commands = h.commands.slice(since);
    const selections = commands.filter((command) => command.kind === 'selectObject');
    const resizes = commands.filter((command) => command.kind === 'setShapeRect');
    expect(selections).toHaveLength(2);
    expect(outcomes.find((outcome) => outcome.seq === selections[1]?.seq)).toMatchObject({ status: 'applied' });
    expect(resizes).toHaveLength(2);
    expect(outcomes.filter((outcome) => resizes.some((resize) => resize.seq === outcome.seq))
      .map((outcome) => outcome.status)).toEqual(['applied', 'applied']);
    expect(saved.text).toBe('al!pha beta\nx\nalpha beta');
    expect(h.view.queryByRole('alert')).toBeNull();
  } finally {
    unsubscribe();
    await act(async () => { replyControl.release(); });
  }
});

test('a handle drag between two text inputs preserves the resize and both characters', async () => {
  const h = await open();
  await editExisting(h);
  const line = h.box.lines[0]!;
  const textAt = h.point(line.caretStops[1]!.x, line.y + line.height / 2);
  const edge = h.point(h.box.x + h.box.w, h.box.y + h.box.h / 2);
  const input = h.view.getByRole('textbox', { name: 'Slide text' });
  const outcomes: Array<{ status: string }> = [];
  const unsubscribe = h.client.subscribe((ack) => { outcomes.push(...ack.outcomes); });
  replyControl.hold();
  try {
    act(() => { type(input, '!'); });
    await waitFor(() => expect(replyControl.count()).toBe(1));
    act(() => {
      fireEvent.pointerDown(h.view.getByTestId('pptx-resize-e'), edge);
      fireEvent.pointerMove(h.stage, { ...edge, clientX: edge.clientX + 10, buttons: 1 });
      fireEvent.pointerUp(h.stage, { ...edge, clientX: edge.clientX + 30 });
      fireEvent.pointerDown(h.canvas, textAt);
      fireEvent.pointerUp(h.stage, textAt);
      type(input, '?');
    });
    await act(async () => { replyControl.release(); });
    const saved = await h.api.readContent();
    expect(saved.ok).toBe(true);
    if (!saved.ok) throw new Error(saved.failure.message);
    expect(saved.slides[0]!.shapes[0]!.width)
      .toBe(6000000 + Math.round(30 * 9144000 / h.list.width));
    expect(outcomes.some((outcome) => outcome.status === 'refused')).toBe(false);
    expect(saved.stories[0]!.text).toBe('a?l!pha beta\nx\nalpha beta');
  } finally {
    unsubscribe();
    await act(async () => { replyControl.release(); });
  }
});

test('a text click between two handle drags keeps both resizes and the typing after them', async () => {
  const h = await open();
  await editExisting(h);
  const line = h.box.lines[0]!;
  const textAt = h.point(line.caretStops[2]!.x, line.y + line.height / 2);
  const edge = h.point(h.box.x + h.box.w, h.box.y + h.box.h / 2);
  const since = h.commands.length;
  const outcomes: Array<{ seq: number; status: string }> = [];
  const unsubscribe = h.client.subscribe((ack) => { outcomes.push(...ack.outcomes); });
  replyControl.hold();
  try {
    act(() => {
      const handle = h.view.getByTestId('pptx-resize-e');
      fireEvent.pointerDown(handle, edge);
      fireEvent.pointerMove(h.stage, { ...edge, clientX: edge.clientX + 10, buttons: 1 });
      fireEvent.pointerUp(h.stage, { ...edge, clientX: edge.clientX + 30 });
      fireEvent.pointerDown(h.canvas, textAt);
      fireEvent.pointerUp(h.stage, textAt);
      fireEvent.pointerDown(h.view.getByTestId('pptx-resize-e'), edge);
      fireEvent.pointerMove(h.stage, { ...edge, clientX: edge.clientX + 10, buttons: 1 });
      fireEvent.pointerUp(h.stage, { ...edge, clientX: edge.clientX + 20 });
      fireEvent.pointerDown(h.canvas, textAt);
      fireEvent.pointerUp(h.stage, textAt);
      type(h.view.getByRole('textbox', { name: 'Slide text' }), '!');
    });
    await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
    await act(async () => { replyControl.release(); });
    const saved = await story(h.api);
    const commands = h.commands.slice(since);
    const selections = commands.filter((command) => command.kind === 'selectObject');
    const resizes = commands.filter((command) => command.kind === 'setShapeRect');
    expect(selections).toHaveLength(2);
    expect(outcomes.find((outcome) => outcome.seq === selections[1]?.seq)).toMatchObject({ status: 'applied' });
    expect(resizes).toHaveLength(2);
    expect(outcomes.filter((outcome) => resizes.some((resize) => resize.seq === outcome.seq))
      .map((outcome) => outcome.status)).toEqual(['applied', 'applied']);
    expect(saved.text).toBe('al!pha beta\nx\nalpha beta');
    expect(h.view.queryByRole('alert')).toBeNull();
  } finally {
    unsubscribe();
    await act(async () => { replyControl.release(); });
  }
});

test('a collaboration cursor sent before pointer-up preserves resize and subsequent typing', async () => {
  let replica!: AsyncCollaborationReplica;
  const h = await open({ collaboration: { clientId: 1001, onReplica: (value) => { if (value) replica = value; } } });
  await editExisting(h);
  const endpoint = collaborationPipeFor(replica)!.open(() => {}, () => {});
  const line = h.box.lines[0]!;
  const textAt = h.point(line.caretStops[2]!.x, line.y + line.height / 2);
  const edge = h.point(h.box.x + h.box.w, h.box.y + h.box.h / 2);
  const outcomes: Array<{ seq: number; status: string; committed?: boolean }> = [];
  const unsubscribe = h.client.subscribe((ack) => { outcomes.push(...ack.outcomes); });
  try {
    await endpoint.post([{ op: 'create', seq: 1, options: {} }], []);
    const since = h.commands.length;
    replyControl.hold();
    act(() => { fireEvent.pointerDown(h.view.getByTestId('pptx-resize-e'), edge); });
    const cursor = endpoint.post([{ op: 'cursor', seq: 2, cursor: { slideId: h.client.state.activeSlideId! },
      broadcast: false }], []);
    await waitFor(() => expect(replyControl.count()).toBe(2));
    act(() => {
      fireEvent.pointerMove(h.stage, { ...edge, clientX: edge.clientX + 10, buttons: 1 });
      fireEvent.pointerUp(h.stage, { ...edge, clientX: edge.clientX + 30 });
      fireEvent.pointerDown(h.canvas, textAt);
      fireEvent.pointerUp(h.stage, textAt);
      type(h.view.getByRole('textbox', { name: 'Slide text' }), '!');
    });
    await act(async () => { replyControl.release(); await cursor; });
    const saved = await story(h.api);
    const [selection, pipe, resize] = h.commands.slice(since);
    expect([selection?.kind, pipe?.kind, resize?.kind]).toEqual(['selectObject', 'collabPipe', 'setShapeRect']);
    expect(resize).toMatchObject({ afterSeq: selection?.seq });
    expect(outcomes.find((outcome) => outcome.seq === pipe?.seq)).toMatchObject({ status: 'noop', committed: false });
    expect(outcomes.find((outcome) => outcome.seq === resize?.seq)).toMatchObject({ status: 'applied', committed: true });
    expect(saved.text).toBe('al!pha beta\nx\nalpha beta');
  } finally {
    unsubscribe();
    await act(async () => { replyControl.release(); });
    endpoint.close();
  }
});

test('a point-based object selection keeps the frame the user clicked', async () => {
  const h = await open();
  await editExisting(h);
  const { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation } = h.client.state.frame!;
  const edge = h.point(h.box.x, h.box.y + h.box.h / 2);
  const click = h.point();
  const since = h.commands.length;
  replyControl.hold();
  try {
    act(() => {
      fireEvent.pointerDown(h.view.getByTestId('pptx-resize-w'), edge);
      fireEvent.pointerMove(h.stage, { ...edge, clientX: edge.clientX + 10, buttons: 1 });
      fireEvent.pointerUp(h.stage, { ...edge, clientX: edge.clientX + 30 });
      fireEvent.pointerDown(h.canvas, click);
      fireEvent.pointerUp(h.stage, click);
    });
    await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
    await act(async () => { replyControl.release(); });
    await h.api.readContent();
    const [, resize, selection] = h.commands.slice(since);
    expect(selection).toMatchObject({ kind: 'selectObject', afterSeq: resize?.seq, payload: {
      scale: 1, frame: { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation },
    } });
  } finally {
    await act(async () => { replyControl.release(); });
  }
});

test('grouped child text hides resize handles while top-level text retains them', async () => {
  const h = await open({ file: await groupedDeck(await deck()) });
  const line = h.box.lines[0]!;
  const at = h.point(line.caretStops[2]!.x, line.y + line.height / 2);
  act(() => { fireEvent.pointerDown(h.canvas, at); fireEvent.pointerUp(h.stage, at); });
  await waitFor(() => expect(h.client.state.selection).toMatchObject({
    shapeId: h.box.shapeId, anchor: 2, focus: 2, canResize: false,
  }));
  await waitFor(() => outline(h.view, false));
});

for (const [start, end] of [[[200.123, 150.456], [400.789, 220.123]],
  [[900.123, 500.456], [980.789, 550.123]], [[200, 150], [201, 151]]] as const) {
  test(`textbox outline survives pointerup at ${start.join(',')}`, async () => {
    const h = await open();
    const scale = parseFloat(h.canvas.style.width) / h.list.width;
    let rect: ShapeRect | undefined;
    const input = presentationSharedInputs.get(h.client)!;
    const submit = input.submit.bind(input);
    const capture = spyOn(input, 'submit').mockImplementation((intent, basis, after) => {
      if (intent.kind === 'insertTextBox') rect = intent.payload.draft.rect;
      return submit(intent, basis, after);
    });
    try {
      fireEvent.click(h.view.getByRole('button', { name: 'Text box' }));
      act(() => {
        fireEvent.pointerDown(h.canvas, h.point(...start));
        fireEvent.pointerMove(h.stage, h.point(...end));
      });
      expect(h.view.getByTestId('pptx-insertion-preview')).toBeDefined();
      replyControl.hold();
      act(() => { fireEvent.pointerUp(h.stage, h.point(...end)); });
      await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
      expect(h.view.queryByTestId('pptx-insertion-preview') === null).toBe(true);
      expect(rect).toBeDefined();
      const echo = outline(h.view, false);
      bounds(echo, { x: rect!.x * h.list.width / h.client.state.size.width,
        y: rect!.y * h.list.height / h.client.state.size.height,
        width: rect!.width * h.list.width / h.client.state.size.width,
        height: rect!.height * h.list.height / h.client.state.size.height }, scale);
      const echoStyle = echo.getAttribute('style');
      await act(async () => { replyControl.release(); });
      await waitFor(() => expect(h.client.state.selection).toMatchObject({ anchor: 0, focus: 0 }));
      await waitFor(() => {
        const selected = h.client.state.selection!;
        const box = decodeCommandDisplayList(h.client.state.frame!).primitives.find((p) =>
          p.kind === 'textBox' && p.shapeId === selected.shapeId) as TextBoxPrimitive;
        expect(box).toBeDefined();
        bounds(outline(h.view), { x: box.x, y: box.y, width: box.w, height: box.h }, scale);
        expect(outline(h.view).getAttribute('style')).toBe(echoStyle);
        const caret = h.view.getByTestId('pptx-caret');
        expect(caret.style.visibility).toBe('visible');
        expect(parseFloat(caret.style.left)).toBeGreaterThanOrEqual(box.x);
        expect(parseFloat(caret.style.left)).toBeLessThan(box.x + box.w);
      });
      fireEvent.click(h.view.getByTestId('pptx-undo'));
      await waitFor(() => expect(h.view.queryByTestId('pptx-text-edit-outline') === null).toBe(true));
    } finally { capture.mockRestore(); replyControl.release(); }
  });
}

test('existing text editing survives Escape and leaves with object selection and slide navigation', async () => {
  const h = await open();
  await editExisting(h);
  bounds(outline(h.view), { x: h.box.x, y: h.box.y, width: h.box.w, height: h.box.h },
    parseFloat(h.canvas.style.width) / h.list.width);
  await h.select();
  expect(h.view.queryByTestId('pptx-text-edit-outline') === null).toBe(true);
  await editExisting(h);
  const selection = { ...h.client.state.selection! };
  const style = outline(h.view).getAttribute('style');
  const submit = spyOn(presentationSharedInputs.get(h.client)!, 'submit');
  try {
    await act(async () => {
      fireEvent.keyDown(h.view.getByRole('textbox', { name: 'Slide text' }), { key: 'Escape' });
      await new Promise<void>((resolve) => setImmediate(resolve));
    });
    expect(submit).not.toHaveBeenCalled();
    expect(h.client.state.selection).toEqual(selection);
    expect(outline(h.view).getAttribute('style')).toBe(style);
  } finally { submit.mockRestore(); }
  fireEvent.click(h.view.getByRole('button', { name: 'Slide 2 of 2' }));
  await waitFor(() => expect(h.client.state.activeSlideId).toBe(h.client.state.slides[1]!.id));
  await waitFor(() => expect(h.view.queryByTestId('pptx-text-edit-outline') === null).toBe(true));
});

test('Escape cancels a textbox drag preview without inserting', async () => {
  const h = await open();
  const submit = spyOn(presentationSharedInputs.get(h.client)!, 'submit');
  try {
    fireEvent.click(h.view.getByRole('button', { name: 'Text box' }));
    act(() => {
      fireEvent.pointerDown(h.canvas, h.point(200, 150));
      fireEvent.pointerMove(h.stage, h.point(400, 220));
    });
    expect(h.view.queryByTestId('pptx-insertion-preview') !== null).toBe(true);
    fireEvent.keyDown(h.stage, { key: 'Escape' });
    expect(h.view.queryByTestId('pptx-insertion-preview') === null).toBe(true);
    expect(h.view.queryByTestId('pptx-text-edit-outline') === null).toBe(true);
    act(() => { fireEvent.pointerUp(h.stage, h.point(400, 220)); });
    expect(submit).not.toHaveBeenCalled();
  } finally { submit.mockRestore(); }
});

for (const change of ['failure', 'tool', 'slide', 'readOnly', 'replacement', 'Escape'] as const) {
  test(`pending textbox outline clears on ${change}`, async () => {
    const h = await open();
    const input = presentationSharedInputs.get(h.client)!;
    const submit = input.submit.bind(input);
    const refusal = spyOn(input, 'submit').mockImplementation((intent, basis, after) => submit(
      change === 'failure' && intent.kind === 'insertTextBox'
        ? { ...intent, payload: { ...intent.payload, draft: { ...intent.payload.draft, text: '\u0000' } } } : intent,
      basis, after));
    try {
      fireEvent.click(h.view.getByRole('button', { name: 'Text box' }));
      replyControl.hold();
      act(() => {
        fireEvent.pointerDown(h.canvas, h.point(200, 150));
        fireEvent.pointerMove(h.stage, h.point(400, 220));
        fireEvent.pointerUp(h.stage, h.point(400, 220));
      });
      await waitFor(() => expect(replyControl.count()).toBeGreaterThan(0));
      outline(h.view, false);
      if (change === 'tool') await act(async () => { expect(await h.api.commands.execute('tool', { value: 'textBox' })).toMatchObject({ ok: true }); });
      if (change === 'slide') fireEvent.click(h.view.getByRole('button', { name: 'Slide 2 of 2' }));
      if (change === 'readOnly') h.view.rerender(h.ui(true));
      if (change === 'replacement') h.view.rerender(<div />);
      if (change === 'Escape') fireEvent.keyDown(h.view.getByRole('textbox', { name: 'Slide text' }), { key: 'Escape' });
      if (change === 'tool' || change === 'readOnly' || change === 'replacement' || change === 'Escape') {
        await waitFor(() => expect(h.view.queryByTestId('pptx-text-edit-outline') === null).toBe(true));
      }
      if (change === 'Escape') expect(h.view.queryByTestId('pptx-insertion-preview') === null).toBe(true);
      await act(async () => { replyControl.release(); });
      if (change === 'failure') await waitFor(() => expect(h.client.state.outcomes).toEqual(
        expect.arrayContaining([expect.objectContaining({ status: 'refused' })])));
      if (change === 'failure' || change === 'slide') {
        await waitFor(() => expect(h.view.queryByTestId('pptx-text-edit-outline') === null).toBe(true));
      }
      if (change === 'Escape') await waitFor(() => {
        expect(h.client.state.selection).toMatchObject({ anchor: 0, focus: 0 });
        outline(h.view);
      });
    } finally { refusal.mockRestore(); replyControl.release(); }
  });
}

test('existing rotated and flipped text uses the displayed shape bounds', async () => {
  const seed = await open();
  const zip = await JSZip.loadAsync(seed.bytes);
  const source = await zip.file('ppt/slides/slide1.xml')!.async('string');
  zip.file('ppt/slides/slide1.xml', source.replace('<a:xfrm>', '<a:xfrm rot="1800000" flipH="1">'));
  seed.view.unmount();
  await seed.client.dispose();
  const h = await open({ file: await zip.generateAsync({ type: 'uint8array' }) });
  await act(async () => { await h.client.dispatch({ kind: 'setSelection', payload: {
    selection: { slideId: h.client.state.slides[0]!.id, shapeId: h.box.shapeId!, storyId: h.box.storyId!,
      anchor: 2, focus: 2, affinity: 'downstream' }, inputTime: 1, undoGroup: 'selection' } }); });
  await waitFor(() => expect(h.client.state.selection).toMatchObject({ anchor: 2, focus: 2 }));
  const angle = Math.PI / 6;
  const width = Math.abs(h.box.w * Math.cos(angle)) + Math.abs(h.box.h * Math.sin(angle));
  const height = Math.abs(h.box.w * Math.sin(angle)) + Math.abs(h.box.h * Math.cos(angle));
  await waitFor(() => bounds(outline(h.view, false), { x: h.box.x + (h.box.w - width) / 2,
    y: h.box.y + (h.box.h - height) / 2, width, height }, parseFloat(h.canvas.style.width) / h.list.width));
});

test('textbox echo waits for display and yields to changed confirmed bounds', async () => {
  const factory = peerlessPainterFactories.frame;
  let holding = false;
  let paint: (() => void) | undefined;
  const painter = spyOn(peerlessPainterFactories, 'frame').mockImplementation((canvas, options) => {
    const value = factory(canvas, options);
    return { ...value, paintFrame(frame, view, media) {
      if (holding && canvas.dataset.testid === 'pptx-slide-canvas') {
        paint = () => value.paintFrame(frame, view, media);
      } else value.paintFrame(frame, view, media);
    } };
  });
  let restore = () => {};
  try {
    const h = await open();
    const input = presentationSharedInputs.get(h.client)!;
    const submit = input.submit.bind(input);
    const capture = spyOn(input, 'submit').mockImplementation((intent, basis, after) => submit(
      intent.kind === 'insertTextBox' ? { ...intent, payload: { ...intent.payload, draft: { ...intent.payload.draft,
        rect: { ...intent.payload.draft.rect, width: intent.payload.draft.rect.width + 952500 } } } } : intent, basis, after));
    restore = () => capture.mockRestore();
    holding = true;
    fireEvent.click(h.view.getByRole('button', { name: 'Text box' }));
    act(() => {
      fireEvent.pointerDown(h.canvas, h.point(200, 150));
      fireEvent.pointerMove(h.stage, h.point(400, 220));
      fireEvent.pointerUp(h.stage, h.point(400, 220));
    });
    await waitFor(() => expect(h.client.state.selection).toMatchObject({ anchor: 0, focus: 0 }));
    await waitFor(() => expect(paint).toBeDefined());
    const scale = parseFloat(h.canvas.style.width) / h.list.width;
    bounds(outline(h.view, false), { x: 200, y: 150, width: 200, height: 70 }, scale);
    holding = false;
    await act(async () => { paint!(); });
    await waitFor(() => bounds(outline(h.view), { x: 200, y: 150, width: 300, height: 70 }, scale));
  } finally { restore(); painter.mockRestore(); }
});
