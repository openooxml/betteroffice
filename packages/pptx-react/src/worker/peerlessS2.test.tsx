import './peerlessTestCapabilities';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type { PptxWorkerSession, PresentationCommandAck, PresentationCommandClient,
  PresentationCommandEvents, PresentationCommandMethods } from '@betteroffice/pptx';
import { createSessionHost } from '../../../../shared/office-session/host';
import { isHostMessage } from '../../../../shared/office-session/protocol';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { createSharedPresentationCommandClient } from '../../../pptx/src/session/commandClient';
import { registerPresentationCommandHandlers } from '../../../pptx/src/session/commandHost';
import * as wasm from '../../../pptx/src/wasm/loader';
import { PEERLESS_FLAG, PptxEditor } from '../PptxEditor';
import { peerlessPainterFactories, peerlessSessionOpener } from './usePeerlessPresentation';
import type { PeerlessViewSnapshot } from './peerlessViewStore';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
let owner: PptxWorkerSession;
let client: PresentationCommandClient;
let handle: wasm.PresentationHandle;
let hold = false;
let replies: Map<number, () => void>;
let paints: PeerlessViewSnapshot[];
let previews: string[];
let previewViews: PeerlessViewSnapshot[];

async function deck(): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  const ids = [1, 2, 3];
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${ids.map((id) => `<Override PartName="/ppt/slides/slide${id}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst>${ids.map((id) => `<p:sldId id="${255 + id}" r:id="rId${id}"/>`).join('')}</p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${ids.map((id) => `<Relationship Id="rId${id}" Type="${rel}/slide" Target="slides/slide${id}.xml"/>`).join('')}</Relationships>`);
  for (const id of ids) zip.file(`ppt/slides/slide${id}.xml`, `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>slide ${id}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}

beforeEach(() => {
  hold = false;
  replies = new Map(); paints = []; previews = []; previewViews = [];
  const worker = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
  Object.defineProperty(globalThis, 'Worker', { configurable: true, value: class {} });
  restorers.push(() => {
    if (worker) Object.defineProperty(globalThis, 'Worker', worker);
    else Reflect.deleteProperty(globalThis, 'Worker');
  });
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (this: HTMLCanvasElement) {
    return new Proxy({ canvas: this, measureText: (text: string) => ({ width: text.length * 10 }) }, {
      get(target, key) { return key in target ? Reflect.get(target, key) : () => {}; },
    }) as unknown as CanvasRenderingContext2D;
  } as unknown as HTMLCanvasElement['getContext']);
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
  const openPresentation = wasm.openPresentation;
  const opened = spyOn(wasm, 'openPresentation').mockImplementation((...args) => {
    handle = openPresentation(...args);
    return handle;
  });
  restorers.push(() => opened.mockRestore());
  const factory = peerlessPainterFactories.frame;
  const painter = spyOn(peerlessPainterFactories, 'frame').mockImplementation((canvas, options) => {
    return factory(canvas, { ...options, onCommit(frame, view) {
      if (canvas.dataset.testid === 'pptx-slide-canvas') paints.push(view);
      else if (canvas.closest('aside')) { previews.push(frame.frame.slideId); previewViews.push(view); }
      options?.onCommit?.(frame, view);
    } });
  });
  restorers.push(() => painter.mockRestore());
  const opener = spyOn(peerlessSessionOpener, 'open').mockImplementation((bytes, options = {}) => {
    const pair = createInProcessPair();
    const registration = registerPresentationCommandHandlers({ initWasm: async () => {
      await wasm.initWasm(await readFile(new URL('../../../pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
    } });
    createSessionHost<PresentationCommandMethods, PresentationCommandEvents, null>({ ...pair.host,
      post(message, transfer) {
        if (hold && isHostMessage(message) && message.kind === 'reply' && message.ok &&
          message.value && typeof message.value === 'object' && 'outcomes' in message.value) {
          const ack = message.value as PresentationCommandAck;
          if (ack.outcomes.length) {
            replies.set(ack.throughSeq, () => { replies.delete(ack.throughSeq); pair.host.post(message, transfer); });
            return;
          }
        }
        pair.host.post(message, transfer);
      },
    }, { handlers: registration.handlers, policies: registration.policies,
      context: null, onDispose: registration.dispose });
    const ready = readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url))
      .then((font) => createSharedPresentationCommandClient(bytes, { session: 's2', incarnation: 's2-worker',
        readOnly: options.readOnly, activeSlideIndex: options.activeSlideIndex,
        fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }] }, pair.client))
      .then((value) => { client = value; return value; });
    owner = { ready, failure: undefined, subscribeFailure: () => () => {},
      dispose: async () => { await ready.then((value) => value.dispose(), () => {}); registration.dispose(); } };
    return owner;
  });
  restorers.push(() => opener.mockRestore());
});

afterEach(async () => {
  hold = false;
  for (const deliver of replies.values()) deliver();
  replies.clear();
  cleanup();
  await owner?.dispose();
  for (const restore of restorers.splice(0).reverse()) restore();
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

async function surface(onError?: (error: Error) => void) {
  const ready = mock(() => {});
  const view = render(<PptxEditor fonts={[]} file={await deck()} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} onError={onError} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  return { ...view, items: Array.from(view.container.querySelectorAll('aside button')) as HTMLButtonElement[] };
}

test('strip_click', async () => {
  const view = await surface();
  const id = client.state.slides[1]!.id;
  fireEvent.click(view.items[1]);
  await waitFor(() => expect(paints[paints.length - 1]?.confirmed?.frame.slideId).toBe(id));
  expect(view.items[1].getAttribute('aria-current')).toBe('page');
  expect(client.state.selection).toBeNull();
});

test('rapid_navigation', async () => {
  const view = await surface();
  const before = paints[paints.length - 1]!.confirmed;
  await waitFor(() => expect(previews).toContain(client.state.slides[2]!.id));
  hold = true;
  fireEvent.click(view.items[1]);
  fireEvent.click(view.items[2]);
  await waitFor(() => expect(replies.size).toBe(2));
  expect(paints[paints.length - 1]?.confirmed).toBe(before);
  expect(view.items[0].getAttribute('aria-current')).toBe('page');
  expect(view.items[1].getAttribute('aria-current')).toBeNull();
  expect(view.items[2].getAttribute('aria-current')).toBeNull();
  const sequences = [...replies.keys()].sort((a, b) => a - b);
  await act(async () => { replies.get(sequences[1])!(); });
  await waitFor(() => expect(paints[paints.length - 1]?.confirmed?.frame.slideId).toBe(client.state.slides[2]!.id));
  expect(paints[paints.length - 1]?.confirmed?.frame.coversThroughSeq).toBe(sequences[1]);
  await act(async () => { replies.get(sequences[0])!(); });
  expect(paints[paints.length - 1]?.confirmed?.frame.slideId).toBe(client.state.slides[2]!.id);
  expect(client.state.activeSlideId).toBe(client.state.slides[2]!.id);
  expect(view.items[2].getAttribute('aria-current')).toBe('page');
});

test('active_preview', async () => {
  const view = await surface();
  await waitFor(() => expect(previews).toContain(client.state.activeSlideId!));
  expect(view.items[0].querySelector('canvas')!.width).toBeGreaterThan(0);
});

test('edit_then_navigate', async () => {
  const view = await surface();
  const read = (await client.read()).result;
  if (!read.ok) throw new Error(read.failure.message);
  const story = read.stories[0]!;
  await act(async () => { await client.dispatch({ kind: 'setSelection', payload: {
    selection: { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
      anchor: 0, focus: 0, affinity: 'downstream' }, inputTime: 1, undoGroup: 's2' } }); });
  hold = true;
  const input = view.getByRole('textbox', { name: 'Slide text' });
  act(() => {
    input.focus();
    input.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, cancelable: true,
      inputType: 'insertText', data: 'typed ' }));
    fireEvent.click(view.items[1]);
  });
  await waitFor(() => expect(replies.size).toBe(1));
  await act(async () => { replies.values().next().value!(); replies.clear(); });
  await waitFor(() => expect(replies.size).toBe(1));
  await act(async () => { replies.values().next().value!(); replies.clear(); });
  hold = false;
  await waitFor(() => expect(paints[paints.length - 1]?.confirmed?.frame.slideId).toBe(client.state.slides[1]!.id));
  const navigationSeq = paints[paints.length - 1]!.confirmed!.frame.coversThroughSeq;
  expect(view.items[0].getAttribute('aria-current')).toBeNull();
  await waitFor(() => expect(previewViews.some((entry) => entry.confirmed?.frame.slideId === story.slideId &&
    entry.confirmed.frame.coversThroughSeq >= navigationSeq &&
    entry.confirmed.displayList.primitives.some((primitive) => primitive.kind === 'textBox' &&
      primitive.lines.some((line) => line.runs.map((run) => run.text).join('') === 'typed slide 1')))).toBe(true));
  const result = (await client.read()).result;
  if (!result.ok) throw new Error(result.failure.message);
  expect(result.stories.find((entry) => entry.slideId === story.slideId)?.text).toBe('typed slide 1');
  expect(result.stories.find((entry) => entry.slideId === client.state.slides[1]!.id)?.text).toBe('slide 2');
});

test('invalid_target', async () => {
  const view = await surface();
  const active = client.state.activeSlideId!;
  const target = client.state.slides[2]!.id;
  await waitFor(() => expect(previews).toContain(target));
  handle.deleteSlide(target);
  let reply: PresentationCommandAck | undefined;
  restorers.push(client.subscribe((ack) => { reply = ack; }));
  fireEvent.click(view.items[2]);
  await waitFor(() => expect(reply?.outcomes).toEqual([
    expect.objectContaining({ status: 'refused', reason: 'bad-target', committed: false }),
  ]));
  expect(client.state.activeSlideId).toBe(active);
  expect(paints[paints.length - 1]?.confirmed?.frame.slideId).toBe(active);
});

for (const target of ['edited_active', 'unloaded_inactive'] as const) {
  test(`deleted_${target}_thumbnail_does_not_fail_editor`, async () => {
    const callbacks = new Map<number, () => void>();
    let nextId = 0;
    for (const [property, value] of [
      ['requestIdleCallback', (callback: () => void) => { callbacks.set(++nextId, callback); return nextId; }],
      ['cancelIdleCallback', (id: number) => { callbacks.delete(id); }],
    ] as const) {
      const original = Object.getOwnPropertyDescriptor(window, property);
      Object.defineProperty(window, property, { configurable: true, value });
      restorers.push(() => {
        if (original) Object.defineProperty(window, property, original);
        else Reflect.deleteProperty(window, property);
      });
    }
    const errors = mock((_error: Error) => {});
    const view = await surface(errors);
    const removedId = client.state.slides[target === 'edited_active' ? 0 : 2]!.id;
    const liveId = client.state.slides[1]!.id;
    if (target === 'edited_active') {
      const read = (await client.read()).result;
      if (!read.ok) throw new Error(read.failure.message);
      const story = read.stories[0]!;
      await act(async () => { await client.dispatch({ kind: 'insertText', payload: {
        at: { source: 'confirmed', selection: { slideId: story.slideId, shapeId: story.shapeId,
          storyId: story.storyId, anchor: 0, focus: 0, affinity: 'downstream' } },
        text: 'typed ', inputTime: 1, undoGroup: 's2-delete',
      } }); });
    }
    expect(previews).not.toContain(client.state.slides[2]!.id);
    handle.addUndoBoundary();
    handle.deleteSlide(removedId);
    handle.addUndoBoundary();
    fireEvent.click(view.items[1]);
    await waitFor(() => expect(paints[paints.length - 1]?.confirmed?.frame.slideId).toBe(liveId));
    const requestThumbnail = client.thumbnail.bind(client);
    let settled = 0;
    const thumbnail = spyOn(client, 'thumbnail').mockImplementation((...args) =>
      requestThumbnail(...args).finally(() => { settled += 1; }));
    restorers.push(() => thumbnail.mockRestore());
    const drainIdle = async () => {
      await act(async () => {
        const queued = [...callbacks.values()];
        callbacks.clear();
        for (const callback of queued) callback();
      });
      await waitFor(() => expect(settled).toBe(thumbnail.mock.calls.length));
      await act(async () => {});
    };
    for (let turn = 0; turn < 4; turn += 1) await drainIdle();
    if (target === 'edited_active') {
      await waitFor(() => expect(previews).toContain(client.state.slides.find((slide) =>
        slide.id !== liveId && slide.id !== removedId)!.id));
    }
    expect(errors).not.toHaveBeenCalled();
    expect(callbacks.size).toBe(0);
    expect(thumbnail.mock.calls.filter(([slideId]) => slideId === removedId)).toHaveLength(1);
    expect(view.getByRole('application').getAttribute('aria-readonly')).toBe('false');
    expect(client.state.activeSlideId).toBe(liveId);
    const survivor = client.state.slides.find((slide) => slide.id !== liveId && slide.id !== removedId)!;
    fireEvent.click(view.items[survivor.index]);
    await waitFor(() => expect(paints[paints.length - 1]?.confirmed?.frame.slideId).toBe(survivor.id));
    for (let turn = 0; turn < 4; turn += 1) await drainIdle();
    expect(callbacks.size).toBe(0);
    expect(thumbnail.mock.calls.filter(([slideId]) => slideId === removedId)).toHaveLength(1);
    expect(errors).not.toHaveBeenCalled();
    const restored = await act(async () => client.dispatch({ kind: 'undo', payload: { inputTime: 2, undoGroup: 'history' } }));
    expect(restored.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    for (let turn = 0; turn < 5; turn += 1) await drainIdle();
    expect(client.state.activeSlideId).toBe(removedId);
    await waitFor(() => expect(paints[paints.length - 1]?.confirmed?.frame.slideId).toBe(removedId));
    expect(previews).toContain(removedId);
    expect(errors).not.toHaveBeenCalled();
  });
}

test('stage_keyboard', async () => {
  const view = await surface();
  const stage = view.getByRole('application');
  for (const [key, index] of [['PageDown', 1], ['End', 2], ['PageUp', 1], ['Home', 0]] as const) {
    stage.focus();
    fireEvent.keyDown(stage, { key });
    await waitFor(() => expect(client.state.activeSlideId).toBe(client.state.slides[index]!.id));
  }
  const input = view.getByRole('textbox', { name: 'Slide text' });
  const before = { seq: client.state.throughSeq, active: client.state.activeSlideId };
  input.focus();
  fireEvent.keyDown(input, { key: 'PageDown' });
  await act(async () => { await client.flush(); });
  expect(client.state.throughSeq).toBe(before.seq);
  expect(client.state.activeSlideId).toBe(before.active);
});
