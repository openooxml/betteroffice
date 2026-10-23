import './peerlessTestCapabilities';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type { PptxWorkerSession, PresentationCommandClient, PresentationCommandEnvelope,
  PresentationCommandEvents, PresentationCommandMethods } from '@betteroffice/pptx';
import { createSessionHost } from '../../../../shared/office-session/host';
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
let owner: PptxWorkerSession | undefined;
let client: PresentationCommandClient;
let commands: PresentationCommandEnvelope[];
let painted: PeerlessViewSnapshot | undefined;
let handle: wasm.PresentationHandle;

async function deck(): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>alpha</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}

beforeEach(() => {
  owner = undefined; commands = []; painted = undefined;
  const original = wasm.openPresentation;
  const opened = spyOn(wasm, 'openPresentation').mockImplementation((...args) => {
    handle = original(...args);
    return handle;
  });
  restorers.push(() => opened.mockRestore());
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
  const factory = peerlessPainterFactories.frame;
  const painter = spyOn(peerlessPainterFactories, 'frame').mockImplementation((canvas, options) =>
    factory(canvas, { ...options, onCommit(frame, view) {
      if (canvas.dataset.testid === 'pptx-slide-canvas') painted = view;
      options?.onCommit?.(frame, view);
    } }));
  restorers.push(() => painter.mockRestore());
  const echoFactory = peerlessPainterFactories.echo;
  const echo = spyOn(peerlessPainterFactories, 'echo').mockImplementation((canvas, options) => {
    const value = echoFactory(canvas, options);
    return { ...value, paintEcho(view) {
      if (canvas.dataset.testid === 'pptx-slide-canvas') painted = view;
      value.paintEcho(view);
    } };
  });
  restorers.push(() => echo.mockRestore());
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
    const ready = readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url))
      .then((font) => createSharedPresentationCommandClient(bytes, { session: 's4', incarnation: 's4-worker',
        readOnly: options.readOnly, activeSlideIndex: options.activeSlideIndex,
        fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }] }, pair.client))
      .then((value) => { client = value; return value; });
    owner = { ready, failure: undefined, subscribeFailure: () => () => {},
      dispose: async () => { try { await ready.then((value) => value.dispose(), () => {}); }
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

test('history_chords', async () => {
  const ready = mock(() => {});
  const view = render(<PptxEditor fonts={[]} file={await deck()} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const undo = view.getByTestId('pptx-undo') as HTMLButtonElement;
  const redo = view.getByTestId('pptx-redo') as HTMLButtonElement;
  expect((undo.getAttribute('aria-disabled') === 'true')).toBe(true);
  expect((redo.getAttribute('aria-disabled') === 'true')).toBe(true);
  const content = (await client.read()).result;
  if (!content.ok) throw new Error(content.failure.message);
  const story = content.stories[0]!;
  await act(async () => { await client.dispatch({ kind: 'insertText', payload: {
    at: { source: 'confirmed', selection: { slideId: story.slideId, shapeId: story.shapeId,
      storyId: story.storyId, anchor: 0, focus: 0, affinity: 'downstream' } },
    text: 'Z', inputTime: 1, undoGroup: 'typing' } }); });
  await waitFor(() => expect((undo.getAttribute('aria-disabled') === 'true')).toBe(false));
  commands.length = 0;
  const input = view.getByRole('textbox', { name: 'Slide text' });
  for (const [key, shiftKey, kind] of [['z', false, 'undo'], ['z', true, 'redo'],
    ['z', false, 'undo'], ['y', false, 'redo']] as const) {
    const count = commands.length;
    act(() => { input.focus(); fireEvent.keyDown(input, { key, shiftKey, ctrlKey: true }); });
    await waitFor(() => expect(commands).toHaveLength(count + 1));
    expect(commands[commands.length - 1]?.kind).toBe(kind);
    await waitFor(() => expect(client.state.throughSeq).toBe(commands[commands.length - 1]!.seq));
    await waitFor(() => {
      expect((undo.getAttribute('aria-disabled') === 'true')).toBe(!client.state.editState.canUndo);
      expect((redo.getAttribute('aria-disabled') === 'true')).toBe(!client.state.editState.canRedo);
    });
  }
  act(() => { fireEvent.click(undo); });
  await waitFor(() => expect(commands).toHaveLength(5));
  expect(commands[commands.length - 1]?.kind).toBe('undo');
  await waitFor(() => expect((redo.getAttribute('aria-disabled') === 'true')).toBe(false));
  act(() => { fireEvent.click(redo); });
  await waitFor(() => expect(commands).toHaveLength(6));
  expect(commands[commands.length - 1]?.kind).toBe('redo');
  await waitFor(() => expect((undo.getAttribute('aria-disabled') === 'true')).toBe(false));
  await act(async () => { await client.dispatch({ kind: 'setSelection', payload: {
    selection: { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
      anchor: 0, focus: 0, affinity: 'downstream' }, inputTime: 2, undoGroup: 'selection' } }); });
  act(() => {
    fireEvent.compositionStart(input);
    fireEvent(input, Object.assign(new Event('compositionupdate', { bubbles: true }), { data: 'preedit' }));
  });
  await waitFor(() => expect(commands[commands.length - 1]?.kind).toBe('compositionStart'));
  await waitFor(() => expect(client.state.throughSeq).toBe(commands[commands.length - 1]!.seq));
  await waitFor(() => expect(painted?.composition?.text).toBe('preedit'));
  const count = commands.length;
  act(() => { fireEvent.keyDown(input, { key: 'z', metaKey: true }); });
  await waitFor(() => expect(commands.filter((command) => command.kind === 'undo')).toHaveLength(4));
  expect(commands.slice(count).map((command) => command.kind)).toEqual(['compositionEnd', 'undo']);
  await waitFor(() => expect(painted?.composition).toBeNull());
  act(() => {
    input.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertFromComposition',
      data: 'preedit', bubbles: true, cancelable: true }));
    fireEvent(input, Object.assign(new Event('compositionend', { bubbles: true }), { data: 'preedit' }));
    input.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertText',
      data: 'preedit', bubbles: true, cancelable: true }));
    fireEvent.keyDown(input, { key: 'z', ctrlKey: true, repeat: true });
  });
  expect(commands.slice(count).some((command) => command.kind === 'compositionCommit')).toBe(false);
  expect(commands.slice(count).map((command) => command.kind)).toEqual(['compositionEnd', 'undo']);
});

test('preedit_survives_ordinary_acknowledgments', async () => {
  const ready = mock(() => {});
  const view = render(<PptxEditor fonts={[]} file={await deck()} {...{ [PEERLESS_FLAG]: true as const }} onReady={ready} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const content = (await client.read()).result;
  if (!content.ok) throw new Error(content.failure.message);
  const story = content.stories[0]!;
  const selection = { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
    anchor: 0, focus: 0, affinity: 'downstream' as const };
  act(() => { void client.dispatch({ kind: 'insertText', payload: {
    at: { source: 'confirmed', selection }, text: 'Z', inputTime: 1, undoGroup: 'typing' } }).catch(() => {}); });
  await waitFor(() => expect(client.state.throughSeq).toBe(1));
  act(() => { void client.dispatch({ kind: 'undo', payload: { inputTime: 2, undoGroup: 'history' } }).catch(() => {}); });
  await waitFor(() => expect(client.state.throughSeq).toBe(2));
  act(() => { void client.dispatch({ kind: 'setSelection', payload: {
    selection, inputTime: 3, undoGroup: 'selection' } }).catch(() => {}); });
  await waitFor(() => expect(client.state.throughSeq).toBe(3));
  const input = view.getByRole('textbox', { name: 'Slide text' });
  act(() => {
    fireEvent.compositionStart(input);
    fireEvent(input, Object.assign(new Event('compositionupdate', { bubbles: true }), { data: 'alpha' }));
  });
  await waitFor(() => expect(client.state.throughSeq).toBe(4));
  expect(commands[commands.length - 1]?.kind).toBe('compositionStart');
  await waitFor(() => expect(painted?.composition?.text).toBe('alpha'));
  act(() => { void client.dispatch({ kind: 'setReadOnly', payload: { readOnly: false, inputTime: 3, undoGroup: 'mode' } }).catch(() => {}); });
  await waitFor(() => expect(client.state.throughSeq).toBe(5));
  await waitFor(() => expect(painted?.composition?.text).toBe('alpha'));
  act(() => { fireEvent(input, Object.assign(new Event('compositionend', { bubbles: true }), { data: 'alpha' })); });
  await waitFor(() => expect(commands[commands.length - 1]?.kind).toBe('compositionEnd'));
  await waitFor(() => expect(client.state.throughSeq).toBe(commands[commands.length - 1]!.seq));
  expect(commands.some((command) => command.kind === 'compositionCommit')).toBe(true);
  const committed = (await client.read()).result;
  expect(committed.ok && committed.stories[0]?.text).toBe('alphaalpha');
});

test('history_to_empty_deck_notifies_change', async () => {
  const ready = mock(() => {});
  const changed = mock(() => {});
  const view = render(<PptxEditor fonts={[]} file={await deck()} {...{ [PEERLESS_FLAG]: true as const }}
    onReady={ready} onChange={changed} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  handle.addUndoBoundary();
  handle.deleteSlide(client.state.slides[0]!.id);
  handle.addUndoBoundary();
  act(() => { void client.dispatch({ kind: 'undo', payload: { inputTime: 1, undoGroup: 'history' } }).catch(() => {}); });
  await waitFor(() => expect(changed).toHaveBeenCalledTimes(1));
  changed.mockClear();
  act(() => { fireEvent.click(view.getByTestId('pptx-redo')); });
  await waitFor(() => expect(client.state.slides).toHaveLength(0));
  await waitFor(() => expect(changed).toHaveBeenCalledTimes(1));
  expect(changed).toHaveBeenCalledWith(expect.objectContaining({ slides: [] }));
  expect(view.getByRole('status').textContent).toContain('No slides');
});
