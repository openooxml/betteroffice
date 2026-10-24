import './peerlessTestCapabilities';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type {
  PptxWorkerSession, PresentationCommandAck, PresentationCommandClient, PresentationCommandEnvelope,
  PresentationCommandEvents, PresentationCommandMethods,
} from '@betteroffice/pptx/worker';
import { createSessionHost } from '../../../../shared/office-session/host';
import { isHostMessage } from '../../../../shared/office-session/protocol';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { createSharedPresentationCommandClient } from '../../../pptx/src/session/commandClient';
import { registerPresentationCommandHandlers } from '../../../pptx/src/session/commandHost';
import * as wasm from '../../../pptx/src/wasm/loader';
import { PptxEditor } from '../PptxEditor';
import { peerlessSessionOpener } from './usePeerlessPresentation';
import * as evaluation from '../commands/evaluate';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
let owner: PptxWorkerSession;
let client: PresentationCommandClient;
let hold = false;
let replies: Map<number, () => void>;
let commands: PresentationCommandEnvelope[];

async function deck(): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  const ids = [1];
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${ids.map((id) => `<Override PartName="/ppt/slides/slide${id}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst>${ids.map((id) => `<p:sldId id="${255 + id}" r:id="rId${id}"/>`).join('')}</p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${ids.map((id) => `<Relationship Id="rId${id}" Type="${rel}/slide" Target="slides/slide${id}.xml"/>`).join('')}</Relationships>`);
  for (const id of ids) zip.file(`ppt/slides/slide${id}.xml`, `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>slide ${id}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}

beforeEach(() => {
  hold = false;
  replies = new Map(); commands = [];
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
  const originalBounds = HTMLElement.prototype.getBoundingClientRect;
  const bounds = spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
    return this.getAttribute('role') === 'toolbar' ? new DOMRect(0, 0, 1400, 36) : originalBounds.call(this);
  });
  restorers.push(() => bounds.mockRestore());
  for (const [property, extent] of [['clientWidth', 1000], ['clientHeight', 580]] as const) {
    const original = Object.getOwnPropertyDescriptor(HTMLElement.prototype, property);
    Object.defineProperty(HTMLElement.prototype, property, { configurable: true,
      get(this: HTMLElement) { return this.getAttribute('role') === 'application' ? extent : property === 'clientWidth' ? 1400 : 0; } });
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
      .then((font) => createSharedPresentationCommandClient(bytes, { session: 's3', incarnation: 's3-worker',
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
  try { cleanup(); await owner?.dispose(); }
  finally { for (const restore of restorers.splice(0).reverse()) restore(); }
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

async function surface(onSave = mock((_bytes: Uint8Array) => {})) {
  const ready = mock(() => {});
  const view = render(<PptxEditor fonts={[]} file={await deck()} onReady={ready} onSave={onSave} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  return view;
}

async function selectText() {
  const result = (await client.read()).result;
  if (!result.ok) throw new Error(result.failure.message);
  const story = result.stories[0]!;
  await act(async () => { await client.dispatch({ kind: 'setSelection', payload: {
    selection: { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
      anchor: 1, focus: 1, affinity: 'downstream' }, inputTime: 1, undoGroup: 's3' } }); });
  commands.length = 0;
  return story;
}

function textSelection(): import('@betteroffice/pptx/worker').PresentationCommandSelection {
  const at = client.state.selection;
  if (!at || at.kind === 'object') throw new Error('Expected text selection');
  return at;
}

test('hidden_input_shortcuts', async () => {
  const saved = mock((_bytes: Uint8Array) => {});
  const view = await surface(saved);
  await selectText();
  const input = view.getByRole('textbox', { name: 'Slide text' });
  for (const [key, patch] of [['b', { bold: true }], ['i', { italic: true }],
    ['u', { underline: 'sng' }]] as const) {
    const count = commands.length;
    act(() => { input.focus(); fireEvent.keyDown(input, { key, ctrlKey: true }); });
    await waitFor(() => expect(commands).toHaveLength(count + 1));
    expect(commands[commands.length - 1]).toMatchObject({ kind: 'formatText', payload: { patch } });
    await waitFor(() => expect(client.state.throughSeq).toBe(commands[commands.length - 1]!.seq));
    act(() => { fireEvent.keyDown(input, { key, ctrlKey: true, repeat: true }); });
    expect(commands).toHaveLength(count + 1);
  }
  act(() => { fireEvent.keyDown(input, { key: 's', metaKey: true }); });
  await waitFor(() => expect(saved).toHaveBeenCalledTimes(1));
  act(() => { fireEvent.keyDown(input, { key: 's', metaKey: true, repeat: true }); });
  expect(commands.filter((command) => command.kind === ('formatText' as string))).toHaveLength(3);
  expect(saved).toHaveBeenCalledTimes(1);
});

test('pending_target', async () => {
  const view = await surface();
  const story = await selectText();
  hold = true;
  const input = view.getByRole('textbox', { name: 'Slide text' });
  act(() => {
    input.focus();
    fireEvent.keyDown(input, { key: 'ArrowRight' });
    fireEvent.click(view.getByTestId('pptx-bold'));
    fireEvent.keyDown(input, { key: 'ArrowRight' });
  });
  await waitFor(() => expect(commands.length).toBeGreaterThanOrEqual(2));
  const predecessor = commands[0]!;
  expect(predecessor).toMatchObject({ kind: 'setSelection', payload: { selection: { anchor: 2, focus: 2 } } });
  expect(commands[1]).toMatchObject({ kind: 'formatText', afterSeq: predecessor.seq,
    payload: { at: { source: 'afterSeq', seq: predecessor.seq }, patch: { bold: true } } });
  hold = false;
  await act(async () => { for (const deliver of [...replies.values()]) deliver(); });
  await waitFor(() => expect(client.state.selection).toMatchObject({ storyId: story.storyId, anchor: 3, focus: 3 }));
  expect(client.state.editState.toolbar.storedMarks).toEqual({});
});

test('toolbar_font_size_commit_restores_input_focus_and_marks', async () => {
  const view = await surface();
  const story = await selectText();
  const input = view.getByRole('textbox', { name: 'Slide text' });
  const size = view.getByRole('combobox', { name: 'Font size' });
  for (const points of [36, 48]) {
    act(() => { size.focus(); });
    if (points === 36) act(() => { fireEvent.click(view.getByRole('option', { name: '36' })); });
    else {
      act(() => { fireEvent.change(size, { target: { value: String(points) } }); });
      act(() => { fireEvent.keyDown(size, { key: 'Enter' }); });
    }
    await waitFor(() => expect(client.state.editState.toolbar.storedMarks).toEqual({ fontSizePt: points }));
    await waitFor(() => expect(document.activeElement).toBe(input));
  }
  expect(client.state.selection).toMatchObject({ storyId: story.storyId, anchor: 1, focus: 1 });
  expect(commands.every((command) => command.kind === 'formatText')).toBe(true);
  act(() => {
    document.activeElement!.dispatchEvent(new InputEvent('beforeinput', {
      bubbles: true, cancelable: true, inputType: 'insertText', data: 'Z' }));
  });
  await waitFor(() => expect(client.state.selection?.focus).toBe(2));
  const result = (await client.read()).result;
  expect(result.ok && result.stories[0]?.text).toBe('sZlide 1');
  expect(client.state.editState.toolbar.formatting?.fontSizePt).toBe(48);
});

test('toolbar_font_size_step_keeps_pending_target_intent', async () => {
  const view = await surface();
  const story = await selectText();
  await act(async () => {
    await client.dispatch({ kind: 'formatText', payload: { inputTime: 1, undoGroup: 's3',
      at: { source: 'confirmed', selection: { ...textSelection(), anchor: 2, focus: 7 } },
      patch: { fontSizePt: 36 } } });
    await client.dispatch({ kind: 'formatText', payload: { inputTime: 1, undoGroup: 's3',
      at: { source: 'confirmed', selection: { ...textSelection(), anchor: 0, focus: 2 } },
      patch: { fontSizePt: 10 } } });
  });
  await selectText();
  hold = true;
  const input = view.getByRole('textbox', { name: 'Slide text' });
  act(() => {
    input.focus();
    fireEvent.keyDown(input, { key: 'ArrowRight' });
    fireEvent.click(view.getByRole('button', { name: 'Increase font size' }));
  });
  await waitFor(() => expect(commands).toHaveLength(2));
  expect(commands[1]).toMatchObject({ kind: 'formatText', afterSeq: commands[0]!.seq,
    payload: { at: { source: 'afterSeq', seq: commands[0]!.seq }, patch: {},
      fontSizeStep: { direction: 'increase', sizes: [...evaluation.DEFAULT_FONT_SIZES] } } });
  hold = false;
  await act(async () => { for (const deliver of [...replies.values()]) deliver(); });
  await waitFor(() => expect(client.state.editState.toolbar.storedMarks).toEqual({ fontSizePt: 48 }));
  expect(client.state.selection).toMatchObject({ storyId: story.storyId, anchor: 2, focus: 2 });
});

test('toolbar_projection_ignores_focus_and_unchanged_formatting_frames', async () => {
  const view = await surface();
  await selectText();
  await act(async () => {
    await client.dispatch({ kind: 'formatText', payload: { inputTime: 1, undoGroup: 's3',
      at: { source: 'confirmed', selection: { ...textSelection(), anchor: 0, focus: 7 } },
      patch: { color: '#111827' } } });
  });
  await selectText();
  const evaluated = spyOn(evaluation, 'evaluatePptxCommand');
  try {
    const input = view.getByRole('textbox', { name: 'Slide text' });
    act(() => { input.focus(); });
    act(() => { input.blur(); });
    await act(async () => {
      await client.dispatch({ kind: 'setSelection', payload: { inputTime: 1, undoGroup: 's3',
        selection: { ...textSelection(), anchor: 2, focus: 2 } } });
    });
    expect(evaluated).not.toHaveBeenCalled();
    hold = true;
    act(() => {
      input.dispatchEvent(new InputEvent('beforeinput', {
        bubbles: true, cancelable: true, inputType: 'insertText', data: 'Z' }));
    });
    await waitFor(() => expect(replies.size).toBe(1));
    expect(evaluated).not.toHaveBeenCalled();
    hold = false;
    await act(async () => { for (const deliver of [...replies.values()]) deliver(); });
    expect(evaluated).not.toHaveBeenCalled();
    await act(async () => {
      await client.dispatch({ kind: 'formatText', payload: { inputTime: 1, undoGroup: 's3',
        at: { source: 'confirmed', selection: textSelection() }, patch: { bold: true } } });
    });
    expect(evaluated).toHaveBeenCalled();
    expect(view.getByTestId('pptx-bold').getAttribute('aria-pressed')).toBe('true');
  } finally {
    hold = false;
    for (const deliver of [...replies.values()]) deliver();
    evaluated.mockRestore();
  }
});
