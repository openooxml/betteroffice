import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, jest, mock, spyOn, test } from 'bun:test';
import type {
  DocxCommandAck, DocxCommandEnvelope, DocxCommandFrame, ResidentCommandConnection,
} from '@betteroffice/docx/peerless';
import {
  applyFrameDeltaOwned, buildMirrorPage, displayPageRevision, FRAME_DELTA_VERSION,
  type DecodedFrameDelta, type DisplayPage, type FramePageOperation,
} from '@betteroffice/docx/layout/render';
import type { PeerlessEditSessionBinding, UsePeerlessEditSessionOptions } from '../hooks/usePeerlessEditSession';
import * as composition from './composition';
import * as display from '@betteroffice/docx/layout/render';
import { installCompositionEventPolyfill } from './compositionEventPolyfill.testing';
import { PeerlessPagedEditor, type PeerlessPagedEditorProps } from './PeerlessPagedEditor';
import { glyphChar, installGlyphPaths, testEchoGlyphLine, testEchoGlyphSet } from './echoGlyphs.testing';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const restoreCompositionEvent = installCompositionEventPolyfill();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
afterEach(() => {
  cleanup();
  for (const restore of restorers.splice(0).reverse()) restore();
});
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  restoreCompositionEvent();
  if (ownsDom) await GlobalRegistrator.unregister();
});

const identity = { session: 'integration', incarnation: 1, epoch: 1 };
const page = { pageId: '1', pageIndex: 0 };

function frame(revision: number, coverage: number, text = ''): DocxCommandFrame {
  const bytes = new ArrayBuffer(88);
  const data = new DataView(bytes);
  new Uint8Array(bytes).set([0x46, 0x44, 0x56, 0x31]);
  data.setUint16(4, 1, true);
  data.setUint16(6, 80, true);
  data.setUint32(8, 88, true);
  data.setUint32(12, 1, true);
  data.setBigUint64(16, 1n, true);
  data.setBigUint64(24, 1n, true);
  data.setBigUint64(32, BigInt(revision), true);
  for (const offset of [56, 60]) data.setUint32(offset, 80, true);
  data.setUint32(64, 4, true);
  data.setUint32(68, 88, true);
  const rect = { x: 0, y: 0, width: 200, height: 40 };
  return { revision, coversThroughSeq: coverage, payload: {
    frame: bytes, layoutRevision: revision, caret: { frameEpoch: revision, caretRect: null },
    stagedPresentation: { frameEpoch: revision, pages: [{ page, frameEpoch: revision,
      bitmap: { width: 200, height: 40, close: mock(() => {}) } as unknown as ImageBitmap,
      backgrounds: [{ id: 'background', rect,
        bitmap: { width: 200, height: 40, close: mock(() => {}) } as unknown as ImageBitmap }],
    }] },
    echoContext: { page, frameEpoch: revision, layoutRevision: revision,
      serializedBytes: 900, text, shaping: 'simple-ltr', transform: [1, 0, 0, 1, 0, 0],
      styleRuns: [{ startUtf16: 0, endUtf16: text.length, fontFamily: 'Arial', fontSizePx: 12,
        color: '#000', bold: false, italic: false, underline: false, strike: false,
        transform: [1, 0, 0, 1, 0, 0], baseline: { x: 0, y: 14 } }],
      boundaries: Array.from({ length: text.length + 1 }, (_, utf16) => ({ utf16, grapheme: utf16,
        caret: { x: utf16 * 10, y: 0, width: 1, height: 40 } })),
      selection: { anchorGrapheme: text.length, focusGrapheme: text.length },
      damage: [{ page, rect, backgroundId: 'background', omittedPrimitiveIds: [] }],
      glyphs: testEchoGlyphLine({ device: { x: 0, y: 0, width: 200, height: 40 } }),
    },
    echoGlyphSets: [testEchoGlyphSet],
  } };
}

async function mount(selection = false, accessible = false, table = false,
  prepare?: UsePeerlessEditSessionOptions['prepareGeometry'], caretX: (revision: number) => number = () => 10,
  props: Partial<PeerlessPagedEditorProps> = {}, multiCell = false) {
  const events: string[] = [];
  restorers.push(installGlyphPaths());
  const canvas = document.createElement('canvas');
  const ctx = {
    save() {}, restore() {}, setTransform() {}, clearRect() {}, putImageData() {},
    getImageData: () => ({ width: canvas.width, height: canvas.height }) as ImageData,
    transform() {}, beginPath() {}, rect() {}, clip() {}, fillRect() {},
    translate() {}, scale() {}, fill: (path: unknown) => { events.push(`echo:${glyphChar(path)}`); },
    drawImage: () => { events.push('pixels'); },
  };
  Object.defineProperty(canvas, 'getContext', { value: () => ctx });
  const position = { story: 'body', paraId: 'worker', offset: 0 };
  const initial: ResidentCommandConnection['initial'] = {
    ...identity, version: 0, throughSeq: 0, outcomes: [], selection: { anchor: position, focus: position },
    editState: { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
      canUndo: false, canRedo: false, hyperlink: null, tableContext: table ? {
        story: 'body', tableIndex: 0, rowIndex: 0, columnIndex: 0, rowCount: 1, columnCount: 2,
        hasMultiCellSelection: multiCell, canSplitCell: false,
      } : null },
    frame: frame(1, 0),
  };
  const commands: DocxCommandEnvelope[] = [];
  const replies: Array<(ack: DocxCommandAck) => void> = [];
  let receive: ((candidate: DocxCommandFrame) => void) | undefined;
  const connection: ResidentCommandConnection = {
    initial, requestFullFrame: async () => initial,
    client: {
      failure: undefined, onFailure: () => () => {}, dispose: async () => {},
      onFrame(listener) { receive = listener; return () => { receive = undefined; }; },
      call: {
        checkpoint: async () => ({ state: new Uint8Array(), epoch: 1, version: 0, throughSeq: 0 }),
        edit(command) {
          commands.push(command);
          return new Promise<DocxCommandAck>((resolve) => { replies.push(resolve); });
        },
      },
    },
  };
  const sessionOptions: UsePeerlessEditSessionOptions = {
    source: Uint8Array.of(1), openConnection: async () => connection, canvasForPage: () => canvas,
    prepareGeometry: prepare ?? function (candidate) {
      return { identity, coversThroughSeq: candidate.coversThroughSeq, frameEpoch: candidate.revision,
        layoutRevision: candidate.payload.layoutRevision,
        pages: accessible ? Array.from({ length: 20 }, (_, pageIndex) => ({
          page: { pageId: String(pageIndex + 1), pageIndex },
          bounds: { x: 0, y: pageIndex * 64, width: 200, height: 40 }, transform: [1, 0, 0, 1, 0, 0] as const,
          displayPage: { pageIndex, width: 200, height: 40, primitives: [{
            kind: 'text' as const, text: `Readable page ${pageIndex + 1}`, x: 0, baselineY: 14, width: 100, font: '12px Arial', color: '#000',
          }] },
        })) : [{ page, bounds: { x: 0, y: 0, width: 200, height: 40 }, transform: [1, 0, 0, 1, 0, 0] }],
        caret: { x: caretX(candidate.revision), y: 0, width: 1, height: 20 },
        selectionRects: selection ? [{ x: 0, y: 0, width: 10, height: 20 }] : [],
      };
    },
  };
  let binding: PeerlessEditSessionBinding | undefined;
  const view = render(<PeerlessPagedEditor {...props} sessionOptions={sessionOptions} onReady={(ready) => { binding = ready; }} />);
  await waitFor(() => { expect(binding).toBeDefined(); });
  return { ...view, binding: binding!, commands, canvas, events,
    textarea: view.getByRole('textbox', { name: 'Document input' }) as HTMLTextAreaElement,
    async acknowledge(index: number, status: 'applied' | 'refused' | 'noop', extra: Partial<DocxCommandAck> = {}) {
      await waitFor(() => { expect(commands[index]).toBeDefined(); });
      const command = commands[index]!;
      await act(async () => { replies[index]!({ ...initial, frame: undefined, version: index + 1,
        throughSeq: command.seq, outcomes: [{ seq: command.seq, status,
          ...(status === 'refused' ? { reason: 'locked-target' } : {}) }],
        ...extra,
      }); });
    },
    present(candidate: DocxCommandFrame) { act(() => { receive?.(candidate); }); },
  };
}

test('peerless accessibility exposes retained offscreen pages without allocating their canvases', async () => {
  const editor = await mount(false, true);
  await waitFor(() => expect(editor.getByText('Readable page 13')).toBeTruthy());
  expect(editor.container.querySelector('[data-page-id="13"] canvas')).toBeNull();
  expect(editor.getByRole('status').getAttribute('aria-live')).toBe('polite');
});

test('a page whose positions changed in place rebuilds its accessibility mirror', async () => {
  const displayPage: DisplayPage = { pageIndex: 0, width: 200, height: 40, positionSpan: [1, 9], primitives: [{
    kind: 'text', text: 'Shifted page', x: 0, baselineY: 14, width: 100, font: '12px Arial', color: '#000', docStart: 1, docEnd: 9,
  }] };
  const delta = (frameEpoch: number, operations: FramePageOperation[]): DecodedFrameDelta => ({
    protocolVersion: FRAME_DELTA_VERSION, full: frameEpoch === 1, docEpoch: frameEpoch, layoutEpoch: frameEpoch, frameEpoch,
    baseFrameEpoch: frameEpoch - 1, pageCount: 1, bytes: new Uint8Array(), operations });
  let retained = applyFrameDeltaOwned(null, delta(1, [{ kind: 'upsert', pageId: 1n, pageIndex: 0, fingerprint: 1n,
    primitiveIds: new BigUint64Array([1n]), page: displayPage }]));
  const editor = await mount(false, false, false, (candidate) => {
    if (candidate.revision > retained.frameEpoch) {
      retained = applyFrameDeltaOwned(retained, delta(candidate.revision,
        [{ kind: 'shift-range', pageIndex: 0, pageId: 1n, count: 1, delta: 2, salt: 3n }]));
    }
    return { identity, coversThroughSeq: candidate.coversThroughSeq, frameEpoch: candidate.revision,
      layoutRevision: candidate.payload.layoutRevision, caret: null, selectionRects: [],
      pages: [{ page, bounds: { x: 0, y: 0, width: 200, height: 40 }, transform: [1, 0, 0, 1, 0, 0],
        displayPage: retained.pages[0]!.page }] };
  });
  const stamped = (root: Element | null | undefined) => [...root?.querySelectorAll<HTMLElement>('[data-doc-start]') ?? []]
    .map((el) => [el.dataset.docStart, el.dataset.docEnd]);
  const mirror = () => editor.container.querySelector('[data-page-id="1"] .canvas-page-mirror');
  await waitFor(() => expect(stamped(mirror())).toContainEqual(['1', '9']));
  editor.present(frame(2, 0));
  expect(retained.pages[0]!.page).toBe(displayPage);
  expect(displayPageRevision(displayPage)).toBe(1);
  const fresh = stamped(buildMirrorPage(displayPage));
  expect(fresh).toContainEqual(['3', '11']);
  expect(fresh).not.toContainEqual(['1', '9']);
  await waitFor(() => expect(stamped(mirror())).toEqual(fresh), { timeout: 5_000 });
}, 15_000);

function preedit(textarea: HTMLTextAreaElement, type: 'compositionstart' | 'compositionupdate' | 'compositionend', data = '') {
  const event = type === 'compositionstart' ? fireEvent.compositionStart
    : type === 'compositionupdate' ? fireEvent.compositionUpdate : fireEvent.compositionEnd;
  event(textarea, { data });
}

function beforeinput(textarea: HTMLTextAreaElement, inputType: string, data: string) {
  act(() => { textarea.dispatchEvent(new InputEvent('beforeinput', {
    bubbles: true, cancelable: true, inputType, data,
  })); });
}

test('table_typing_queues_behind_the_worker_inflight_window', async () => {
  const editor = await mount(false, false, true);
  beforeinput(editor.textarea, 'insertText', 'X');
  fireEvent.keyDown(editor.textarea, { key: 'Tab', code: 'Tab' });
  fireEvent.keyUp(editor.textarea, { key: 'Tab', code: 'Tab' });
  beforeinput(editor.textarea, 'insertText', 'Y');
  fireEvent.keyDown(editor.textarea, { key: 'Shift', code: 'ShiftLeft', shiftKey: true });
  fireEvent.keyDown(editor.textarea, { key: 'Tab', code: 'Tab', shiftKey: true });
  fireEvent.keyUp(editor.textarea, { key: 'Tab', code: 'Tab', shiftKey: true });
  fireEvent.keyUp(editor.textarea, { key: 'Shift', code: 'ShiftLeft' });
  beforeinput(editor.textarea, 'insertText', 'Z');
  await waitFor(() => expect(editor.commands).toHaveLength(4));
  expect(editor.binding.admission.getSnapshot().originalUnackedIntents).toBe(5);
  expect(editor.commands.map(({ kind }) => kind)).toEqual(['insertText', 'navigateCell', 'insertText', 'navigateCell']);
  await editor.acknowledge(0, 'applied');
  await waitFor(() => expect(editor.commands).toHaveLength(5));
  expect(editor.commands[4]!.payload).toEqual({ text: 'Z' });
  for (let index = 1; index < editor.commands.length; index += 1) {
    expect(editor.commands[index]!.afterSeq).toBe(editor.commands[index - 1]!.seq);
    await editor.acknowledge(index, 'applied');
  }
  await editor.binding.session.flush();
  expect(editor.binding.admission.getSnapshot().originalUnackedIntents).toBe(0);
});

test('refused_ime_retains_recoverable_draft', async () => {
  const editor = await mount();
  preedit(editor.textarea, 'compositionstart');
  preedit(editor.textarea, 'compositionupdate', '日本語');
  fireEvent.input(editor.textarea, { inputType: 'insertCompositionText', isComposing: true,
    target: { value: '日本語' } });
  expect(editor.commands).toEqual([]);
  preedit(editor.textarea, 'compositionend', '日本語');
  await editor.acknowledge(0, 'refused');
  expect(editor.commands.map((command) => command.kind)).toEqual(['replaceSelection']);
  expect(editor.binding.admission.getSnapshot().compositionSlot).toBeNull();
  const recovered = editor.queryByRole('textbox', { name: 'Recoverable document text' }) as HTMLTextAreaElement | null;
  expect(recovered?.value).toBe('日本語');
  expect(recovered?.readOnly).toBe(true);
  expect(editor.textarea.value).toBe('');
});

test('synthetic_composition_payload_commits_and_recovers_through_paged_editor', async () => {
  const editor = await mount();
  fireEvent.compositionStart(editor.textarea);
  fireEvent.compositionUpdate(editor.textarea, { data: 'あ' });
  expect(editor.textarea.value).toBe('');
  fireEvent.compositionEnd(editor.textarea, { data: 'あ' });
  beforeinput(editor.textarea, 'insertFromComposition', 'あ');
  fireEvent.input(editor.textarea, { inputType: 'insertFromComposition', target: { value: 'あ' } });
  await waitFor(() => { expect(editor.commands).toHaveLength(1); });
  expect(editor.commands[0]!.kind).toBe('replaceSelection');
  expect(editor.commands[0]!.payload).toEqual({ text: 'あ', selection: {
    anchor: { story: 'body', paraId: 'worker', offset: 0 },
    focus: { story: 'body', paraId: 'worker', offset: 0 },
  } });
  await editor.acknowledge(0, 'refused');
  expect((editor.getByRole('textbox', { name: 'Recoverable document text' }) as HTMLTextAreaElement).value).toBe('あ');
  beforeinput(editor.textarea, 'insertText', 'あ');
  fireEvent.input(editor.textarea, { inputType: 'insertText', target: { value: 'あ' } });
  await waitFor(() => { expect(editor.commands).toHaveLength(2); });
  expect(editor.commands[1]!.kind).toBe('insertText');
});

test('delayed_composition_refusal_retains_draft_without_exposing_active_preedit_overlays', async () => {
  const editor = await mount(true);
  preedit(editor.textarea, 'compositionstart');
  preedit(editor.textarea, 'compositionupdate', '先');
  preedit(editor.textarea, 'compositionend', '先');
  await waitFor(() => { expect(editor.commands).toHaveLength(1); });
  preedit(editor.textarea, 'compositionstart');
  preedit(editor.textarea, 'compositionupdate', '未確定');
  fireEvent.input(editor.textarea, { inputType: 'insertCompositionText', isComposing: true,
    target: { value: '未確定' } });
  expect(editor.container.querySelectorAll('[aria-hidden="true"]')).toHaveLength(0);
  await editor.acknowledge(0, 'refused');
  expect((editor.getByRole('textbox', { name: 'Recoverable document text' }) as HTMLTextAreaElement).value).toBe('先');
  editor.present(frame(2, editor.commands[0]!.seq));
  expect(editor.binding.viewStore.getSnapshot().geometry?.frameEpoch).toBe(2);
  expect(editor.binding.admission.getSnapshot().compositionSlot).not.toBeNull();
  expect(editor.textarea.value).toBe('未確定');
  expect(editor.commands).toHaveLength(1);
  expect(editor.container.querySelectorAll('[aria-hidden="true"]')).toHaveLength(0);
});

test('recoverable_drafts_remain_until_individually_dismissed', async () => {
  const editor = await mount();
  for (const [index, text] of ['先', '後'].entries()) {
    preedit(editor.textarea, 'compositionstart');
    preedit(editor.textarea, 'compositionupdate', text);
    preedit(editor.textarea, 'compositionend', text);
    await editor.acknowledge(index, 'refused');
  }
  const drafts = () => editor.queryAllByRole('textbox', { name: 'Recoverable document text' })
    .map((draft) => (draft as HTMLTextAreaElement).value);
  expect(drafts()).toEqual(['先', '後']);
  expect(editor.queryAllByRole('button', { name: 'Dismiss draft' })).toHaveLength(2);
  fireEvent.click(editor.getAllByRole('button', { name: 'Dismiss draft' })[0]!);
  expect(drafts()).toEqual(['後']);
  preedit(editor.textarea, 'compositionstart');
  preedit(editor.textarea, 'compositionupdate', '新');
  preedit(editor.textarea, 'compositionend', '新');
  await editor.acknowledge(2, 'refused');
  expect(drafts()).toEqual(['後', '新']);
  fireEvent.click(editor.getAllByRole('button', { name: 'Dismiss draft' })[1]!);
  expect(drafts()).toEqual(['後']);
  fireEvent.click(editor.getByRole('button', { name: 'Dismiss draft' }));
  expect(drafts()).toEqual([]);
  expect(editor.commands).toHaveLength(3);
});

test('composition_commits_as_one_replace_selection_through_paged_editor', async () => {
  const begin = spyOn(composition, 'beginComposition');
  const update = spyOn(composition, 'updateComposition');
  const end = spyOn(composition, 'endComposition');
  restorers.push(() => begin.mockRestore(), () => update.mockRestore(), () => end.mockRestore());
  const editor = await mount();
  beforeinput(editor.textarea, 'insertText', 'x');
  preedit(editor.textarea, 'compositionstart');
  preedit(editor.textarea, 'compositionupdate', 'に');
  preedit(editor.textarea, 'compositionupdate', '日本語');
  fireEvent.input(editor.textarea, { inputType: 'insertCompositionText', isComposing: true,
    target: { value: '日本語' } });
  await waitFor(() => { expect(editor.commands).toHaveLength(1); });
  preedit(editor.textarea, 'compositionend', '日本語');
  beforeinput(editor.textarea, 'insertFromComposition', '日本語');
  fireEvent.input(editor.textarea, { inputType: 'insertFromComposition', target: { value: '日本語' } });
  fireEvent.blur(editor.textarea);
  await waitFor(() => { expect(editor.commands).toHaveLength(2); });
  expect(editor.commands.map((command) => command.kind)).toEqual(['insertText', 'replaceSelection']);
  expect(editor.commands[1]!.payload).toEqual({ text: '日本語' });
  expect(editor.commands[1]!.afterSeq).toBe(editor.commands[0]!.seq);
  expect(begin).toHaveBeenCalledTimes(1);
  expect(update).toHaveBeenCalled();
  expect(end).toHaveBeenCalledTimes(1);
  await editor.acknowledge(0, 'applied');
  await editor.acknowledge(1, 'applied');
  expect(editor.queryByRole('textbox', { name: 'Recoverable document text' })).toBeNull();
});

test('worker_frame_installs_through_frame_presentation_and_retires_echo', async () => {
  const editor = await mount();
  const reconcile = spyOn(editor.binding.echoPainter, 'reconcile');
  restorers.push(() => reconcile.mockRestore());
  beforeinput(editor.textarea, 'insertText', 'x');
  const ticket = editor.binding.admission.getSnapshot().tail!;
  expect(editor.events).toContain('echo:x');
  await editor.acknowledge(0, 'applied');
  expect(reconcile).not.toHaveBeenCalled();
  expect(editor.binding.viewStore.getSnapshot().presentedThroughSeq).toBe(0);
  preedit(editor.textarea, 'compositionstart');
  preedit(editor.textarea, 'compositionupdate', '未確定');
  fireEvent.input(editor.textarea, { inputType: 'insertCompositionText', isComposing: true,
    target: { value: '未確定' } });
  const candidate = frame(2, ticket.seq!, 'x');
  editor.events.length = 0;
  editor.present(candidate);
  expect(editor.events[0]).toBe('pixels');
  expect(reconcile).toHaveBeenCalledTimes(1);
  expect(reconcile.mock.calls[0]![0]).toBe(candidate);
  expect(reconcile.mock.calls[0]![1].map((entry) => entry.ticket)).toEqual([ticket]);
  expect(reconcile.mock.calls[0]![2]).toEqual([]);
  expect(await ticket.presented).toBe('presented');
  expect(editor.binding.viewStore.getSnapshot().geometry?.frameEpoch).toBe(2);
  expect(editor.container.querySelector('[data-page-id="1"] canvas')).toBe(editor.canvas);
  expect(editor.textarea.value).toBe('未確定');
  expect(editor.commands).toHaveLength(1);
  expect(editor.container.querySelectorAll('[aria-hidden="true"]')).toHaveLength(0);
});

test('worker caret blinks on the hybrid interval only while the input is focused', async () => {
  const editor = await mount(false, false, false, undefined, (revision) => revision * 10);
  jest.useFakeTimers();
  restorers.push(() => jest.useRealTimers());
  const caret = () => {
    const overlays = editor.container.querySelectorAll<HTMLElement>('[aria-hidden="true"]');
    expect(overlays).toHaveLength(1);
    return overlays[0]!.style;
  };
  const tick = (ms: number) => act(() => { jest.advanceTimersByTime(ms); });
  expect(caret().opacity).toBe('0');
  act(() => { editor.textarea.focus(); });
  expect(caret().opacity).toBe('1');
  tick(529);
  expect(caret().opacity).toBe('1');
  tick(1);
  expect(caret().opacity).toBe('0');
  tick(530);
  expect(caret().opacity).toBe('1');
  tick(530);
  expect(caret().opacity).toBe('0');
  editor.present(frame(2, 0));
  expect(caret().left).toBe('20px');
  expect(caret().opacity).toBe('1');
  tick(529);
  expect(caret().opacity).toBe('1');
  tick(1);
  expect(caret().opacity).toBe('0');
  tick(530);
  expect(caret().opacity).toBe('1');
  act(() => { editor.textarea.blur(); });
  expect(caret().opacity).toBe('0');
  tick(530 * 3);
  expect(caret().opacity).toBe('0');
  act(() => { editor.textarea.focus(); });
  expect(caret().opacity).toBe('1');
});

function pointerHost(container: HTMLElement) {
  const host = container.querySelector<HTMLElement>('[data-peerless-pages-host]')!;
  const capture = mock(() => {});
  const release = mock(() => {});
  Object.defineProperties(host, {
    getBoundingClientRect: { value: () => ({ x: 0, y: 0, left: 0, top: 0, right: 200, bottom: 40, width: 200, height: 40 }) },
    setPointerCapture: { value: capture },
    hasPointerCapture: { value: () => true },
    releasePointerCapture: { value: release },
  });
  return { host, capture, release };
}

async function provisionalEditor(region: display.DisplayListRegionHit['region'] = 'body', loaded = true) {
  const native = {} as display.RustDisplayListQueryEngine;
  const hitTestRegions = mock(() => ({ region, pos: 12 }));
  const caretRect = mock(() => ({ pageIndex: 0, x: 47, y: 6, width: 0, height: 18 }));
  const dispose = mock(() => {});
  const getLoaded = spyOn(display, 'loadedRustDisplayListQueryEngine').mockReturnValue(loaded ? native : null);
  let resolveEngine!: (engine: display.RustDisplayListQueryEngine) => void;
  const loading = new Promise<display.RustDisplayListQueryEngine>((resolve) => { resolveEngine = resolve; });
  const load = spyOn(display, 'loadRustDisplayListQueryEngine').mockReturnValue(loading);
  const queries = spyOn(display, 'createDisplayListQueries').mockReturnValue({ hitTestRegions, caretRect, dispose } as unknown as display.DisplayListQueries);
  restorers.push(() => getLoaded.mockRestore(), () => load.mockRestore(), () => queries.mockRestore());
  const displayPage: DisplayPage = { pageIndex: 0, width: 200, height: 40, primitives: [] };
  const editor = await mount(false, false, false, candidate => ({
    identity, coversThroughSeq: candidate.coversThroughSeq, frameEpoch: candidate.revision,
    layoutRevision: candidate.payload.layoutRevision,
    pages: [{ page, bounds: { x: 30, y: 1200, width: 200, height: 40 },
      transform: [1, 0, 0, 1, 30, 1200], displayPage }],
    caret: { x: 95, y: 1208, width: 1, height: 21 }, selectionRects: [],
  }), undefined, { zoom: 2 });
  const { host } = pointerHost(editor.container);
  const point = { pointerId: 1, pointerType: 'mouse', isPrimary: true, button: 0, buttons: 1,
    detail: 1, clientX: 160, clientY: 2420 };
  const caret = () => host.querySelector<HTMLElement>('[aria-hidden="true"]');
  return { ...editor, host, point, caret, hitTestRegions, caretRect, queries, dispose, load, native, displayPage,
    async finishEngineLoad() {
      await act(async () => {
        getLoaded.mockReturnValue(native);
        resolveEngine(native);
        await loading;
      });
    },
  };
}

test('mount requests the query engine before input and the first click uses it after loading', async () => {
  const editor = await provisionalEditor('body', false);
  expect(editor.load).toHaveBeenCalledTimes(1);
  expect(editor.commands).toHaveLength(0);
  expect(editor.queries).not.toHaveBeenCalled();
  await editor.finishEngineLoad();
  fireEvent.pointerDown(editor.host, editor.point);
  expect(editor.caret()?.style.left).toBe('77px');
  expect(editor.caret()?.style.top).toBe('1206px');
  expect(editor.queries).toHaveBeenCalledWith({ pages: [editor.displayPage] }, editor.native);
  expect(editor.load).toHaveBeenCalledTimes(1);
});

test('a cross-paragraph click ack updates toolbar and formatting before its frame', async () => {
  const editor = await provisionalEditor();
  const before = editor.binding.viewStore.getSnapshot();
  const at = { story: 'body', paraId: 'another-paragraph', offset: 2 };
  const editState: NonNullable<DocxCommandAck['editState']> = {
    ...before.editState!,
    toolbar: {
      paraId: at.paraId, bold: true, italic: false, underline: false, strike: false, superscript: false, subscript: false,
      fontFamily: 'Arial', fontSize: 24, color: null, highlight: null, styleId: null, alignment: null,
      paragraphProperties: {}, hasSelection: false, isMultiParagraph: false, inTable: false,
      isSingleEmbed: false, embedKind: null, isImage: false, inInsertion: false, inDeletion: false,
    },
    effectiveFormatting: {
      bold: true, italic: false, underline: false, strike: false, superscript: false, subscript: false,
      fontFamily: 'Arial', fontSizePt: 12, color: null, highlight: null, styleId: null, alignment: null,
    },
  };
  fireEvent.pointerDown(editor.host, editor.point);
  await editor.acknowledge(0, 'noop', { version: 0, selection: { anchor: at, focus: at }, editState });
  const after = editor.binding.viewStore.getSnapshot();
  expect(after.workerSelection?.focus).toEqual(at);
  expect(after.editState).toEqual(editState);
  expect(after.version).toBe(before.version);
  expect(after.presentedThroughSeq).toBe(before.presentedThroughSeq);
  expect(after.geometry).toBe(before.geometry);
  expect(after.provisionalCaret).not.toBeNull();
});

for (const detail of [0, 1]) {
  test(`pointerdown detail ${detail} paints a provisional caret synchronously at the presented body hit rect`, async () => {
    const editor = await provisionalEditor();
    const before = editor.binding.viewStore.getSnapshot();
    fireEvent.pointerDown(editor.host, { ...editor.point, detail });
    expect(editor.caret()?.style.left).toBe('77px');
    expect(editor.caret()?.style.top).toBe('1206px');
    expect(editor.caret()?.style.height).toBe('18px');
    expect(editor.caret()?.style.opacity).toBe('1');
    expect(editor.binding.viewStore.getSnapshot().geometry).toBe(before.geometry);
    expect(editor.binding.viewStore.getSnapshot().workerSelection).toEqual(before.workerSelection);
    expect(editor.queries).toHaveBeenCalledWith({ pages: [editor.displayPage] }, editor.native);
    expect(editor.hitTestRegions).toHaveBeenCalledWith(0, 50, 10);
    expect(editor.caretRect).toHaveBeenCalledWith(12);
    expect(editor.dispose).not.toHaveBeenCalled();
    expect(editor.load).not.toHaveBeenCalled();
    await waitFor(() => expect(editor.commands).toHaveLength(1));
    expect(editor.binding.viewStore.getSnapshot().provisionalCaret?.throughSeq).toBe(editor.commands[0]!.seq);
    expect(editor.commands[0]).toMatchObject({ kind: 'selectPoint', payload: {
      page, x: 50, y: 10, frameEpoch: 1, layoutRevision: 1, clickCount: 1,
      modifiers: { shift: false, alt: false, ctrl: false, meta: false },
    } });
  });
}

test('repeated clicks on a presented page reuse its queries until unmount', async () => {
  const editor = await provisionalEditor();
  for (let click = 0; click < 8; click++) {
    fireEvent.pointerDown(editor.host, { ...editor.point, clientX: 160 + click * 6 });
    fireEvent.pointerUp(editor.host, { ...editor.point, clientX: 160 + click * 6, buttons: 0 });
    await editor.acknowledge(click, 'noop', { version: 0 });
    editor.present(frame(click + 2, editor.commands[click]!.seq));
    await waitFor(() => expect(editor.binding.viewStore.getSnapshot().geometry?.frameEpoch).toBe(click + 2));
  }
  expect(editor.caretRect).toHaveBeenCalledTimes(8);
  expect(editor.queries).toHaveBeenCalledTimes(1);
  expect(editor.dispose).not.toHaveBeenCalled();
  editor.unmount();
  expect(editor.dispose).toHaveBeenCalledTimes(1);
});

test('consecutive detail-zero presses exclude double and triple clicks from the provisional caret', async () => {
  const editor = await provisionalEditor();
  let timeStamp = 100;
  const press = () => {
    const down = new PointerEvent('pointerdown', { ...editor.point, detail: 0, bubbles: true, cancelable: true });
    Object.defineProperty(down, 'timeStamp', { value: timeStamp });
    timeStamp += 100;
    fireEvent(editor.host, down);
    fireEvent.pointerUp(editor.host, { ...editor.point, detail: 0 });
  };
  press();
  expect(editor.binding.viewStore.getSnapshot().provisionalCaret).not.toBeNull();
  await waitFor(() => expect(editor.commands).toHaveLength(1));
  await editor.acknowledge(0, 'applied');
  editor.present(frame(2, editor.commands[0]!.seq));
  await waitFor(() => expect(editor.binding.viewStore.getSnapshot().geometry?.frameEpoch).toBe(2));
  editor.caretRect.mockClear();
  press();
  expect(editor.binding.viewStore.getSnapshot().provisionalCaret).toBeNull();
  expect(editor.caretRect).not.toHaveBeenCalled();
  await waitFor(() => expect(editor.commands).toHaveLength(2));
  expect(editor.commands[1]).toMatchObject({ payload: { clickCount: 2 } });
  await editor.acknowledge(1, 'applied');
  editor.present(frame(3, editor.commands[1]!.seq));
  await waitFor(() => expect(editor.binding.viewStore.getSnapshot().geometry?.frameEpoch).toBe(3));
  press();
  expect(editor.binding.viewStore.getSnapshot().provisionalCaret).toBeNull();
  expect(editor.caretRect).not.toHaveBeenCalled();
  await waitFor(() => expect(editor.commands).toHaveLength(3));
  expect(editor.commands[2]).toMatchObject({ payload: { clickCount: 3 } });
});

test('a covering frame replaces the provisional caret with its authoritative caret', async () => {
  const editor = await provisionalEditor();
  fireEvent.pointerDown(editor.host, editor.point);
  expect(editor.caret()?.style.left).toBe('77px');
  await editor.acknowledge(0, 'applied', { version: 0 });
  expect(editor.caret()?.style.left).toBe('77px');
  editor.present(frame(2, editor.commands[0]!.seq));
  await waitFor(() => expect(editor.caret()?.style.left).toBe('95px'));
  expect(editor.caret()?.style.top).toBe('1208px');
  expect(editor.binding.viewStore.getSnapshot().provisionalCaret).toBeNull();
});

for (const excluded of ['shift-click', 'non-body hit', 'stale geometry', 'pending input', 'unloaded engine', 'double-click', 'triple-click', 'nonprimary press', 'secondary button', 'read-only', 'composition', 'modified click']) {
  test(`provisional caret does not engage for ${excluded}`, async () => {
    const editor = await provisionalEditor(excluded === 'non-body hit' ? 'header' : 'body', excluded !== 'unloaded engine');
    const queryCount = editor.queries.mock.calls.length;
    if (excluded === 'stale geometry') {
      act(() => editor.binding.viewStore.acceptWorkerState(identity, {
        epoch: 1, version: 1, throughSeq: 1, presentedThroughSeq: 0,
        selection: undefined, editState: undefined,
      }));
    }
    if (excluded === 'read-only') act(() => { editor.binding.admission.setReadOnly(true); });
    if (excluded === 'composition') preedit(editor.textarea, 'compositionstart');
    if (excluded === 'pending input') beforeinput(editor.textarea, 'insertText', 'x');
    fireEvent.pointerDown(editor.host, { ...editor.point, shiftKey: excluded === 'shift-click',
      isPrimary: excluded !== 'nonprimary press', button: excluded === 'secondary button' ? 2 : 0,
      ctrlKey: excluded === 'modified click', detail: excluded === 'double-click' ? 2 : excluded === 'triple-click' ? 3 : 0 });
    expect(editor.binding.viewStore.getSnapshot().provisionalCaret ?? null).toBeNull();
    expect(editor.caret()?.style.left).not.toBe('77px');
    expect(editor.caretRect).not.toHaveBeenCalled();
    if (excluded !== 'non-body hit') expect(editor.queries).toHaveBeenCalledTimes(queryCount);
    expect(editor.load).toHaveBeenCalledTimes(excluded === 'unloaded engine' ? 1 : 0);
  });
}

test('a navigate intent drops the provisional caret before another frame', async () => {
  const editor = await provisionalEditor();
  fireEvent.pointerDown(editor.host, editor.point);
  expect(editor.caret()?.style.left).toBe('77px');
  await editor.acknowledge(0, 'applied');
  fireEvent.keyDown(editor.textarea, { key: 'ArrowRight', code: 'ArrowRight' });
  expect(editor.binding.viewStore.getSnapshot().provisionalCaret).toBeNull();
  expect(editor.caret()).toBeNull();
  await waitFor(() => expect(editor.commands[1]?.kind).toBe('navigate'));
});

test('composition start drops the provisional caret', async () => {
  const editor = await provisionalEditor();
  fireEvent.pointerDown(editor.host, editor.point);
  expect(editor.caret()?.style.left).toBe('77px');
  preedit(editor.textarea, 'compositionstart');
  expect(editor.binding.viewStore.getSnapshot().provisionalCaret).toBeNull();
  expect(editor.caret()).toBeNull();
});

test('typing waits for the click frame before echo replaces the provisional caret', async () => {
  const editor = await provisionalEditor();
  fireEvent.pointerDown(editor.host, editor.point);
  expect(editor.caret()?.style.left).toBe('77px');
  beforeinput(editor.textarea, 'insertText', 'x');
  expect(editor.events).not.toContain('echo:x');
  expect(editor.caret()?.style.left).toBe('77px');
  await editor.acknowledge(0, 'applied');
  await editor.acknowledge(1, 'applied');
  const clear = spyOn(editor.binding.viewStore, 'setProvisionalCaret');
  restorers.push(() => clear.mockRestore());
  editor.present(frame(2, editor.commands[0]!.seq));
  await waitFor(() => expect(editor.events).toContain('echo:x'));
  expect(clear).toHaveBeenCalledWith(null);
  expect(editor.binding.viewStore.getSnapshot().provisionalCaret).toBeNull();
  expect(editor.caret()).toBeNull();
});

test('touch presses allow native panning without selecting or focusing the input', async () => {
  const editor = await mount();
  const { host, capture } = pointerHost(editor.container);
  const surface = editor.canvas.parentElement!;
  expect(host.style.touchAction).toBe('pan-x pan-y');
  expect(surface.style.touchAction).toBe('pan-x pan-y');
  const down = new PointerEvent('pointerdown', { pointerId: 1, pointerType: 'touch', isPrimary: true,
    button: 0, buttons: 1, clientX: 50, clientY: 10, bubbles: true, cancelable: true });
  fireEvent(surface, down);
  expect(down.defaultPrevented).toBe(false);
  expect(capture).not.toHaveBeenCalled();
  expect(editor.binding.admission.getSnapshot().originalUnackedIntents).toBe(0);
  expect(document.activeElement).not.toBe(editor.textarea);
});

test('a touch tap places the caret on release and keeps the document input focused', async () => {
  const editor = await mount();
  const { host } = pointerHost(editor.container);
  const point = { pointerId: 1, pointerType: 'touch', isPrimary: true, button: 0, clientX: 50, clientY: 10 };
  fireEvent.pointerDown(host, { ...point, buttons: 1 });
  expect(editor.binding.admission.getSnapshot().originalUnackedIntents).toBe(0);
  fireEvent.pointerMove(host, { ...point, buttons: 1, clientX: 52 });
  fireEvent.pointerUp(host, { ...point, clientX: 52 });
  await waitFor(() => expect(editor.commands).toHaveLength(1));
  expect(editor.commands[0]).toMatchObject({ kind: 'selectPoint', payload: { page, x: 52, y: 10, clickCount: 1 } });
  const mouse = new MouseEvent('mousedown', { button: 0, bubbles: true, cancelable: true });
  fireEvent(host, mouse);
  expect(mouse.defaultPrevented).toBe(true);
  expect(document.activeElement).toBe(editor.textarea);
});

for (const end of ['up', 'cancel', 'lost-capture'] as const) {
  test(`a touch long press extends selection until pointer ${end}`, async () => {
    const editor = await mount();
    const { host, capture, release } = pointerHost(editor.container);
    jest.useFakeTimers();
    restorers.push(() => jest.useRealTimers());
    const point = { pointerId: 1, pointerType: 'touch', isPrimary: true, button: 0, clientX: 50, clientY: 10 };
    fireEvent.pointerDown(host, { ...point, buttons: 1 });
    fireEvent.pointerMove(host, { ...point, buttons: 1, clientX: 57 });
    act(() => { jest.advanceTimersByTime(499); });
    expect(editor.binding.admission.getSnapshot().originalUnackedIntents).toBe(0);
    expect(capture).not.toHaveBeenCalled();
    expect(host.style.touchAction).toBe('pan-x pan-y');
    act(() => { jest.advanceTimersByTime(1); });
    expect(capture).toHaveBeenCalledWith(1);
    expect(host.style.touchAction).toBe('none');
    const move = new Event('touchmove', { bubbles: true, cancelable: true });
    fireEvent(host, move);
    expect(move.defaultPrevented).toBe(true);
    fireEvent.pointerMove(host, { ...point, buttons: 1, clientX: 80 });
    if (end === 'up') fireEvent.pointerUp(host, { ...point, clientX: 80 });
    else if (end === 'cancel') fireEvent.pointerCancel(host, point);
    else fireEvent.lostPointerCapture(host, point);
    expect(release).toHaveBeenCalledWith(1);
    expect(host.style.touchAction).toBe('pan-x pan-y');
    fireEvent.pointerMove(host, { ...point, buttons: 1, clientX: 90 });
    jest.useRealTimers();
    await waitFor(() => expect(editor.commands).toHaveLength(2));
    expect(editor.commands[0]).toMatchObject({ kind: 'selectPoint', payload: { page, x: 50, y: 10 } });
    expect(editor.commands[1]).toMatchObject({ kind: 'extendSelection', payload: { page, x: 80, y: 10 } });
  });
}

test('a touch long press keeps selecting when capture transfers from the canvas to the host', async () => {
  const editor = await mount();
  const { host, capture, release } = pointerHost(editor.container);
  jest.useFakeTimers();
  restorers.push(() => jest.useRealTimers());
  const point = { pointerId: 1, pointerType: 'touch', isPrimary: true, button: 0, clientX: 50, clientY: 10 };
  fireEvent.pointerDown(editor.canvas, { ...point, buttons: 1 });
  act(() => { jest.advanceTimersByTime(500); });
  expect(capture).toHaveBeenCalledWith(1);
  fireEvent.lostPointerCapture(editor.canvas, { ...point, bubbles: true });
  expect(release).not.toHaveBeenCalled();
  expect(host.style.touchAction).toBe('none');
  fireEvent.pointerMove(host, { ...point, buttons: 1, clientX: 80 });
  fireEvent.pointerUp(host, { ...point, clientX: 80 });
  expect(release).toHaveBeenCalledWith(1);
  expect(host.style.touchAction).toBe('pan-x pan-y');
  fireEvent.pointerMove(host, { ...point, buttons: 1, clientX: 90 });
  jest.useRealTimers();
  await waitFor(() => expect(editor.commands).toHaveLength(2));
  expect(editor.commands[0]).toMatchObject({ kind: 'selectPoint', payload: { page, x: 50, y: 10 } });
  expect(editor.commands[1]).toMatchObject({ kind: 'extendSelection', payload: { page, x: 80, y: 10 } });
});

test('a touch swipe cancels long press and keeps native scrolling', async () => {
  const editor = await mount();
  const { host, capture } = pointerHost(editor.container);
  jest.useFakeTimers();
  restorers.push(() => jest.useRealTimers());
  const point = { pointerId: 1, pointerType: 'touch', isPrimary: true, button: 0, clientX: 50, clientY: 10 };
  fireEvent.pointerDown(host, { ...point, buttons: 1 });
  fireEvent.pointerMove(host, { ...point, buttons: 1, clientX: 80 });
  act(() => { jest.advanceTimersByTime(500); });
  const move = new Event('touchmove', { bubbles: true, cancelable: true });
  fireEvent(host, move);
  expect(move.defaultPrevented).toBe(false);
  fireEvent.pointerUp(host, { ...point, clientX: 80 });
  expect(capture).not.toHaveBeenCalled();
  expect(host.style.touchAction).toBe('pan-x pan-y');
  expect(editor.binding.admission.getSnapshot().originalUnackedIntents).toBe(0);
  expect(editor.commands).toEqual([]);
});

test('touch taps seven pixels apart select a word then a paragraph', async () => {
  const editor = await mount();
  const { host } = pointerHost(editor.container);
  for (const clientX of [50, 57, 64]) {
    const point = { pointerId: 1, pointerType: 'touch', isPrimary: true, button: 0, clientX, clientY: 10 };
    fireEvent.pointerDown(host, { ...point, buttons: 1 });
    fireEvent.pointerUp(host, point);
  }
  await waitFor(() => expect(editor.commands).toHaveLength(3));
  expect(editor.commands).toMatchObject([
    { kind: 'selectPoint', payload: { clickCount: 1 } },
    { kind: 'selectPoint', payload: { clickCount: 2 } },
    { kind: 'selectPoint', payload: { clickCount: 3 } },
  ]);
});

for (const end of ['move', 'release-away', 'cancel', 'lost-capture'] as const) {
  test(`a touch ${end} does not place the caret or extend selection`, async () => {
    const editor = await mount();
    const { host } = pointerHost(editor.container);
    const point = { pointerId: 1, pointerType: 'touch', isPrimary: true, button: 0, clientX: 50, clientY: 10 };
    fireEvent.pointerDown(host, { ...point, buttons: 1 });
    if (end === 'move') {
      const move = new PointerEvent('pointermove', { ...point, buttons: 1, clientY: 30, bubbles: true, cancelable: true });
      fireEvent(host, move);
      expect(move.defaultPrevented).toBe(false);
      fireEvent.pointerMove(host, { ...point, buttons: 1 });
    } else if (end === 'cancel') fireEvent.pointerCancel(host, point);
    else if (end === 'lost-capture') fireEvent.lostPointerCapture(host, point);
    fireEvent.pointerUp(host, { ...point, clientY: end === 'release-away' ? 30 : 10 });
    expect(editor.binding.admission.getSnapshot().originalUnackedIntents).toBe(0);
    expect(editor.commands).toEqual([]);
    expect(document.activeElement).not.toBe(editor.textarea);
  });
}

for (const target of ['pages', 'outside']) {
  test(`a second touch on ${target} cancels pending caret placement`, async () => {
    const editor = await mount();
    const { host } = pointerHost(editor.container);
    const second = target === 'pages' ? host : document.body;
    const point = { pointerType: 'touch', button: 0, clientX: 50, clientY: 10 };
    fireEvent.pointerDown(host, { ...point, pointerId: 1, isPrimary: true, buttons: 1 });
    fireEvent.pointerDown(second, { ...point, pointerId: 2, isPrimary: false, buttons: 1 });
    fireEvent.pointerUp(second, { ...point, pointerId: 2, isPrimary: false });
    fireEvent.pointerUp(host, { ...point, pointerId: 1, isPrimary: true });
    expect(editor.binding.admission.getSnapshot().originalUnackedIntents).toBe(0);
    expect(editor.commands).toEqual([]);
    expect(document.activeElement).not.toBe(editor.textarea);
  });
}

for (const pointerType of ['mouse', 'pen']) {
  test(`${pointerType} dragging selects text and stops extending after release`, async () => {
    const editor = await mount();
    const { host, capture, release } = pointerHost(editor.container);
    const point = { pointerId: 1, pointerType, isPrimary: true, button: 0, clientX: 50, clientY: 10 };
    const down = new PointerEvent('pointerdown', { ...point, buttons: 1, bubbles: true, cancelable: true });
    fireEvent(host, down);
    expect(down.defaultPrevented).toBe(true);
    expect(capture).toHaveBeenCalledWith(1);
    fireEvent.pointerMove(host, { ...point, buttons: 1, clientX: 80 });
    fireEvent.pointerUp(host, { ...point, clientX: 80 });
    expect(release).toHaveBeenCalledWith(1);
    fireEvent.pointerMove(host, { ...point, buttons: 1, clientX: 90 });
    await waitFor(() => expect(editor.commands).toHaveLength(2));
    expect(editor.commands[0]).toMatchObject({ kind: 'selectPoint', payload: { page, x: 50, y: 10 } });
    expect(editor.commands[1]).toMatchObject({ kind: 'extendSelection', payload: { page, x: 80, y: 10 } });
    expect(document.activeElement).toBe(editor.textarea);
  });
}

for (const pointerType of ['mouse', 'touch']) {
  test(`a ${pointerType} press outside the pages ends their click sequence`, async () => {
    const counts: number[] = [];
    const editor = await mount(false, false, false, undefined, undefined, {
      onPartPoint: ({ clickCount }) => { counts.push(clickCount); return true; },
      pageOverlays: <button type="button">Close header editing</button>,
    });
    const { host } = pointerHost(editor.container);
    const press = (target: Element) => {
      fireEvent.pointerDown(target, { pointerId: 1, pointerType, isPrimary: true, button: 0, clientX: 50, clientY: 10 });
      fireEvent.pointerUp(target, { pointerId: 1, pointerType, isPrimary: true, button: 0, clientX: 50, clientY: 10 });
    };
    press(host);
    press(host);
    press(editor.getByRole('button', { name: 'Close header editing' }));
    press(host);
    press(host);
    press(document.body);
    press(host);
    expect(counts).toEqual([1, 2, 1, 2, 1]);
  });
}

test('single_cell selection keeps the worker caret next to its highlight', async () => {
  const editor = await mount(true, false, true);
  await waitFor(() => expect(editor.container.querySelectorAll('[aria-hidden="true"]')).toHaveLength(2));
});

test('multi_cell selection hides the worker caret but keeps its highlight', async () => {
  const editor = await mount(true, false, true, undefined, () => 10, {}, true);
  await waitFor(() => expect(editor.container.querySelectorAll('[aria-hidden="true"]')).toHaveLength(1));
});

for (const change of ['scroll', 'ancestor zoom'] as const) {
  test(`picture drag tracks current page coordinates after ${change}`, async () => {
    const move = mock(() => {});
    const editor = await mount(false, false, false, candidate => ({
      identity, documentVersion: 'v1', coversThroughSeq: candidate.coversThroughSeq, frameEpoch: candidate.revision,
      layoutRevision: candidate.payload.layoutRevision, caret: null, selectionRects: [],
      pages: [{ page, bounds: { x: 0, y: 0, width: 200, height: 40 }, transform: [1, 0, 0, 1, 0, 0] }],
      objects: { version: 'v1', frameEpoch: candidate.revision, layoutRevision: candidate.payload.layoutRevision,
        selection: [], unbuiltPages: [], objects: [{ handle: 'picture', kind: 'image', story: 'body', floating: true,
          rects: [{ ...page, x: 20, y: 5, width: 40, height: 20 }] }] },
    }), undefined, { pictures: { canMove: () => true, move } });
    const host = editor.container.querySelector<HTMLElement>('[data-peerless-pages-host]')!;
    let top = 100;
    let scale = 1;
    Object.defineProperties(host, {
      getBoundingClientRect: { value: () => ({ left: 0, top, width: 200 * scale, height: 40 * scale }) },
      setPointerCapture: { value() {} }, hasPointerCapture: { value: () => true }, releasePointerCapture: { value() {} },
    });
    const point = { pointerId: 1, pointerType: 'mouse', button: 0, buttons: 1, clientX: 30, clientY: 110 };
    fireEvent.pointerDown(host, point);
    if (change === 'scroll') top = 80;
    else scale = 2;
    const end = { ...point, clientX: 50 * scale, clientY: top + 15 * scale };
    fireEvent.pointerMove(host, end);
    const preview = host.querySelector<HTMLElement>('.paged-editor__picture-drag-preview')!;
    expect(preview.style.left).toBe('40px');
    expect(preview.style.top).toBe('10px');
    fireEvent.pointerUp(host, { ...end, buttons: 0 });
    expect(move).toHaveBeenCalledWith({ kind: 'moveObject', payload: {
      handle: 'picture', expectVersion: 'v1', hOffsetEmu: 40 * 9525, vOffsetEmu: 10 * 9525,
    } });
  });
}
