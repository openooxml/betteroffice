import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, jest, mock, spyOn, test } from 'bun:test';
import type {
  DocxCommandAck, DocxCommandEnvelope, DocxCommandFrame, ResidentCommandConnection,
} from '@betteroffice/docx/peerless';
import {
  applyFrameDeltaOwned, displayPageRevision, FRAME_DELTA_VERSION,
  type DecodedFrameDelta, type DisplayPage, type FramePageOperation,
} from '@betteroffice/docx/layout/render';
import type { PeerlessEditSessionBinding, UsePeerlessEditSessionOptions } from '../hooks/usePeerlessEditSession';
import * as composition from './composition';
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
    async acknowledge(index: number, status: 'applied' | 'refused') {
      await waitFor(() => { expect(commands[index]).toBeDefined(); });
      const command = commands[index]!;
      await act(async () => { replies[index]!({ ...initial, frame: undefined, version: index + 1,
        throughSeq: command.seq, outcomes: [{ seq: command.seq, status,
          ...(status === 'refused' ? { reason: 'locked-target' } : {}) }],
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
  const mirror = () => editor.container.querySelector('[data-page-id="1"] .canvas-page-mirror')?.firstElementChild;
  await waitFor(() => expect(mirror()).toBeTruthy());
  const before = mirror();
  editor.present(frame(2, 0));
  expect(retained.pages[0]!.page).toBe(displayPage);
  expect(displayPageRevision(displayPage)).toBe(1);
  await waitFor(() => expect(mirror()).not.toBe(before), { timeout: 5_000 });
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

test('a press outside the pages ends their click sequence', async () => {
  const counts: number[] = [];
  const editor = await mount(false, false, false, undefined, undefined, {
    onPartPoint: ({ clickCount }) => { counts.push(clickCount); return true; },
    pageOverlays: <button type="button">Close header editing</button>,
  });
  const host = editor.container.querySelector<HTMLElement>('[data-peerless-pages-host]')!;
  Object.defineProperty(host, 'getBoundingClientRect', { value: () => ({ x: 0, y: 0, left: 0, top: 0, right: 200, bottom: 40, width: 200, height: 40 }) });
  const press = (target: Element) => {
    fireEvent.pointerDown(target, { pointerId: 1, button: 0, clientX: 50, clientY: 10 });
    fireEvent.pointerUp(target, { pointerId: 1, button: 0, clientX: 50, clientY: 10 });
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

test('single_cell selection keeps the worker caret next to its highlight', async () => {
  const editor = await mount(true, false, true);
  await waitFor(() => expect(editor.container.querySelectorAll('[aria-hidden="true"]')).toHaveLength(2));
});

test('multi_cell selection hides the worker caret but keeps its highlight', async () => {
  const editor = await mount(true, false, true, undefined, () => 10, {}, true);
  await waitFor(() => expect(editor.container.querySelectorAll('[aria-hidden="true"]')).toHaveLength(1));
});
