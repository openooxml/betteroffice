import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, spyOn, test } from 'bun:test';
import type { DocxCommandPayload, DocxCommandSelection, DocxWorkerEditSession } from '@betteroffice/docx/peerless';
import type { CommandTicket, SubmitOptions, TicketOutcome } from '../../../../../../shared/office-session/commands/types';
import type { PeerlessEchoIntent, PeerlessEchoPainter } from './echoPainter';
import type { PeerlessFrameGeometry, PeerlessViewSnapshot, PeerlessViewStore } from './viewStore';
import { PeerlessInput, type PeerlessInputProps } from './PeerlessInput';
import { clipboardFragmentHtml } from '../internals/clipboardHtml';
import type { DocxClipboardFragment } from '@betteroffice/docx/yrs';
import { admitIntent, createInputAdmission, submitFormattingIntent } from './inputAdmission';
import { installCompositionEventPolyfill } from './compositionEventPolyfill.testing';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const restoreCompositionEvent = installCompositionEventPolyfill();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  restoreCompositionEvent();
  if (ownsDom) await GlobalRegistrator.unregister();
});

function mount(readOnly = false, callbacks: Pick<PeerlessInputProps, 'onCompositionChange' | 'onRecoverableText' | 'onClipboard' | 'onNavigateCell' | 'onPasteFragment'> = {}) {
  const submitted: string[] = [];
  const entries: Array<{
    payload: DocxCommandPayload; options: SubmitOptions<PeerlessEchoIntent>; ticket: CommandTicket;
    settle(outcome: TicketOutcome, selection?: DocxCommandSelection): void; reject(error: unknown): void;
  }> = [];
  const waiters: Array<{ count: number; resolve(): void }> = [];
  const waitForSubmitted = (count: number) => submitted.length >= count ? Promise.resolve()
    : new Promise<void>((resolve) => { waiters.push({ count, resolve }); });
  const listeners = new Set<() => void>();
  const snapshot: PeerlessViewSnapshot = {
    identity: { session: 'test', incarnation: 1, epoch: 1 }, version: 0,
    throughSeq: 0, presentedThroughSeq: 0, workerSelection: null, anchors: null, geometry: null,
    editState: { readOnly, toolbar: null, effectiveFormatting: null, storedMarks: null,
      canUndo: false, canRedo: false, hyperlink: null, tableContext: null },
  };
  const viewStore: PeerlessViewStore = {
    subscribe: (listener) => { listeners.add(listener); return () => { listeners.delete(listener); }; },
    getSnapshot: () => snapshot, getServerSnapshot: () => snapshot,
    acceptWorkerState() {}, installFrameGeometry(geometry) {
      snapshot.geometry = geometry;
      for (const listener of listeners) listener();
    }, dispose() {},
  };
  const session: DocxWorkerEditSession<PeerlessEchoIntent> = {
    backend: 'worker', state: { epoch: 1, version: 0, throughSeq: 0,
      presentedThroughSeq: 0, selection: null, editState: snapshot.editState },
    submit(kind, payload, options = {}) {
      submitted.push(kind);
      let settle!: (outcome: TicketOutcome) => void;
      let reject!: (error: unknown) => void;
      const ticket: CommandTicket = { id: submitted.length, seq: submitted.length,
        ack: new Promise((resolve, fail) => { settle = resolve; reject = fail; }), presented: new Promise(() => {}) };
      entries.push({ payload, options, ticket, reject, settle(outcome, selection) {
        if (selection) {
          snapshot.workerSelection = selection;
          snapshot.throughSeq = outcome.seq ?? 0;
          session.state.selection = selection;
          session.state.throughSeq = outcome.seq ?? 0;
        }
        settle(outcome);
      } });
      if ((kind === 'insertText' || kind === 'replaceSelection') && 'text' in payload &&
        /[\r\n\u000B\u2028\u2029]/u.test(payload.text)) {
        settle({ status: 'refused', seq: ticket.seq, reason: 'invalid-text' });
      }
      for (const waiter of waiters) if (submitted.length >= waiter.count) waiter.resolve();
      return ticket;
    },
    subscribe: () => () => {}, whenPresented: async () => 'presented',
    flush: async () => {}, dispose: async () => {},
  };
  const echoPainter: PeerlessEchoPainter = {
    installContext() {}, restoreBackground: () => false, preview: () => undefined,
    paint() {}, drop() {}, reconcile() {}, disable() {}, dispose() {},
  };
  const admission = createInputAdmission({ session, viewStore, echoPainter });
  const view = render(<PeerlessInput admission={admission} viewStore={viewStore} {...callbacks} />);
  return { textarea: view.getByRole('textbox') as HTMLTextAreaElement, admission, submitted, entries,
    snapshot, session, viewStore, waitForSubmitted, unmount: view.unmount,
    refresh: () => view.rerender(<PeerlessInput admission={admission} viewStore={viewStore} {...callbacks} />) };
}

function beforeinput(textarea: HTMLTextAreaElement, inputType: string, data: string | null = null) {
  act(() => { textarea.dispatchEvent(new InputEvent('beforeinput', {
    bubbles: true, cancelable: true, inputType, data,
  })); });
}

test('a held submission failure does not report applied composition text as recoverable', async () => {
  const recovered: string[] = [];
  const phases: string[] = [];
  const input = mount(false, {
    onRecoverableText: text => recovered.push(text),
    onCompositionChange: state => { if (state) phases.push(state.phase); },
  });
  fireEvent.compositionStart(input.textarea);
  fireEvent.compositionUpdate(input.textarea, { data: '漢字' });
  const failed = admitIntent(input.admission, { kind: 'insertText', payload: { text: 'failed' } });
  const last = admitIntent(input.admission, { kind: 'insertText', payload: { text: 'last' } });
  if (failed.status !== 'held' || last.status !== 'held') throw new Error('intents were not held');
  const failure = new Error('held submit failed');
  const send = input.session.submit.bind(input.session);
  const submit = spyOn(input.session, 'submit').mockImplementation((kind, payload, options) => {
    if ('text' in payload && payload.text === 'failed') throw failure;
    return send(kind, payload, options);
  });
  try {
    fireEvent.compositionEnd(input.textarea, { data: '漢字' });
    expect(input.submitted).toEqual(['replaceSelection']);
    expect(recovered).toEqual([]);
    expect(phases).toEqual(['active', 'active', 'committed']);
    await expect(failed.released).rejects.toBe(failure);
    await expect(last.released).rejects.toBe(failure);
    expect(input.admission.getSnapshot().originalUnackedIntents).toBe(1);
    await act(async () => {
      input.entries[0]!.settle({ status: 'applied', seq: 1 });
      await input.entries[0]!.ticket.ack;
    });
    expect(recovered).toEqual([]);
    expect(phases).not.toContain('refused');
    expect(input.admission.getSnapshot()).toMatchObject({ originalUnackedIntents: 0, tail: undefined, echoEnabled: true });
  } finally { submit.mockRestore(); input.unmount(); input.admission.dispose(); }
});

test('d2b_overlapping_browser_events_commit_once', () => {
  const { textarea, submitted } = mount();
  fireEvent.keyDown(textarea, { key: 'Backspace', code: 'Backspace' });
  beforeinput(textarea, 'deleteContentBackward');
  fireEvent.input(textarea, { inputType: 'deleteContentBackward', target: { value: '' } });
  fireEvent.keyUp(textarea, { key: 'Backspace' });
  fireEvent.keyDown(textarea, { key: 'x', code: 'KeyX' });
  beforeinput(textarea, 'insertText', 'x');
  fireEvent.input(textarea, { data: 'x', target: { value: 'x' } });
  fireEvent.keyUp(textarea, { key: 'x' });
  fireEvent.input(textarea, { inputType: 'insertText', data: 'y', target: { value: 'y' } });
  fireEvent.paste(textarea, { clipboardData: { getData: () => 'paste' } });
  beforeinput(textarea, 'insertFromPaste');
  fireEvent.input(textarea, { inputType: 'insertFromPaste', target: { value: 'paste' } });
  fireEvent.keyDown(textarea, { key: 'b', ctrlKey: true });
  beforeinput(textarea, 'formatBold');
  fireEvent.input(textarea, { inputType: 'formatBold' });
  fireEvent.keyUp(textarea, { key: 'b' });
  expect(submitted).toEqual(['deleteBackward', 'insertText', 'insertText', 'replaceSelection', 'toggleMark']);
  expect(textarea.value).toBe('');
});

test('Shift+Enter inserts one line break across overlapping browser events', () => {
  for (const inputType of ['insertLineBreak', 'insertParagraph']) {
    const input = mount();
    beforeinput(input.textarea, 'insertText', 'a');
    const enter = new KeyboardEvent('keydown', {
      key: 'Enter', code: 'Enter', shiftKey: true, bubbles: true, cancelable: true,
    });
    act(() => { input.textarea.dispatchEvent(enter); });
    expect(enter.defaultPrevented).toBe(true);
    beforeinput(input.textarea, inputType);
    fireEvent.input(input.textarea, { inputType, target: { value: '\n' } });
    fireEvent.keyUp(input.textarea, { key: 'Enter', shiftKey: true });
    beforeinput(input.textarea, 'insertText', 'b');
    fireEvent.keyDown(input.textarea, { key: 'Enter' });
    beforeinput(input.textarea, 'insertParagraph');
    fireEvent.input(input.textarea, { inputType: 'insertParagraph', target: { value: '\n' } });
    fireEvent.keyUp(input.textarea, { key: 'Enter' });
    expect(input.submitted).toEqual(['insertText', 'lineBreak', 'insertText', 'paragraphBreak']);
    expect(input.entries[1]!.payload).toEqual({});
    expect(input.entries[1]!.options.fence).toBe(true);
    for (let i = 1; i < input.entries.length; i += 1) {
      expect(input.entries[i]!.options.after).toBe(input.entries[i - 1]!.ticket);
    }
    input.unmount();
    input.admission.dispose();
  }
});

test('d2b_identical_input_after_trailing_composition_is_admitted', () => {
  for (const trailingBeforeinput of [true, false]) {
    const input = mount();
    fireEvent.compositionStart(input.textarea);
    fireEvent.compositionUpdate(input.textarea, { data: 'あ' });
    fireEvent.compositionEnd(input.textarea, { data: 'あ' });
    if (trailingBeforeinput) beforeinput(input.textarea, 'insertFromComposition', 'あ');
    fireEvent.input(input.textarea, { inputType: 'insertFromComposition', target: { value: 'あ' } });
    expect(input.submitted).toEqual(['replaceSelection']);
    beforeinput(input.textarea, 'insertText', 'あ');
    fireEvent.input(input.textarea, { inputType: 'insertText', target: { value: 'あ' } });
    fireEvent.input(input.textarea, { inputType: 'insertText', target: { value: 'あ' } });
    expect(input.submitted).toEqual(['replaceSelection', 'insertText', 'insertText']);
    expect(input.entries.map(({ payload }) => payload)).toEqual([{ text: 'あ' }, { text: 'あ' }, { text: 'あ' }]);
    input.unmount();
  }
});

test('d2b_vertical_navigation_uses_geometry_after_pending_typing', async () => {
  for (const horizontal of [true, false]) {
    const input = mount();
    const geometry = (coversThroughSeq: number, x: number) => ({
      identity: input.snapshot.identity, coversThroughSeq, frameEpoch: coversThroughSeq + 1,
      layoutRevision: coversThroughSeq + 1, pages: [], caret: { x, y: 10, width: 1, height: 12 }, selectionRects: [],
    });
    act(() => { input.viewStore.installFrameGeometry(geometry(0, 4)); });
    beforeinput(input.textarea, 'insertText', 'abc');
    if (horizontal) {
      fireEvent.keyDown(input.textarea, { key: 'ArrowLeft' });
      fireEvent.keyUp(input.textarea, { key: 'ArrowLeft' });
    }
    const count = input.entries.length;
    fireEvent.keyDown(input.textarea, { key: 'ArrowDown' });
    fireEvent.keyUp(input.textarea, { key: 'ArrowDown' });
    beforeinput(input.textarea, 'insertText', 'z');
    expect(input.entries).toHaveLength(count);
    for (let i = 0; i < count; i += 1) input.entries[i]!.settle({ status: 'applied', seq: i + 1 });
    for (let i = 0; i < 8; i += 1) await Promise.resolve();
    act(() => { input.viewStore.installFrameGeometry(geometry(count - 1, 7)); });
    for (let i = 0; i < 4; i += 1) await Promise.resolve();
    expect(input.entries).toHaveLength(count);
    act(() => { input.viewStore.installFrameGeometry(geometry(count, 19)); });
    await input.waitForSubmitted(count + 2);
    expect(input.entries[count]!.payload).toEqual({ unit: 'line', direction: 'down', extend: false, goalX: 19 });
    expect(input.entries[count]!.options.after).toBe(input.entries[count - 1]!.ticket);
    expect(input.entries[count + 1]!.payload).toEqual({ text: 'z' });
    expect(input.entries[count + 1]!.options.after).toBe(input.entries[count]!.ticket);
    input.unmount();
    input.admission.dispose();
  }
});

test('d2b_null_caret_frame_releases_vertical_navigation_and_typing', async () => {
  for (const story of ['body', 'header', 'cell']) {
    const input = mount();
    try {
      const key = story === 'body' ? 'ArrowDown' : 'ArrowUp';
      admitIntent(input.admission, { kind: 'selectPoint', payload: {
        page: { pageId: '1', pageIndex: 0 }, x: 12, y: 20, layoutRevision: 1,
        frameEpoch: 1, clickCount: 1, story,
        modifiers: { shift: story === 'body', alt: false, ctrl: false, meta: false },
      } });
      fireEvent.keyDown(input.textarea, { key });
      fireEvent.keyUp(input.textarea, { key });
      beforeinput(input.textarea, 'insertText', 'z');
      input.entries[0]!.settle({ status: 'applied', seq: 1 }, {
        anchor: { story, paraId: 'worker-selection', offset: 2 },
        focus: { story, paraId: 'worker-selection', offset: story === 'body' ? 5 : 2 },
      });
      for (let i = 0; i < 8; i += 1) await Promise.resolve();
      expect(input.entries).toHaveLength(1);
      act(() => { input.viewStore.installFrameGeometry({
        identity: input.snapshot.identity, coversThroughSeq: 1, frameEpoch: 2,
        layoutRevision: 2, pages: [], caret: null, selectionRects: [],
      }); });
      for (let i = 0; i < 12; i += 1) await Promise.resolve();
      expect(input.submitted).toEqual(['selectPoint', 'navigate', 'insertText']);
      expect(input.entries[1]!.payload).toEqual({
        unit: 'line', direction: story === 'body' ? 'down' : 'up', extend: false, goalX: 0,
      });
      expect(input.entries[1]!.options.after).toBe(input.entries[0]!.ticket);
      expect(input.entries[2]!.options.after).toBe(input.entries[1]!.ticket);
      expect(input.entries[2]!.payload).toEqual({ text: 'z' });
    } finally {
      input.unmount();
      input.admission.dispose();
    }
  }
});

async function navigateNullCaret(options: {
  key: 'ArrowUp' | 'ArrowDown'; story: string; goalX: number; shiftKey?: boolean;
  selectionRects?: PeerlessFrameGeometry['selectionRects']; pages?: PeerlessFrameGeometry['pages'];
}) {
  const input = mount();
  try {
    admitIntent(input.admission, { kind: 'selectPoint', payload: {
      page: { pageId: '1', pageIndex: 0 }, x: 12, y: 20, layoutRevision: 1,
      frameEpoch: 1, clickCount: 1, story: options.story,
      modifiers: { shift: options.story === 'body', alt: false, ctrl: false, meta: false },
    } });
    fireEvent.keyDown(input.textarea, { key: options.key, shiftKey: options.shiftKey });
    fireEvent.keyUp(input.textarea, { key: options.key, shiftKey: options.shiftKey });
    beforeinput(input.textarea, 'insertText', 'z');
    input.entries[0]!.settle({ status: 'applied', seq: 1 }, {
      anchor: { story: options.story, paraId: 'worker-selection', offset: 2 },
      focus: { story: options.story, paraId: 'worker-selection', offset: options.story === 'body' ? 5 : 2 },
    });
    const geometry: PeerlessFrameGeometry = {
      identity: input.snapshot.identity, coversThroughSeq: 0, frameEpoch: 2, layoutRevision: 2,
      pages: options.pages ?? [{ page: { pageId: '1', pageIndex: 0 },
        bounds: { x: 31, y: 0, width: 200, height: 300 }, transform: [1, 0, 0, 1, 0, 0] }],
      caret: null, selectionRects: options.selectionRects ?? [],
    };
    act(() => { input.viewStore.installFrameGeometry(geometry); });
    for (let i = 0; i < 12; i += 1) await Promise.resolve();
    expect(input.entries).toHaveLength(1);
    act(() => { input.viewStore.installFrameGeometry({ ...geometry, coversThroughSeq: 1, frameEpoch: 3 }); });
    await input.waitForSubmitted(3);
    expect(input.submitted).toEqual(['selectPoint', 'navigate', 'insertText']);
    expect(input.entries[1]!.payload).toEqual({
      unit: 'line', direction: options.key === 'ArrowDown' ? 'down' : 'up',
      extend: options.shiftKey ?? false, goalX: options.goalX,
    });
    expect(input.entries[1]!.options.after).toBe(input.entries[0]!.ticket);
    expect(input.entries[2]!.options.after).toBe(input.entries[1]!.ticket);
    expect(input.entries[2]!.payload).toEqual({ text: 'z' });
  } finally {
    input.unmount();
    input.admission.dispose();
  }
}

test('d2b_null_caret_range_arrow_down_uses_last_right_edge_then_types', async () => {
  for (const shiftKey of [false, true]) await navigateNullCaret({
    key: 'ArrowDown', story: 'body', shiftKey, goalX: 68,
    selectionRects: [{ x: 12, y: 10, width: 80, height: 12 }, { x: 23, y: 30, width: 45, height: 12 }],
  });
});

test('d2b_null_caret_range_arrow_up_uses_first_left_edge_then_types', async () => {
  for (const shiftKey of [false, true]) await navigateNullCaret({
    key: 'ArrowUp', story: 'body', shiftKey, goalX: 12,
    selectionRects: [{ x: 12, y: 10, width: 80, height: 12 }, { x: 23, y: 30, width: 45, height: 12 }],
  });
});

test('d2b_null_caret_header_footer_arrow_down_then_types', async () => {
  for (const story of ['header', 'footer']) await navigateNullCaret({ key: 'ArrowDown', story, goalX: 31 });
});

test('d2b_null_caret_table_cell_arrow_down_then_types', async () => {
  await navigateNullCaret({ key: 'ArrowDown', story: 'cell', goalX: 68,
    selectionRects: [{ x: 23, y: 30, width: 45, height: 12 }],
  });
});

test('d2b_null_caret_empty_rects_without_retained_goal_uses_first_page_left_bound', async () => {
  await navigateNullCaret({ key: 'ArrowDown', story: 'body', goalX: 31, pages: [
    { page: { pageId: '1', pageIndex: 0 }, bounds: { x: 31, y: 0, width: 200, height: 300 },
      transform: [1, 0, 0, 1, 0, 0] },
    { page: { pageId: '2', pageIndex: 1 }, bounds: { x: 100, y: 300, width: 200, height: 300 },
      transform: [1, 0, 0, 1, 0, 0] },
  ] });
});

test('d2b_null_caret_rect_edges_precede_retained_goal_and_empty_rects_retain_it', async () => {
  for (const selectionRects of [[], [{ x: 12, y: 10, width: 30, height: 12 }]]) {
    const input = mount();
    try {
      act(() => { input.viewStore.installFrameGeometry({
        identity: input.snapshot.identity, coversThroughSeq: 0, frameEpoch: 1, layoutRevision: 1,
        pages: [], caret: { x: 80, y: 10, width: 1, height: 12 }, selectionRects: [],
      }); });
      for (let i = 0; i < 2; i += 1) {
        fireEvent.keyDown(input.textarea, { key: 'ArrowDown' });
        fireEvent.keyUp(input.textarea, { key: 'ArrowDown' });
      }
      beforeinput(input.textarea, 'insertText', 'z');
      await input.waitForSubmitted(1);
      expect(input.entries[0]!.payload).toEqual({ unit: 'line', direction: 'down', extend: false, goalX: 80 });
      input.entries[0]!.settle({ status: 'applied', seq: 1 });
      act(() => { input.viewStore.installFrameGeometry({
        identity: input.snapshot.identity, coversThroughSeq: 1, frameEpoch: 2, layoutRevision: 2,
        pages: [{ page: { pageId: '1', pageIndex: 0 }, bounds: { x: 31, y: 0, width: 200, height: 300 },
          transform: [1, 0, 0, 1, 0, 0] }], caret: null, selectionRects,
      }); });
      await input.waitForSubmitted(3);
      expect(input.entries[1]!.payload).toEqual({
        unit: 'line', direction: 'down', extend: false, goalX: selectionRects.length ? 42 : 80,
      });
      expect(input.entries[1]!.options.after).toBe(input.entries[0]!.ticket);
      expect(input.entries[2]!.options.after).toBe(input.entries[1]!.ticket);
      expect(input.entries[2]!.payload).toEqual({ text: 'z' });
    } finally {
      input.unmount();
      input.admission.dispose();
    }
  }
});

test('d2b_vertical_goal_survives_short_lines_and_resets_after_typing_or_horizontal_navigation', async () => {
  for (const reset of ['typing', 'horizontal']) {
    const input = mount();
    const geometry = (coversThroughSeq: number, x: number) => ({
      identity: input.snapshot.identity, coversThroughSeq, frameEpoch: coversThroughSeq + 1,
      layoutRevision: coversThroughSeq + 1, pages: [], caret: { x, y: 10, width: 1, height: 12 }, selectionRects: [],
    });
    const down = () => {
      fireEvent.keyDown(input.textarea, { key: 'ArrowDown' });
      fireEvent.keyUp(input.textarea, { key: 'ArrowDown' });
    };
    act(() => { input.viewStore.installFrameGeometry(geometry(0, 80)); });
    down();
    down();
    await input.waitForSubmitted(1);
    expect(input.entries[0]!.payload).toEqual({ unit: 'line', direction: 'down', extend: false, goalX: 80 });
    input.entries[0]!.settle({ status: 'applied', seq: 1 }, {
      anchor: { story: 'body', paraId: 'short-line', offset: 2 },
      focus: { story: 'body', paraId: 'short-line', offset: 2 },
    });
    act(() => { input.viewStore.installFrameGeometry(geometry(1, 20)); });
    await input.waitForSubmitted(2);
    expect(input.entries[1]!.payload).toEqual({ unit: 'line', direction: 'down', extend: false, goalX: 80 });
    expect(input.entries[1]!.options.after).toBe(input.entries[0]!.ticket);
    if (reset === 'typing') beforeinput(input.textarea, 'insertText', 'z');
    else {
      fireEvent.keyDown(input.textarea, { key: 'ArrowLeft' });
      fireEvent.keyUp(input.textarea, { key: 'ArrowLeft' });
    }
    down();
    await input.waitForSubmitted(3);
    input.entries[1]!.settle({ status: 'applied', seq: 2 }, {
      anchor: { story: 'body', paraId: 'long-line', offset: 8 },
      focus: { story: 'body', paraId: 'long-line', offset: 8 },
    });
    act(() => { input.viewStore.installFrameGeometry(geometry(2, 80)); });
    input.entries[2]!.settle({ status: 'applied', seq: 3 }, {
      anchor: { story: 'body', paraId: 'long-line', offset: 4 },
      focus: { story: 'body', paraId: 'long-line', offset: 4 },
    });
    for (let i = 0; i < 12; i += 1) await Promise.resolve();
    expect(input.entries).toHaveLength(3);
    act(() => { input.viewStore.installFrameGeometry(geometry(3, 35)); });
    await input.waitForSubmitted(4);
    expect(input.entries[3]!.payload).toEqual({ unit: 'line', direction: 'down', extend: false, goalX: 35 });
    expect(input.entries[3]!.options.after).toBe(input.entries[2]!.ticket);
    input.unmount();
    input.admission.dispose();
  }
});

test('d2b_explicit_vertical_goal_wins_over_geometry_and_is_retained', async () => {
  for (const goalX of [0, 41]) {
    const input = mount();
    const geometry = (coversThroughSeq: number, x: number) => ({
      identity: input.snapshot.identity, coversThroughSeq, frameEpoch: coversThroughSeq + 1,
      layoutRevision: coversThroughSeq + 1, pages: [], caret: { x, y: 10, width: 1, height: 12 }, selectionRects: [],
    });
    act(() => { input.viewStore.installFrameGeometry(geometry(0, 80)); });
    admitIntent(input.admission, { kind: 'navigate', payload: { unit: 'line', direction: 'down', extend: false, goalX } });
    fireEvent.keyDown(input.textarea, { key: 'ArrowUp' });
    fireEvent.keyUp(input.textarea, { key: 'ArrowUp' });
    await input.waitForSubmitted(1);
    expect(input.entries[0]!.payload).toEqual({ unit: 'line', direction: 'down', extend: false, goalX });
    input.entries[0]!.settle({ status: 'applied', seq: 1 });
    act(() => { input.viewStore.installFrameGeometry(geometry(1, 20)); });
    await input.waitForSubmitted(2);
    expect(input.entries[1]!.payload).toEqual({ unit: 'line', direction: 'up', extend: false, goalX });
    admitIntent(input.admission, { kind: 'navigate', payload: {
      unit: 'line', direction: 'down', extend: false, goalX: 13,
    } });
    input.entries[1]!.settle({ status: 'applied', seq: 2 });
    act(() => { input.viewStore.installFrameGeometry(geometry(2, 20)); });
    await input.waitForSubmitted(3);
    expect(input.entries[2]!.payload).toEqual({ unit: 'line', direction: 'down', extend: false, goalX: 13 });
    input.unmount();
    input.admission.dispose();
  }
});

test('table_tab_and_shift_tab_navigate_behind_pending_edits', () => {
  const input = mount();
  input.snapshot.editState!.tableContext = {
    story: 'cell', tableIndex: 0, rowIndex: 0, columnIndex: 0, rowCount: 1, columnCount: 1,
    hasMultiCellSelection: false, canSplitCell: false,
  };
  beforeinput(input.textarea, 'insertText', 'x');
  const tab = new KeyboardEvent('keydown', { key: 'Tab', code: 'Tab', bubbles: true, cancelable: true });
  act(() => { input.textarea.dispatchEvent(tab); });
  expect(tab.defaultPrevented).toBe(true);
  beforeinput(input.textarea, 'insertText', '\t');
  fireEvent.input(input.textarea, { inputType: 'insertText', target: { value: '\t' } });
  fireEvent.keyUp(input.textarea, { key: 'Tab' });
  beforeinput(input.textarea, 'insertText', 'y');
  const shiftTab = new KeyboardEvent('keydown', {
    key: 'Tab', code: 'Tab', shiftKey: true, bubbles: true, cancelable: true,
  });
  act(() => { input.textarea.dispatchEvent(shiftTab); });
  expect(shiftTab.defaultPrevented).toBe(true);
  expect(input.submitted).toEqual(['insertText', 'navigateCell', 'insertText', 'navigateCell']);
  expect(input.entries.map(({ payload }) => payload)).toEqual([{ text: 'x' },
    { direction: 'forward', extend: false, createLastRow: true }, { text: 'y' },
    { direction: 'backward', extend: false, createLastRow: false }]);
  expect(input.entries[1]!.options.after).toBe(input.entries[0]!.ticket);
  expect(input.entries[2]!.options.after).toBe(input.entries[1]!.ticket);
  input.unmount();
});

test('d2b_readonly_admits_navigation_only', () => {
  const { textarea, admission, submitted } = mount(true);
  fireEvent.keyDown(textarea, { key: 'ArrowRight' });
  fireEvent.keyUp(textarea, { key: 'ArrowRight' });
  fireEvent.keyDown(textarea, { key: 'Backspace' });
  beforeinput(textarea, 'deleteContentBackward');
  fireEvent.keyUp(textarea, { key: 'Backspace' });
  beforeinput(textarea, 'insertText', 'x');
  fireEvent.input(textarea, { inputType: 'insertText', target: { value: 'x' } });
  fireEvent.paste(textarea, { clipboardData: { getData: () => 'paste' } });
  fireEvent.keyDown(textarea, { key: 'b', ctrlKey: true });
  submitFormattingIntent(admission, { kind: 'toggleMark', payload: { mark: { type: 'bold' } } });
  admitIntent(admission, { kind: 'selectPoint', payload: {
    page: { pageId: '1', pageIndex: 0 }, x: 0, y: 0, layoutRevision: 1, frameEpoch: 1,
    clickCount: 1, modifiers: { shift: false, alt: false, ctrl: false, meta: false },
  } });
  expect(submitted).toEqual(['navigate', 'selectPoint']);
});

test('readonly_transition_preserves_navigation_and_blocks_browser_mutations', () => {
  const input = mount();
  act(() => { input.admission.setReadOnly(true); });
  expect(input.textarea.readOnly).toBe(true);
  fireEvent.keyDown(input.textarea, { key: 'ArrowRight' });
  fireEvent.keyUp(input.textarea, { key: 'ArrowRight' });
  fireEvent.keyDown(input.textarea, { key: 'ArrowRight', shiftKey: true });
  fireEvent.keyUp(input.textarea, { key: 'ArrowRight', shiftKey: true });
  beforeinput(input.textarea, 'insertText', 'blocked');
  fireEvent.keyDown(input.textarea, { key: 'Backspace' });
  fireEvent.keyUp(input.textarea, { key: 'Backspace' });
  fireEvent.paste(input.textarea, { clipboardData: { getData: () => 'blocked' } });
  fireEvent.keyDown(input.textarea, { key: 'b', ctrlKey: true });
  expect(input.submitted).toEqual(['setReadOnly', 'navigate', 'navigate']);
  expect(input.entries[1]!.payload).toEqual({ unit: 'character', direction: 'forward', extend: false });
  expect(input.entries[2]!.payload).toEqual({ unit: 'character', direction: 'forward', extend: true });
  expect(input.entries[1]!.options.after).toBe(input.entries[0]!.ticket);
  expect(input.entries[2]!.options.after).toBe(input.entries[1]!.ticket);
  input.unmount();
  input.admission.dispose();
});

test('d2b_modifier_deletion_preserves_selected_ranges', async () => {
  const selection: DocxCommandSelection = {
    anchor: { story: 'body', paraId: 'worker-range', offset: 2 },
    focus: { story: 'body', paraId: 'worker-range', offset: 5 },
  };
  for (const modifier of ['ctrlKey', 'altKey', 'metaKey']) {
    for (const key of ['Backspace', 'Delete']) {
      const input = mount();
      input.snapshot.workerSelection = input.session.state.selection = selection;
      fireEvent.keyDown(input.textarea, { key, [modifier]: true });
      beforeinput(input.textarea, key === 'Backspace' ? 'deleteWordBackward' : 'deleteWordForward');
      fireEvent.input(input.textarea, { inputType: key === 'Backspace' ? 'deleteWordBackward' : 'deleteWordForward' });
      expect(input.submitted).toEqual([key === 'Backspace' ? 'deleteBackward' : 'deleteForward']);
      expect(input.snapshot.workerSelection).toBe(selection);
      input.unmount();
    }
  }
  for (const inputType of ['deleteWordBackward', 'deleteWordForward', 'deleteSoftLineBackward', 'deleteSoftLineForward']) {
    const input = mount();
    input.snapshot.workerSelection = input.session.state.selection = selection;
    beforeinput(input.textarea, inputType);
    expect(input.submitted).toEqual([inputType.endsWith('Backward') ? 'deleteBackward' : 'deleteForward']);
    input.unmount();
  }
  for (const selected of [false, true]) {
    const input = mount();
    input.snapshot.workerSelection = input.session.state.selection = {
      anchor: selection.anchor, focus: selection.anchor,
    };
    admitIntent(input.admission, { kind: 'insertText', payload: { text: 'x' } });
    fireEvent.keyDown(input.textarea, { key: 'Backspace', ctrlKey: true });
    fireEvent.keyUp(input.textarea, { key: 'Backspace' });
    fireEvent.keyDown(input.textarea, { key: 'ArrowRight' });
    expect(input.submitted).toEqual(['insertText']);
    input.entries[0]!.settle({ status: 'applied', seq: 1 }, selected ? selection : {
      anchor: selection.focus, focus: selection.focus,
    });
    await input.waitForSubmitted(selected ? 3 : 4);
    expect(input.submitted).toEqual(selected ? ['insertText', 'deleteBackward', 'navigate']
      : ['insertText', 'navigate', 'deleteBackward', 'navigate']);
    for (let i = 1; i < input.entries.length; i += 1) {
      expect(input.entries[i]!.options.after).toBe(input.entries[i - 1]!.ticket);
    }
    input.unmount();
  }
});

test('d2b_multiline_paste_forms_one_ordered_chain', () => {
  for (const source of ['paste', 'beforeinput', 'input']) {
    const input = mount();
    admitIntent(input.admission, { kind: 'insertText', payload: { text: 'x' } });
    const text = 'a\r\nb\n\nc\rd\u2028e\u2029f\u000B';
    if (source === 'paste') fireEvent.paste(input.textarea, { clipboardData: { getData: () => text } });
    else if (source === 'beforeinput') beforeinput(input.textarea, 'insertFromPaste', text);
    else fireEvent.input(input.textarea, { inputType: 'insertFromPaste', target: { value: text } });
    beforeinput(input.textarea, 'insertText', 'z');
    expect(input.submitted).toEqual([
      'insertText', 'replaceSelection', 'paragraphBreak', 'insertText', 'paragraphBreak',
      'paragraphBreak', 'insertText', 'paragraphBreak', 'insertText', 'paragraphBreak',
      'insertText', 'paragraphBreak', 'insertText', 'paragraphBreak', 'insertText',
    ]);
    expect(input.entries.flatMap(({ payload }) => 'text' in payload ? [payload.text] : [])).toEqual([
      'x', 'a', 'b', 'c', 'd', 'e', 'f', 'z',
    ]);
    for (let i = 1; i < input.entries.length; i += 1) {
      expect(input.entries[i]!.options.after).toBe(input.entries[i - 1]!.ticket);
    }
    expect(input.entries.slice(1, -1).every(({ options }) => options.fence)).toBe(true);
    input.unmount();
  }
});

test('d2b_terminal_failure_settles_admission_and_recovers_composition_text', async () => {
  const recovered: string[] = [];
  const phases: string[] = [];
  const input = mount(false, {
    onRecoverableText: (text) => recovered.push(text),
    onCompositionChange: (state) => { if (state) phases.push(state.phase); },
  });
  fireEvent.compositionStart(input.textarea);
  fireEvent.compositionUpdate(input.textarea, { data: '日本語' });
  fireEvent.compositionEnd(input.textarea, { data: '日本語' });
  expect(input.submitted).toEqual(['replaceSelection']);
  expect(input.admission.getSnapshot().originalUnackedIntents).toBe(1);
  await act(async () => { input.entries[0]!.reject(new Error('Terminal session failure')); });
  expect(input.admission.getSnapshot().originalUnackedIntents).toBe(0);
  expect(input.admission.getSnapshot().tail).toBeUndefined();
  expect(recovered).toEqual(['日本語']);
  expect(phases.at(-1)).toBe('refused');
  expect(input.submitted).toEqual(['replaceSelection']);
});

test('d2b_disposed_composition_recovers_text_and_removes_listeners', () => {
  const recovered: string[] = [];
  const input = mount(false, { onRecoverableText: (text) => recovered.push(text) });
  fireEvent.compositionStart(input.textarea);
  fireEvent.compositionUpdate(input.textarea, { data: 'draft' });
  const remove = spyOn(input.textarea, 'removeEventListener');
  input.admission.dispose();
  expect(() => input.unmount()).not.toThrow();
  expect(recovered).toEqual(['draft']);
  expect(remove.mock.calls.map(([type]) => type)).toEqual([
    'keydown', 'keyup', 'beforeinput', 'input', 'paste', 'copy', 'cut', 'compositionstart',
    'compositionupdate', 'compositionend', 'blur',
  ]);
  remove.mockRestore();
  expect(() => {
    fireEvent.keyDown(input.textarea, { key: 'Backspace' });
    beforeinput(input.textarea, 'insertText', 'x');
    fireEvent.compositionStart(input.textarea);
    fireEvent.compositionEnd(input.textarea, { data: 'draft' });
    fireEvent.blur(input.textarea);
  }).not.toThrow();
  expect(input.submitted).toEqual([]);
  expect(recovered).toEqual(['draft']);
});

test('copy and cut use the worker clipboard adapter instead of the empty input value', () => {
  const calls: boolean[] = [];
  const input = mount(false, { onClipboard: async (cut) => { calls.push(cut); } });
  const copy = new Event('copy', { bubbles: true, cancelable: true });
  const cut = new Event('cut', { bubbles: true, cancelable: true });
  act(() => { input.textarea.dispatchEvent(copy); input.textarea.dispatchEvent(cut); });
  expect(calls).toEqual([false, true]);
  expect(copy.defaultPrevented).toBe(true);
  expect(cut.defaultPrevented).toBe(true);
  expect(input.submitted).toEqual([]);
});

test('copy and cut hand the event clipboardData to the adapter while the event is dispatching', () => {
  const seen: Array<DataTransfer | null | undefined> = [];
  const input = mount(false, { onClipboard: async (_cut, data) => { seen.push(data); } });
  const clipboardData = { setData() {}, getData: () => '' } as unknown as DataTransfer;
  for (const type of ['copy', 'cut']) {
    const event = Object.assign(new Event(type, { bubbles: true, cancelable: true }), { clipboardData });
    act(() => { input.textarea.dispatchEvent(event); });
  }
  expect(seen).toEqual([clipboardData, clipboardData]);
});

test('worker input enables spelling and follows acknowledged caret geometry', () => {
  const input = mount();
  expect(input.textarea.getAttribute('spellcheck')).toBe('true');
  act(() => { input.viewStore.installFrameGeometry({ identity: input.snapshot.identity, coversThroughSeq: 0,
    frameEpoch: 1, layoutRevision: 1, pages: [], selectionRects: [], caret: { x: 32, y: 72, width: 1, height: 18 } }); });
  input.refresh();
  expect(input.textarea.style.left).toBe('32px');
  expect(input.textarea.style.top).toBe('72px');
  fireEvent.compositionStart(input.textarea);
  fireEvent.input(input.textarea, { inputType: 'insertCompositionText', isComposing: true, target: { value: 'かな' } });
  act(() => { input.viewStore.installFrameGeometry({ ...input.snapshot.geometry!, frameEpoch: 2,
    caret: { x: 90, y: 120, width: 1, height: 18 } }); });
  input.refresh();
  expect(input.textarea.value).toBe('かな');
  expect(input.textarea.style.left).toBe('32px');
});

test('reads wait for active composition to settle before flushing accepted input', async () => {
  const input = mount();
  fireEvent.compositionStart(input.textarea);
  let settled = false;
  const wait = input.admission.whenCompositionSettled!().then(() => { settled = true; });
  await Promise.resolve();
  expect(settled).toBe(false);
  fireEvent.compositionEnd(input.textarea, { data: '語' });
  await wait;
  expect(settled).toBe(true);
  expect(input.submitted).toEqual(['replaceSelection']);
});

test('readonly_table_tab_navigates_without_requesting_a_new_row', () => {
  const input = mount(true);
  input.snapshot.editState!.tableContext = {
    story: 'body', tableIndex: 0, rowIndex: 0, columnIndex: 0, rowCount: 1, columnCount: 2,
    hasMultiCellSelection: false, canSplitCell: false,
  };
  fireEvent.keyDown(input.textarea, { key: 'Tab', code: 'Tab' });
  expect(input.submitted).toEqual(['navigateCell']);
  expect(input.entries[0]!.payload).toEqual({ direction: 'forward', extend: false, createLastRow: false });
  input.unmount();
});

test('history keyboard and browser events share admission without duplicate commands', () => {
  const input = mount();
  fireEvent.keyDown(input.textarea, { key: 'z', ctrlKey: true });
  beforeinput(input.textarea, 'historyUndo');
  fireEvent.keyUp(input.textarea, { key: 'z' });
  fireEvent.keyDown(input.textarea, { key: 'z', metaKey: true, shiftKey: true });
  beforeinput(input.textarea, 'historyRedo');
  fireEvent.keyUp(input.textarea, { key: 'z' });
  fireEvent.keyDown(input.textarea, { key: 'y', ctrlKey: true });
  fireEvent.keyUp(input.textarea, { key: 'y' });
  beforeinput(input.textarea, 'historyUndo');
  expect(input.submitted).toEqual(['undo', 'redo', 'redo', 'undo']);
});

test('table Tab delegates to the shell controller and leaves modified or composing Tab alone', () => {
  const directions: string[] = [];
  const input = mount(false, { onNavigateCell: direction => { directions.push(direction); } });
  input.snapshot.editState!.tableContext = { story: 'body', tableIndex: 0, rowIndex: 0, columnIndex: 0,
    rowCount: 1, columnCount: 2, hasMultiCellSelection: false, canSplitCell: false };
  fireEvent.keyDown(input.textarea, { key: 'Tab' });
  fireEvent.keyDown(input.textarea, { key: 'Tab', shiftKey: true });
  fireEvent.keyDown(input.textarea, { key: 'Tab', ctrlKey: true });
  fireEvent.compositionStart(input.textarea);
  fireEvent.keyDown(input.textarea, { key: 'Tab', isComposing: true });
  expect(directions).toEqual(['forward', 'backward']);
  expect(input.submitted).toEqual([]);
});

function tableFragment(): DocxClipboardFragment {
  return {
    paragraphs: [{
      tables: [{
        tblPr: { width: 9360, widthType: 'dxa' },
        grid: [],
        rows: [{ trPr: {}, cells: [{ tcPr: {}, content: [{ kind: 'text', text: 'Cell' }, { kind: 'pilcrow' }] }] }],
      }],
      text: '',
    }],
  };
}

function pasteWith(input: { textarea: HTMLTextAreaElement }, html: string, text: string) {
  fireEvent.paste(input.textarea, { clipboardData: { getData: (type: string) => type === 'text/html' ? html : text } });
}

test('paste_html_table_calls_onPasteFragment_and_suppresses_text_intents', () => {
  const seen: Array<{ fragment: DocxClipboardFragment; inputTime: number }> = [];
  const input = mount(false, { onPasteFragment: (fragment, inputTime) => { seen.push({ fragment, inputTime }); return true; } });
  try {
    const fragment = tableFragment();
    pasteWith(input, clipboardFragmentHtml(fragment), 'fallback');
    expect(seen).toHaveLength(1);
    expect(seen[0]!.fragment).toEqual(fragment);
    expect(input.submitted).toEqual([]);
  } finally {
    input.unmount();
    input.admission.dispose();
  }
});

test('paste_text_only_html_keeps_text_intents', () => {
  let calls = 0;
  const input = mount(false, { onPasteFragment: () => { calls += 1; return true; } });
  try {
    pasteWith(input, '<html><body><p>Hello</p></body></html>', 'Hello');
    expect(calls).toBe(0);
    expect(input.submitted).toEqual(['replaceSelection']);
    expect(input.entries.map(({ payload }) => payload)).toEqual([{ text: 'Hello' }]);
  } finally {
    input.unmount();
    input.admission.dispose();
  }
});

test('paste_fragment_callback_false_falls_back_to_text', () => {
  let calls = 0;
  const input = mount(false, { onPasteFragment: () => { calls += 1; return false; } });
  try {
    pasteWith(input, clipboardFragmentHtml(tableFragment()), 'fallback');
    expect(calls).toBe(1);
    expect(input.submitted).toEqual(['replaceSelection']);
    expect(input.entries.map(({ payload }) => payload)).toEqual([{ text: 'fallback' }]);
  } finally {
    input.unmount();
    input.admission.dispose();
  }
});
