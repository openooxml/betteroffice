import { expect, mock, spyOn, test } from 'bun:test';
import { createPeerlessShellController, LEGACY_SYNC_API_PROP } from './shellController';
import { createPeerlessViewStore } from './viewStore';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import { PLAIN_CONTEXT } from '../../../commands/testing';
import { admitIntent, createInputAdmission } from './inputAdmission';
import type { DocxWorkerEditSession } from '@betteroffice/docx/peerless';
import type { PeerlessEchoIntent, PeerlessEchoPainter } from './echoPainter';
import type { CommandTicket } from '../../../../../../shared/office-session/commands/types';
import { normalizeInputEvent } from './inputEvents';
import { toolbarFixture } from './toolbarCommands.testing';

function fixture(presentedThroughSeq = 0) {
  const viewStore = createPeerlessViewStore({
    identity: { session: 'worker', incarnation: 1, epoch: 1 }, version: 0,
    throughSeq: 0, presentedThroughSeq, workerSelection: null, anchors: null,
    editState: { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
      canUndo: true, canRedo: false, hyperlink: null, tableContext: null }, geometry: null,
  });
  const flush = mock(async () => {});
  const binding = { viewStore, admission: {}, session: { flush } } as unknown as PeerlessEditSessionBinding;
  const read = mock(async () => ({ status: 'ok', version: '0', value: { version: '0', paragraphs: [] } }));
  const controller = createPeerlessShellController({ read: read as never, focus() {}, zoom: 1 });
  return { controller, binding, viewStore, flush, read };
}

test('font selection prepares measurement before formatting and rejects replacement', async () => {
  const { controller, binding, viewStore } = fixture();
  const state = viewStore.getSnapshot();
  viewStore.acceptWorkerState(state.identity, { ...state, epoch: 1, selection: null,
    editState: { ...state.editState!, toolbar: PLAIN_CONTEXT } });
  const order: string[] = [];
  const format = mock(() => { order.push('format'); return { status: 'submitted', ticket: { ack: Promise.resolve({ status: 'applied' }) } }; });
  const fonts = { faces: [], families: ['Times New Roman'], chains: { 'times new roman|0|0': [0] } };
  const submit = mock(() => { order.push('environment'); return { status: 'submitted', ticket: { ack: Promise.resolve({ status: 'noop' }) } }; });
  binding.commands = { format, submit } as never;
  const prepareFontFamily = mock(async (family: string): Promise<typeof fonts | void> => { order.push(family); return fonts; });
  controller.update({ read: async () => ({} as never), focus() {}, zoom: 1, prepareFontFamily });
  controller.attach(binding);
  expect(await controller.commands.store.execute('fontFamily', { family: 'Times New Roman' })).toMatchObject({ ok: true });
  expect(order).toEqual(['Times New Roman', 'environment', 'format']);
  expect(submit).toHaveBeenCalledWith({ kind: 'setFontEnvironment', payload: fonts }, { fence: true });
  prepareFontFamily.mockImplementation(async () => { controller.detach(binding); });
  expect(await controller.commands.store.execute('fontFamily', { family: 'Arial' })).toMatchObject({ ok: false });
  expect(format).toHaveBeenCalledTimes(1);
});

test('preview remains read-only until the command frame is presented', () => {
  const { controller, binding, viewStore } = fixture(-1);
  controller.opening();
  const preview = { frame: Uint8Array.of(1), documentPreview: true } as never;
  controller.preview(preview);
  expect(controller.getSnapshot()).toMatchObject({ status: 'opening', preview, previewPainted: false });
  controller.previewPresented(preview);
  expect(controller.getSnapshot().previewPainted).toBe(true);
  expect(controller.commands.store.getState('undo').enabled).toBe(false);
  const state = viewStore.getSnapshot();
  viewStore.acceptWorkerState(state.identity, { ...state, epoch: 1, presentedThroughSeq: -1, selection: null });
  controller.attach(binding);
  expect(controller.getSnapshot().preview).toBe(preview);
  viewStore.installFrameGeometry({ identity: state.identity, coversThroughSeq: 0, frameEpoch: 1, layoutRevision: 1, pages: [], caret: null, selectionRects: [] });
  expect(controller.getSnapshot()).toMatchObject({ status: 'ready', preview: null });
  controller.empty();
  controller.previewPresented(preview);
  expect(controller.getSnapshot()).toMatchObject({ status: 'empty', preview: null, previewPainted: false });
});

test('save includes navigation and typing deferred behind the first acknowledgement', async () => {
  const { controller, binding, viewStore } = fixture();
  const order: string[] = [];
  const tickets: CommandTicket[] = [];
  let release!: () => void;
  let text = '';
  const session = {
    state: { epoch: 1, version: 0, throughSeq: 0, presentedThroughSeq: 0, selection: null, editState: undefined },
    submit(kind, payload, options) {
      const seq = tickets.length + 1;
      const ack = (seq === 1 ? new Promise<void>((resolve) => { release = resolve; }) : Promise.resolve()).then(() => {
        order.push(kind);
        if (kind === 'insertText') text += (payload as { text: string }).text;
        session.state.throughSeq = seq;
        viewStore.installFrameGeometry({ identity: viewStore.getSnapshot().identity, coversThroughSeq: seq,
          frameEpoch: seq, layoutRevision: seq, pages: [], caret: { x: 12, y: 10, width: 1, height: 20 }, selectionRects: [] });
        return { status: 'applied' as const, seq };
      });
      if (options?.after) expect(options.after.seq).toBe(seq - 1);
      const ticket = { id: seq, seq, ack, presented: Promise.resolve('presented' as const) };
      tickets.push(ticket);
      return ticket;
    },
    flush: async () => { await Promise.all(tickets.map((ticket) => ticket.ack)); },
  } satisfies Pick<DocxWorkerEditSession<PeerlessEchoIntent>, 'state' | 'submit' | 'flush'>;
  binding.session = session as DocxWorkerEditSession<PeerlessEchoIntent>;
  binding.admission = createInputAdmission({ session: binding.session, viewStore,
    echoPainter: { disable() {}, preview: () => undefined } as unknown as PeerlessEchoPainter });
  controller.update({ read: async () => ({} as never), focus() {}, zoom: 1,
    save: async () => { order.push('save'); return new TextEncoder().encode(text).buffer; } });
  controller.attach(binding);
  admitIntent(binding.admission, { kind: 'insertText', payload: { text: 'x' } });
  admitIntent(binding.admission, normalizeInputEvent({ eventId: 1, inputTime: 0,
    source: 'keydown', key: 'ArrowUp', code: 'ArrowUp', shift: false, alt: false, ctrl: false, meta: false },
    { composing: false, compositionId: null, lastHandledEventId: null, trailingCompositionText: null }).intents[0]!);
  admitIntent(binding.admission, { kind: 'insertText', payload: { text: 'y' } });
  expect(tickets).toHaveLength(1);
  const saved = controller.ref.save();
  await Promise.resolve();
  release();
  expect(new TextDecoder().decode((await saved)!)).toBe('xy');
  expect(order).toEqual(['insertText', 'navigate', 'insertText', 'save']);
  binding.admission.dispose();
});

test('layout waits issued while the worker opens settle on the first complete layout, a failure or unmount', async () => {
  const { controller, binding, viewStore } = fixture(-1);
  const subscribe = viewStore.subscribe;
  let subscriptions = 0;
  viewStore.subscribe = (listener) => { subscriptions += 1; const stop = subscribe(listener); return () => { subscriptions -= 1; stop(); }; };
  const turns = () => new Promise((resolve) => setTimeout(resolve, 0));
  const page = (pageIndex: number) => ({ page: { pageId: String(pageIndex), pageIndex },
    bounds: { x: 0, y: pageIndex * 800, width: 600, height: 800 }, transform: [1, 0, 0, 1, 0, 0] as const });
  const identity = viewStore.getSnapshot().identity;
  controller.opening();
  let pages: number | Error | undefined;
  controller.ref.whenLayoutComplete().then((count) => { pages = count; }, (error: Error) => { pages = error; });
  await turns();
  expect(pages).toBeUndefined();
  controller.attach(binding);
  viewStore.installFrameGeometry({ identity, coversThroughSeq: 0, frameEpoch: 1, layoutRevision: 1, layoutProvisional: true,
    pages: [page(0)], caret: null, selectionRects: [] });
  expect(controller.getSnapshot().status).toBe('ready');
  await turns();
  expect(pages).toBeUndefined();
  viewStore.installFrameGeometry({ identity, coversThroughSeq: 0, frameEpoch: 2, layoutRevision: 2,
    pages: [page(0), page(1)], caret: null, selectionRects: [] });
  await turns();
  expect(pages).toBe(2);
  expect(subscriptions).toBe(1);

  controller.opening();
  const failed = controller.ref.whenLayoutComplete();
  controller.fail(new Error('Broken package'));
  await expect(failed).rejects.toThrow('Broken package');
  await expect(controller.ref.whenLayoutComplete()).rejects.toThrow('Broken package');

  controller.opening();
  const closed = controller.ref.whenLayoutComplete();
  await turns();
  controller.dispose();
  await expect(closed).rejects.toThrow('closed');
});

test('worker state opens shell command gates without a legacy document or session', () => {
  const { controller, binding } = fixture();
  expect(controller.commands.store.getState('undo').enabled).toBe(false);
  controller.attach(binding);
  expect(controller.getSnapshot().status).toBe('ready');
  expect(controller.commands.store.getState('undo').enabled).toBe(true);
  expect(controller.commands.store.getState('redo').enabled).toBe(false);
  controller.detach(binding);
  expect(controller.commands.store.getState('undo').enabled).toBe(false);
});

test('worker text color preserves theme shades and automatic clearing', async () => {
  const { controller, binding, viewStore } = fixture();
  const state = viewStore.getSnapshot();
  viewStore.acceptWorkerState(state.identity, { ...state, epoch: 1, selection: null,
    editState: { ...state.editState!, toolbar: PLAIN_CONTEXT } });
  const format = mock((_intent: unknown, _options: unknown) => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'applied' }) } }));
  binding.commands = { format } as never;
  controller.attach(binding);
  expect(await controller.commands.store.execute('textColor', { color: { themeColor: 'accent1', themeTint: '80' } })).toMatchObject({ ok: true });
  expect(await controller.commands.store.execute('textColor', { color: 'auto' })).toMatchObject({ ok: true });
  expect(format.mock.calls.map(([intent]) => intent)).toEqual([
    { kind: 'formatRange', payload: { delta: { color: { themeColor: 'accent1' }, other: { textColor: { themeColor: 'accent1', themeTint: '80' } } } } },
    { kind: 'formatRange', payload: { delta: { color: null, other: { textColor: null } } } },
  ]);
});

test('worker highlightColor formats and clears through formatRange', async () => {
  const { controller, format, parity } = toolbarFixture();
  parity('highlightColor');
  for (const color of ['yellow', 'none']) {
    expect(await controller.commands.store.execute('highlightColor', { color })).toMatchObject({ ok: true });
    expect(format).toHaveBeenLastCalledWith({ kind: 'formatRange', payload: { delta: { highlight: color === 'none' ? null : color } } }, { fence: true });
  }
});

test('worker insertBreak stays a loud sync-API failure', () => {
  const { controller } = toolbarFixture();
  expect(() => controller.ref.insertBreak({ paraId: 'p1', type: 'page' })).toThrow(LEGACY_SYNC_API_PROP);
});

test('font readiness updates choices and callbacks while screen paint waits for accepted geometry', async () => {
  const { controller, binding, viewStore } = fixture(-1);
  const onFontsLoaded = mock(() => {});
  const onFirstPagePainted = mock(() => {});
  const onMemoryPressure = mock(() => {});
  controller.update({ read: async () => ({} as never), focus() {}, zoom: 1, onFontsLoaded, onFirstPagePainted, onMemoryPressure,
    fontFamilies: ['Host face'], memoryBudget: { warningBytes: 10, criticalBytes: 20 } });
  controller.opening();
  controller.fontEnvironment({ phase: 'ready', revision: 1, families: ['Custom face'] });
  expect(onFontsLoaded).toHaveBeenCalledTimes(1);
  expect(onFirstPagePainted).not.toHaveBeenCalled();
  expect(controller.environment().fonts.map(option => option.label)).toEqual(['Host face', 'Custom face']);
  controller.fontEnvironment({ phase: 'presentation', revision: 1, families: ['Custom face'], frameEpoch: 2 });
  expect(onFirstPagePainted).not.toHaveBeenCalled();
  controller.memory([{ label: 'docx-edit', bufferBytes: 30, liveBytes: 12 }]);
  expect(controller.ref.getMemoryStats()).toMatchObject({ main: [], worker: [{ label: 'docx-edit', liveBytes: 12 }] });
  expect(onMemoryPressure).toHaveBeenCalledWith(expect.objectContaining({ level: 'warning' }));
  controller.attach(binding);
  viewStore.installFrameGeometry({ identity: viewStore.getSnapshot().identity, coversThroughSeq: 0,
    frameEpoch: 2, layoutRevision: 2, pages: [{ page: { pageId: '1', pageIndex: 0 },
      bounds: { x: 0, y: 0, width: 600, height: 800 }, transform: [1, 0, 0, 1, 0, 0] }], caret: null, selectionRects: [] });
  expect(onFirstPagePainted).toHaveBeenCalledTimes(1);
  controller.fontEnvironment({ phase: 'ready', revision: 1, families: ['Custom face'] });
  expect(onFontsLoaded).toHaveBeenCalledTimes(1);
  const state = viewStore.getSnapshot();
  viewStore.acceptWorkerState(state.identity, { ...state, epoch: 1, selection: null,
    editState: { ...state.editState!, toolbar: PLAIN_CONTEXT } });
  const format = mock(() => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'applied' }) } }));
  binding.commands = { format } as never;
  expect(await controller.commands.store.execute('fontFamily', { family: 'Custom face' })).toMatchObject({ ok: true });
  expect(format).toHaveBeenCalledWith({ kind: 'formatRange', payload: { delta: { fontFamily: { ascii: 'Custom face', hAnsi: 'Custom face' } } } }, { fence: true });
});

test('sync model APIs fail loudly and async reads flush worker input', async () => {
  const { controller, binding, flush, read } = fixture();
  controller.attach(binding);
  for (const member of ['getDocument', 'getEditorRef', 'getSelectionInfo', 'findInDocument', 'getPageContent', 'onContentChange']) {
    expect(() => (controller.ref as unknown as Record<string, () => unknown>)[member]!()).toThrow(LEGACY_SYNC_API_PROP);
  }
  expect(await controller.ref.readParagraphs({} as never)).toEqual({ version: '0', paragraphs: [] } as never);
  expect(flush).toHaveBeenCalledTimes(1);
  expect(read).toHaveBeenCalledTimes(1);
});

test('replacement during a worker read rejects instead of returning the prior document', async () => {
  const { controller, binding, read } = fixture();
  let finish!: () => void;
  let begin!: () => void;
  const started = new Promise<void>((resolve) => { begin = resolve; });
  read.mockImplementation(async () => { await new Promise<void>((resolve) => { finish = resolve; begin(); }); return { status: 'ok', version: '0', value: { version: '0', paragraphs: [] } }; });
  controller.attach(binding);
  const pending = controller.ref.readParagraphs({} as never);
  await started;
  controller.detach(binding);
  finish();
  await expect(pending).rejects.toThrow('replaced');
});

test('worker selection and version events ignore geometry-only updates and unsubscribe independently', () => {
  const { controller, binding, viewStore } = fixture();
  const onDocumentChange = mock(() => {});
  const onSelectionChange = mock(() => {});
  controller.update({ read: async () => ({} as never), focus() {}, zoom: 1, onDocumentChange, onSelectionChange });
  controller.attach(binding);
  const documentListener = mock(() => {});
  const selectionListener = mock(() => {});
  controller.ref.onDocumentChange(documentListener);
  const remove = controller.ref.onSelectionChange(selectionListener);
  const selection = { anchor: { story: 'body', paraId: 'a', offset: 1 }, focus: { story: 'body', paraId: 'b', offset: 2 } };
  const identity = viewStore.getSnapshot().identity;
  viewStore.acceptWorkerState(identity, { ...viewStore.getSnapshot(), epoch: 1, version: 1, throughSeq: 1, selection });
  expect(onDocumentChange).toHaveBeenCalledWith({ version: '1' });
  expect(documentListener).toHaveBeenCalledWith({ version: '1' });
  expect(selectionListener).toHaveBeenCalledWith(expect.objectContaining({ hasSelection: true, isMultiParagraph: true,
    startParagraphIndex: -1, endParagraphIndex: -1 }));
  expect(onSelectionChange).toHaveBeenCalledTimes(1);
  viewStore.installFrameGeometry({ identity, coversThroughSeq: 1, frameEpoch: 1, layoutRevision: 1, pages: [], caret: null, selectionRects: [] });
  expect(documentListener).toHaveBeenCalledTimes(1);
  expect(selectionListener).toHaveBeenCalledTimes(1);
  remove();
  viewStore.acceptWorkerState(identity, { ...viewStore.getSnapshot(), epoch: 1, version: 1, throughSeq: 2, selection: null });
  expect(selectionListener).toHaveBeenCalledTimes(1);
  expect(onSelectionChange).toHaveBeenLastCalledWith(null);
});

test('a throwing change subscriber cannot block worker readiness or other subscribers', () => {
  const { controller, binding, viewStore } = fixture();
  const onError = mock(() => {});
  controller.update({ read: async () => ({} as never), focus() {}, zoom: 1, onError,
    onDocumentChange: () => { throw new Error('host callback failed'); } });
  controller.attach(binding);
  const listener = mock(() => {});
  controller.ref.onDocumentChange(listener);
  viewStore.acceptWorkerState(viewStore.getSnapshot().identity, { ...viewStore.getSnapshot(), epoch: 1, version: 1, throughSeq: 1,
    selection: null });
  expect(listener).toHaveBeenCalledWith({ version: '1' });
  expect(onError).toHaveBeenCalledTimes(1);
  expect(controller.getSnapshot().status).toBe('ready');
});

test('shell history commands use the worker command binding and inspect refusals', async () => {
  const { controller, binding } = fixture();
  const submit = mock(() => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'applied' }) } }));
  binding.commands = { submit } as never;
  controller.attach(binding);
  expect(await controller.commands.store.execute('undo', null)).toMatchObject({ ok: true, status: 'executed' });
  expect(submit).toHaveBeenCalledWith({ kind: 'undo', payload: {} }, { fence: true });
  submit.mockImplementation(() => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'refused', reason: 'read-only' }) } }));
  expect(await controller.commands.store.execute('undo', null)).toMatchObject({ ok: false });
});

test('history availability follows authoritative worker state', () => {
  const { controller, binding, viewStore } = fixture();
  controller.attach(binding);
  const state = viewStore.getSnapshot();
  viewStore.acceptWorkerState(state.identity, { epoch: 1, version: 1, throughSeq: 1, presentedThroughSeq: 0, selection: null,
    editState: { ...state.editState!, canUndo: false, canRedo: true } });
  expect(controller.commands.store.getState('undo').enabled).toBe(false);
  expect(controller.commands.store.getState('redo').enabled).toBe(true);
});

test('toolbar marks enter the shared input tail before preceding acknowledgements', async () => {
  const { controller, binding, viewStore, flush } = fixture();
  const state = viewStore.getSnapshot();
  viewStore.acceptWorkerState(state.identity, { ...state, epoch: 1, selection: null,
    editState: { ...state.editState!, toolbar: PLAIN_CONTEXT } });
  const format = mock(() => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'applied' }) } }));
  binding.commands = { format } as never;
  flush.mockImplementation(() => new Promise<void>(() => {}));
  controller.attach(binding);
  for (const mark of ['bold', 'italic', 'underline'] as const) {
    expect(controller.commands.store.getState(mark).enabled).toBe(true);
    expect(await controller.commands.store.execute(mark, null)).toEqual({ ok: true, status: 'executed' });
    expect(format).toHaveBeenLastCalledWith({ kind: 'toggleMark', payload: { mark: { type: mark } } }, { fence: true });
  }
  expect(flush).not.toHaveBeenCalled();
  format.mockImplementation(() => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'refused', reason: 'read-only' }) } }));
  expect(await controller.commands.store.execute('bold', null)).toMatchObject({ ok: false, failure: { code: 'read-only' } });
});

test('table navigation uses the controller admission and creates a final row only when editing forward', () => {
  const { controller, binding } = fixture();
  const submit = mock((_intent: unknown) => ({ status: 'ignored', reason: 'empty-composition' }));
  binding.commands = { submit } as never;
  binding.admission.getSnapshot = () => ({ readOnly: false } as never);
  controller.attach(binding);
  controller.navigateCell('forward');
  controller.navigateCell('backward');
  expect(submit.mock.calls.map(([intent]) => intent)).toEqual([
    { kind: 'navigateCell', payload: { direction: 'forward', extend: false, createLastRow: true } },
    { kind: 'navigateCell', payload: { direction: 'backward', extend: false, createLastRow: false } },
  ]);
  controller.update({ read: async () => ({} as never), focus() {}, zoom: 1, readOnly: true });
  controller.navigateCell('forward');
  expect(submit.mock.calls.at(-1)![0]).toMatchObject({ payload: { createLastRow: false } });
});

test('table selection dispatches selectTableCells as fenced selection input', async () => {
  const { controller, binding, viewStore } = fixture();
  const state = viewStore.getSnapshot();
  viewStore.acceptWorkerState(state.identity, { ...state, epoch: 1, selection: null,
    editState: { ...state.editState!, tableContext: { story: 'body', tableIndex: 0, rowIndex: 0, columnIndex: 0,
      rowCount: 2, columnCount: 2, hasMultiCellSelection: false, canSplitCell: true } } });
  const submit = mock(() => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'noop' }) } }));
  binding.commands = { submit } as never;
  controller.attach(binding);
  expect(controller.environment().supportedCommands?.has('tableAction')).toBe(true);
  expect(controller.environment().tableActions).toEqual(new Set(['selectTable', 'selectRow', 'selectColumn',
    'addRowAbove', 'addRowBelow', 'addColumnLeft', 'addColumnRight', 'deleteRow', 'deleteColumn', 'deleteTable',
    'mergeCells', 'splitCell', 'cellFillColor', 'borderAll', 'borderOutside', 'borderInside', 'borderNone',
    'borderTop', 'borderBottom', 'borderLeft', 'borderRight', 'borderColor', 'borderWidth', 'cellBorder', 'applyTableStyle',
    'tableProperties', 'openTableProperties']));
  expect(await controller.commands.store.execute('tableAction', 'selectTable')).toMatchObject({ ok: true, status: 'noop' });
  expect(await controller.commands.store.execute('tableAction', 'selectRow')).toMatchObject({ ok: true });
  expect(await controller.commands.store.execute('tableAction', 'selectColumn')).toMatchObject({ ok: true });
  const calls = submit.mock.calls as unknown as Array<[unknown, unknown]>;
  expect(calls.map(([intent, options]) => ({ intent, options }))).toEqual([
    { intent: { kind: 'selectTableCells', payload: { target: 'table' } }, options: { source: 'selection', fence: true } },
    { intent: { kind: 'selectTableCells', payload: { target: 'row' } }, options: { source: 'selection', fence: true } },
    { intent: { kind: 'selectTableCells', payload: { target: 'column' } }, options: { source: 'selection', fence: true } },
  ]);
  expect(await controller.commands.store.execute('tableAction', 'deleteRow')).toMatchObject({ ok: true, status: 'noop' });
  expect(calls.at(-1)).toEqual([{ kind: 'deleteTableRow', payload: {} }, { fence: true }]);
});

function tableFixture(tableContext: Record<string, unknown> = {}) {
  const { controller, binding, viewStore, read } = fixture();
  const state = viewStore.getSnapshot();
  viewStore.acceptWorkerState(state.identity, { ...state, epoch: 1, selection: null,
    editState: { ...state.editState!, tableContext: { story: 'body', tableIndex: 0, rowIndex: 0, columnIndex: 0,
      rowCount: 3, columnCount: 3, hasMultiCellSelection: true, canSplitCell: true, rowSpan: 1, columnSpan: 1, ...tableContext } } });
  const submit = mock(() => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'applied' }) } }));
  binding.commands = { submit } as never;
  controller.attach(binding);
  return { controller, binding, viewStore, submit, read };
}

test('table structure actions submit the worker table commands without a selection', async () => {
  const { controller, submit } = tableFixture();
  expect(controller.environment().tableActions).toEqual(new Set(['selectTable', 'selectRow', 'selectColumn',
    'addRowAbove', 'addRowBelow', 'addColumnLeft', 'addColumnRight', 'deleteRow', 'deleteColumn', 'deleteTable',
    'mergeCells', 'splitCell', 'cellFillColor', 'borderAll', 'borderOutside', 'borderInside', 'borderNone',
    'borderTop', 'borderBottom', 'borderLeft', 'borderRight', 'borderColor', 'borderWidth', 'cellBorder', 'applyTableStyle',
    'tableProperties', 'openTableProperties']));
  const cases: Array<[unknown, unknown]> = [
    ['addRowAbove', { kind: 'insertTableRow', payload: { side: 'above' } }],
    ['addRowBelow', { kind: 'insertTableRow', payload: { side: 'below' } }],
    ['addColumnLeft', { kind: 'insertTableColumn', payload: { side: 'left' } }],
    ['addColumnRight', { kind: 'insertTableColumn', payload: { side: 'right' } }],
    ['deleteRow', { kind: 'deleteTableRow', payload: {} }],
    ['deleteColumn', { kind: 'deleteTableColumn', payload: {} }],
    ['deleteTable', { kind: 'deleteTable', payload: {} }],
    ['mergeCells', { kind: 'mergeTableCells', payload: {} }],
    [{ type: 'cellFillColor', color: '#FF0000' }, { kind: 'setTableCellShading', payload: { color: 'FF0000' } }],
    [{ type: 'cellFillColor', color: null }, { kind: 'setTableCellShading', payload: { color: null } }],
  ];
  for (const [action] of cases) {
    expect(await controller.commands.store.execute('tableAction', action as never)).toMatchObject({ ok: true, status: 'executed' });
  }
  expect((submit.mock.calls as unknown as Array<[unknown, unknown]>).map(([intent, options]) => ({ intent, options })))
    .toEqual(cases.map(([, intent]) => ({ intent, options: { fence: true } })));
  submit.mockImplementation(() => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'refused', reason: 'invalid-cell-target' }) } }));
  expect(await controller.commands.store.execute('tableAction', 'deleteRow')).toMatchObject({ ok: false, failure: { code: 'command-failed' } });
});

test('table border actions submit hybrid presets and pen state', async () => {
  const { controller, submit, viewStore } = tableFixture();
  const pen = { style: 'single', size: 4, color: { rgb: '000000' } };
  const all = (border: typeof pen) => ({ top: border, bottom: border, left: border, right: border, insideH: border, insideV: border });
  const cases: Array<[unknown, unknown]> = [
    ['borderAll', { kind: 'setTableBorders', payload: { borders: all(pen) } }],
    ['borderOutside', { kind: 'setTableBorders', payload: { borders: { top: pen, bottom: pen, left: pen, right: pen } } }],
    ['borderInside', { kind: 'setTableBorders', payload: { borders: { insideH: pen, insideV: pen } } }],
    ['borderNone', { kind: 'setTableBorders', payload: { borders: all({ style: 'none', size: 0, color: { rgb: '000000' } }) } }],
    ['borderTop', { kind: 'setTableBorders', payload: { borders: { top: pen } } }],
    ['borderBottom', { kind: 'setTableBorders', payload: { borders: { bottom: pen } } }],
    ['borderLeft', { kind: 'setTableBorders', payload: { borders: { left: pen } } }],
    ['borderRight', { kind: 'setTableBorders', payload: { borders: { right: pen } } }],
  ];
  for (const [action] of cases) {
    expect(await controller.commands.store.execute('tableAction', action as never)).toMatchObject({ ok: true, status: 'executed' });
  }
  expect((submit.mock.calls as unknown as Array<[unknown, unknown]>).map(([intent, options]) => ({ intent, options })))
    .toEqual(cases.map(([, intent]) => ({ intent, options: { fence: true } })));
  submit.mockClear();
  const red = { style: 'single', size: 4, color: { rgb: 'FF0000' } };
  expect(await controller.commands.store.execute('tableAction', { type: 'borderColor', color: '#FF0000' })).toMatchObject({ ok: true });
  expect((submit.mock.calls.at(-1) as unknown[])[0]).toEqual({ kind: 'setTableBorders', payload: { borders: all(red) } });
  const wide = { style: 'single', size: 12, color: { rgb: 'FF0000' } };
  expect(await controller.commands.store.execute('tableAction', { type: 'borderWidth', size: 12 })).toMatchObject({ ok: true });
  expect((submit.mock.calls.at(-1) as unknown[])[0]).toEqual({ kind: 'setTableBorders', payload: { borders: all(wide) } });
  expect(await controller.commands.store.execute('tableAction', 'borderTop')).toMatchObject({ ok: true });
  expect((submit.mock.calls.at(-1) as unknown[])[0]).toEqual({ kind: 'setTableBorders', payload: { borders: { top: wide } } });
  expect(await controller.commands.store.execute('tableAction', { type: 'cellBorder', side: 'left', style: 'double', size: 6, color: '#00FF00' })).toMatchObject({ ok: true });
  expect((submit.mock.calls.at(-1) as unknown[])[0]).toEqual({ kind: 'setTableBorders',
    payload: { borders: { left: { style: 'double', size: 6, color: { rgb: '00FF00' } } } } });
  expect(await controller.commands.store.execute('tableAction', 'borderTop')).toMatchObject({ ok: true });
  expect((submit.mock.calls.at(-1) as unknown[])[0]).toEqual({ kind: 'setTableBorders', payload: { borders: { top: wide } } });
  const grid = { style: 'single', size: 4, color: { rgb: '000000' } };
  expect(await controller.commands.store.execute('tableAction', { type: 'applyTableStyle', styleId: 'TableGrid' })).toMatchObject({ ok: true });
  expect((submit.mock.calls.at(-1) as unknown[])[0]).toEqual({ kind: 'setTableBorders', payload: { borders: all(grid) } });
  submit.mockClear();
  expect(await controller.commands.store.execute('tableAction', { type: 'applyTableStyle', styleId: 'TableNormal' })).toMatchObject({ ok: true, status: 'noop' });
  expect(await controller.commands.store.execute('tableAction', { type: 'applyTableStyle', styleId: 'MissingStyle' })).toMatchObject({ ok: true, status: 'noop' });
  expect(submit).not.toHaveBeenCalled();
  const current = viewStore.getSnapshot();
  viewStore.acceptWorkerState(current.identity, { ...current, epoch: 1, selection: null,
    editState: { ...current.editState!, tableContext: { ...current.editState!.tableContext!,
      cellBorderColor: { rgb: '112233' } } } });
  const synced = { style: 'single', size: 12, color: { rgb: '112233' } };
  expect(await controller.commands.store.execute('tableAction', { type: 'borderWidth', size: 12 })).toMatchObject({ ok: true });
  expect((submit.mock.calls.at(-1) as unknown[])[0]).toEqual({ kind: 'setTableBorders', payload: { borders: all(synced) } });
  expect(await controller.commands.store.execute('tableAction', { type: 'borderColor', color: '#FF0000' })).toMatchObject({ ok: true });
  expect(await controller.commands.store.execute('tableAction', 'borderTop')).toMatchObject({ ok: true });
  expect((submit.mock.calls.at(-1) as unknown[])[0]).toEqual({ kind: 'setTableBorders',
    payload: { borders: { top: { style: 'single', size: 12, color: { rgb: 'FF0000' } } } } });
});

test('worker route refuses the eight hybrid-unsupported table actions', async () => {
  const { controller, submit } = tableFixture();
  const actions = [
    { type: 'cellVerticalAlign', align: 'center' }, { type: 'cellMargins', margins: { top: 100 } },
    { type: 'cellTextDirection', direction: 'btLr' }, { type: 'toggleNoWrap' },
    { type: 'rowHeight', height: 400, rule: 'exact' }, { type: 'toggleHeaderRow' },
    { type: 'distributeColumns' }, { type: 'autoFitContents' },
  ] as const;
  for (const action of actions) {
    expect(controller.commands.store.getState('tableAction', action as never))
      .toMatchObject({ enabled: false, disabledReason: { code: 'unsupported-command' } });
    expect(await controller.commands.store.execute('tableAction', action as never)).toMatchObject({ ok: false });
  }
  expect(submit).not.toHaveBeenCalled();
});

test('table structure gates follow last row and column, merge selection and suggesting mode', async () => {
  const single = tableFixture({ rowCount: 1, columnCount: 1, hasMultiCellSelection: false });
  expect(single.controller.commands.store.getState('tableAction', 'deleteRow'))
    .toMatchObject({ enabled: false, disabledReason: { code: 'last-row' } });
  expect(single.controller.commands.store.getState('tableAction', 'deleteColumn'))
    .toMatchObject({ enabled: false, disabledReason: { code: 'last-column' } });
  expect(single.controller.commands.store.getState('tableAction', 'mergeCells'))
    .toMatchObject({ enabled: false, disabledReason: { code: 'multiple-cells-required' } });

  const { controller } = tableFixture();
  expect(await controller.commands.store.execute('editingMode', { mode: 'suggesting' })).toMatchObject({ ok: true });
  for (const action of ['addRowAbove', 'addRowBelow', 'deleteRow'] as const) {
    expect(controller.commands.store.getState('tableAction', action)).toMatchObject({ enabled: true });
  }
  for (const action of ['addColumnLeft', 'addColumnRight', 'deleteColumn', 'deleteTable', 'mergeCells', 'splitCell'] as const) {
    expect(controller.commands.store.getState('tableAction', action))
      .toMatchObject({ enabled: false, disabledReason: { code: 'suggesting-unsupported' } });
  }
  expect(controller.commands.store.getState('tableAction', { type: 'cellFillColor', color: 'FF0000' }))
    .toMatchObject({ enabled: false, disabledReason: { code: 'suggesting-unsupported' } });
});

test('splitCell opens the dialog with span defaults and submits splitTableCell on apply', async () => {
  const { controller, submit } = tableFixture({ rowSpan: 1, columnSpan: 2 });
  let captured: { defaults: unknown; apply(rows: number, columns: number): Promise<unknown> } | undefined;
  const openSplitCell = mock((defaults: unknown, apply: (rows: number, columns: number) => Promise<unknown>) => {
    captured = { defaults, apply };
  });
  controller.update({ read: async () => ({} as never), focus() {}, zoom: 1, openSplitCell } as never);
  expect(await controller.commands.store.execute('tableAction', 'splitCell')).toMatchObject({ ok: true, status: 'opened' });
  expect(openSplitCell).toHaveBeenCalledTimes(1);
  expect(captured!.defaults).toEqual({ initialRows: 1, initialCols: 2, minRows: 1, minCols: 2 });
  expect(submit).not.toHaveBeenCalled();
  expect(await captured!.apply(1, 3)).toMatchObject({ ok: true, status: 'executed' });
  expect((submit.mock.calls as unknown as Array<[unknown, unknown]>).map(([intent, options]) => ({ intent, options })))
    .toEqual([{ intent: { kind: 'splitTableCell', payload: { rows: 1, columns: 3 } }, options: { fence: true } }]);
});

test('table properties open waits for admitted input', async () => {
  const { controller, binding, submit, read, viewStore } = tableFixture();
  const view = viewStore.getSnapshot();
  viewStore.acceptWorkerState(view.identity, { ...view, epoch: 1, selection: { anchor: { story: 'body', paraId: 'p1', offset: 0 }, focus: { story: 'body', paraId: 'p1', offset: 0 } } });
  const order: string[] = [];
  let releaseTyping!: () => void;
  const typing = new Promise<void>((resolve) => { releaseTyping = resolve; });
  let submissions = 0;
  const session = {
    state: { epoch: 1, version: 0, throughSeq: 0, presentedThroughSeq: 0, selection: null, editState: undefined },
    submit(kind: string) {
      order.push(kind);
      submissions += 1;
      const seq = submissions;
      return { id: seq, seq,
        ack: (seq === 1 ? typing : Promise.resolve()).then(() => ({ status: 'applied' as const, seq })),
        presented: Promise.resolve('presented' as const) };
    },
    flush: async () => {},
  };
  binding.session = session as never;
  binding.admission = createInputAdmission({ session: session as never, viewStore,
    echoPainter: { disable() {}, preview: () => undefined } as never });
  admitIntent(binding.admission, { kind: 'insertText', payload: { text: 'x' } });
  const composition = binding.admission.reserveComposition();
  admitIntent(binding.admission, { kind: 'insertText', payload: { text: 'y' } });
  let captured: { initial: unknown; apply(props: unknown): Promise<unknown> } | undefined;
  const openTableProperties = mock((initial: unknown, apply: (props: unknown) => Promise<unknown>) => {
    captured = { initial, apply };
  });
  read.mockImplementation(async () => { order.push('read'); return { status: 'ok', version: 'native', value: { version: 'native', paragraphs: [] } }; });
  const readAt = mock(async () => { order.push('readAt'); return { status: 'ok', version: 'native',
    value: { kind: 'table', properties: { width: 2500, widthType: 'pct', justification: 'center' } } }; });
  controller.update({ read: read as never, readAt: readAt as never, focus() {}, zoom: 1, openTableProperties } as never);
  const opening = controller.commands.store.execute('tableAction', { type: 'openTableProperties' });
  await new Promise((resolve) => setTimeout(resolve, 0));
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(openTableProperties).not.toHaveBeenCalled();
  expect(readAt).not.toHaveBeenCalled();
  releaseTyping();
  binding.admission.settleComposition(composition.slot, { text: 'z', target: composition.target });
  expect(await opening).toMatchObject({ ok: true, status: 'opened' });
  expect(order).toEqual(['insertText', 'replaceSelection', 'insertText', 'read', 'readAt']);
  expect(openTableProperties).toHaveBeenCalledTimes(1);
  expect(captured!.initial).toEqual({ width: 2500, widthType: 'pct', justification: 'center' });
  expect(readAt).toHaveBeenCalledWith({ kind: 'properties', target: { kind: 'tableSelection', selection: { anchor: { story: 'body', paraId: 'p1', offset: 0 }, focus: { story: 'body', paraId: 'p1', offset: 0 } } } }, 'native');
  expect(submit).not.toHaveBeenCalled();
  expect(await captured!.apply({ width: null, widthType: 'auto', justification: 'left' })).toMatchObject({ ok: true });
  expect((submit.mock.calls as unknown as Array<[unknown, unknown]>).map(([intent, options]) => ({ intent, options }))).toEqual([{
    intent: { kind: 'setTableProperties', payload: { op: 'patch', patch: { width: null, widthType: 'auto', justification: 'left' }, expectVersion: 'native', selection: { anchor: { story: 'body', paraId: 'p1', offset: 0 }, focus: { story: 'body', paraId: 'p1', offset: 0 } } } },
    options: { fence: true },
  }]);
  binding.admission.dispose();
});

test('table properties apply rejects changed origin', async () => {
  const { controller, binding, submit, read, viewStore } = tableFixture();
  const view = viewStore.getSnapshot();
  const at = (offset: number) => ({ anchor: { story: 'body', paraId: 'p1', offset }, focus: { story: 'body', paraId: 'p1', offset } });
  const moveSelection = (offset: number) => {
    const current = viewStore.getSnapshot();
    viewStore.acceptWorkerState(current.identity, { ...current, epoch: 1, selection: at(offset) });
  };
  viewStore.acceptWorkerState(view.identity, { ...view, epoch: 1, selection: at(0) });
  let captured: { apply(props: unknown): Promise<unknown> } | undefined;
  const openTableProperties = mock((_initial: unknown, apply: (props: unknown) => Promise<unknown>) => { captured = { apply }; });
  read.mockImplementation(async () => ({ status: 'ok', version: 'native', value: { version: 'native', paragraphs: [] } }));
  const readAt = mock(async () => ({ status: 'ok', version: 'native',
    value: { kind: 'table', properties: { width: 2500, widthType: 'pct', justification: 'center' } } }));
  controller.update({ read: read as never, readAt: readAt as never, focus() {}, zoom: 1, openTableProperties } as never);
  expect(await controller.commands.store.execute('tableAction', { type: 'openTableProperties' })).toMatchObject({ ok: true, status: 'opened' });
  moveSelection(1);
  expect(await captured!.apply({ width: 100, widthType: 'dxa', justification: 'right' }))
    .toMatchObject({ ok: false, failure: { code: 'target-changed' } });
  expect(submit).not.toHaveBeenCalled();
  moveSelection(0);
  readAt.mockImplementation(async () => {
    moveSelection(1);
    return { status: 'ok', version: 'native',
      value: { kind: 'table', properties: { width: 2500, widthType: 'pct', justification: 'center' } } };
  });
  expect(await controller.commands.store.execute('tableAction', { type: 'openTableProperties' }))
    .toMatchObject({ ok: false, failure: { code: 'target-changed' } });
  expect(openTableProperties).toHaveBeenCalledTimes(1);
  expect(submit).not.toHaveBeenCalled();
  moveSelection(0);
  readAt.mockImplementation(async () => ({ status: 'ok', version: 'native',
    value: { kind: 'table', properties: { width: 2500, widthType: 'pct', justification: 'center' } } }));
  expect(await controller.commands.store.execute('tableAction', { type: 'openTableProperties' })).toMatchObject({ ok: true, status: 'opened' });
  controller.detach(binding);
  controller.attach({ ...binding } as never);
  expect(await captured!.apply({ width: 100, widthType: 'dxa', justification: 'right' }))
    .toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
  expect(submit).not.toHaveBeenCalled();
});

test('internal modes remain switchable after viewing and update mutation author context', async () => {
  const { controller, binding } = fixture();
  const author = mock((_name: string | null) => {});
  binding.admission.setSuggesting = author;
  controller.attach(binding);
  expect(await controller.commands.store.execute('editingMode', { mode: 'suggesting' })).toMatchObject({ ok: true });
  expect(controller.environment().mode).toBe('suggesting');
  expect(author).toHaveBeenLastCalledWith('User');
  expect(await controller.commands.store.execute('editingMode', { mode: 'viewing' })).toMatchObject({ ok: true });
  expect(controller.commands.store.getState('undo').enabled).toBe(false);
  expect(controller.commands.store.getState('editingMode', { mode: 'editing' }).enabled).toBe(true);
  expect(await controller.commands.store.execute('editingMode', { mode: 'editing' })).toMatchObject({ ok: true });
  expect(author).toHaveBeenLastCalledWith(null);
});

test('controlled mode requests notify the host and preserve the committed prop', async () => {
  const { controller, binding } = fixture();
  const onModeChange = mock(() => {});
  controller.update({ read: async () => ({} as never), focus() {}, zoom: 1, mode: 'editing', onModeChange });
  controller.attach(binding);
  expect(await controller.commands.store.execute('editingMode', { mode: 'suggesting' })).toMatchObject({ ok: true, status: 'requested' });
  expect(onModeChange).toHaveBeenCalledWith('suggesting');
  expect(controller.environment().mode).toBe('editing');
});

test('queries pin the native version and refuse commands after replacement', async () => {
  const { controller, binding, read } = fixture();
  const readAt = mock(async () => ({ status: 'ok', version: 'native', value: { metadata: {}, geometry: [], frameEpoch: 1, layoutRevision: 1 } }));
  read.mockImplementation(async () => ({ status: 'ok', version: 'native', value: { version: 'native', paragraphs: [] } }));
  const submit = mock(() => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'noop' }) } }));
  binding.commands = { submit } as never;
  controller.update({ read: read as never, readAt: readAt as never, focus() {}, zoom: 1 });
  controller.attach(binding);
  const query = await controller.query({ kind: 'partSnapshot' });
  expect(readAt).toHaveBeenCalledWith({ kind: 'partSnapshot' }, 'native');
  expect(await controller.applyQuery(query, { kind: 'setWatermark', payload: { watermark: null, expectVersion: query.version } }))
    .toMatchObject({ ok: true, status: 'noop' });
  expect((submit.mock.calls[0] as unknown[])[0]).toMatchObject({ kind: 'setWatermark', payload: { watermark: null, expectVersion: 'native' } });
  controller.detach(binding);
  controller.attach({ ...binding } as PeerlessEditSessionBinding);
  expect(await controller.applyQuery(query, { kind: 'setWatermark', payload: { watermark: null, expectVersion: query.version } }))
    .toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
  expect(submit).toHaveBeenCalledTimes(1);
});

test('part queries pin the native version and refuse commands after replacement', async () => {
  const { controller, binding, read } = fixture();
  const readAt = mock(async () => ({ status: 'ok', version: 'native', value: { metadata: {}, geometry: [], frameEpoch: 1, layoutRevision: 1 } }));
  read.mockImplementation(async () => ({ status: 'ok', version: 'native', value: { version: 'native', paragraphs: [] } }));
  const submit = mock(() => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'noop' }) } }));
  binding.commands = { submit } as never;
  controller.update({ read: read as never, readAt: readAt as never, focus() {}, zoom: 1 });
  controller.attach(binding);
  const query = await controller.query({ kind: 'partSnapshot' });
  expect(readAt).toHaveBeenCalledWith({ kind: 'partSnapshot' }, 'native');
  expect(await controller.applyQuery(query, { kind: 'openPart', payload: { kind: 'footnote', noteId: 4, page: 1, section: 'final', expectVersion: query.version } }))
    .toMatchObject({ ok: true, status: 'noop' });
  expect((submit.mock.calls[0] as unknown[])[0]).toMatchObject({ kind: 'openPart', payload: { noteId: 4, expectVersion: 'native' } });
  controller.detach(binding);
  controller.attach({ ...binding } as PeerlessEditSessionBinding);
  expect(await controller.applyQuery(query, { kind: 'ensurePart', payload: { kind: 'header', page: 0, section: 'final', expectVersion: query.version } }))
    .toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
  expect(submit).toHaveBeenCalledTimes(1);
});

test('superseded part queries do not expose chrome metadata', async () => {
  const { controller, binding, read } = fixture();
  const readAt = mock(async () => ({ status: 'superseded' }));
  controller.update({ read: read as never, readAt: readAt as never, focus() {}, zoom: 1 });
  controller.attach(binding);
  await expect(controller.query({ kind: 'partSnapshot' })).rejects.toThrow('superseded');
});

test('control projection reads survive geometry refreshes in the same command epoch', async () => {
  const { controller, binding, viewStore, read } = fixture();
  read.mockImplementation(async () => ({ status: 'ok', version: 'native', value: { version: 'native', paragraphs: [] } }));
  const readAt = mock(async () => {
    const view = viewStore.getSnapshot();
    viewStore.acceptWorkerState({ ...view.identity }, { epoch: view.identity.epoch, version: view.version,
      throughSeq: view.throughSeq, presentedThroughSeq: view.presentedThroughSeq, selection: null, editState: { ...view.editState! } });
    return { status: 'ok', version: 'native', value: { version: 'native', frameEpoch: 2, layoutRevision: 1, controls: [] } };
  });
  controller.update({ read: read as never, readAt: readAt as never, focus() {}, zoom: 1 });
  controller.attach(binding);
  await expect(controller.query({ kind: 'contentControlHandles' })).resolves.toMatchObject({ version: 'native', value: { controls: [] } });
});

test('table gridlines toggle as view state and report it as active', async () => {
  const { controller, binding } = fixture();
  const onTableGridlines = mock((_visible: boolean) => {});
  controller.update({ read: async () => ({} as never), focus() {}, zoom: 1, onTableGridlines } as never);
  controller.attach(binding);
  expect(controller.commands.store.getState('tableGridlines')).toMatchObject({ enabled: true, active: false });
  expect(await controller.commands.store.execute('tableGridlines', null)).toMatchObject({ ok: true });
  expect(onTableGridlines.mock.calls).toEqual([[true]]);
  expect(controller.commands.store.getState('tableGridlines').active).toBe(true);
  expect(await controller.commands.store.execute('tableGridlines', null)).toMatchObject({ ok: true });
  expect(onTableGridlines.mock.calls).toEqual([[true], [false]]);
});

for (const lifecycle of ['empty', 'dispose'] as const) {
  test(`${lifecycle} releases unacknowledged batch commit callbacks`, async () => {
    const { controller, binding } = fixture();
    const applied = mock(() => {});
    let retained: Map<unknown, unknown> | undefined;
    const original = Map.prototype.set;
    const set = spyOn(Map.prototype, 'set').mockImplementation(function (this: Map<unknown, unknown>, key: unknown, value: unknown) {
      if ((value as { applied?: unknown } | null)?.applied === applied) retained = this;
      return original.call(this, key, value);
    });
    let submitted!: () => void;
    const started = new Promise<void>(resolve => { submitted = resolve; });
    binding.commands = { submit: () => {
      submitted();
      return { status: 'submitted', ticket: { seq: 1, ack: new Promise(() => {}) } };
    } } as never;
    try {
      controller.attach(binding);
      void controller.applyEdits({ expectVersion: '0', steps: [] }, () => null, applied);
      await started;
      expect(retained?.size).toBe(1);
      controller[lifecycle]();
      expect(retained?.size).toBe(0);
      expect(applied).not.toHaveBeenCalled();
    } finally { set.mockRestore(); controller.dispose(); }
  });
}

test('font readiness revisions restart with the recovered worker incarnation', () => {
  const { controller, binding, viewStore } = fixture();
  const onFontsLoaded = mock(() => {});
  controller.update({ read: async () => ({} as never), focus() {}, zoom: 1, onFontsLoaded });
  controller.attach(binding);
  const state = viewStore.getSnapshot();
  controller.fontEnvironment({ ...state.identity, phase: 'ready', revision: 8, families: ['Old face'] });
  const identity = { ...state.identity, incarnation: 2 };
  viewStore.acceptWorkerState(identity, { ...state, epoch: 1, selection: null });
  controller.fontEnvironment({ ...identity, phase: 'ready', revision: 1, families: ['Recovered face'] });
  controller.fontEnvironment({ ...identity, phase: 'ready', revision: 1, families: ['Recovered face'] });
  expect(onFontsLoaded).toHaveBeenCalledTimes(2);
});
