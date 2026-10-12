import { expect, mock, test } from 'bun:test';
import { createPeerlessShellController, LEGACY_SYNC_API_PROP } from './shellController';
import { createPeerlessViewStore } from './viewStore';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import { PLAIN_CONTEXT } from '../../../commands/testing';
import { admitIntent, createInputAdmission } from './inputAdmission';
import type { DocxWorkerEditSession } from '@betteroffice/docx/peerless';
import type { PeerlessEchoIntent, PeerlessEchoPainter } from './echoPainter';
import type { CommandTicket } from '../../../../../../shared/office-session/commands/types';
import { normalizeInputEvent } from './inputEvents';

function fixture() {
  const viewStore = createPeerlessViewStore({
    identity: { session: 'worker', incarnation: 1, epoch: 1 }, version: 0,
    throughSeq: 0, presentedThroughSeq: 0, workerSelection: null, anchors: null,
    editState: { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
      canUndo: true, canRedo: false, hyperlink: null, tableContext: null }, geometry: null,
  });
  const flush = mock(async () => {});
  const binding = { viewStore, admission: {}, session: { flush } } as unknown as PeerlessEditSessionBinding;
  const read = mock(async () => ({ status: 'ok', version: '0', value: { version: '0', paragraphs: [] } }));
  const controller = createPeerlessShellController({ read: read as never, focus() {}, zoom: 1 });
  return { controller, binding, viewStore, flush, read };
}

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
