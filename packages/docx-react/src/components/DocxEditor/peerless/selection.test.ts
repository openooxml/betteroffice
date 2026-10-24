import { expect, test } from 'bun:test';
import { createDocxWorkerEditSession } from '@betteroffice/docx/worker';
import type {
  DocxCommandAck, DocxCommandEnvelope, DocxCommandSelection, DocxSelectionPointIntent, ResidentCommandConnection,
} from '@betteroffice/docx/worker';
import type { CommandTicket } from '../../../../../../shared/office-session/commands/types';
import { captureSelectionIntent } from './selection';
import type { PeerlessCommandBinding } from './commandBinding';
import type { PeerlessEchoIntent } from './echoPainter';
import { createPeerlessViewStore, type PeerlessFrameGeometry } from './viewStore';

const geometry: PeerlessFrameGeometry = {
  identity: { session: 'selection', incarnation: 1, epoch: 1 },
  coversThroughSeq: 0, frameEpoch: 7, layoutRevision: 4,
  pages: [{ page: { pageId: '10', pageIndex: 0 }, bounds: { x: 0, y: 0, width: 600, height: 800 }, transform: [1, 0, 0, 1, 0, 0] }],
  caret: null, selectionRects: [],
};

function point(): DocxSelectionPointIntent {
  return {
    page: { pageId: '10', pageIndex: 0 }, x: 32, y: 48,
    frameEpoch: 7, layoutRevision: 4, clickCount: 1,
    modifiers: { shift: false, alt: false, ctrl: false, meta: false },
    story: 'body', cellId: 'cell-a',
  };
}

function position(paraId: string, offset: number, byte: number) {
  return { story: 'body', paraId, offset, sticky: new Uint8Array([byte, byte + 1]) };
}

function orderedSelection() {
  const old = position('old', 2, 8);
  const initial: ResidentCommandConnection['initial'] = {
    ...geometry.identity, version: 0, throughSeq: 0, outcomes: [], selection: { anchor: old, focus: old },
    editState: { readOnly: false, toolbar: null, effectiveFormatting: null, storedMarks: null,
      canUndo: false, canRedo: false, hyperlink: null, tableContext: null },
    frame: { coversThroughSeq: 0, revision: 7, payload: {
      frame: new ArrayBuffer(0), caret: { frameEpoch: 7, caretRect: null }, layoutRevision: 4,
    } },
  };
  const selections = new Map<number, DocxCommandSelection>([[0, initial.selection!]]);
  const replies = new Map<number, { ack: DocxCommandAck; resolve(ack: DocxCommandAck): void }>();
  const received: { command: DocxCommandEnvelope; predecessor: DocxCommandSelection; selection: DocxCommandSelection }[] = [];
  let version = 0;
  const connection: ResidentCommandConnection = {
    initial, requestFullFrame: async () => initial,
    client: { failure: undefined, onFailure: () => () => {}, dispose: async () => {},
      call: {
        checkpoint: async () => ({ state: new Uint8Array(), epoch: 1, version, throughSeq: received.length }),
        edit(command) {
          const predecessor = selections.get(command.afterSeq ?? 0)!;
          expect(predecessor).toBeDefined();
          let selection: DocxCommandSelection;
          if (command.kind === 'insertText') {
            const at = position('post-edit', 3, 20);
            selection = { anchor: at, focus: at };
            version++;
          } else if (command.kind === 'selectPoint') {
            const at = position(predecessor.focus.paraId, 7, 30);
            selection = { anchor: at, focus: at };
          } else if (command.kind === 'extendSelection') {
            selection = { anchor: predecessor.anchor, focus: position('dragged', 6, 40) };
          } else if (command.kind === 'navigate') {
            selection = { anchor: predecessor.anchor, focus: position('navigated', 9, 50) };
          } else throw new Error('Unexpected command');
          selections.set(command.seq, selection);
          received.push({ command, predecessor, selection });
          const ack: DocxCommandAck = { ...initial, frame: undefined, version, throughSeq: command.seq,
            outcomes: [{ seq: command.seq, status: 'applied' }], selection };
          return new Promise<DocxCommandAck>((resolve) => { replies.set(command.seq, { ack, resolve }); });
        },
      },
    },
  };
  const viewStore = createPeerlessViewStore({ identity: geometry.identity, version: 0, throughSeq: 0,
    presentedThroughSeq: 0, workerSelection: initial.selection, anchors: null, editState: initial.editState, geometry });
  let echoPaints = 0;
  const session = createDocxWorkerEditSession<Uint8Array, PeerlessEchoIntent>({
    source: new Uint8Array([1]), connection, maxRecoveryAttempts: 0,
    recovery: { reopen: async () => { throw new Error('Unexpected recovery'); } },
    echo: { paint: () => { echoPaints++; }, drop() {}, present: () => true },
  });
  const unsubscribe = session.subscribe((state) => viewStore.acceptWorkerState(geometry.identity, state));
  let tail: CommandTicket | undefined;
  const commands: PeerlessCommandBinding = {
    submit(intent, options) {
      if (intent.kind === 'acceptRevision' || intent.kind === 'rejectRevision') throw new Error('Unexpected review intent');
      const echo: PeerlessEchoIntent | undefined = intent.kind === 'insertText'
        ? { intent, originalIntentCount: 1, contextFrameEpoch: 7 } : undefined;
      tail = session.submit(intent.kind, intent.payload, { ...options, after: tail, echo });
      return { status: 'submitted', ticket: tail, echoEnabled: echo !== undefined };
    },
    format: () => { throw new Error('Unexpected formatting'); },
  };
  const submit = (intent: Parameters<PeerlessCommandBinding['submit']>[0]) => {
    const result = commands.submit(intent, { fence: intent.kind !== 'insertText' });
    if (result.status !== 'submitted') throw new Error('Intent not submitted');
    return result.ticket;
  };
  return { received, selections, viewStore, submit, echoPaints: () => echoPaints,
    async ack(ticket: CommandTicket) {
      const reply = replies.get(ticket.seq!)!;
      reply.resolve(reply.ack);
      await ticket.ack;
      await Promise.resolve();
    },
    async dispose() { unsubscribe(); viewStore.dispose(); await session.dispose(); },
  };
}

test('d2a_click_during_unacked_echo_queues_coordinates', async () => {
  const run = orderedSelection();
  try {
    const insert = run.submit({ kind: 'insertText', payload: { text: 'edit' } });
    let acked = false;
    void insert.ack.then(() => { acked = true; });
    const captured = point();
    const click = run.submit(captureSelectionIntent({ kind: 'point', point: captured }, geometry));
    captured.x = 999;
    captured.page.pageId = 'echo-page';
    captured.modifiers.shift = true;
    await Promise.resolve();
    expect(acked).toBe(false);
    expect(run.echoPaints()).toBeGreaterThan(0);
    expect(run.received.map(({ command }) => command.kind)).toEqual(['insertText', 'selectPoint']);
    expect(run.received[1]!.command).toMatchObject({ seq: click.seq, afterSeq: insert.seq, payload: point() });
    expect(run.received[1]!.command.payload).not.toHaveProperty('selection');
    expect(run.received[1]!.predecessor.focus).toMatchObject({ paraId: 'post-edit', offset: 3 });
    expect(run.viewStore.getSnapshot().workerSelection?.focus.paraId).toBe('old');
    await run.ack(insert);
    await run.ack(click);
    expect(run.viewStore.getSnapshot().workerSelection?.focus).toEqual(position('post-edit', 7, 30));
    expect(run.viewStore.getSnapshot().anchors).toEqual({ anchor: new Uint8Array([30, 31]), focus: new Uint8Array([30, 31]) });
  } finally { await run.dispose(); }
});

test('d2a_drag_and_shift_navigation_keep_worker_anchors', async () => {
  const run = orderedSelection();
  try {
    const insert = run.submit({ kind: 'insertText', payload: { text: 'edit' } });
    const pointIntent = point();
    pointIntent.modifiers.shift = true;
    const dragIntent = captureSelectionIntent({ kind: 'drag', point: pointIntent }, geometry);
    const drag = run.submit(dragIntent);
    const navigation = run.submit(captureSelectionIntent({ kind: 'navigation',
      navigation: { unit: 'word', direction: 'forward', extend: true } }, null));
    await Promise.resolve();
    expect(run.received.map(({ command }) => [command.kind, command.afterSeq])).toEqual([
      ['insertText', undefined], ['extendSelection', insert.seq], ['navigate', drag.seq],
    ]);
    expect(run.received[1]!.command.payload).toEqual(pointIntent);
    expect(dragIntent.payload).not.toBe(pointIntent);
    expect(run.received[2]!.command.payload).toEqual({ unit: 'word', direction: 'forward', extend: true });
    expect(run.received.every(({ command }) => !('selection' in command.payload))).toBe(true);
    expect(run.received[1]!.predecessor.anchor).toEqual(position('post-edit', 3, 20));
    expect(run.received[2]!.predecessor).toEqual({ anchor: position('post-edit', 3, 20), focus: position('dragged', 6, 40) });
    await run.ack(insert);
    await run.ack(drag);
    expect(run.viewStore.getSnapshot().anchors).toEqual({ anchor: new Uint8Array([20, 21]), focus: new Uint8Array([40, 41]) });
    await run.ack(navigation);
    expect(run.viewStore.getSnapshot().workerSelection).toEqual({ anchor: position('post-edit', 3, 20), focus: position('navigated', 9, 50) });
    expect(run.viewStore.getSnapshot().anchors).toEqual({ anchor: new Uint8Array([20, 21]), focus: new Uint8Array([50, 51]) });
    run.selections.get(navigation.seq!)!.anchor.sticky![0] = 255;
    expect(run.viewStore.getSnapshot().anchors!.anchor).toEqual(new Uint8Array([20, 21]));
  } finally { await run.dispose(); }
});
