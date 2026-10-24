import { expect, mock, test } from 'bun:test';
import type { ResidentReviewSnapshot } from '@betteroffice/docx/yrs';
import type { DocxWorkerEditSession } from '@betteroffice/docx/worker';
import type { CommandTicket, TicketOutcome } from '../../../../../../shared/office-session/commands/types';
import { toolbarFixture } from './toolbarCommands.testing';
import { admitIntent, createInputAdmission } from './inputAdmission';
import type { PeerlessInputIntent } from './inputEvents';
import { createPeerlessCommandBinding } from './commandBinding';
import type { PeerlessEchoIntent, PeerlessEchoPainter } from './echoPainter';

const site = (revisionId: string, paraId: string, start: number, end: number) => ({ revisionId, kind: 'insertion', author: 'Editor',
  date: '2026-10-01T00:00:00Z', preview: 'Typed', range: { story: 'body', start: { paraId, offset: start }, end: { paraId, offset: end } } });
const entry = (revisionId: number, from: number) => ({ revisionId, type: 'insertion', text: 'Typed', author: 'Editor', from, to: from + 4 });

function reviewFixture() {
  const fixture = toolbarFixture();
  const review = { version: 'native', previewVersion: 0, layoutRevision: 1, frameEpoch: 1, comments: [],
    revisions: [site('9', 'p2', 0, 3), site('7', 'p1', 2, 6), site('11', 'p3', 0, 2)],
    proposals: { version: 'native', previewVersion: 0, proposals: [{ id: 'host', revisionIds: ['11'] }] },
    revisionGeometry: { 7: [{ pageIndex: 0, x: 1, y: 2, width: 3, height: 4 }] },
    proposalGeometry: {}, commentGeometry: {}, proposalHighlights: {} } as unknown as ResidentReviewSnapshot;
  const sidebar = { comments: [], revisions: [{ key: 'revision-9', start: { story: 'body', position: 20 } },
    { key: 'revision-7', start: { story: 'body', position: 4 } }, { key: 'revision-11', start: { story: 'body', position: 30 } }],
  trackedChanges: { entries: [entry(9, 20), entry(7, 4), entry(11, 30)], commentToRevision: [] } };
  const adjacent: { value: { revisionId: string } | null } = { value: null };
  const readAt = mock(async (request: { kind: string } & Record<string, unknown>) => ({ status: 'ok', version: 'native',
    value: request.kind === 'sidebar' ? sidebar : request.kind === 'adjacentRevision' ? adjacent.value : review }));
  const view = fixture.viewStore.getSnapshot();
  fixture.viewStore.acceptWorkerState(view.identity, { ...view, epoch: 1, selection: fixture.selection,
    editState: { ...view.editState!, documentVersion: 'native', currentRevisionId: '7' } });
  const reveal = mock(async (_target: { kind: 'revision'; revisionId: string } | { kind: 'comment'; id: number }) => true);
  fixture.controller.update({ read: fixture.read as never, readAt: readAt as never, focus() {}, zoom: 1, scrollToReviewTarget: reveal });
  const feed = async () => {
    const reviewRead = await fixture.controller.query({ kind: 'reviewSnapshot' });
    fixture.controller.revisions(reviewRead, await fixture.controller.query({ kind: 'sidebar', commentIds: [], expectVersion: reviewRead.version }));
  };
  return { ...fixture, feed, reveal, readAt, adjacent };
}

test('worker review commands target the tracked change the worker reports at the selection', async () => {
  const { controller, submit, parity, viewStore, feed } = reviewFixture();
  await feed();
  const env = controller.environment();
  expect(env.revisions).toEqual([{ revisionId: '7', position: 4 }, { revisionId: '9', position: 20 }]);
  expect([...env.revisionIds].sort()).toEqual(['11', '7', '9']);
  expect(env.currentRevisionId).toBe('7');
  parity('reviewAccept');
  parity('reviewReject', { revisionId: '9' });
  expect(controller.commands.store.getState('reviewAccept')).toMatchObject({ enabled: true });
  expect(controller.commands.store.getState('reviewReject', { revisionId: '12' }))
    .toMatchObject({ enabled: false, disabledReason: { code: 'revision-not-found' } });
  expect(await controller.commands.store.execute('reviewAccept', null)).toMatchObject({ ok: true });
  expect(submit).toHaveBeenLastCalledWith({ kind: 'acceptRevision', payload: {} }, { fence: true });
  expect(await controller.commands.store.execute('reviewReject', { revisionId: '9' })).toMatchObject({ ok: true });
  expect(submit).toHaveBeenLastCalledWith({ kind: 'rejectRevision', payload: { revisionId: '9' } }, { fence: true });
  const view = viewStore.getSnapshot();
  const outside = { anchor: { story: 'body', paraId: 'p1', offset: 7 }, focus: { story: 'body', paraId: 'p1', offset: 7 } };
  viewStore.acceptWorkerState(view.identity, { ...view, epoch: 1, selection: outside, editState: { ...view.editState!, currentRevisionId: undefined } });
  expect(controller.environment().currentRevisionId).toBeNull();
  expect(controller.commands.store.getState('reviewAccept')).toMatchObject({ enabled: false, disabledReason: { code: 'revision-required' } });
});

test('worker tracked changes skip host proposals and survive until a newer read replaces them', async () => {
  const { controller, viewStore, feed } = reviewFixture();
  await feed();
  const changes = controller.trackedChanges()!;
  expect(changes.entries.map(change => change.revisionId)).toEqual([9, 7]);
  expect([...changes.anchors.keys()]).toEqual(['revision-7']);
  const view = viewStore.getSnapshot();
  viewStore.acceptWorkerState(view.identity, { ...view, epoch: 1, version: view.version + 1, selection: null });
  controller.revisions(null, null);
  expect(controller.trackedChanges()).toBe(changes);
  expect([...controller.environment().revisionIds].sort()).toEqual(['11', '7', '9']);
  controller.empty();
  expect(controller.trackedChanges()).toBeNull();
  expect(controller.environment().revisions).toEqual([]);
});

test('worker next and previous reveal the change the worker finds from the selection', async () => {
  const { controller, viewStore, parity, reveal, feed, readAt, adjacent, selection } = reviewFixture();
  expect(controller.commands.store.getState('reviewNext')).toMatchObject({ enabled: false, disabledReason: { code: 'no-revisions' } });
  await feed();
  parity('reviewNext');
  parity('reviewPrevious');
  const step = async (id: 'reviewNext' | 'reviewPrevious', revisionId: string) => {
    adjacent.value = { revisionId };
    expect(await controller.commands.store.execute(id, null)).toMatchObject({ ok: true, status: 'executed' });
    expect(reveal.mock.calls.at(-1)![0]).toEqual({ kind: 'revision', revisionId });
    return readAt.mock.calls.at(-1)![0];
  };
  expect(await step('reviewNext', '9')).toEqual({ kind: 'adjacentRevision', selection, direction: 'next' });
  const view = viewStore.getSnapshot();
  const outside = { anchor: { story: 'body', paraId: 'p9', offset: 0 }, focus: { story: 'body', paraId: 'p9', offset: 0 } };
  viewStore.acceptWorkerState(view.identity, { ...view, epoch: 1, selection: outside });
  expect(await step('reviewPrevious', '7')).toEqual({ kind: 'adjacentRevision', selection: outside, direction: 'previous' });
  adjacent.value = null;
  const reveals = reveal.mock.calls.length;
  expect(await controller.commands.store.execute('reviewNext', null)).toMatchObject({ ok: false, failure: { code: 'command-failed' } });
  expect(reveal.mock.calls.length).toBe(reveals);
  adjacent.value = { revisionId: '9' };
  reveal.mockImplementationOnce(async () => false);
  expect(await controller.commands.store.execute('reviewNext', null)).toMatchObject({ ok: false, failure: { code: 'command-failed' } });
});

test('a queued review command settles as replaced instead of reaching the next document', async () => {
  const { controller, submit, feed } = reviewFixture();
  await feed();
  let release!: () => void;
  submit.mockImplementationOnce(() => ({ status: 'submitted', ticket: { ack: new Promise(resolve => { release = () => resolve({ status: 'applied' }); }) } }));
  const first = controller.commands.store.execute('reviewAccept', { revisionId: '7' });
  const second = controller.commands.store.execute('reviewAccept', { revisionId: '9' });
  while (!submit.mock.calls.length) await Promise.resolve();
  const next = toolbarFixture();
  controller.attach(next.binding);
  release();
  expect(await first).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
  expect(await second).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
  expect(next.submit).not.toHaveBeenCalled();
});

function workerAdmission({ binding, viewStore }: ReturnType<typeof reviewFixture>) {
  const applied: string[] = [];
  const decided = new Set<string>();
  const revisionAt: Record<string, string | undefined> = { p1: '7', p2: '9' };
  const tickets: CommandTicket[] = [];
  let release!: () => void;
  const gate = new Promise<void>(resolve => { release = resolve; });
  const session = {
    state: { epoch: 1, version: 0, throughSeq: 0, presentedThroughSeq: 0, selection: null, editState: viewStore.getSnapshot().editState },
    submit(kind, payload, options) {
      const seq = tickets.length + 1;
      const ack: Promise<TicketOutcome> = Promise.all([options?.after?.ack, seq === 1 ? gate : undefined]).then(([previous]) => {
        if (previous?.status === 'refused' || previous?.status === 'cancelled') return { seq, status: 'cancelled', reason: 'dependency-refused' };
        if ('expectVersion' in payload && payload.expectVersion !== session.state.editState?.documentVersion) return { seq, status: 'refused', reason: 'stale-version' };
        const revisionId = (payload as { revisionId?: string }).revisionId;
        if (revisionId !== undefined && decided.has(revisionId)) return { seq, status: 'refused', reason: 'unknown-revision' };
        if (revisionId !== undefined) decided.add(revisionId);
        applied.push(kind === 'insertText' ? (payload as { text: string }).text : revisionId ? `${kind} ${revisionId}` : kind);
        const editState = { ...session.state.editState!, documentVersion: `v${seq}`,
          ...(kind === 'select' ? { currentRevisionId: revisionAt[(payload as { selection: { anchor: { paraId: string } } }).selection.anchor.paraId] } : {}) };
        session.state = { ...session.state, throughSeq: seq, editState };
        viewStore.acceptWorkerState(viewStore.getSnapshot().identity, { ...session.state, selection: viewStore.getSnapshot().workerSelection });
        return { seq, status: 'applied' };
      });
      const ticket = { id: seq, seq, ack, presented: Promise.resolve('presented' as const) };
      tickets.push(ticket);
      return ticket;
    },
    flush: async () => { await Promise.all(tickets.map(ticket => ticket.ack)); },
  } satisfies Pick<DocxWorkerEditSession<PeerlessEchoIntent>, 'state' | 'submit' | 'flush'>;
  binding.session = session as unknown as DocxWorkerEditSession<PeerlessEchoIntent>;
  binding.admission = createInputAdmission({ session: binding.session, viewStore,
    echoPainter: { disable() {}, preview: () => undefined } as unknown as PeerlessEchoPainter });
  binding.commands = createPeerlessCommandBinding(binding.admission);
  const admit = (intent: PeerlessInputIntent) => { const admitted = admitIntent(binding.admission, intent); if (admitted.status !== 'submitted') throw new Error('not admitted'); return admitted.ticket.ack; };
  const type = (text: string) => admit({ kind: 'insertText', payload: { text } });
  const select = (paraId: string) => admit({ kind: 'select', payload: { selection: { anchor: { story: 'body', paraId, offset: 1 }, focus: { story: 'body', paraId, offset: 1 } } } });
  return { applied, release, type, select };
}

test('review commands take their admission place between keystrokes and expect the version their predecessor left', async () => {
  const fixture = reviewFixture();
  await fixture.feed();
  const { controller, binding } = fixture;
  const { applied, release, type } = workerAdmission(fixture);
  const keys = [type('a')];
  const first = controller.commands.store.execute('reviewAccept', { revisionId: '7' });
  keys.push(type('b'));
  const second = controller.commands.store.execute('reviewReject', { revisionId: '9' });
  keys.push(type('c'));
  release();
  expect(await first).toEqual({ ok: true, status: 'executed' });
  expect(await second).toEqual({ ok: true, status: 'executed' });
  expect((await Promise.all(keys)).map(outcome => outcome.status)).toEqual(['applied', 'applied', 'applied']);
  expect(applied).toEqual(['a', 'acceptRevision 7', 'b', 'rejectRevision 9', 'c']);
  binding.admission.dispose();
});

test('a review decision the engine refuses leaves later typing and decisions to run from the unchanged document', async () => {
  const fixture = reviewFixture();
  await fixture.feed();
  const { controller, binding } = fixture;
  const { applied, release, type } = workerAdmission(fixture);
  const first = controller.commands.store.execute('reviewAccept', { revisionId: '7' });
  const second = controller.commands.store.execute('reviewAccept', { revisionId: '7' });
  const typed = type('x');
  const rejected = controller.commands.store.execute('reviewReject', { revisionId: '9' });
  release();
  expect(await first).toEqual({ ok: true, status: 'executed' });
  expect(await second).toMatchObject({ ok: false, failure: { code: 'revision-not-found' } });
  expect((await typed).status).toBe('applied');
  expect(await rejected).toEqual({ ok: true, status: 'executed' });
  expect(applied).toEqual(['acceptRevision 7', 'x', 'rejectRevision 9']);
  binding.admission.dispose();
});

test('an implicit review decision targets the change at the selection its admitted predecessor leaves', async () => {
  const fixture = reviewFixture();
  await fixture.feed();
  const { controller, binding } = fixture;
  const { applied, release, type, select } = workerAdmission(fixture);
  const moves = [select('p2')];
  const rejected = controller.commands.store.execute('reviewReject', null);
  moves.push(select('p9'));
  const accepted = controller.commands.store.execute('reviewAccept', null);
  const typed = type('x');
  release();
  expect(await rejected).toEqual({ ok: true, status: 'executed' });
  expect(await accepted).toMatchObject({ ok: false, failure: { code: 'revision-required' } });
  expect((await Promise.all([...moves, typed])).map(outcome => outcome.status)).toEqual(['applied', 'applied', 'applied']);
  expect(applied).toEqual(['select', 'rejectRevision 9', 'select', 'x']);
  binding.admission.dispose();
});

test('host proposal revisions stay addressable by id and a vanished revision answers revision-not-found', async () => {
  const { controller, submit, feed } = reviewFixture();
  await feed();
  expect([...controller.environment().revisionIds].sort()).toEqual(['11', '7', '9']);
  expect(await controller.commands.store.execute('reviewReject', { revisionId: '11' })).toMatchObject({ ok: true });
  expect(submit.mock.lastCall![0]).toMatchObject({ kind: 'rejectRevision', payload: { revisionId: '11' } });
  submit.mockImplementationOnce(() => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'refused', reason: 'unknown-revision' }) } }) as never);
  expect(await controller.commands.store.execute('reviewAccept', { revisionId: '7' })).toMatchObject({ ok: false, failure: { code: 'revision-not-found' } });
});
