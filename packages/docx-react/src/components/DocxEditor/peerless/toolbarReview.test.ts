import { expect, mock, test } from 'bun:test';
import type { ResidentReviewSnapshot } from '@betteroffice/docx/yrs';
import { toolbarFixture } from './toolbarCommands.testing';

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
  const readAt = mock(async (request: { kind: string }) => ({ status: 'ok', version: 'native', value: request.kind === 'sidebar' ? sidebar : review }));
  const view = fixture.viewStore.getSnapshot();
  fixture.viewStore.acceptWorkerState(view.identity, { ...view, epoch: 1, selection: fixture.selection,
    editState: { ...view.editState!, documentVersion: 'native' } });
  fixture.controller.update({ read: fixture.read as never, readAt: readAt as never, focus() {}, zoom: 1 });
  const feed = async () => {
    const reviewRead = await fixture.controller.query({ kind: 'reviewSnapshot' });
    fixture.controller.revisions(reviewRead, await fixture.controller.query({ kind: 'sidebar', commentIds: [], expectVersion: reviewRead.version }));
  };
  return { ...fixture, feed };
}

test('worker review commands target the tracked change at the selection', async () => {
  const { controller, submit, parity, viewStore, feed } = reviewFixture();
  expect(controller.commands.store.getState('reviewAccept')).toMatchObject({ enabled: false, disabledReason: { code: 'revision-required' } });
  await feed();
  const env = controller.environment();
  expect(env.revisions).toEqual([{ revisionId: '7', position: 4 }, { revisionId: '9', position: 20 }]);
  expect([...env.revisionIds].sort()).toEqual(['7', '9']);
  expect(env.currentRevisionId).toBe('7');
  parity('reviewAccept');
  parity('reviewReject', { revisionId: '9' });
  expect(controller.commands.store.getState('reviewAccept')).toMatchObject({ enabled: true });
  expect(controller.commands.store.getState('reviewReject', { revisionId: '11' }))
    .toMatchObject({ enabled: false, disabledReason: { code: 'revision-not-found' } });
  expect(await controller.commands.store.execute('reviewAccept')).toMatchObject({ ok: true });
  expect(submit).toHaveBeenLastCalledWith({ kind: 'acceptRevision', payload: { revisionId: '7', expectVersion: 'native' } }, { fence: true });
  expect(await controller.commands.store.execute('reviewReject', { revisionId: '9' })).toMatchObject({ ok: true });
  expect(submit).toHaveBeenLastCalledWith({ kind: 'rejectRevision', payload: { revisionId: '9', expectVersion: 'native' } }, { fence: true });
  const view = viewStore.getSnapshot();
  const outside = { anchor: { story: 'body', paraId: 'p1', offset: 7 }, focus: { story: 'body', paraId: 'p1', offset: 7 } };
  viewStore.acceptWorkerState(view.identity, { ...view, epoch: 1, selection: outside });
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
  expect([...controller.environment().revisionIds].sort()).toEqual(['7', '9']);
  controller.empty();
  expect(controller.trackedChanges()).toBeNull();
  expect(controller.environment().revisions).toEqual([]);
});
