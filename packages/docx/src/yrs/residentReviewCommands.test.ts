import { afterEach, expect, test } from 'bun:test';
import { closeCommandEngines, command, execute, openCommandEngine, snapshot, TIMEOUT } from './residentCommandTestSupport';
import { createProposalRegistry } from './proposals';
import { readResidentReviewSnapshot } from './residentReview';
import { applyFrameDelta, decodeFrameDelta } from '../layout/render/frameDelta';

afterEach(closeCommandEngines);

test('review_refusals_leave_document_history_and_registry_unchanged', async () => {
  const arm = await openCommandEngine();
  arm.state.proposals = createProposalRegistry(arm.engine.proposalEngine);
  const range = { story: 'body', start: { paraId: arm.state.selection!.anchor.paraId, offset: 0 }, end: { paraId: arm.state.selection!.anchor.paraId, offset: 5 } };
  const version = arm.engine.proposalEngine.version();
  const before = snapshot(arm);
  expect(execute(arm, command(arm, 'addComment', { ranges: [range], author: 'Ada', date: '2026-01-01', body: 'Review', expectVersion: 'stale' }))).toMatchObject({ status: 'refused', reason: 'stale-version' });
  expect(snapshot(arm)).toEqual(before);
  expect(execute(arm, command(arm, 'addComment', { ranges: [{ ...range, end: { ...range.end, offset: 500 } }], author: 'Ada', date: '2026-01-01', body: 'Review', expectVersion: version }))).toMatchObject({ status: 'refused', reason: 'invalid-range' });
  expect(snapshot(arm)).toEqual(before);
  expect(execute(arm, command(arm, 'removeComment', { id: 'missing', expectVersion: version }))).toMatchObject({ status: 'refused', reason: 'unknown-comment' });
  expect(snapshot(arm)).toEqual(before);
  expect(execute(arm, command(arm, 'acceptRevision', { revisionId: 'missing', expectVersion: version }))).toMatchObject({ status: 'refused', reason: 'unknown-revision' });
  arm.state.readOnly = true;
  expect(execute(arm, command(arm, 'addComment', { ranges: [range], author: 'Ada', date: '2026-01-01', body: 'Review', expectVersion: version }))).toMatchObject({ status: 'refused', reason: 'read-only' });
  expect(arm.engine.listComments()).toEqual([]);
}, TIMEOUT);

test('review_commands_support_reanchoring_and_resolve_all_revision_sites', async () => {
  const arm = await openCommandEngine(['Alpha beta']);
  const paraId = arm.state.selection!.anchor.paraId;
  const range = { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 5 } };
  const current = () => arm.engine.proposalEngine.version();
  expect(execute(arm, command(arm, 'addComment', { ranges: [range], author: 'Ada', date: '2026-01-01', body: 'Review', expectVersion: current() }))).toMatchObject({ status: 'applied' });
  const id = arm.engine.listComments()[0]!.id;
  expect(execute(arm, command(arm, 'setCommentRanges', { id, ranges: [{ ...range, start: { paraId, offset: 6 }, end: { paraId, offset: 10 } }], expectVersion: current() }))).toMatchObject({ status: 'applied' });
  expect(arm.engine.resolveComment(id)).toMatchObject([{ story: 'body', start: 6, end: 10 }]);
  execute(arm, command(arm, 'insertText', { text: '!', suggesting: { name: 'Ada', date: '2026-01-01' } }));
  const revisionId = arm.engine.geometryReader.listRevisions()[0]!.revisionId;
  expect(execute(arm, command(arm, 'acceptRevision', { revisionId, expectVersion: current() }))).toMatchObject({ status: 'applied' });
  expect(arm.engine.geometryReader.listRevisions()).toEqual([]);
}, TIMEOUT);

test('review snapshots of a preview succeed without a comment projection', async () => {
  const arm = await openCommandEngine(['Alpha beta', 'Second'], { preview: 1 });
  const frame = applyFrameDelta(null, decodeFrameDelta(arm.engine.buildDisplayListFrame('{}', 0)));
  const review = await readResidentReviewSnapshot(arm.resident, createProposalRegistry(arm.engine.proposalEngine), frame, 0);
  expect(review.comments).toEqual([]);
  expect(review).not.toHaveProperty('uiComments');
  expect(review).not.toHaveProperty('commentIds');
}, TIMEOUT);

test('a read-only state applies granted host proposals and refuses ungranted proposals and edits', async () => {
  const arm = await openCommandEngine(['Alpha beta', 'Second']);
  arm.state.proposals = createProposalRegistry(arm.engine.proposalEngine);
  arm.state.readOnly = true;
  const paragraph = arm.engine.paragraphIdentities().paragraphs[0]!.session!;
  const proposals = [{ id: 'one', paragraph, op: 'replaceText' as const, search: 'Alpha', replaceWith: 'Omega', suggest: { author: 'Ada', date: '2026-01-01' } }];
  const version = () => arm.engine.proposalEngine.version();
  const before = snapshot(arm);
  expect(execute(arm, command(arm, 'propose', { expectVersion: version(), proposals }))).toMatchObject({ status: 'refused', reason: 'read-only' });
  expect(execute(arm, command(arm, 'propose', { expectVersion: version(), proposals, hostProposal: false }))).toMatchObject({ status: 'refused', reason: 'read-only' });
  expect(execute(arm, command(arm, 'propose', { expectVersion: version(), proposals, hostProposal: 'yes' as never }))).toMatchObject({ status: 'refused', reason: 'invalid-payload' });
  expect(execute(arm, command(arm, 'insertText', { text: '!', hostProposal: true } as never))).toMatchObject({ status: 'refused', reason: 'invalid-payload' });
  expect(snapshot(arm)).toEqual(before);
  expect(arm.state.proposals.snapshot().proposals).toEqual([]);
  expect(execute(arm, command(arm, 'propose', { expectVersion: version(), proposals, hostProposal: true }))).toMatchObject({ status: 'applied' });
  const proposed = arm.state.proposals.snapshot();
  expect(proposed.proposals.map(({ id, state }) => ({ id, state }))).toEqual([{ id: 'one', state: 'proposed' }]);
  expect(execute(arm, command(arm, 'setProposalStates', { expectVersion: version(), expectPreviewVersion: proposed.previewVersion,
    changes: [{ id: 'one', state: 'accepted' }] }))).toMatchObject({ status: 'refused', reason: 'read-only' });
  expect(execute(arm, command(arm, 'setProposalStates', { expectVersion: version(), expectPreviewVersion: proposed.previewVersion,
    changes: [{ id: 'one', state: 'accepted' }], hostProposal: true }))).toMatchObject({ status: 'applied' });
  expect(arm.state.proposals.snapshot().proposals[0]!.state).toBe('accepted');
  expect(execute(arm, command(arm, 'insertText', { text: '!' }))).toMatchObject({ status: 'refused', reason: 'read-only' });
  expect(execute(arm, command(arm, 'undo', {}))).toMatchObject({ status: 'refused', reason: 'read-only' });
  expect(execute(arm, command(arm, 'withdrawProposals', { expectVersion: version(), ids: ['one'] }))).toMatchObject({ status: 'refused', reason: 'read-only' });
  expect(execute(arm, command(arm, 'withdrawProposals', { expectVersion: version(), ids: ['one'], hostProposal: true }))).toMatchObject({ status: 'applied' });
  expect(arm.state.proposals.snapshot().proposals).toEqual([]);
  expect(arm.state.readOnly).toBe(true);
}, TIMEOUT);
