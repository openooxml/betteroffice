import { commandSelectionRange } from './residentCommandAnchors';
import type { DocxCommandIntent, DocxCommentResult, ResidentDocxCommandEngine, ResidentDocxCommandState } from './residentCommandProtocol';
import type { DocxProposalRequest, DocxProposalStateRequest, DocxProposalWithdrawRequest } from './proposals';

export const reviewCommandKeys = {
  addComment: ['ranges', 'author', 'date', 'body', 'expectVersion'],
  updateComment: ['id', 'patch', 'expectVersion'],
  setCommentRanges: ['id', 'ranges', 'expectVersion'],
  removeComment: ['id', 'expectVersion'],
  acceptRevision: ['revisionId', 'expectVersion'],
  rejectRevision: ['revisionId', 'expectVersion'],
  propose: ['proposals', 'expectVersion', 'hostProposal'],
  setProposalStates: ['changes', 'expectVersion', 'expectPreviewVersion', 'hostProposal'],
  withdrawProposals: ['ids', 'expectVersion', 'hostProposal'],
};

export type ReviewCommand = Extract<DocxCommandIntent, { kind: keyof typeof reviewCommandKeys }>;

export function grantedHostProposal(command: ReviewCommand): boolean {
  return (command.kind === 'propose' || command.kind === 'setProposalStates' || command.kind === 'withdrawProposals') &&
    command.payload.hostProposal === true;
}

export function preflightReviewCommand(engine: ResidentDocxCommandEngine, state: ResidentDocxCommandState, command: ReviewCommand): string | null {
  const payload = command.payload;
  if (payload.expectVersion !== engine.proposalEngine.version()) return 'stale-version';
  if (command.kind === 'withdrawProposals') {
    if (!state.proposals) return 'review-unavailable';
  }
  switch (command.kind) {
    case 'addComment': case 'setCommentRanges': {
      if (!Array.isArray(command.payload.ranges) || !command.payload.ranges.length) return 'invalid-range';
      for (const range of command.payload.ranges) {
        if (!range || !range.start || !range.end) return 'invalid-range';
        const resolved = commandSelectionRange(engine, {
          anchor: { story: range.story, ...range.start }, focus: { story: range.story, ...range.end },
        });
        if (!resolved || resolved.start.paraId !== range.start.paraId || resolved.start.offset !== range.start.offset || (resolved.start.paraId === resolved.end.paraId && resolved.start.offset === resolved.end.offset)) return 'invalid-range';
      }
      break;
    }
  }
  if ('id' in payload && !engine.listComments().some(comment => comment.id === payload.id)) return 'unknown-comment';
  if ('revisionId' in payload && state.proposals?.snapshot().proposals.some(proposal => proposal.revisionIds.includes(payload.revisionId))) return 'tracked-revision-conflict';
  if ('revisionId' in payload && !engine.geometryReader.listRevisions().some(revision => revision.revisionId === payload.revisionId)) return 'unknown-revision';
  return null;
}

export function applyReviewCommand(engine: ResidentDocxCommandEngine, state: ResidentDocxCommandState, command: ReviewCommand):
  { status: 'applied' | 'noop'; comment?: Omit<DocxCommentResult, 'seq'> } | { status: 'refused'; reason: string } {
  const version = engine.proposalEngine.version();
  const before = state.proposals?.snapshot();
  if (command.kind === 'propose' || command.kind === 'setProposalStates' || command.kind === 'withdrawProposals') {
    if (!state.proposals) return { status: 'refused', reason: 'review-unavailable' };
    const { hostProposal: _grant, ...request } = command.payload;
    const result = command.kind === 'propose' ? state.proposals.propose(request as DocxProposalRequest)
      : command.kind === 'setProposalStates' ? state.proposals.setStates(request as DocxProposalStateRequest)
      : state.proposals.withdraw(request as DocxProposalWithdrawRequest);
    if (!result.ok) return { status: 'refused', reason: result.failure.code };
    return { status: JSON.stringify(before) === JSON.stringify(result.snapshot) ? 'noop' : 'applied' };
  }
  engine.beginUndoCapture();
  engine.addUndoBoundary();
  let added: string | undefined;
  switch (command.kind) {
    case 'addComment': added = engine.addComment(command.payload.ranges, command.payload.author, command.payload.date, command.payload.body).commentId; break;
    case 'updateComment': engine.updateComment(command.payload.id, command.payload.patch); break;
    case 'setCommentRanges': engine.setCommentRanges(command.payload.id, command.payload.ranges); break;
    case 'removeComment': engine.applyRawOps('body', [{ op: 'removeComment', id: command.payload.id }]); break;
    case 'acceptRevision': engine.acceptChange({ revisionId: command.payload.revisionId }); break;
    case 'rejectRevision': engine.rejectChange({ revisionId: command.payload.revisionId }); break;
  }
  engine.addUndoBoundary();
  engine.markProjectionStories(engine.geometryReader.storyIds());
  const status = engine.proposalEngine.version() === version ? 'noop' : 'applied';
  const projection = added === undefined ? null : engine.commentProjection();
  const id = added === undefined ? undefined : projection?.ids[added];
  return projection && id !== undefined ? { status, comment: { id, comments: projection.comments } } : { status };
}
