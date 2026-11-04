import type { Proposal, ProposalChange } from '@betteroffice/pptx';
import type { TFunction } from '@betteroffice/pptx-i18n';

export function proposalSlideLabels(change: ProposalChange, proposal: Proposal | undefined, t: TFunction): string[] {
  if (change.shapeId !== null || !proposal) return [];
  const labels = new Set<string>();
  for (const edit of proposal.edits) {
    if (edit.type !== 'setSlideLayout' && edit.type !== 'setSlideBackground') continue;
    const slideId = edit.slideId.startsWith('@') ? proposal.refs?.[edit.slideId.slice(1)] : edit.slideId;
    if (slideId === change.slideId) labels.add(t(edit.type === 'setSlideLayout' ? 'proposals.layoutChanged' : 'proposals.backgroundChanged'));
  }
  return [...labels];
}
