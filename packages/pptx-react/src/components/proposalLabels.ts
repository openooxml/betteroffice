import type { Proposal, ProposalChange } from '@betteroffice/pptx';
import type { TFunction } from '@betteroffice/pptx-i18n';

export function proposalSlideLabels(change: ProposalChange, proposal: Proposal | undefined, t: TFunction): string[] {
  if (!proposal) return [];
  const labels: string[] = [];
  for (const edit of proposal.edits) {
    if (!('slideId' in edit)) continue;
    const slideId = edit.slideId.startsWith('@') ? proposal.refs?.[edit.slideId.slice(1)] : edit.slideId;
    if (slideId !== change.slideId) continue;
    if (edit.type === 'insertTableRow' || edit.type === 'deleteTableRow' || edit.type === 'insertTableColumn' || edit.type === 'deleteTableColumn') {
      const shapeId = edit.shapeId.startsWith('@') ? proposal.refs?.[edit.shapeId.slice(1)] : edit.shapeId;
      if (shapeId !== change.shapeId) continue;
      const key = edit.type === 'insertTableRow' ? 'proposals.rowInserted'
        : edit.type === 'deleteTableRow' ? 'proposals.rowDeleted'
          : edit.type === 'insertTableColumn' ? 'proposals.columnInserted' : 'proposals.columnDeleted';
      labels.push(t(key));
    } else if (change.shapeId === null && (edit.type === 'setSlideLayout' || edit.type === 'setSlideBackground')) {
      const label = t(edit.type === 'setSlideLayout' ? 'proposals.layoutChanged' : 'proposals.backgroundChanged');
      if (!labels.includes(label)) labels.push(label);
    }
  }
  return labels;
}
