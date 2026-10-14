import { extractTrackedChangesFromYrs, type YrsSidebarProjection } from '@betteroffice/docx/layout/render';
import type { DocxProposalState, ResidentReviewSnapshot, YrsLoc } from '@betteroffice/docx/yrs';
import type { ReactSidebarItem } from '../../../plugin-api/types';
import type { DocxPluginGeometry } from '../../../plugins/types';
import { TrackedChangeCard } from '../../sidebar/TrackedChangeCard';
import type { createPeerlessShellController } from './shellController';

export function workerProposalSidebarItems(controller: ReturnType<typeof createPeerlessShellController>, review: ResidentReviewSnapshot | null,
  geometry: DocxPluginGeometry | null, enabled: boolean, report?: (error: Error) => void): ReactSidebarItem[] {
  if (!enabled || !review) return [];
  return review.proposals.proposals.flatMap(proposal => {
    const target = review.proposalGeometry.targets[proposal.id];
    if (!target?.ok) return [];
    const projection = { locToDisplayPoint: (loc: YrsLoc) => target.paragraph === null || loc.paraId !== proposal.paragraph.paraId || loc.story !== proposal.paragraph.story
      ? null : { position: target.paragraph + loc.offset } } as YrsSidebarProjection;
    const revisions = review.revisions.filter(revision => proposal.revisionIds.includes(revision.revisionId));
    const entries = extractTrackedChangesFromYrs(revisions, projection).entries;
    const placed = geometry?.getAnchorGeometry({ kind: 'proposal', id: proposal.id });
    const rect = review.proposalHighlights[proposal.id]?.[0];
    const frame = controller.getSnapshot().view?.geometry;
    const page = rect && frame?.pages.find(page => page.page.pageIndex === rect.pageIndex);
    const bounds = rect && (geometry?.dom.getPageBounds(rect.pageIndex) ?? (page ? { y: page.bounds.y + 48 / controller.ref.getZoom() } : null));
    const fixedY = bounds && rect ? bounds.y + rect.y : placed?.ok ? placed.anchor.y / controller.ref.getZoom() : undefined;
    const decide = (state: DocxProposalState) => {
      void (async () => {
        const current = await controller.ref.getProposals();
        const result = await controller.ref.setProposalStates({ expectVersion: current.version, expectPreviewVersion: current.previewVersion,
          changes: [{ id: proposal.id, state }] });
        if (!result.ok) throw new Error(result.failure.message ?? result.failure.code);
      })().catch(error => report?.(error instanceof Error ? error : new Error(String(error))));
    };
    return entries.map<ReactSidebarItem>((change, index) => ({ id: `host-proposal-${proposal.id}-${index}`, anchorPos: change.from,
      anchorKey: `revision-${change.revisionId}`, priority: 1, estimatedHeight: 80, fixedY, hidden: fixedY === undefined,
      render: props => <TrackedChangeCard {...props} change={change} replies={[]} onAccept={() => decide('accepted')} onReject={() => decide('rejected')} />,
    }));
  });
}
