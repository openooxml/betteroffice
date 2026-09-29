import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { DocxSessionParagraphAnchor } from '@betteroffice/docx/yrs';

export type DocxOccurrence = 'first' | 'all' | number;

export interface DocxProposalRecord {
  id: string;
  state: 'proposed' | 'accepted' | 'rejected';
  paragraph: DocxSessionParagraphAnchor;
  revisionIds: readonly string[];
  changed: boolean;
}

export interface DocxProposalSnapshot {
  version: string;
  previewVersion: number;
  proposals: readonly DocxProposalRecord[];
}

interface ProposalRegistry {
  getProposals(): DocxProposalSnapshot;
  onProposalChange(listener: (snapshot: DocxProposalSnapshot) => void): () => void;
}

/** The session's proposal registry, or null for a session without one. */
export function proposalSnapshot(session: object | null): DocxProposalSnapshot | null {
  const registry = session as Partial<ProposalRegistry> | null;
  return typeof registry?.getProposals === 'function' ? registry.getProposals() : null;
}

export function observeProposals(
  session: object,
  listener: (snapshot: DocxProposalSnapshot) => void
): () => void {
  const registry = session as Partial<ProposalRegistry>;
  return typeof registry.onProposalChange === 'function'
    ? registry.onProposalChange(listener)
    : () => {};
}

/** Each decided proposal's revisions, by revision id. */
export function revisionPreview(
  snapshot: DocxProposalSnapshot
): Readonly<Record<string, 'accepted' | 'rejected'>> | undefined {
  const decided: Record<string, 'accepted' | 'rejected'> = {};
  for (const proposal of snapshot.proposals) {
    if (proposal.state === 'proposed') continue;
    for (const revisionId of proposal.revisionIds) decided[revisionId] = proposal.state;
  }
  return Object.keys(decided).length > 0 ? decided : undefined;
}

/** The preview key a layout of the session's current preview carries. */
export function currentPreviewKey(_session: object | null): string {
  return '';
}

/** The preview key of the layout `queries` answer for. */
export function renderedPreviewKey(_queries: DisplayListQueries): string {
  return '';
}
