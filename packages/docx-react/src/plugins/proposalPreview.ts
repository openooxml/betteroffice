import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { DocxProposalSnapshot } from '@betteroffice/docx/yrs';

export type {
  DocxOccurrence,
  DocxProposalRecord,
  DocxProposalSnapshot,
} from '@betteroffice/docx/yrs';

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

/** The preview key a layout of the session's current preview carries. */
export function currentPreviewKey(_session: object | null): string {
  return '';
}

/** The preview key of the layout `queries` answer for. */
export function renderedPreviewKey(_queries: DisplayListQueries): string {
  return '';
}
