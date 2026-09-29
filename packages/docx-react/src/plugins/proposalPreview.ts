import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { proposalRevisionPreview, type DocxProposalSnapshot } from '@betteroffice/docx/yrs';
import {
  revisionPreviewKey,
  revisionPreviewKeyOf,
} from '../components/DocxEditor/internals/layoutProvenance';

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
export function currentPreviewKey(session: object | null): string {
  const snapshot = proposalSnapshot(session);
  return revisionPreviewKey(snapshot ? proposalRevisionPreview(snapshot) : undefined);
}

/** The preview key of the layout `queries` answer for. */
export function renderedPreviewKey(queries: DisplayListQueries): string {
  return revisionPreviewKeyOf(queries) ?? '';
}
