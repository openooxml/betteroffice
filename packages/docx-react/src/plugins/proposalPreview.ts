import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { proposalRevisionPreview, type DocxProposalSnapshot, type YrsSession } from '@betteroffice/docx/yrs';
import { workerOpenDocumentHeld } from '../components/DocxEditor/internals/workerOpenReplica';
import { workerProposalRoundAuthority } from '../components/DocxEditor/internals/workerProposalAuthority';
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
  workerDocumentMirrored(): boolean;
}

/** The session's proposal registry, or null for a session without one. */
export function proposalSnapshot(session: object | null): DocxProposalSnapshot | null {
  const worker = session ? workerProposalRoundAuthority(session as YrsSession) : null;
  if (worker && (worker.initialized || !workerOpenDocumentHeld(session as YrsSession))) return worker.snapshot();
  const registry = session as Partial<ProposalRegistry> | null;
  return typeof registry?.getProposals === 'function' ? registry.getProposals() : null;
}

export function observeProposals(
  session: object,
  listener: (snapshot: DocxProposalSnapshot) => void
): () => void {
  const registry = session as Partial<ProposalRegistry>;
  const worker = workerProposalRoundAuthority(session as YrsSession);
  const update = () => { const snapshot = proposalSnapshot(session); if (snapshot) listener(snapshot); };
  const unsubscribe = typeof registry.onProposalChange === 'function'
    ? registry.onProposalChange(update)
    : () => {};
  const unsubscribeWorker = typeof registry.workerDocumentMirrored === 'function'
    ? worker?.subscribe(() => { if (!workerOpenDocumentHeld(session as YrsSession)) update(); })
    : undefined;
  return () => { unsubscribe(); unsubscribeWorker?.(); };
}

/** The preview key a layout of the session's current preview carries. */
export function currentPreviewKey(session: object | null): string {
  const worker = session ? workerProposalRoundAuthority(session as YrsSession) : null;
  if (worker && (worker.initialized || !workerOpenDocumentHeld(session as YrsSession))) {
    return revisionPreviewKey(worker.revisionPreview());
  }
  const snapshot = proposalSnapshot(session);
  return revisionPreviewKey(snapshot ? proposalRevisionPreview(snapshot) : undefined);
}

/** The preview key of the layout `queries` answer for. */
export function renderedPreviewKey(queries: DisplayListQueries): string {
  return revisionPreviewKeyOf(queries) ?? '';
}
