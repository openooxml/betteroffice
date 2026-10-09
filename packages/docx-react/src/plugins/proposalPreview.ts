import { hasEditorWorkerProposalRounds, registeredWorkerProposalAuthority, subscribeEditorWorkerProposalAuthority } from '../components/DocxEditor/internals/workerProposalAuthority';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { proposalRevisionPreview, type YrsSession, type DocxProposalSnapshot } from '@betteroffice/docx/yrs';
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
  if (session && hasEditorWorkerProposalRounds(session as YrsSession)) return registeredWorkerProposalAuthority(session as YrsSession)!.snapshot();
  const registry = session as Partial<ProposalRegistry> | null;
  return typeof registry?.getProposals === 'function' ? registry.getProposals() : null;
}

export function observeProposals(
  session: object,
  listener: (snapshot: DocxProposalSnapshot) => void
): () => void {
  const registry = session as Partial<ProposalRegistry>;
  const unsubscribe = typeof registry.onProposalChange === 'function'
    ? registry.onProposalChange((snapshot) => {
        if (hasEditorWorkerProposalRounds(session as YrsSession)) {
          const current = proposalSnapshot(session);
          if (current) listener(current);
        } else listener(snapshot);
      })
    : () => {};
  const unsubscribeWorker = subscribeEditorWorkerProposalAuthority(session as YrsSession, () => {
    const snapshot = proposalSnapshot(session);
    if (snapshot) listener(snapshot);
  });
  return () => { unsubscribe(); unsubscribeWorker(); };
}

/** The preview key a layout of the session's current preview carries. */
export function currentPreviewKey(session: object | null): string {
  if (session && hasEditorWorkerProposalRounds(session as YrsSession)) return revisionPreviewKey(registeredWorkerProposalAuthority(session as YrsSession)!.revisionPreview());
  const snapshot = proposalSnapshot(session);
  return revisionPreviewKey(snapshot ? proposalRevisionPreview(snapshot) : undefined);
}

/** The preview key of the layout `queries` answer for. */
export function renderedPreviewKey(queries: DisplayListQueries): string {
  return revisionPreviewKeyOf(queries) ?? '';
}
