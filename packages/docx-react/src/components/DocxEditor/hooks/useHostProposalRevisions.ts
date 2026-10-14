import { useEffect, useState } from 'react';
import type { DocxProposalSnapshot, YrsSession } from '@betteroffice/docx/yrs';
import { yrsIdToNumericId } from '@betteroffice/docx/layout/render';
import { hasEditorWorkerProposalRounds, registeredWorkerProposalAuthority, subscribeEditorWorkerProposalAuthority } from '../internals/workerProposalAuthority';

const NO_REVISIONS = new Set<string>();

/** Sidebar anchor keys of the session's host-proposal revisions, updated as proposals change. */
export function useHostProposalRevisions(session: YrsSession | null): Set<string> {
  const [state, setState] = useState<{ session: YrsSession; keys: Set<string> } | null>(null);
  useEffect(() => {
    if (!session) return;
    const update = (snapshot: DocxProposalSnapshot): void => {
      const keys = new Set<string>();
      const workerSnapshot = hasEditorWorkerProposalRounds(session)
        ? registeredWorkerProposalAuthority(session)?.snapshot()
        : null;
      const proposals = workerSnapshot ? [...snapshot.proposals, ...workerSnapshot.proposals] : snapshot.proposals;
      for (const proposal of proposals) {
        for (const revisionId of proposal.revisionIds) {
          keys.add(`revision-${yrsIdToNumericId(revisionId)}`);
        }
      }
      setState((previous) =>
        previous?.session === session &&
        previous.keys.size === keys.size &&
        [...keys].every((key) => previous.keys.has(key))
          ? previous
          : { session, keys }
      );
    };
    update(session.getProposals());
    const unsubscribe = session.onProposalChange(update);
    const unsubscribeWorker = subscribeEditorWorkerProposalAuthority(session, () => update(session.getProposals()));
    return () => { unsubscribe(); unsubscribeWorker(); };
  }, [session]);
  return state && state.session === session && state.keys.size > 0 ? state.keys : NO_REVISIONS;
}
