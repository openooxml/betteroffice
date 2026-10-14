import { useEffect, useState } from 'react';
import {
  proposalRevisionPreview,
  type DocxProposalSnapshot,
  type YrsRenderEnv,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { hasEditorWorkerProposalRounds, registeredWorkerProposalAuthority, subscribeEditorWorkerProposalAuthority } from '../internals/workerProposalAuthority';

export interface RevisionPreviewState {
  previewVersion: number;
  revisionPreview: YrsRenderEnv['revisionPreview'];
}

const NO_PREVIEW: RevisionPreviewState = Object.freeze({
  previewVersion: 0,
  revisionPreview: undefined,
});

/** The session's proposal decisions for the render env; a new object only when they change. */
export function useRevisionPreview(session: YrsSession | null): RevisionPreviewState {
  const [state, setState] = useState<{ session: YrsSession; preview: RevisionPreviewState } | null>(
    null
  );
  useEffect(() => {
    if (!session) return;
    const update = (snapshot: DocxProposalSnapshot): void => {
      const authority = hasEditorWorkerProposalRounds(session) ? registeredWorkerProposalAuthority(session) : null;
      const previewVersion = authority ? authority.previewVersion() : snapshot.previewVersion;
      const revisionPreview = authority ? authority.revisionPreview() : proposalRevisionPreview(snapshot);
      setState((previous) =>
        previous?.session === session &&
        previous.preview.previewVersion === previewVersion &&
        JSON.stringify(previous.preview.revisionPreview) === JSON.stringify(revisionPreview)
          ? previous
          : { session, preview: { previewVersion, revisionPreview } }
      );
    };
    update(session.getProposals());
    const unsubscribe = session.onProposalChange(update);
    const unsubscribeWorker = subscribeEditorWorkerProposalAuthority(session, () => update(session.getProposals()));
    return () => { unsubscribe(); unsubscribeWorker(); };
  }, [session]);
  return state && state.session === session ? state.preview : NO_PREVIEW;
}
