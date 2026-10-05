import { useEffect, useState } from 'react';
import {
  proposalRevisionPreview,
  type YrsRenderEnv,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { workerProposalRoundAuthority } from '../internals/workerProposalAuthority';

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
    const authority = workerProposalRoundAuthority(session);
    const update = (): void => {
      const snapshot = session.getProposals();
      const previewVersion = authority?.initialized ? authority.previewVersion() : snapshot.previewVersion;
      const revisionPreview = authority?.initialized ? authority.revisionPreview() : proposalRevisionPreview(snapshot);
      setState((previous) =>
        previous?.session === session &&
        previous.preview.previewVersion === previewVersion &&
        JSON.stringify(previous.preview.revisionPreview) === JSON.stringify(revisionPreview)
          ? previous
          : { session, preview: { previewVersion, revisionPreview } }
      );
    };
    update();
    const unsubscribe = session.onProposalChange(update);
    const unsubscribeWorker = authority?.subscribe(update);
    return () => { unsubscribe(); unsubscribeWorker?.(); };
  }, [session]);
  return state && state.session === session ? state.preview : NO_PREVIEW;
}
