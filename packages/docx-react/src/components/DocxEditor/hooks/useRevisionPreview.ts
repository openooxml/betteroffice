import { useEffect, useState } from 'react';
import {
  proposalRevisionPreview,
  type DocxProposalSnapshot,
  type YrsRenderEnv,
  type YrsSession,
} from '@betteroffice/docx/yrs';

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
      const revisionPreview = proposalRevisionPreview(snapshot);
      setState((previous) =>
        previous?.session === session &&
        previous.preview.previewVersion === snapshot.previewVersion &&
        JSON.stringify(previous.preview.revisionPreview) === JSON.stringify(revisionPreview)
          ? previous
          : { session, preview: { previewVersion: snapshot.previewVersion, revisionPreview } }
      );
    };
    update(session.getProposals());
    return session.onProposalChange(update);
  }, [session]);
  return state && state.session === session ? state.preview : NO_PREVIEW;
}
