import {
  hiddenRanges as readHiddenRanges,
  hiddenRangesForPreview,
  resolveAnchorTarget as readAnchorTarget,
  type AnchorReader,
  type AnchorResolutionFailure,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { hasEditorWorkerProposalRounds, registeredWorkerProposalAuthority } from '../components/DocxEditor/internals/workerProposalAuthority';
import { proposalSnapshot } from './proposalPreview';
import type { DocxGeometryTarget } from './types';

export {
  anchorFailure,
  isBodyStory,
  textRangeToRaw,
  type RawAnchorRange,
} from '@betteroffice/docx/yrs';
export type AnchorFailure = AnchorResolutionFailure;
export type AnchorSession = AnchorReader;

export function hiddenRanges(session: AnchorSession, version: string) {
  if (hasEditorWorkerProposalRounds(session as YrsSession)) {
    return hiddenRangesForPreview(
      session,
      version,
      registeredWorkerProposalAuthority(session as YrsSession)!.revisionPreview()
    );
  }
  return readHiddenRanges(session, version, proposalSnapshot(session));
}

export function resolveAnchorTarget(
  session: AnchorSession,
  target: DocxGeometryTarget,
  version: string
) {
  return readAnchorTarget(
    session,
    target,
    version,
    target.kind === 'proposal' ? proposalSnapshot(session) : null
  );
}
