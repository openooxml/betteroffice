import {
  hiddenRanges as readHiddenRanges,
  resolveAnchorTarget as readAnchorTarget,
  type AnchorReader,
  type AnchorResolutionFailure,
} from '@betteroffice/docx/yrs';
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
