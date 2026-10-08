import {
  anchorFailure,
  hiddenRangesForPreview,
  resolveAnchorTarget as readAnchorTarget,
  sessionAnchorTarget,
  type AnchorDisplayTarget,
  type AnchorReader,
  type AnchorResolutionFailure,
  type ResidentEngineWorkerClient,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { revisionPreviewKey } from '../components/DocxEditor/internals/layoutProvenance';
import { workerOpenReplicaReady } from '../components/DocxEditor/internals/workerOpenReplica';
import { proposalSnapshot, revisionPreviewOf } from './proposalPreview';
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
  return hiddenRangesForPreview(session, version, revisionPreviewOf(session));
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

/**
 * Resolves `targets` in the worker at its `version` under the session's revision preview with one
 * read, refusing when that preview is not the rendered `previewKey`; null once the worker moved past
 * `version`. Persisted anchors resolve on the main thread when it holds the document, else in the worker.
 */
export async function readWorkerAnchorTargets(
  read: ResidentEngineWorkerClient['documentRead'],
  session: YrsSession,
  targets: readonly Exclude<DocxGeometryTarget, { kind: 'proposal' }>[],
  version: string,
  previewVersion: number,
  previewKey: string
): Promise<AnchorDisplayTarget[] | null> {
  const revisionPreview = revisionPreviewOf(session);
  if (
    (proposalSnapshot(session)?.previewVersion ?? 0) !== previewVersion ||
    revisionPreviewKey(revisionPreview) !== previewKey
  ) {
    return targets.map(() =>
      anchorFailure('layout-unavailable', 'No rendered layout shows this target yet')
    );
  }
  const onMain = workerOpenReplicaReady(session);
  const converted = targets.map((target) =>
    onMain ? sessionAnchorTarget(session, target, session.version()) : target
  );
  const posted = converted.filter((target): target is (typeof targets)[number] => !('ok' in target));
  let answers: Iterator<AnchorDisplayTarget> = [].values();
  if (posted.length > 0) {
    const reply = await read({ kind: 'anchorTargets', targets: posted, revisionPreview, expectVersion: version });
    if (reply.version !== version || !reply.value) return null;
    answers = reply.value.values();
  }
  return converted.map((target) => ('ok' in target ? target : answers.next().value!));
}
