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
 * Resolves `target` in the worker at its `version` under the session's revision preview, refusing
 * when that preview is not the rendered `previewKey`; null once the worker moved past `version`.
 * Persisted anchors resolve on the main thread when it holds the document, else in the worker.
 */
export async function readWorkerAnchorTarget(
  read: ResidentEngineWorkerClient['documentRead'],
  session: YrsSession,
  target: Exclude<DocxGeometryTarget, { kind: 'proposal' }>,
  version: string,
  previewVersion: number,
  previewKey: string
): Promise<AnchorDisplayTarget | null> {
  const rendered = () =>
    (proposalSnapshot(session)?.previewVersion ?? 0) === previewVersion &&
    revisionPreviewKey(revisionPreviewOf(session)) === previewKey;
  const unrendered = () =>
    anchorFailure('layout-unavailable', 'No rendered layout shows this target yet');
  if (!rendered()) return unrendered();
  let resolver: Pick<AnchorReader, 'resolveParagraphAnchor'> = session;
  if (
    (target.kind === 'paragraph' || target.kind === 'search') &&
    target.paragraph.kind !== 'session' &&
    !workerOpenReplicaReady(session)
  ) {
    const resolved = await read({ kind: 'resolveParagraphAnchors', anchors: [target.paragraph] });
    if (resolved.version !== version) return null;
    resolver = { resolveParagraphAnchor: () => resolved.value.results[0]! };
  }
  const posted = sessionAnchorTarget(resolver, target);
  if ('ok' in posted) return posted;
  const revisionPreview = revisionPreviewOf(session);
  if (
    (proposalSnapshot(session)?.previewVersion ?? 0) !== previewVersion ||
    revisionPreviewKey(revisionPreview) !== previewKey
  )
    return unrendered();
  const reply = await read({ kind: 'anchorTarget', target: posted, revisionPreview, expectVersion: version });
  return reply.version === version ? reply.value : null;
}
