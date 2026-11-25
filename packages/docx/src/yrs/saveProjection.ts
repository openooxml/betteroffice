import type { RepackOptions, RepackResult } from '../docx/rezip';
import { modelSaveFields, sha256Hex } from '../docx/rustSaveFacade';
import { withSessionRevisionIds } from '../docx/sessionRevisionIds';
import type { Document } from '../types/document';
import type { YrsSession } from './index';
import { sessionMaterializedSource, sessionSourcePackage } from './sessionInternals';

const projections = new WeakMap<Document, {
  session: YrsSession;
  version: string;
  sourceDigest: string;
  baseline: string;
}>();

function saveSnapshot(document: Document): string {
  const fields = modelSaveFields(document);
  return withSessionRevisionIds(fields, () => JSON.stringify(fields));
}

function sameBytes(bytes: Uint8Array, buffer: ArrayBuffer): boolean {
  const other = new Uint8Array(buffer);
  return bytes.length === other.length && bytes.every((byte, index) => byte === other[index]);
}

export function rememberSaveProjection(session: YrsSession, base: Document, projection: Document): void {
  const source = sessionSourcePackage(session);
  if (!source || !base.originalBuffer) return;
  if (base.originalBuffer !== source.buffer &&
      !sameBytes(new Uint8Array(source.buffer), base.originalBuffer)) return;
  const materialized = sessionMaterializedSource(session);
  if (!materialized || saveSnapshot(base) !== saveSnapshot(materialized)) return;
  const sourceDigest = session.paragraphIdentities().packageSha256;
  if (!sourceDigest) return;
  projections.set(projection, {
    session,
    version: session.version(),
    sourceDigest,
    baseline: saveSnapshot(projection),
  });
}

export async function saveCurrentProjection(
  document: Document,
  options: RepackOptions
): Promise<RepackResult | undefined> {
  const provenance = projections.get(document);
  const buffer = document.originalBuffer;
  if (!provenance || !buffer) return undefined;
  const { session, version, sourceDigest, baseline } = provenance;
  if (!sessionSourcePackage(session) || version !== session.version()) return undefined;
  const bytes = new Uint8Array(buffer).slice();
  const [digest, { saveYrsDocxWithWarnings }] = await Promise.all([
    sha256Hex(bytes),
    import('./saveYrsDocx'),
  ]);
  const source = sessionSourcePackage(session);
  if (!source || version !== session.version() || digest !== sourceDigest ||
      document.originalBuffer !== buffer || !sameBytes(bytes, buffer) ||
      (source.buffer !== buffer && !sameBytes(bytes, source.buffer)) ||
      baseline !== saveSnapshot(document)) return undefined;
  const { saved, warnings } = await saveYrsDocxWithWarnings(session, options);
  return {
    buffer: Uint8Array.from(saved.bytes).buffer,
    warnings: [...(document.warnings ?? []).filter(warning => !warnings.includes(warning)), ...warnings],
  };
}
