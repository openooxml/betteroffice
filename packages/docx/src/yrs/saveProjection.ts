import type { RepackOptions, RepackResult } from '../docx/rezip';
import { modelSaveFields, sha256Hex } from '../docx/rustSaveFacade';
import type { Document } from '../types/document';
import type { YrsSession } from './index';
import { sessionMaterializedSource, sessionSourceDigest, sessionSourcePackage } from './sessionInternals';

const projections = new WeakMap<Document, {
  session: YrsSession;
  version: string;
  sourceDigest: string;
  baseline: Promise<string | undefined>;
}>();

const snapshotDigest = (document: Document): Promise<string | undefined> =>
  sha256Hex(new TextEncoder().encode(JSON.stringify(modelSaveFields(document), (key, value) =>
    key === 'sourceOrdinal' ? undefined : value
  ))).catch(() => undefined);

function sameSaveTree(left: unknown, right: unknown): boolean {
  if (left === right) return true;
  if (!left || !right || typeof left !== 'object' || typeof right !== 'object') return false;
  if (Array.isArray(left) || Array.isArray(right)) {
    return Array.isArray(left) && Array.isArray(right) && left.length === right.length
      && left.every((value, index) => sameSaveTree(value, right[index]));
  }
  const a = left as Record<string, unknown>;
  const b = right as Record<string, unknown>;
  const keys = (value: Record<string, unknown>) =>
    Object.keys(value).filter(key => key !== 'sourceOrdinal' && value[key] !== undefined);
  const fields = keys(a);
  return fields.length === keys(b).length && fields.every(key => sameSaveTree(a[key], b[key]));
}

function sameBytes(bytes: Uint8Array, buffer: ArrayBuffer): boolean {
  const other = new Uint8Array(buffer);
  return bytes.length === other.length && bytes.every((byte, index) => byte === other[index]);
}

export function rememberSaveProjection(session: YrsSession, base: Document, projection: Document): void {
  const source = sessionSourcePackage(session);
  const sourceDigest = sessionSourceDigest(session);
  if (!source || !sourceDigest || !base.originalBuffer) return;
  if (base.originalBuffer !== source.buffer &&
      !sameBytes(new Uint8Array(source.buffer), base.originalBuffer)) return;
  const materialized = sessionMaterializedSource(session);
  if (!materialized || !sameSaveTree(modelSaveFields(base), modelSaveFields(materialized))) return;
  projections.set(projection, {
    session,
    version: session.version(),
    sourceDigest,
    baseline: snapshotDigest(projection),
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
  const [digest, savedModel, currentModel, save] = await Promise.all([
    sha256Hex(bytes), baseline, snapshotDigest(document), import('./saveYrsDocx'),
  ]);
  const source = sessionSourcePackage(session);
  if (!source || version !== session.version() || digest !== sourceDigest ||
      document.originalBuffer !== buffer || !sameBytes(bytes, buffer) ||
      (source.buffer !== buffer && !sameBytes(bytes, source.buffer)) ||
      savedModel === undefined || savedModel !== currentModel) return undefined;
  const { saved, warnings } = await save.saveYrsDocxWithWarnings(session, options);
  const published = save.publishedSaveVersion(saved);
  if (published === session.version() && provenance.version === version) provenance.version = published;
  return {
    buffer: Uint8Array.from(saved.bytes).buffer,
    warnings: [...(document.warnings ?? []).filter(warning => !warnings.includes(warning)), ...warnings],
  };
}
