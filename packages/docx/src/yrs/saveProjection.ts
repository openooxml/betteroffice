import type { RepackOptions, RepackResult } from '../docx/rezip';
import {
  assertSafeSaveTree, modelSaveFields, parsedModelUnchanged, sameSaveTree, sha256Hex, snapshotSaveTree,
} from '../docx/rustSaveFacade';
import type { Document } from '../types/document';
import type { YrsSession } from './index';
import { sessionMaterializedSource, sessionSourceDigest, sessionSourcePackage } from './sessionInternals';

type SaveFields = ReturnType<typeof modelSaveFields>;

const hostEnvelopes = new WeakMap<Document, SaveFields>();
const projectedModels = new WeakSet<Document>();
const projections = new WeakMap<Document, {
  session: YrsSession;
  version: string;
  sourceDigest: string;
  base: Document;
  baseAtProjection: SaveFields | undefined;
  baseline: SaveFields;
}>();

export function rememberHostEnvelope(document: Document): void {
  hostEnvelopes.set(document, snapshotSaveTree(modelSaveFields(document)));
}

export function isHostEnvelope(document: Document): boolean {
  return hostEnvelopes.has(document);
}

export function isModelProjection(document: Document): boolean {
  return projectedModels.has(document);
}

function sameBytes(bytes: Uint8Array, buffer: ArrayBuffer): boolean {
  const other = new Uint8Array(buffer);
  return bytes.length === other.length && bytes.every((byte, index) => byte === other[index]);
}

export function rememberSaveProjection(session: YrsSession, base: Document, projection: Document): void {
  const source = sessionSourcePackage(session);
  const sourceDigest = sessionSourceDigest(session);
  if (!base.originalBuffer) return;
  if (!source) {
    projectedModels.add(projection);
    return;
  }
  if (!sourceDigest) return;
  const host = hostEnvelopes.get(base);
  if (host && !sameSaveTree(host, modelSaveFields(base))) return;
  projections.set(projection, {
    session,
    version: session.version(),
    sourceDigest,
    base,
    baseAtProjection: host ? undefined : snapshotSaveTree(modelSaveFields(base)),
    baseline: snapshotSaveTree(modelSaveFields(projection)),
  });
}

export async function saveCurrentProjection(
  document: Document,
  options: RepackOptions,
  buffer = document.originalBuffer
): Promise<RepackResult | undefined> {
  const provenance = projections.get(document);
  if (!provenance || !buffer) return undefined;
  const { session, version, sourceDigest, base, baseAtProjection, baseline } = provenance;
  if (!sessionSourcePackage(session) || version !== session.version()) return undefined;
  const bytes = new Uint8Array(buffer).slice();
  const [digest, save] = await Promise.all([sha256Hex(bytes), import('./saveYrsDocx')]);
  const source = sessionSourcePackage(session);
  if (!source || version !== session.version() || digest !== sourceDigest ||
      !sameBytes(bytes, buffer) || (source.buffer !== buffer && !sameBytes(bytes, source.buffer))) return undefined;
  const fields = modelSaveFields(document);
  assertSafeSaveTree(fields, 'save');
  if (!sameSaveTree(baseline, fields)) return undefined;
  if (baseAtProjection) {
    const parsed = parsedModelUnchanged(base);
    if (!parsed || await parsed.source !== sourceDigest) {
      const materialized = sessionMaterializedSource(session);
      if (!materialized || (materialized !== base &&
          !sameSaveTree(baseAtProjection, modelSaveFields(materialized)))) return undefined;
    }
    if (!sameSaveTree(baseAtProjection, modelSaveFields(base)) || version !== session.version() ||
        !sameSaveTree(baseline, modelSaveFields(document))) return undefined;
  }
  const { saved, warnings } = await save.saveYrsDocxWithWarnings(session, options, true);
  const published = save.publishedSaveVersion(saved);
  if (published === session.version() && provenance.version === version) provenance.version = published;
  return {
    buffer: Uint8Array.from(saved.bytes).buffer,
    warnings: [...(document.warnings ?? []).filter(warning => !warnings.includes(warning)), ...warnings],
  };
}
