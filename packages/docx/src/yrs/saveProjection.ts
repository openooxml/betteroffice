import type { RepackOptions, RepackResult } from '../docx/rezip';
import {
  modelSaveFields, parsedModelUnchanged, saveTreeFingerprint, sha256Hex,
} from '../docx/rustSaveFacade';
import type { Document } from '../types/document';
import type { YrsSession } from './index';
import { sessionMaterializedSource, sessionSourceDigest, sessionSourcePackage } from './sessionInternals';

const hostEnvelopes = new WeakMap<Document, string>();
const projections = new WeakMap<Document, {
  session: YrsSession;
  version: string;
  sourceDigest: string;
  base: Document;
  baseAtProjection: string | undefined;
  baseline: string;
}>();

export function rememberHostEnvelope(document: Document): void {
  hostEnvelopes.set(document, saveTreeFingerprint(modelSaveFields(document)));
}

function sameBytes(bytes: Uint8Array, buffer: ArrayBuffer): boolean {
  const other = new Uint8Array(buffer);
  return bytes.length === other.length && bytes.every((byte, index) => byte === other[index]);
}

export function rememberSaveProjection(session: YrsSession, base: Document, projection: Document): void {
  const source = sessionSourcePackage(session);
  const sourceDigest = sessionSourceDigest(session);
  if (!base.originalBuffer || !source || !sourceDigest) return;
  const host = hostEnvelopes.get(base);
  const baseFingerprint = saveTreeFingerprint(modelSaveFields(base));
  if (host && host !== baseFingerprint) return;
  projections.set(projection, {
    session,
    version: session.version(),
    sourceDigest,
    base,
    baseAtProjection: host ? undefined : baseFingerprint,
    baseline: saveTreeFingerprint(modelSaveFields(projection)),
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
  if (baseline !== saveTreeFingerprint(modelSaveFields(document))) return undefined;
  if (baseAtProjection) {
    const parsed = parsedModelUnchanged(base);
    if (parsed?.sourceDigest !== sourceDigest) {
      const materialized = sessionMaterializedSource(session);
      if (!materialized || (materialized !== base &&
          baseAtProjection !== saveTreeFingerprint(modelSaveFields(materialized)))) return undefined;
    }
    if (baseAtProjection !== saveTreeFingerprint(modelSaveFields(base))) return undefined;
  }
  const { saved, warnings } = await save.saveYrsDocxWithWarnings(session, options);
  const published = save.publishedSaveVersion(saved);
  if (published === session.version() && provenance.version === version) provenance.version = published;
  return {
    buffer: Uint8Array.from(saved.bytes).buffer,
    warnings: [...(document.warnings ?? []).filter(warning => !warnings.includes(warning)), ...warnings],
  };
}
