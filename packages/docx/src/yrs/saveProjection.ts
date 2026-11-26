import type { RepackOptions, RepackResult } from '../docx/rezip';
import {
  markProjectedBodyRewrite, modelSaveFields, parsedModelUnchanged, saveTreeFingerprint, sha256Hex,
} from '../docx/rustSaveFacade';
import type { Document } from '../types/document';
import type { YrsSession } from './index';
import type { ProjectionFormatEdit } from './saveYrsDocx';
import { sessionMaterializedSource, sessionSourceDigest, sessionSourcePackage } from './sessionInternals';

const hostEnvelopes = new WeakMap<Document, { fingerprint: string; sourceDigest: string | undefined }>();
const formatEdits = new WeakMap<YrsSession, readonly ProjectionFormatEdit[]>();
const projections = new WeakMap<Document, {
  session: YrsSession;
  version: string;
  sourceDigest: string;
  base: Document;
  baseAtProjection: string | undefined;
  baseline: string;
  formatEdits: readonly ProjectionFormatEdit[];
}>();

export function rememberProjectionFormatEdit(session: YrsSession, edit: ProjectionFormatEdit): void {
  const previous = formatEdits.get(session);
  formatEdits.set(session, [...(previous ?? []), edit]);
}

export function clearProjectionFormatEdits(session: YrsSession): void {
  formatEdits.delete(session);
}

export function rememberHostEnvelope(document: Document, sourceDigest: string | undefined): void {
  hostEnvelopes.set(document, { fingerprint: saveTreeFingerprint(modelSaveFields(document)), sourceDigest });
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
    if (projection.package.document.content.some(block => block.type === 'paragraph' && block.renderedPageBreakBefore)) {
      markProjectedBodyRewrite(projection);
    }
    return;
  }
  if (!sourceDigest) return;
  const host = hostEnvelopes.get(base);
  const baseFingerprint = saveTreeFingerprint(modelSaveFields(base));
  if (host && (host.sourceDigest !== sourceDigest || host.fingerprint !== baseFingerprint)) return;
  const version = session.version();
  projections.set(projection, {
    session,
    version,
    sourceDigest,
    base,
    baseAtProjection: host ? undefined : baseFingerprint,
    baseline: saveTreeFingerprint(modelSaveFields(projection)),
    formatEdits: formatEdits.get(session) ?? [],
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
  const { saved, warnings } = await save.saveYrsDocxWithWarnings(session, options, provenance.formatEdits);
  if (saved.keptUnchanged) {
    const { structure, paragraphs } = saved.keptUnchanged;
    const parts = [...new Set(paragraphs.map(({ part }) => part))];
    const location = parts.length ? ` in ${parts.join(', ')}` : '';
    warnings.push(`DOCX_KEPT_UNCHANGED Edits were skipped to preserve source structure in ${structure} ${structure === 1 ? 'paragraph' : 'paragraphs'}${location}.`);
  }
  for (const { part, dropped } of saved.structureDropped ?? []) {
    warnings.push(`DOCX_PART_REWRITTEN ${part} was rewritten from the document model. Content the model does not carry may be lost. Saving edits through a session (createYrsSession, saveYrsDocx) patches changed paragraphs in place where it can. Dropped elements include ${dropped.join(', ')}.`);
  }
  const published = save.publishedSaveVersion(saved);
  if (published === session.version() && provenance.version === version) {
    provenance.version = published;
  }
  return {
    buffer: Uint8Array.from(saved.bytes).buffer,
    warnings: [...(document.warnings ?? []).filter(warning => !warnings.includes(warning)), ...warnings],
  };
}
